"""SMTP delivery for faculty duty reports and daily change reports."""

from __future__ import annotations

import html
import logging
import os
import smtplib
import ssl
import threading
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

from bson.objectid import ObjectId

from assignments import get_timetable
from reports import (
    DOCX_MIME,
    XLSX_MIME,
    daily_change_filename,
    faculty_report_filename,
    generate_daily_change_reports,
    generate_faculty_reports,
    report_identity,
    report_output_root,
)

logger = logging.getLogger(__name__)


def _env_bool(name, default=False):
    raw = os.getenv(name)
    if raw is None or raw == "":
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name, default):
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


def _split_recipients(value):
    return [item.strip() for item in str(value or "").replace(";", ",").split(",") if item.strip()]


def mail_config():
    host = (os.getenv("SMTP_HOST") or "").strip()
    port = _env_int("SMTP_PORT", 587)
    username = (os.getenv("SMTP_USERNAME") or "").strip()
    password = os.getenv("SMTP_PASSWORD") or ""
    from_email = (os.getenv("SMTP_FROM_EMAIL") or "").strip() or username
    from_name = (os.getenv("SMTP_FROM_NAME") or "PICT Examination Cell").strip()
    use_ssl = _env_bool("SMTP_USE_SSL", port == 465)
    use_tls = _env_bool("SMTP_USE_TLS", not use_ssl)
    test_mode = _env_bool("MAIL_TEST_MODE", False)
    test_recipient = (os.getenv("MAIL_TEST_RECIPIENT") or "").strip()
    dry_run = _env_bool("MAIL_DRY_RUN", False)
    if not from_email and dry_run:
        from_email = "examcell@pict.edu"
    async_send = _env_bool("MAIL_ASYNC", True)
    send_daily_changes = _env_bool("MAIL_SEND_DAILY_CHANGES", True)
    report_recipients = _split_recipients(
        os.getenv("REPORT_RECIPIENTS") or os.getenv("MAIL_REPORT_RECIPIENTS")
    )
    fallback_report_recipients = _split_recipients(
        ",".join([
            os.getenv("COE_EMAIL") or "",
            os.getenv("PRINCIPAL_EMAIL") or "",
            os.getenv("COORDINATOR_EMAIL") or "",
        ])
    )
    if not report_recipients:
        report_recipients = fallback_report_recipients

    if test_mode and not test_recipient:
        test_mode = False

    raw_enabled = (os.getenv("MAIL_ENABLED") or "auto").strip().lower()
    configured = bool(host and from_email)
    if raw_enabled in {"false", "0", "no", "off"}:
        enabled = False
    elif raw_enabled in {"true", "1", "yes", "on"}:
        enabled = configured
    else:
        enabled = configured

    missing = []
    if not host:
        missing.append("SMTP_HOST")
    if not from_email:
        missing.append("SMTP_FROM_EMAIL (or SMTP_USERNAME)")
    if test_mode and not test_recipient:
        missing.append("MAIL_TEST_RECIPIENT")

    if not enabled and raw_enabled in {"false", "0", "no", "off"}:
        reason = "Mail is disabled by MAIL_ENABLED=false."
    elif not configured:
        reason = "Mail is not configured yet. Missing: " + ", ".join(missing)
    else:
        reason = ""

    return {
        "enabled": enabled,
        "configured": configured,
        "reason": reason,
        "missing": missing,
        "host": host,
        "port": port,
        "username": username,
        "password": password,
        "from_email": from_email,
        "from_name": from_name,
        "reply_to": (os.getenv("MAIL_REPLY_TO") or from_email).strip(),
        "use_ssl": use_ssl,
        "use_tls": use_tls,
        "timeout": max(5, _env_int("SMTP_TIMEOUT", 30)),
        "test_mode": test_mode,
        "test_recipient": test_recipient,
        "dry_run": dry_run,
        "async_send": async_send,
        "send_daily_changes": send_daily_changes,
        "report_recipients": report_recipients,
    }


def public_mail_status():
    config = mail_config()
    return {
        "enabled": config["enabled"],
        "configured": config["configured"],
        "test_mode": config["test_mode"],
        "test_recipient": config["test_recipient"] if config["test_mode"] else None,
        "dry_run": config["dry_run"],
        "from_email": config["from_email"] or None,
        "reason": config["reason"],
        "missing": config["missing"],
    }


def _delivery_recipient(intended, config):
    if config["test_mode"]:
        return config["test_recipient"]
    return str(intended or "").strip()


def _signature_html(identity):
    coe_image = ""
    principal_image = ""
    if identity.get("coe_image"):
        coe_image = "<div style=\"font-size:11px;color:#64748b;margin-top:4px;\">Signature image attached in report</div>"
    if identity.get("principal_image"):
        principal_image = "<div style=\"font-size:11px;color:#64748b;margin-top:4px;\">Signature image attached in report</div>"
    return f"""
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:22px;border-collapse:collapse;">
        <tr>
          <td width="50%" style="padding:12px 16px;border:1px solid #bfdbfe;background:#eff6ff;border-radius:8px 0 0 8px;vertical-align:top;">
            <div style="font-size:10px;letter-spacing:1.2px;color:#1d4ed8;font-weight:700;">DIGITALLY SIGNED</div>
            {coe_image}
            <div style="font-size:14px;font-weight:700;color:#0f172a;margin-top:8px;">{html.escape(identity['coe_name'])}</div>
            <div style="font-size:11px;color:#475569;">{html.escape(identity['coe_title'])}</div>
            <div style="font-size:10px;color:#64748b;margin-top:6px;">{html.escape(identity['signed_at'])}</div>
          </td>
          <td width="50%" style="padding:12px 16px;border:1px solid #bfdbfe;border-left:none;background:#eff6ff;border-radius:0 8px 8px 0;vertical-align:top;">
            <div style="font-size:10px;letter-spacing:1.2px;color:#1d4ed8;font-weight:700;">DIGITALLY SIGNED</div>
            {principal_image}
            <div style="font-size:14px;font-weight:700;color:#0f172a;margin-top:8px;">{html.escape(identity['principal_name'])}</div>
            <div style="font-size:11px;color:#475569;">{html.escape(identity['principal_title'])}</div>
            <div style="font-size:10px;color:#64748b;margin-top:6px;">{html.escape(identity['signed_at'])}</div>
          </td>
        </tr>
      </table>
    """


def _email_shell(identity, title, intro_html, content_html, footer_html=""):
    return f"""<!doctype html>
<html>
  <body style="margin:0;padding:0;background:#eef2f7;font-family:Segoe UI,Arial,sans-serif;color:#0f172a;">
    <div style="max-width:760px;margin:24px auto;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #dbe4ef;">
      <div style="background:linear-gradient(135deg,#1e3a5f,#0f766e);padding:24px 28px;color:#ffffff;">
        <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;opacity:.8;">Examination Section</div>
        <div style="font-size:22px;font-weight:700;margin-top:6px;">{html.escape(identity['institute_name'])}</div>
        <div style="font-size:14px;margin-top:8px;color:#dbeafe;">{html.escape(title)}</div>
      </div>
      <div style="padding:26px 28px;">
        {intro_html}
        {content_html}
        <div style="font-size:12px;color:#475569;margin-top:18px;line-height:1.7;">
          {footer_html}
        </div>
        {_signature_html(identity)}
        <div style="margin-top:18px;padding-top:14px;border-top:1px solid #e2e8f0;font-size:10px;color:#94a3b8;line-height:1.6;">
          This is a system-generated communication from the PICT Examination Section.
          The attached Word and Excel reports contain digitally signed approval blocks for the COE and Principal.
        </div>
      </div>
    </div>
  </body>
</html>"""


def _attach_report(message, filename, payload, mime):
    subtype = mime.split("/", 1)[1]
    message.add_attachment(
        payload,
        maintype="application",
        subtype=subtype,
        filename=filename,
    )


def build_faculty_message(timetable, teacher, report, config):
    context = report["context"]
    actual_email = str(teacher.get("email") or "").strip()
    if not actual_email:
        raise ValueError("Faculty email is missing")
    delivered_to = _delivery_recipient(actual_email, config)
    if not delivered_to:
        raise ValueError("No delivery recipient configured")

    identity = context["identity"]
    teacher_name = context["teacher"]["name"]
    timetable_dept = context["timetable"]["department"]
    revision = context["timetable"]["revision"]
    subject = f"Invigilation Duty Report - {timetable_dept} - {teacher_name} (Rev {revision})"
    if config["test_mode"]:
        subject = f"[TEST to {actual_email}] {subject}"

    duty_rows = "".join(
        f"<tr>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;'>{html.escape(duty['date'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;text-align:center;'>{html.escape(duty['slot'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;'>{html.escape(duty['session'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;'>{html.escape(duty['course_id'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;text-align:center;'>{html.escape(duty['year'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;text-align:center;'>{html.escape(duty['room'])}</td>"
        f"</tr>"
        for duty in context["duties"]
    )
    intro = (
        f"<p style='margin:0 0 10px;'>Dear {html.escape(teacher_name)},</p>"
        f"<p style='margin:0 0 16px;line-height:1.7;'>Your invigilation duties for the "
        f"<strong>{html.escape(timetable_dept)}</strong> examination timetable are published. "
        f"The complete Word and Excel reports are attached.</p>"
    )
    test_notice = ""
    if config["test_mode"]:
        test_notice = (
            "<div style='background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;"
            "padding:10px 12px;border-radius:8px;font-size:12px;margin-bottom:14px;'>"
            f"TEST DELIVERY: this email was intended for <strong>{html.escape(actual_email)}</strong> "
            f"and redirected to <strong>{html.escape(delivered_to)}</strong>.</div>"
        )
    content = f"""
      {test_notice}
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:12px;">
        <thead>
          <tr style="background:#1e3a5f;color:#ffffff;">
            <th style="padding:9px;text-align:left;border:1px solid #1e3a5f;">Date</th>
            <th style="padding:9px;border:1px solid #1e3a5f;">Slot</th>
            <th style="padding:9px;text-align:left;border:1px solid #1e3a5f;">Session</th>
            <th style="padding:9px;text-align:left;border:1px solid #1e3a5f;">Course</th>
            <th style="padding:9px;border:1px solid #1e3a5f;">Year</th>
            <th style="padding:9px;border:1px solid #1e3a5f;">Room</th>
          </tr>
        </thead>
        <tbody>{duty_rows}</tbody>
      </table>
      <div style="margin-top:16px;padding:12px 14px;background:#f8fafc;border-left:4px solid #0f766e;border-radius:6px;font-size:12px;line-height:1.7;">
        <strong>Reporting instructions:</strong>
        report at least {context['reporting_minutes']} minutes before the session for Junior Supervisor duties
        and {context['squad_reporting_minutes']} minutes before the session for Squad duties, to the
        Examination Control Room ({html.escape(context['control_room'])}).
      </div>
    """
    footer = (
        f"Timetable version {context['timetable']['version']} &nbsp;|&nbsp; "
        f"Revision {revision} &nbsp;|&nbsp; Generated {html.escape(context['generated_at'])}"
    )
    html_body = _email_shell(identity, "Faculty Invigilation Duty Report", intro, content, footer)
    text_body = (
        f"Dear {teacher_name},\n\n"
        f"Your invigilation duties for the {timetable_dept} examination timetable are published.\n\n"
        + "\n".join(
            f"- {duty['date']}, Slot {duty['slot']}, {duty['session']}, "
            f"Course {duty['course_id']}, Year {duty['year']}, Room {duty['room']}"
            for duty in context["duties"]
        )
        + f"\n\nReport at least {context['reporting_minutes']} minutes before the session "
          f"for Junior Supervisor duties and {context['squad_reporting_minutes']} minutes before "
          f"the session for Squad duties, to the Examination Control Room "
          f"({context['control_room']}).\n\n"
        + (
            f"TEST DELIVERY: intended for {actual_email}; redirected to {delivered_to}.\n\n"
            if config["test_mode"] else ""
        )
        + f"Digitally signed by {identity['coe_name']} ({identity['coe_title']}) and "
          f"{identity['principal_name']} ({identity['principal_title']}) on {identity['signed_at']}.\n"
    )

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = formataddr((config["from_name"], config["from_email"]))
    message["To"] = delivered_to
    message["Reply-To"] = config["reply_to"] or config["from_email"]
    message["Date"] = formatdate(localtime=True)
    message["Message-ID"] = make_msgid(domain=config["from_email"].split("@")[-1] if "@" in config["from_email"] else None)
    if config["test_mode"]:
        message["X-PICT-Test-Mode"] = "true"
        message["X-PICT-Intended-Recipient"] = actual_email
    message.set_content(text_body)
    message.add_alternative(html_body, subtype="html")
    _attach_report(
        message,
        report["filenames"]["docx"],
        report["buffers"]["docx"],
        DOCX_MIME,
    )
    _attach_report(
        message,
        report["filenames"]["xlsx"],
        report["buffers"]["xlsx"],
        XLSX_MIME,
    )
    return message, {
        "kind": "faculty_duty",
        "teacher_id": context["teacher"]["teacher_id"],
        "teacher_name": teacher_name,
        "intended_to": actual_email,
        "delivered_to": delivered_to,
        "subject": subject,
        "attachments": [report["filenames"]["docx"], report["filenames"]["xlsx"]],
    }


def build_daily_change_message(generated_reports, config):
    if not generated_reports:
        raise ValueError("No daily change reports to send")
    recipients = config["report_recipients"]
    if config["test_mode"]:
        delivered = [config["test_recipient"]] if config["test_recipient"] else []
        intended = recipients or [config["test_recipient"]]
    else:
        delivered = recipients
        intended = recipients
    if not delivered:
        raise ValueError(
            "No daily report recipients configured. Set REPORT_RECIPIENTS or MAIL_TEST_RECIPIENT."
        )

    context = generated_reports[0]["context"]
    identity = context["identity"]
    dates = [report["context"]["report_date"] for report in generated_reports]
    subject = "Daily Faculty Duty Change Report - " + ", ".join(dates)
    if config["test_mode"]:
        subject = f"[TEST to {', '.join(intended)}] {subject}"

    summary_rows = "".join(
        f"<tr>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;'>{html.escape(report['context']['report_date'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;text-align:center;'>{len(report['context']['rows'])}</td>"
        f"<td style='padding:8px;border:1px solid #dbe4ef;'>{html.escape(', '.join(report['filenames'].values()))}</td>"
        f"</tr>"
        for report in generated_reports
    )
    attachments = []
    for report in generated_reports:
        attachments.append((report["filenames"]["docx"], report["buffers"]["docx"], DOCX_MIME))
        attachments.append((report["filenames"]["xlsx"], report["buffers"]["xlsx"], XLSX_MIME))

    intro = (
        "<p style='margin:0 0 16px;line-height:1.7;'>The following faculty duty changes were approved "
        "for the published examination timetable. Word and Excel copies are attached for the "
        "Examination Section records.</p>"
    )
    test_notice = ""
    if config["test_mode"]:
        test_notice = (
            "<div style='background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;"
            "padding:10px 12px;border-radius:8px;font-size:12px;margin-bottom:14px;'>"
            f"TEST DELIVERY: report recipients were redirected to "
            f"<strong>{html.escape(', '.join(delivered))}</strong>.</div>"
        )
    content = f"""
      {test_notice}
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:12px;">
        <thead>
          <tr style="background:#1e3a5f;color:#ffffff;">
            <th style="padding:9px;text-align:left;border:1px solid #1e3a5f;">Examination Date</th>
            <th style="padding:9px;border:1px solid #1e3a5f;">Changes</th>
            <th style="padding:9px;text-align:left;border:1px solid #1e3a5f;">Attachments</th>
          </tr>
        </thead>
        <tbody>{summary_rows}</tbody>
      </table>
    """
    footer = (
        f"Department: {context['timetable']['department']} &nbsp;|&nbsp; "
        f"Revision {context['timetable']['revision']} &nbsp;|&nbsp; Generated {html.escape(context['generated_at'])}"
    )
    html_body = _email_shell(identity, "Daily Faculty Duty Change Report", intro, content, footer)
    text_body = (
        "Daily Faculty Duty Change Report\n\n"
        + "\n".join(
            f"- {report['context']['report_date']}: {len(report['context']['rows'])} change(s)"
            for report in generated_reports
        )
        + "\n\n"
        + (
            f"TEST DELIVERY: redirected to {', '.join(delivered)}.\n\n"
            if config["test_mode"] else ""
        )
        + f"Digitally signed by {identity['coe_name']} and {identity['principal_name']} "
          f"on {identity['signed_at']}.\n"
    )

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = formataddr((config["from_name"], config["from_email"]))
    message["To"] = ", ".join(delivered)
    message["Reply-To"] = config["reply_to"] or config["from_email"]
    message["Date"] = formatdate(localtime=True)
    message["Message-ID"] = make_msgid(domain=config["from_email"].split("@")[-1] if "@" in config["from_email"] else None)
    if config["test_mode"]:
        message["X-PICT-Test-Mode"] = "true"
    message.set_content(text_body)
    message.add_alternative(html_body, subtype="html")
    for filename, payload, mime in attachments:
        _attach_report(message, filename, payload, mime)
    return message, {
        "kind": "daily_change",
        "intended_to": ", ".join(intended),
        "delivered_to": ", ".join(delivered),
        "subject": subject,
        "attachments": [filename for filename, _, _ in attachments],
        "report_dates": dates,
    }


def _open_smtp(config):
    if config["use_ssl"]:
        context = ssl.create_default_context()
        server = smtplib.SMTP_SSL(
            config["host"], config["port"], timeout=config["timeout"], context=context
        )
    else:
        server = smtplib.SMTP(config["host"], config["port"], timeout=config["timeout"])
        server.ehlo()
        if config["use_tls"]:
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
    if config["username"] and config["password"]:
        server.login(config["username"], config["password"])
    return server


def _write_dry_run_message(message, config):
    folder = report_output_root() / "mail_outbox"
    folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    recipient = str(message.get("To") or "recipient").replace("@", "_at_").replace(",", "_")
    safe_recipient = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in recipient)[:80]
    path = folder / f"{timestamp}_{safe_recipient}.eml"
    path.write_bytes(message.as_bytes())
    return str(path)


def send_messages(messages_with_meta, config):
    """Send a batch of messages, reconnecting once after transport errors."""
    results = []
    if config["dry_run"]:
        for message, meta in messages_with_meta:
            try:
                path = _write_dry_run_message(message, config)
                results.append({"meta": meta, "status": "dry_run", "error": None, "eml_path": path})
            except Exception as exc:
                results.append({"meta": meta, "status": "failed", "error": str(exc), "eml_path": None})
        return results

    server = None
    for message, meta in messages_with_meta:
        error = None
        for attempt in range(2):
            try:
                if server is None:
                    server = _open_smtp(config)
                server.send_message(
                    message,
                    from_addr=config["from_email"],
                    to_addrs=_split_recipients(meta.get("delivered_to")),
                )
                error = None
                break
            except Exception as exc:
                error = exc
                try:
                    if server is not None:
                        server.close()
                except Exception:
                    pass
                server = None
                logger.warning(
                    "SMTP delivery failed for %s (attempt %s/2): %s",
                    meta.get("delivered_to"), attempt + 1, exc,
                )
        if error:
            results.append({"meta": meta, "status": "failed", "error": str(error), "eml_path": None})
        else:
            results.append({"meta": meta, "status": "sent", "error": None, "eml_path": None})
    try:
        if server is not None:
            server.quit()
    except Exception:
        pass
    return results


def _record_delivery(database, job_id, result, config, timetable_id):
    meta = result["meta"]
    database["mail_log"].insert_one({
        "job_id": str(job_id) if job_id else None,
        "timetable_id": str(timetable_id),
        "kind": meta.get("kind"),
        "teacher_id": meta.get("teacher_id"),
        "teacher_name": meta.get("teacher_name"),
        "intended_to": meta.get("intended_to"),
        "delivered_to": meta.get("delivered_to"),
        "subject": meta.get("subject"),
        "attachments": meta.get("attachments") or [],
        "status": result["status"],
        "error": result.get("error"),
        "eml_path": result.get("eml_path"),
        "test_mode": bool(config["test_mode"]),
        "dry_run": bool(config["dry_run"]),
        "created_at": datetime.utcnow().isoformat() + "Z",
        "sent_at": datetime.utcnow().isoformat() + "Z" if result["status"] in {"sent", "dry_run"} else None,
    })


def _update_job(database, job_id, **fields):
    if not job_id:
        return
    try:
        database["mail_jobs"].update_one({"_id": ObjectId(str(job_id))}, {"$set": fields})
    except Exception:
        logger.exception("Could not update mail job %s", job_id)


def send_timetable_emails(timetable_id, database=None, job_id=None, only_teacher_id=None,
                          force_test=False, triggered_by=None):
    from db import get_db

    database = database if database is not None else get_db()
    config = mail_config()
    if force_test:
        config = dict(config)
        config["test_mode"] = bool(config["test_recipient"])
    timetable = get_timetable(str(timetable_id), database=database)
    if not timetable:
        raise ValueError("Published timetable not found")
    timetable_id = str(timetable["_id"])

    _update_job(
        database, job_id, status="running", started_at=datetime.utcnow().isoformat() + "Z",
        message="Generating faculty reports and sending emails.",
    )
    assignments = list(database["assignments"].find({
        "timetable_id": timetable_id,
        "active": True,
        "teacher_id": {"$ne": None},
    }))
    teachers = {
        str(doc.get("teacher_id")): doc
        for doc in database["teachers"].find({})
        if doc.get("teacher_id") and not doc.get("is_admin")
    }
    teacher_ids = sorted({str(row.get("teacher_id")) for row in assignments if row.get("teacher_id")})
    messages = []
    skipped = []
    for teacher_id in teacher_ids:
        if only_teacher_id and str(only_teacher_id) != str(teacher_id):
            continue
        teacher = teachers.get(str(teacher_id))
        if not teacher:
            skipped.append({"teacher_id": str(teacher_id), "reason": "Teacher record not found"})
            continue
        if not str(teacher.get("email") or "").strip():
            skipped.append({"teacher_id": str(teacher_id), "reason": "Faculty email is missing"})
            continue
        report = generate_faculty_reports(database, timetable, teacher, persist=True)
        if not report:
            skipped.append({"teacher_id": str(teacher_id), "reason": "No active duties"})
            continue
        try:
            message, meta = build_faculty_message(timetable, teacher, report, config)
        except Exception as exc:
            skipped.append({"teacher_id": str(teacher_id), "reason": str(exc)})
            continue
        messages.append((message, meta))

    if not messages:
        summary = {
            "sent": 0,
            "failed": 0,
            "dry_run": 0,
            "skipped": len(skipped),
            "skipped_details": skipped,
            "message": "No faculty emails were generated.",
        }
        _update_job(
            database, job_id, status="completed", completed_at=datetime.utcnow().isoformat() + "Z",
            summary=summary,
        )
        return summary

    results = send_messages(messages, config)
    for result in results:
        _record_delivery(database, job_id, result, config, timetable_id)

    sent = sum(1 for item in results if item["status"] == "sent")
    dry_run = sum(1 for item in results if item["status"] == "dry_run")
    failed = sum(1 for item in results if item["status"] == "failed")
    status = "completed_with_errors" if failed else "completed"
    summary = {
        "sent": sent,
        "failed": failed,
        "dry_run": dry_run,
        "skipped": len(skipped),
        "skipped_details": skipped,
        "test_mode": config["test_mode"],
        "test_recipient": config["test_recipient"] if config["test_mode"] else None,
        "message": (
            f"{sent} email(s) sent"
            + (f", {dry_run} dry-run message(s) written" if dry_run else "")
            + (f", {failed} failed" if failed else "")
            + (f", {len(skipped)} skipped" if skipped else "")
            + "."
        ),
    }
    _update_job(
        database, job_id, status=status, completed_at=datetime.utcnow().isoformat() + "Z",
        summary=summary,
    )
    return summary


def send_daily_change_emails(timetable_id, report_dates, database=None, job_id=None,
                             triggered_by=None):
    from db import get_db

    database = database if database is not None else get_db()
    config = mail_config()
    timetable = get_timetable(str(timetable_id), database=database)
    if not timetable:
        raise ValueError("Published timetable not found")
    timetable_id = str(timetable["_id"])
    generated = generate_daily_change_reports(
        database, timetable, report_dates, persist=True, generated_by=triggered_by
    )
    if not generated:
        summary = {"sent": 0, "failed": 0, "generated": 0, "message": "No duty changes found for those dates."}
        _update_job(
            database, job_id, status="completed", completed_at=datetime.utcnow().isoformat() + "Z",
            summary=summary,
        )
        return summary

    if not config["report_recipients"] and not config["test_mode"]:
        summary = {
            "sent": 0,
            "failed": 0,
            "generated": len(generated),
            "message": (
                f"Generated {len(generated)} daily report(s), but no REPORT_RECIPIENTS "
                "are configured, so no email was sent."
            ),
        }
        _update_job(
            database, job_id, status="generated_only", completed_at=datetime.utcnow().isoformat() + "Z",
            summary=summary,
        )
        return summary

    _update_job(
        database, job_id, status="running", started_at=datetime.utcnow().isoformat() + "Z",
        message="Sending daily change reports.",
    )
    message, meta = build_daily_change_message(generated, config)
    results = send_messages([(message, meta)], config)
    for result in results:
        _record_delivery(database, job_id, result, config, timetable_id)
    sent = sum(1 for item in results if item["status"] == "sent")
    dry_run = sum(1 for item in results if item["status"] == "dry_run")
    failed = sum(1 for item in results if item["status"] == "failed")
    status = "completed_with_errors" if failed else "completed"
    summary = {
        "sent": sent,
        "failed": failed,
        "dry_run": dry_run,
        "generated": len(generated),
        "report_dates": [report["context"]["report_date"] for report in generated],
        "message": (
            f"Generated {len(generated)} daily report(s); "
            f"{sent} email(s) sent"
            + (f", {dry_run} dry-run message(s) written" if dry_run else "")
            + (f", {failed} failed" if failed else "")
            + "."
        ),
    }
    _update_job(
        database, job_id, status=status, completed_at=datetime.utcnow().isoformat() + "Z",
        summary=summary,
    )
    return summary


def _job_document(kind, timetable_id, triggered_by, extra):
    return {
        "kind": kind,
        "timetable_id": str(timetable_id),
        "triggered_by": str(triggered_by) if triggered_by else None,
        "status": "queued",
        "created_at": datetime.utcnow().isoformat() + "Z",
        "summary": {},
        **extra,
    }


def queue_timetable_emails(timetable_id, database=None, triggered_by=None):
    from db import get_db

    database = database if database is not None else get_db()
    config = mail_config()
    if not config["enabled"] and not config["dry_run"]:
        return {
            "queued": False,
            "reason": config["reason"] or "Mail is not configured.",
            "missing": config["missing"],
        }

    job_doc = _job_document("timetable_publish", timetable_id, triggered_by, {})
    job_id = database["mail_jobs"].insert_one(job_doc).inserted_id
    job_id_str = str(job_id)
    if config["async_send"]:
        thread = threading.Thread(
            target=_run_timetable_job,
            args=(job_id_str, str(timetable_id), triggered_by),
            name=f"pict-mail-{job_id_str[-6:]}",
            daemon=True,
        )
        thread.start()
        return {
            "queued": True,
            "job_id": job_id_str,
            "test_mode": config["test_mode"],
            "dry_run": config["dry_run"],
        }

    summary = send_timetable_emails(
        timetable_id, database=database, job_id=job_id_str, triggered_by=triggered_by
    )
    return {"queued": False, "job_id": job_id_str, "summary": summary}


def _run_timetable_job(job_id, timetable_id, triggered_by):
    from db import get_db

    database = None
    try:
        database = get_db()
        send_timetable_emails(
            timetable_id, database=database, job_id=job_id, triggered_by=triggered_by
        )
    except Exception as exc:
        logger.exception("Timetable email job %s failed", job_id)
        if database is not None:
            _update_job(
                database, job_id, status="failed",
                completed_at=datetime.utcnow().isoformat() + "Z",
                error=str(exc),
            )


def queue_daily_change_emails(timetable_id, report_dates, database=None, triggered_by=None):
    from db import get_db

    database = database if database is not None else get_db()
    config = mail_config()
    if not config["send_daily_changes"]:
        return {"queued": False, "reason": "Daily change emails are disabled."}
    if not config["enabled"] and not config["dry_run"]:
        return {
            "queued": False,
            "reason": config["reason"] or "Mail is not configured.",
            "missing": config["missing"],
        }
    if not config["report_recipients"] and not config["test_mode"]:
        return {
            "queued": False,
            "reason": "Set REPORT_RECIPIENTS to email daily change reports.",
        }

    job_doc = _job_document(
        "daily_change", timetable_id, triggered_by,
        {"report_dates": [str(value) for value in report_dates]},
    )
    job_id = database["mail_jobs"].insert_one(job_doc).inserted_id
    job_id_str = str(job_id)
    if config["async_send"]:
        thread = threading.Thread(
            target=_run_daily_change_job,
            args=(job_id_str, str(timetable_id), list(report_dates), triggered_by),
            name=f"pict-change-mail-{job_id_str[-6:]}",
            daemon=True,
        )
        thread.start()
        return {"queued": True, "job_id": job_id_str}
    summary = send_daily_change_emails(
        timetable_id, report_dates, database=database,
        job_id=job_id_str, triggered_by=triggered_by,
    )
    return {"queued": False, "job_id": job_id_str, "summary": summary}


def _run_daily_change_job(job_id, timetable_id, report_dates, triggered_by):
    from db import get_db

    database = None
    try:
        database = get_db()
        send_daily_change_emails(
            timetable_id, report_dates, database=database,
            job_id=job_id, triggered_by=triggered_by,
        )
    except Exception as exc:
        logger.exception("Daily change email job %s failed", job_id)
        if database is not None:
            _update_job(
                database, job_id, status="failed",
                completed_at=datetime.utcnow().isoformat() + "Z",
                error=str(exc),
            )


def get_mail_job(job_id, database=None):
    from db import get_db

    database = database if database is not None else get_db()
    try:
        job = database["mail_jobs"].find_one({"_id": ObjectId(str(job_id))})
    except Exception:
        return None
    if not job:
        return None
    job["_id"] = str(job["_id"])
    return job


def get_mail_status(database=None, limit=25):
    from db import get_db

    database = database if database is not None else get_db()
    jobs = list(database["mail_jobs"].find({}).sort("created_at", -1).limit(max(1, min(int(limit), 100))))
    logs = list(database["mail_log"].find({}).sort("created_at", -1).limit(max(1, min(int(limit) * 4, 400))))
    for job in jobs:
        job["_id"] = str(job["_id"])
    for log in logs:
        log["_id"] = str(log["_id"])
    return {"config": public_mail_status(), "jobs": jobs, "recent_deliveries": logs}
