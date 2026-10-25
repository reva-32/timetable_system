"""Check the configured outgoing-mail settings without touching timetable data.

Usage:
    python verify_mail.py           # configuration + SMTP connection/auth test
    python verify_mail.py --send    # also send one minimal test email

This script never writes to the timetable, assignments, faculty or request
collections.  With --send it sends only to MAIL_TEST_RECIPIENT, never to a
faculty mailbox.
"""

from __future__ import annotations

import argparse
import sys
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

from mailer import _open_smtp, mail_config, send_messages  # noqa: E402
from reports import report_identity  # noqa: E402


def mask_email(value):
    value = str(value or "").strip()
    if not value or "@" not in value:
        return value or "(empty)"
    local, domain = value.split("@", 1)
    visible = local[:1] if local else ""
    return f"{visible}***@{domain}"


def print_status(cfg):
    print("PICT outgoing mail configuration")
    print("--------------------------------")
    print(f"Mail enabled        : {cfg['enabled']}")
    print(f"SMTP host / port    : {cfg['host'] or '(empty)'} / {cfg['port']}")
    transport = "SSL" if cfg["use_ssl"] else ("STARTTLS" if cfg["use_tls"] else "plain")
    print(f"Transport security  : {transport}")
    print(f"SMTP username       : {mask_email(cfg['username'])}")
    print(f"SMTP password       : {'set' if cfg['password'] else '(empty)'}")
    print(f"From address        : {cfg['from_email'] or '(empty)'}")
    print(f"From name           : {cfg['from_name']}")
    print(f"Demo/test mode      : {cfg['test_mode']}")
    print(f"Demo recipient      : {cfg['test_recipient'] or '(empty)'}")
    print(f"Daily recipients    : {', '.join(cfg['report_recipients']) or '(none)'}")
    print(f"Dry run             : {cfg['dry_run']}")
    print(f"Async send          : {cfg['async_send']}")

    identity = report_identity()
    print(f"COE name / image    : {identity['coe_name']} / {identity['coe_image'] or '(text block)'}")
    print(
        f"Principal / image   : {identity['principal_name']} / "
        f"{identity['principal_image'] or '(text block)'}"
    )
    if cfg["missing"]:
        print(f"Missing settings    : {', '.join(cfg['missing'])}")
    if cfg["reason"]:
        print(f"Configuration note  : {cfg['reason']}")
    print()


def main():
    parser = argparse.ArgumentParser(
        description="Verify SMTP credentials without changing timetable data."
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Send one minimal email to MAIL_TEST_RECIPIENT after connection/auth succeeds.",
    )
    args = parser.parse_args()

    cfg = mail_config()
    print_status(cfg)

    if not cfg["enabled"]:
        print("[FAIL] Mail is not enabled/configured.")
        print("       Fill SMTP_HOST and SMTP_FROM_EMAIL (or SMTP_USERNAME) in backend/.env.")
        return 1

    if not cfg["username"] or not cfg["password"]:
        print(
            "[WARN] SMTP username/password is incomplete. "
            "This is only valid for a relay that allows unauthenticated sending."
        )

    print("Checking SMTP connection, STARTTLS/SSL and authentication...")
    server = None
    try:
        server = _open_smtp(cfg)
        server.noop()
        if cfg["username"] and cfg["password"]:
            print("[OK]   SMTP connection, transport security and login succeeded.")
        else:
            print("[OK]   SMTP connection and transport security succeeded (login not configured).")
    except Exception as exc:
        print(f"[FAIL] SMTP check failed: {exc}")
        print("       Check host, port, TLS/SSL mode, username and App Password.")
        return 2
    finally:
        if server is not None:
            try:
                server.quit()
            except Exception:
                pass

    if not args.send:
        print("[OK]   Credentials verified. No email was sent.")
        print("       Run with --send to send one minimal test message to the demo recipient.")
        return 0

    recipient = str(cfg["test_recipient"] or "").strip()
    if not recipient:
        print("[FAIL] MAIL_TEST_RECIPIENT is empty; refusing to send a test message.")
        return 3

    message = EmailMessage()
    message["Subject"] = "PICT mail configuration test"
    message["From"] = formataddr((cfg["from_name"], cfg["from_email"]))
    message["To"] = recipient
    message["Reply-To"] = cfg["reply_to"] or cfg["from_email"]
    message.set_content(
        "This is a PICT ExamSched mail configuration test.\n\n"
        "If you received this message, the SMTP connection, authentication and "
        "sender details are working.\n\n"
        "No timetable or faculty data was changed by this test.\n"
    )
    meta = {
        "kind": "mail_configuration_test",
        "intended_to": recipient,
        "delivered_to": recipient,
        "subject": message["Subject"],
        "attachments": [],
    }
    results = send_messages([(message, meta)], cfg)
    result = results[0]
    if result["status"] in {"sent", "dry_run"}:
        print(f"[OK]   Test message status: {result['status']}")
        if result.get("eml_path"):
            print(f"       Dry-run file: {result['eml_path']}")
        else:
            print(f"       Delivered to: {recipient}")
        return 0
    print(f"[FAIL] Test message failed: {result.get('error')}")
    return 4


if __name__ == "__main__":
    sys.exit(main())
