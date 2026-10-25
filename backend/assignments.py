"""Single source of truth for published invigilation duties.

One document per duty (seat) is stored in the ``assignments`` collection.
Teacher history, duty counts, exports and rosters are derived from these
documents.  Approved swaps change two assignment documents inside a MongoDB
transaction and are written to ``swap_log``.
"""

from __future__ import annotations

import os
from datetime import datetime

import pandas as pd
from bson.objectid import ObjectId
from pymongo.errors import OperationFailure, PyMongoError

try:
    from db import format_date_to_standard
except Exception:  # pragma: no cover - only hit in isolated imports
    def format_date_to_standard(value):
        return str(value) if value else "TBD"

from eligibility import normalize_role, normalize_year, teacher_can_invigilate


class AssignmentConflict(Exception):
    """Raised when a swap cannot be applied without losing consistency."""


def utcnow():
    return datetime.utcnow().isoformat() + "Z"


def _database(database=None):
    if database is not None:
        return database
    from db import get_db
    return get_db()


def _transactions_unsupported(exc):
    message = str(exc).lower()
    return (
        getattr(exc, "code", None) == 20
        or "transaction numbers are only allowed on a replica set" in message
        or "transactions are not supported" in message
    )


def run_in_transaction(database, operation):
    """Run ``operation(session)`` in a transaction when the server supports it.

    Atlas and any replica-set deployment support transactions.  A local
    standalone mongod does not; in that case the operation runs without a
    session and callers still receive a clear error if an update fails.
    """
    client = getattr(database, "client", None)
    if client is None:
        return operation(None)
    session = client.start_session()
    try:
        try:
            with session.start_transaction():
                return operation(session)
        except OperationFailure as exc:
            if _transactions_unsupported(exc):
                return operation(None)
            raise
    finally:
        session.end_session()


def _teacher_lookup(database):
    lookup = {}
    for doc in database["teachers"].find({}):
        if doc.get("teacher_id"):
            lookup[str(doc["teacher_id"])] = doc
        if doc.get("email"):
            lookup[str(doc["email"]).lower()] = doc
    return lookup


def normalize_allowed_years(raw):
    if not raw:
        return []
    if isinstance(raw, str):
        text = raw.strip()
        if text.startswith("[") and text.endswith("]"):
            try:
                import ast
                parsed = ast.literal_eval(text)
                if isinstance(parsed, (list, tuple, set)):
                    raw = list(parsed)
                else:
                    raw = [text]
            except (ValueError, SyntaxError):
                raw = [part for part in text.replace(";", ",").split(",") if part.strip()]
        else:
            raw = [part for part in text.replace(";", ",").split(",") if part.strip()]
    return sorted({normalize_year(value) for value in raw if value})


def assignment_document(timetable_id, department, duty, teacher=None,
                        confirmed_at=None, active=True):
    role = normalize_role(duty.get("role_required") or duty.get("role") or "Junior")
    teacher = teacher or {}
    teacher_id = duty.get("teacher_id")
    return {
        "timetable_id": str(timetable_id),
        "department": department,
        "duty_id": str(duty.get("duty_id")),
        "date": format_date_to_standard(duty.get("date")),
        "slot": str(duty.get("slot")),
        "session": duty.get("session") or "TBD",
        "course_id": str(duty.get("course_id") or "ALL"),
        "year": normalize_year(duty.get("year", "ALL")),
        "department_required": duty.get("department") or "General",
        "exam_department": duty.get("exam_department") or department,
        "cross_department": bool(duty.get("cross_department", False)),
        "allowed_years": normalize_allowed_years(duty.get("allowed_years")),
        "room": duty.get("room") or duty.get("room_assigned") or "TBD",
        "block": duty.get("block") or "",
        "role": role,
        "teacher_id": str(teacher_id) if teacher_id else None,
        "teacher_name": duty.get("teacher_name") or teacher.get("name"),
        "pinned": bool(duty.get("pinned")),
        "active": bool(active),
        "version": 1,
        "confirmed_at": confirmed_at or utcnow(),
        "created_at": utcnow(),
        "updated_at": utcnow(),
    }


def sync_timetable_assignments(timetable_id, department, duties, teachers=None,
                               confirmed_at=None, database=None):
    """Archive the department's previous active seats and insert this version."""
    database = _database(database)
    teachers = teachers or {}
    docs = []
    for duty in duties:
        teacher = teachers.get(str(duty.get("teacher_id")), {}) if duty.get("teacher_id") else {}
        docs.append(assignment_document(
            timetable_id, department, duty, teacher=teacher,
            confirmed_at=confirmed_at, active=True,
        ))

    def operation(session):
        if docs:
            database["assignments"].insert_many(docs, session=session, ordered=True)
        return len(docs)

    # Insert first, then archive older seats.  If the insert fails, the old
    # active roster is untouched.  The two operations are in one transaction
    # when MongoDB supports transactions.
    def transactional(session):
        previous = list(database["assignments"].find(
            {"department": department, "active": True, "timetable_id": {"$ne": str(timetable_id)}},
            session=session,
        ))
        previous_ids = [doc["_id"] for doc in previous]
        result = operation(session)
        if previous_ids:
            database["assignments"].update_many(
                {"_id": {"$in": previous_ids}},
                {"$set": {"active": False, "archived_at": utcnow()}},
                session=session,
            )
        return result

    try:
        return run_in_transaction(database, transactional)
    except Exception:
        # Best-effort cleanup for standalone deployments without transactions.
        try:
            database["assignments"].delete_many({"timetable_id": str(timetable_id)})
        except Exception:
            pass
        raise


def migrate_legacy_timetables(database=None):
    """Import existing timetable/history data into ``assignments`` once."""
    database = _database(database)
    if database["assignments"].count_documents({}) > 0:
        return 0

    timetables = list(database["timetables"].find({}).sort([("confirmed_at", 1)]))
    latest_by_department = {}
    for doc in timetables:
        dept = doc.get("department") or "ALL"
        latest_by_department[dept] = str(doc["_id"])

    docs = []
    represented = set()
    teacher_lookup = _teacher_lookup(database)
    for tt in timetables:
        tt_id = str(tt["_id"])
        dept = tt.get("department") or "ALL"
        confirmed_at = tt.get("confirmed_at") or utcnow()
        active = latest_by_department.get(dept) == tt_id and not tt.get("cleared")
        for index, duty in enumerate(tt.get("teacher_duties") or []):
            teacher_id = str(duty.get("teacher_id") or "")
            if not teacher_id:
                continue
            teacher = teacher_lookup.get(teacher_id, {})
            duty = dict(duty)
            duty.setdefault("duty_id", duty.get("duty_id") or f"tt-{tt_id}-{index + 1}")
            duty["teacher_name"] = duty.get("teacher_name") or teacher.get("name")
            docs.append(assignment_document(tt_id, dept, duty, teacher=teacher,
                                            confirmed_at=confirmed_at, active=active))
            role = normalize_role(duty.get("role_required") or duty.get("role"))
            represented.add((teacher_id, format_date_to_standard(duty.get("date")),
                             str(duty.get("slot")), role, str(duty.get("course_id") or "ALL")))

    # Preserve history rows that predate the timetables collection.
    counter = 0
    for teacher in database["teachers"].find({}):
        tid = str(teacher.get("teacher_id") or "")
        if not tid:
            continue
        for entry in teacher.get("history") or []:
            date_value = format_date_to_standard(entry.get("exam_date"))
            slot = str(entry.get("slot_id") or entry.get("slot") or "")
            role = normalize_role(entry.get("role_assigned") or entry.get("role") or "Junior")
            course_id = str(entry.get("course_id") or "ALL")
            key = (tid, date_value, slot, role, course_id)
            if key in represented:
                continue
            counter += 1
            duty = {
                "duty_id": f"legacy-{counter}",
                "date": date_value,
                "slot": slot,
                "session": entry.get("session") or "TBD",
                "course_id": course_id,
                "year": entry.get("year") or "ALL",
                "department": entry.get("department") or "General",
                "room": entry.get("room_assigned") or entry.get("room") or "TBD",
                "role_required": role,
                "teacher_id": tid,
                "teacher_name": teacher.get("name"),
            }
            doc = assignment_document("legacy", entry.get("department") or "General", duty,
                                      teacher=teacher, confirmed_at=entry.get("assigned_at"), active=False)
            doc["legacy"] = True
            docs.append(doc)
            represented.add(key)

    if docs:
        database["assignments"].insert_many(docs, ordered=False)
    return len(docs)


def get_active_timetable(database=None):
    database = _database(database)
    # The ``ALL`` document published by the multi-department flow is a view
    # for the UI, not a department roster.  Legacy roster/export code must keep
    # seeing a real department document.
    return database["timetables"].find_one(
        {"cleared": {"$ne": True}, "multi_department": {"$ne": True}},
        sort=[("confirmed_at", -1)],
    )


def get_timetable(timetable_id=None, database=None):
    database = _database(database)
    if timetable_id:
        try:
            return database["timetables"].find_one({"_id": ObjectId(str(timetable_id))})
        except Exception:
            return None
    return get_active_timetable(database)


def derive_history(teacher_id, database=None):
    """Return a teacher's duty history from the assignments collection."""
    database = _database(database)
    rows = list(database["assignments"].find({"teacher_id": str(teacher_id)}))
    if not rows:
        teacher = database["teachers"].find_one(
            {"$or": [{"teacher_id": str(teacher_id)}, {"email": str(teacher_id)}]}
        )
        return list((teacher or {}).get("history") or [])

    deduped = {}
    for row in rows:
        key = (
            row.get("date"), str(row.get("slot")), normalize_role(row.get("role")),
            row.get("course_id"), row.get("room"),
        )
        current = deduped.get(key)
        stamp = str(row.get("updated_at") or row.get("confirmed_at") or row.get("created_at") or "")
        if current is None or stamp > str(current.get("_stamp") or ""):
            deduped[key] = {**row, "_stamp": stamp}

    history = []
    for row in deduped.values():
        history.append({
            "slot_id": str(row.get("slot")),
            "slot": str(row.get("slot")),
            "exam_date": row.get("date"),
            "session": row.get("session") or "TBD",
            "course_id": row.get("course_id") or "ALL",
            "year": row.get("year") or "ALL",
            "department": row.get("department_required") or row.get("department") or "General",
            "role_assigned": normalize_role(row.get("role")),
            "room_assigned": row.get("room") or "TBD",
            "assigned_at": row.get("confirmed_at") or row.get("created_at"),
            "updated_at": row.get("updated_at"),
            "pinned": bool(row.get("pinned")),
        })
    history.sort(key=lambda item: (str(item.get("exam_date")), str(item.get("slot_id"))))
    return history


def rebuild_teacher_history(teacher_ids=None, database=None):
    """Recompute the teacher.history cache from authoritative assignments."""
    database = _database(database)
    query = {"teacher_id": {"$in": [str(tid) for tid in teacher_ids]}} if teacher_ids else {}
    teacher_query = {"teacher_id": {"$in": [str(tid) for tid in teacher_ids]}} if teacher_ids else {}
    for teacher in database["teachers"].find(teacher_query):
        tid = str(teacher.get("teacher_id") or "")
        if not tid:
            continue
        history = derive_history(tid, database=database)
        counts = {"squad": 0, "junior": 0, "senior": 0}
        for entry in history:
            role_key = normalize_role(entry.get("role_assigned")).lower()
            if role_key not in counts:
                role_key = "junior"
            counts[role_key] += 1
        last_role = history[-1].get("role_assigned") if history else "N/A"
        cycle_start = str(teacher.get("fairness_cycle_started_at") or "")
        served_high = False
        for entry in history:
            if normalize_role(entry.get("role_assigned")) not in {"Senior", "Squad"}:
                continue
            stamp = str(entry.get("updated_at") or entry.get("assigned_at") or "")
            if not cycle_start or (stamp and stamp > cycle_start):
                served_high = True
                break
        database["teachers"].update_one(
            {"_id": teacher["_id"]},
            {"$set": {
                "history": history,
                "duty_counts": counts,
                "last_role": last_role,
                "has_served_high_role": served_high,
            }},
        )


def _eligibility_duty(row):
    """Adapt an assignment document to the eligibility check's duty shape."""
    return {
        "year": row.get("year") or "ALL",
        "allowed_years": row.get("allowed_years") or [],
        "department": row.get("department_required") or row.get("department"),
    }


def find_assignment(teacher_id, date_value, slot_value, timetable_id=None,
                    database=None, active_only=True):
    database = _database(database)
    query = {
        "teacher_id": str(teacher_id),
        "date": format_date_to_standard(date_value),
        "slot": str(slot_value),
    }
    if timetable_id:
        query["timetable_id"] = str(timetable_id)
    if active_only:
        query["active"] = True
    return database["assignments"].find_one(query, sort=[("version", -1)])


def _update_timetable_duty(database, timetable_id, duty_id, teacher_id, teacher_name,
                           session=None, course_id=None, room=None, session_obj=None,
                           revision=None):
    timetable = database["timetables"].find_one({"_id": ObjectId(str(timetable_id))}, session=session_obj)
    if not timetable:
        raise AssignmentConflict("Published timetable document was not found")
    duties = list(timetable.get("teacher_duties") or [])
    index = None
    for i, duty in enumerate(duties):
        if str(duty.get("duty_id")) == str(duty_id):
            index = i
            break
    if index is None:
        raise AssignmentConflict("Timetable document is out of sync with assignments")
    duties[index]["teacher_id"] = teacher_id
    duties[index]["teacher_name"] = teacher_name
    if session is not None:
        duties[index]["session"] = session
    if course_id is not None:
        duties[index]["course_id"] = course_id
    if room is not None:
        duties[index]["room_assigned"] = room
    all_duties = list(timetable.get("all_duties") or [])
    for duty in all_duties:
        if str(duty.get("duty_id")) != str(duty_id):
            continue
        duty["teacher_id"] = teacher_id
        duty["teacher_name"] = teacher_name
        if session is not None:
            duty["session"] = session
        if course_id is not None:
            duty["course_id"] = course_id
        if room is not None:
            duty["room"] = room
        break
    if revision is None:
        revision = int(timetable.get("revision") or 1) + 1
    update_fields = {"teacher_duties": duties, "last_modified_at": utcnow(), "revision": revision}
    if all_duties:
        update_fields["all_duties"] = all_duties
    result = database["timetables"].update_one(
        {"_id": timetable["_id"]},
        {"$set": update_fields},
        session=session_obj,
    )
    if result.modified_count != 1:
        raise AssignmentConflict("Unable to update published timetable")
    return revision


def swap_assignments(timetable_id, requester_id, partner_id, requester_date,
                     requester_slot, partner_date, partner_slot, actor_id=None,
                     reason="", database=None):
    """Atomically swap two assignment seats and keep the timetable in sync."""
    database = _database(database)
    timetable = get_timetable(timetable_id, database=database)
    if not timetable:
        raise AssignmentConflict("Published timetable not found")
    timetable_id = str(timetable["_id"])

    teacher_lookup = _teacher_lookup(database)
    requester = teacher_lookup.get(str(requester_id))
    partner = teacher_lookup.get(str(partner_id))
    if not requester or not partner:
        raise AssignmentConflict("Requester or swap partner not found")
    requester_key = str(requester.get("teacher_id") or requester_id)
    partner_key = str(partner.get("teacher_id") or partner_id)
    if requester_key == partner_key:
        raise AssignmentConflict("A duty cannot be swapped with itself")

    req_date = format_date_to_standard(requester_date)
    par_date = format_date_to_standard(partner_date)
    req_slot = str(requester_slot)
    par_slot = str(partner_slot)

    def operation(session):
        first = database["assignments"].find_one({
            "timetable_id": timetable_id, "teacher_id": requester_key,
            "date": req_date, "slot": req_slot, "active": True,
        }, session=session)
        second = database["assignments"].find_one({
            "timetable_id": timetable_id, "teacher_id": partner_key,
            "date": par_date, "slot": par_slot, "active": True,
        }, session=session)
        if not first or not second:
            raise AssignmentConflict("One or both duties no longer exist on the published roster")
        if first.get("_id") == second.get("_id"):
            raise AssignmentConflict("The selected duties are the same seat")
        if normalize_role(first.get("role")) != normalize_role(second.get("role")):
            raise AssignmentConflict("Swap partners must have the same invigilation role")
        if not teacher_can_invigilate(requester, _eligibility_duty(second)):
            raise AssignmentConflict("Requester is not eligible for the partner's duty (year/department rule)")
        if not teacher_can_invigilate(partner, _eligibility_duty(first)):
            raise AssignmentConflict("Partner is not eligible for the requester's duty (year/department rule)")

        # One duty per day, checked again at approval time.
        other_req = database["assignments"].find_one({
            "timetable_id": timetable_id, "teacher_id": requester_key,
            "date": par_date, "active": True, "_id": {"$ne": second["_id"]},
        }, session=session)
        other_par = database["assignments"].find_one({
            "timetable_id": timetable_id, "teacher_id": partner_key,
            "date": req_date, "active": True, "_id": {"$ne": first["_id"]},
        }, session=session)
        if other_req or other_par:
            raise AssignmentConflict("A teacher would have two duties on the same day")

        update1 = database["assignments"].update_one(
            {"_id": first["_id"], "version": first.get("version", 1)},
            {"$set": {
                "teacher_id": partner_key,
                "teacher_name": partner.get("name"),
                "version": int(first.get("version", 1)) + 1,
                "updated_at": utcnow(),
                "swapped_with": requester_key,
            }},
            session=session,
        )
        update2 = database["assignments"].update_one(
            {"_id": second["_id"], "version": second.get("version", 1)},
            {"$set": {
                "teacher_id": requester_key,
                "teacher_name": requester.get("name"),
                "version": int(second.get("version", 1)) + 1,
                "updated_at": utcnow(),
                "swapped_with": partner_key,
            }},
            session=session,
        )
        if update1.modified_count != 1 or update2.modified_count != 1:
            raise AssignmentConflict("The roster changed while approving this swap; reload and try again")

        revision = _update_timetable_duty(
            database, timetable_id, first.get("duty_id"), partner_key, partner.get("name"),
            session_obj=session,
        )
        revision = _update_timetable_duty(
            database, timetable_id, second.get("duty_id"), requester_key, requester.get("name"),
            session_obj=session, revision=revision,
        )

        database["swap_log"].insert_one({
            "timetable_id": timetable_id,
            "requester_id": requester_key,
            "partner_id": partner_key,
            "requester_date": req_date,
            "requester_slot": req_slot,
            "partner_date": par_date,
            "partner_slot": par_slot,
            "role": normalize_role(first.get("role")),
            "actor_id": str(actor_id) if actor_id else None,
            "reason": reason,
            "revision": revision,
            "created_at": utcnow(),
        }, session=session)
        return {
            "timetable_id": timetable_id,
            "requester_assignment": first,
            "partner_assignment": second,
            "revision": revision,
        }

    return run_in_transaction(database, operation)


def move_assignment(timetable_id, requester_id, target_date, target_slot,
                    actor_id=None, reason="", database=None):
    """Move one teacher to a vacant seat of the same role, atomically.

    Direct moves are disabled by default in the API because they can change a
    day's headcount.  This helper is used only when ``ALLOW_DIRECT_MOVE`` is
    enabled, and it always moves into an existing *vacant* seat so the total
    demand does not change.
    """
    database = _database(database)
    timetable = get_timetable(timetable_id, database=database)
    if not timetable:
        raise AssignmentConflict("Published timetable not found")
    timetable_id = str(timetable["_id"])

    teacher_lookup = _teacher_lookup(database)
    requester = teacher_lookup.get(str(requester_id))
    if not requester:
        raise AssignmentConflict("Requester teacher not found")
    requester_key = str(requester.get("teacher_id") or requester_id)

    tgt_date = format_date_to_standard(target_date)
    tgt_slot = str(target_slot)

    def operation(session):
        current = database["assignments"].find_one({
            "timetable_id": timetable_id, "teacher_id": requester_key,
            "active": True,
        }, session=session, sort=[("date", 1), ("slot", 1)])
        if not current:
            raise AssignmentConflict("The requester has no active duty on the published roster")
        if str(current.get("date")) == tgt_date and str(current.get("slot")) == tgt_slot:
            raise AssignmentConflict("The request is already on that duty")

        vacant = database["assignments"].find_one({
            "timetable_id": timetable_id,
            "date": tgt_date,
            "slot": tgt_slot,
            "role": current.get("role"),
            "active": True,
            "$or": [{"teacher_id": None}, {"teacher_id": {"$exists": False}}, {"teacher_id": ""}],
        }, session=session)
        if not vacant:
            raise AssignmentConflict("No vacant seat exists for that date, slot and role")

        if not teacher_can_invigilate(requester, _eligibility_duty(vacant)):
            raise AssignmentConflict("Requester is not eligible for the target duty (year/department rule)")

        other = database["assignments"].find_one({
            "timetable_id": timetable_id,
            "teacher_id": requester_key,
            "date": tgt_date,
            "active": True,
            "_id": {"$ne": current["_id"]},
        }, session=session)
        if other:
            raise AssignmentConflict("The teacher already has a duty on the target date")

        update = database["assignments"].update_one(
            {"_id": vacant["_id"], "version": vacant.get("version", 1)},
            {"$set": {
                "teacher_id": requester_key,
                "teacher_name": requester.get("name"),
                "version": int(vacant.get("version", 1)) + 1,
                "updated_at": utcnow(),
                "moved_from": current.get("duty_id"),
            }},
            session=session,
        )
        if update.modified_count != 1:
            raise AssignmentConflict("The roster changed while booking this seat; reload and try again")

        release = database["assignments"].update_one(
            {"_id": current["_id"], "version": current.get("version", 1)},
            {"$set": {
                "teacher_id": None,
                "teacher_name": None,
                "version": int(current.get("version", 1)) + 1,
                "updated_at": utcnow(),
            }},
            session=session,
        )
        if release.modified_count != 1:
            raise AssignmentConflict("The roster changed while moving this duty; reload and try again")

        # Keep the published timetable's assigned-duty rows in sync.  The
        # teacher's old row disappears (the seat is now vacant) and the newly
        # filled seat is appended with the same duty_id used by assignments.
        timetable_doc = database["timetables"].find_one(
            {"_id": ObjectId(timetable_id)}, session=session
        )
        if not timetable_doc:
            raise AssignmentConflict("Published timetable document was not found")
        duties = [
            duty for duty in (timetable_doc.get("teacher_duties") or [])
            if str(duty.get("duty_id")) != str(current.get("duty_id"))
        ]
        duties.append({
            "duty_id": vacant.get("duty_id"),
            "slot": vacant.get("slot"),
            "date": vacant.get("date"),
            "session": vacant.get("session"),
            "course_id": vacant.get("course_id"),
            "role": normalize_role(vacant.get("role")),
            "teacher_id": requester_key,
            "teacher_name": requester.get("name"),
            "year": vacant.get("year") or "ALL",
            "department": vacant.get("department_required") or "General",
            "room_assigned": vacant.get("room"),
        })
        revision = int(timetable_doc.get("revision") or 1) + 1
        all_duties = list(timetable_doc.get("all_duties") or [])
        for duty in all_duties:
            duty_id = str(duty.get("duty_id"))
            if duty_id == str(current.get("duty_id")):
                duty["teacher_id"] = None
                duty["teacher_name"] = None
            elif duty_id == str(vacant.get("duty_id")):
                duty["teacher_id"] = requester_key
                duty["teacher_name"] = requester.get("name")
        update_fields = {
            "teacher_duties": duties,
            "last_modified_at": utcnow(),
            "revision": revision,
        }
        if all_duties:
            update_fields["all_duties"] = all_duties
        saved = database["timetables"].update_one(
            {"_id": timetable_doc["_id"]},
            {"$set": update_fields},
            session=session,
        )
        if saved.modified_count != 1:
            raise AssignmentConflict("Unable to update published timetable")

        database["swap_log"].insert_one({
            "timetable_id": timetable_id,
            "action": "move",
            "requester_id": requester_key,
            "from_date": current.get("date"),
            "from_slot": str(current.get("slot")),
            "to_date": tgt_date,
            "to_slot": tgt_slot,
            "role": normalize_role(current.get("role")),
            "actor_id": str(actor_id) if actor_id else None,
            "reason": reason,
            "revision": revision,
            "created_at": utcnow(),
        }, session=session)
        return {
            "timetable_id": timetable_id,
            "revision": revision,
            "from": {"date": current.get("date"), "slot": str(current.get("slot"))},
            "to": {"date": tgt_date, "slot": tgt_slot},
        }

    return run_in_transaction(database, operation)


def count_approved_swaps(timetable_id, database=None):
    database = _database(database)
    return database["swap_log"].count_documents({"timetable_id": str(timetable_id)})


def build_all_teacher_duties_payload(timetable_doc, teachers):
    """Build the coordinator duties payload used by PDF exports."""
    teachers_by_id = {}
    for teacher in teachers or []:
        tid = teacher.get("teacher_id") or teacher.get("id")
        if tid:
            teachers_by_id[str(tid)] = teacher

    result = []
    if timetable_doc and timetable_doc.get("teacher_duties"):
        for duty in timetable_doc.get("teacher_duties") or []:
            tid = str(duty.get("teacher_id") or "")
            if not tid:
                continue
            teacher = teachers_by_id.get(tid, {})
            result.append({
                "duty_id": duty.get("duty_id"),
                "teacher_id": tid,
                "teacher_name": duty.get("teacher_name") or teacher.get("name", ""),
                "date": duty.get("date"),
                "slot": str(duty.get("slot")),
                "session": duty.get("session") or "TBD",
                "course_id": duty.get("course_id") or "ALL",
                "role": normalize_role(duty.get("role") or duty.get("role_required")),
                "room_assigned": duty.get("room_assigned") or duty.get("room") or "TBD",
                "year": duty.get("year") or "ALL",
                "department": duty.get("department") or "General",
                "last_role": teacher.get("last_role", "N/A"),
                "has_served_high_role": teacher.get("has_served_high_role"),
            })
        return result

    for teacher in teachers or []:
        tid = str(teacher.get("teacher_id") or teacher.get("id") or "")
        for entry in teacher.get("history") or []:
            result.append({
                "duty_id": entry.get("duty_id"),
                "teacher_id": tid,
                "teacher_name": teacher.get("name", ""),
                "date": entry.get("exam_date"),
                "slot": str(entry.get("slot_id") or entry.get("slot") or ""),
                "session": entry.get("session") or "TBD",
                "course_id": entry.get("course_id") or "ALL",
                "role": normalize_role(entry.get("role_assigned") or entry.get("role")),
                "room_assigned": entry.get("room_assigned") or "TBD",
                "year": entry.get("year") or "ALL",
                "department": entry.get("department") or "General",
                "last_role": teacher.get("last_role", "N/A"),
                "has_served_high_role": teacher.get("has_served_high_role"),
            })
    return result


def select_best_timetable_doc(candidates):
    """Prefer the document with the most published duty rows."""
    candidates = list(candidates or [])
    if not candidates:
        return None
    return max(candidates, key=lambda doc: len(doc.get("teacher_duties") or []))
