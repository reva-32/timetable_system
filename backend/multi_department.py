"""Multi-department examination scheduling.

A generation run schedules every selected department on a common examination
calendar.  Each department keeps its own timetable (and therefore its own PDF
page), while the calendar, rooms and invigilation pool are coordinated so that
no faculty member gets two duties on one date and no student has two papers in
the same session.

The public output is a timetable PDF; rooms and duties are rendered into a
separate detailed PDF by :mod:`department_pdf`.
"""

from __future__ import annotations

import math
import os
import re
from collections import defaultdict
from datetime import date, datetime, timedelta

from department_config import (
    CATEGORY_ORDER,
    DEFAULT_SESSION,
    DEPARTMENT_BLOCKS,
    DEPARTMENT_ORDER,
    block_names,
    category_session,
    department_short_title,
    department_title,
    squads_for_blocks,
)
from eligibility import (
    canonical_department,
    effective_teaching_years,
    is_fy_department,
    normalize_role,
    normalize_year,
    teacher_can_invigilate,
)
from main import (
    assign_teachers,
    build_conflict_graph,
    dsatur_coloring,
    load_data,
    load_room_data,
    load_teacher_data,
)

_CATEGORY_ALIASES = {
    "PC": "PC",
    "CORE": "PC",
    "PROGRAMMECORE": "PC",
    "PROGRAMCORECOURSE": "PC",
    "PE": "PE",
    "ELECTIVE": "PE",
    "PROGRAMMEELECTIVE": "PE",
    "PROGRAMELECTIVE": "PE",
    "MD": "MD",
    "MULTIDISCIPLINARY": "MD",
    "MULTIDISCIPLINARYCOURSE": "MD",
    "HC": "HC",
    "HONOURS": "HC",
    "HONORS": "HC",
    "MINOR": "HC",
    "HONOURSMINOR": "HC",
}


def _text(value, default=""):
    if value is None:
        return default
    text = str(value).strip()
    return text if text and text.lower() not in {"nan", "none", "n/a"} else default


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _text(value)
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Invalid date: {value}")


def _date_label(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%d-%b-%Y")
    return _text(value)


def _category_of(course_id, info):
    raw = _text(info.get("exam_category"))
    if raw:
        key = re.sub(r"[^A-Z]", "", raw.upper())
        if key in _CATEGORY_ALIASES:
            return _CATEGORY_ALIASES[key]
        for category in CATEGORY_ORDER:
            if key.startswith(category):
                return category
    for candidate in (str(course_id), _text(info.get("course_code"))):
        match = re.search(r"(PC|PE|MD|HC)", candidate.upper())
        if match:
            return match.group(1)
    return ""


def _course_sort_key(course_id, info):
    order = _text(info.get("exam_order"))
    try:
        numeric_order = int(order)
    except (TypeError, ValueError):
        numeric_order = 10 ** 6
    return (numeric_order, str(info.get("course_code") or course_id), str(course_id))


def _group_courses(courses, student_courses, exam_type, year_filter):
    """Group one department's courses into conflict-free examination groups.

    Courses with the same paper category are grouped together first.  This is
    what lets five electives share the Thursday slot while keeping all core
    papers on separate days.  DSATUR then splits a category if two of its
    papers really do share students.
    """
    buckets = defaultdict(dict)
    for course_id, info in courses.items():
        category = _category_of(course_id, info)
        year_key = normalize_year(info.get("year")) if year_filter == "ALL" else ""
        key = (year_key, category)
        buckets[key][course_id] = info

    def bucket_order(item):
        (year_key, category), _ = item
        category_index = CATEGORY_ORDER.index(category) if category in CATEGORY_ORDER else len(CATEGORY_ORDER)
        return (category_index, str(year_key))

    groups = []
    for (year_key, category), bucket in sorted(buckets.items(), key=bucket_order):
        course_ids = sorted(bucket, key=lambda cid: _course_sort_key(cid, bucket[cid]))
        graph = build_conflict_graph(student_courses, all_courses=course_ids)
        coloring = dsatur_coloring(graph)
        by_color = defaultdict(list)
        for course_id in course_ids:
            by_color[coloring.get(course_id, 0)].append(course_id)

        def color_order(color):
            return min(_course_sort_key(cid, bucket[cid]) for cid in by_color[color])

        for color in sorted(by_color, key=color_order):
            group_ids = sorted(by_color[color], key=lambda cid: _course_sort_key(cid, bucket[cid]))
            session = _group_session(group_ids, bucket, category, exam_type)
            groups.append({
                "category": category,
                "category_label": category,
                "session": session,
                "courses": [
                    {
                        "course_id": cid,
                        "course_code": _text(bucket[cid].get("course_code"), cid),
                        "course_name": _text(bucket[cid].get("course_name"), "Course"),
                        "year": normalize_year(bucket[cid].get("year")),
                        "students_count": int(bucket[cid].get("students_count") or 0),
                    }
                    for cid in group_ids
                ],
            })
    return groups


def _group_session(course_ids, courses, category, exam_type):
    sessions = {
        _text(courses[cid].get("session"))
        for cid in course_ids
        if _text(courses[cid].get("session"))
        and _text(courses[cid].get("session")).lower() not in {"general", "tbd"}
    }
    if len(sessions) == 1:
        return sessions.pop()
    return category_session(category, exam_type)


def _available_dates(start, end):
    current = start
    dates = []
    while current <= end:
        if current.weekday() != 6:  # Sunday is not an examination day.
            dates.append(current)
        current += timedelta(days=1)
    return dates


def _select_calendar_dates(start, end, count):
    """Pick ``count`` dates spread across the period, including both ends."""
    candidates = _available_dates(start, end)
    if count <= 0:
        return []
    if count > len(candidates):
        raise ValueError(
            f"The selected period has {len(candidates)} examination day(s) "
            f"(Sundays excluded), but the timetable needs {count}. "
            "Please extend the end date."
        )
    if count == 1:
        return [candidates[0]]
    indices = [round(index * (len(candidates) - 1) / (count - 1)) for index in range(count)]
    # Rounding can rarely duplicate an index; fill any gaps deterministically.
    unique = []
    for index in indices:
        while index in unique:
            index += 1
        if index >= len(candidates):
            index = next(candidate for candidate in range(len(candidates)) if candidate not in unique)
        unique.append(index)
    return [candidates[index] for index in unique]


def _workbook_exam_dates(slots, start, end):
    """Return the examination dates declared in the workbook's Slots sheet.

    The college circular is built from the uploaded workbook, so its Slots
    sheet is authoritative whenever it contains enough dates for the generated
    paper groups.  The even-spread calendar remains the fallback for workbooks
    that do not provide a usable Slots sheet.
    """
    def slot_order(item):
        try:
            return int(item.get("display_id") or 0)
        except (TypeError, ValueError):
            return 0

    dates = []
    entries = slots or {}
    if isinstance(entries, dict):
        ordered = sorted(entries.values(), key=lambda item: slot_order(item)
                         if isinstance(item, dict) else 0)
    else:
        ordered = list(entries)
    for entry in ordered:
        if not isinstance(entry, dict):
            continue
        raw = entry.get("date")
        if not raw:
            continue
        try:
            value = _as_date(raw)
        except ValueError:
            continue
        if value < start or value > end:
            continue
        if value.weekday() == 6:  # Sunday
            continue
        if value not in dates:
            dates.append(value)
    dates.sort()
    return dates


def _calendar_for_run(slots, start, end, count):
    """Choose the examination dates for one generated run.

    The coordinator's selected period is the outer boundary.  When the
    workbook's Slots sheet contains enough dates inside that period they are
    used, so the PDF can mirror the college circular exactly.  Otherwise the
    run spreads exams evenly across the selected period (Sundays excluded).
    """
    workbook_dates = _workbook_exam_dates(slots, start, end)
    if workbook_dates and len(workbook_dates) >= count:
        return workbook_dates[:count]
    return _select_calendar_dates(start, end, count)


def _block_definitions(department, rooms):
    """Return exactly the configured number of blocks for a department."""
    dept = canonical_department(department)
    configured = DEPARTMENT_BLOCKS.get(dept, 0)
    labels = block_names(dept)
    department_rooms = [
        room for room in rooms
        if canonical_department(room.get("department")) == dept
    ]
    department_rooms.sort(key=lambda room: str(room.get("room_id") or ""))
    blocks = []
    for index, label in enumerate(labels[:configured]):
        room = department_rooms[index] if index < len(department_rooms) else None
        blocks.append({
            "block": label,
            "room_id": _text(room.get("room_id"), label) if room else label,
            "capacity": max(1, int(room.get("capacity") or 0)) if room else 60,
            "courses": [],
            "students": 0,
        })
    return blocks


def _allocate_blocks(courses, blocks):
    """Spread the papers of one session across the department's blocks."""
    usage = [
        {
            "block": block["block"],
            "room_id": block["room_id"],
            "capacity": block["capacity"],
            "courses": [],
            "students": 0,
        }
        for block in blocks
    ]
    if not usage:
        return usage
    cursor = 0
    for course in courses:
        students = max(0, int(course.get("students_count") or 0))
        remaining = max(1, students)
        attempts = 0
        while remaining > 0 and attempts < len(usage) * 4:
            block = usage[cursor % len(usage)]
            cursor += 1
            attempts += 1
            if course["course_id"] not in block["courses"]:
                block["courses"].append(course["course_id"])
            take = min(remaining, max(1, block["capacity"]))
            block["students"] += take
            remaining -= take
        if remaining > 0:
            # Very large cohort: record the overflow against the first block so
            # the detailed PDF never silently hides students.
            usage[0]["students"] += remaining
            if course["course_id"] not in usage[0]["courses"]:
                usage[0]["courses"].append(course["course_id"])
    return usage


def _unique_students_for_courses(course_ids, student_courses):
    course_set = {str(course_id) for course_id in course_ids}
    students = {
        student_id
        for student_id, enrolled in (student_courses or {}).items()
        if course_set.intersection({str(course_id) for course_id in enrolled})
    }
    return students


def _normalize_departments(departments):
    values = departments or DEPARTMENT_ORDER
    result = []
    for value in values:
        dept = canonical_department(value)
        if dept and dept not in result:
            result.append(dept)
    return result


def _build_sessions(departments, groups_by_department, dates):
    """Place each department's ordered groups onto the common calendar."""
    max_groups = max((len(groups) for groups in groups_by_department.values()), default=0)
    if max_groups and max_groups > len(dates):
        raise ValueError(
            "Not enough examination dates for the generated paper groups. "
            "Please extend the exam period."
        )

    sessions_by_department = {}
    for dept in departments:
        groups = groups_by_department.get(dept) or []
        sessions = []
        for index, group in enumerate(groups):
            session_date = dates[index] if index < len(dates) else dates[-1]
            sessions.append({
                "department": dept,
                "position": index,
                "slot": index + 1,
                "date": _date_label(session_date),
                "date_iso": session_date.isoformat(),
                "date_label": _date_label(session_date),
                "session": group["session"],
                "category": group.get("category") or "",
                "categories": sorted({
                    _category_of(course["course_id"], course)
                    for course in group["courses"]
                    if _category_of(course["course_id"], course)
                }),
                "courses": group["courses"],
            })
        sessions_by_department[dept] = sessions
    return sessions_by_department


def _build_department_duties(department, sessions, blocks, year_filter, start_duty_id):
    duties = []
    duty_id = start_duty_id
    squad_count = squads_for_blocks(len(blocks))
    dept = canonical_department(department)
    for session in sessions:
        course_codes = [course.get("course_code") or course.get("course_id") for course in session["courses"]]
        years = sorted({
            normalize_year(course.get("year"))
            for course in session["courses"]
            if normalize_year(course.get("year")) in {"SY", "TY", "BTECH"}
        })
        allowed_years = years or (
            [normalize_year(year_filter)] if normalize_year(year_filter) in {"SY", "TY", "BTECH"} else []
        )
        if not allowed_years:
            allowed_years = ["SY", "TY", "BTECH"]
        duty_year = allowed_years[0] if len(allowed_years) == 1 else "ALL"
        common = {
            "slot": session["slot"],
            "date": session["date_label"],
            "session": session["session"],
            "exam_department": dept,
            "cross_department": True,
            "department": dept,
            "year": duty_year,
            "allowed_years": list(allowed_years),
            "course_id": ",".join(str(code) for code in course_codes if code) or "ALL",
            "courses": [str(code) for code in course_codes if code],
            "category": session.get("category") or "",
            "exam_group_id": f"{dept}-{session['position'] + 1}",
        }
        for block in blocks:
            duties.append({
                **common,
                "duty_id": duty_id,
                "role_required": "Junior",
                "block": block["block"],
                "room": block["room_id"],
                "room_assigned": block["room_id"],
            })
            duty_id += 1
        for _ in range(squad_count):
            duties.append({
                **common,
                "duty_id": duty_id,
                "role_required": "Squad",
                "block": "Squad",
                "room": "Roaming",
                "room_assigned": "Roaming",
            })
            duty_id += 1
        duties.append({
            **common,
            "duty_id": duty_id,
            "role_required": "Senior",
            "block": "Control",
            "room": "Control",
            "room_assigned": "Control",
        })
        duty_id += 1
    return duties, squad_count


def _history_in_period(history, start, end):
    """Drop duties that belong to the exam period being (re)generated.

    A new timetable replaces the old timetable for the same dates.  Keeping
    those dates in the fairness history would make every faculty member appear
    busy on every examination day and prevent a regeneration.
    """
    if start is None or end is None:
        return list(history or [])
    kept = []
    for entry in history or []:
        raw = entry.get("exam_date") or entry.get("date")
        try:
            value = _as_date(raw)
        except (TypeError, ValueError):
            kept.append(entry)
            continue
        if start <= value <= end:
            continue
        kept.append(entry)
    return kept


def _counts_from_history(history):
    counts = {"squad": 0, "junior": 0, "senior": 0}
    for entry in history or []:
        role = normalize_role(
            entry.get("role_assigned") or entry.get("role") or "Junior"
        ).lower()
        if role in counts:
            counts[role] += 1
    return counts


def _faculty_inputs(teachers, start=None, end=None):
    """Build the faculty/status structures consumed by ``assign_teachers``."""
    from db import get_all_faculty_status, init_faculty

    init_faculty(teachers)
    faculty_status = get_all_faculty_status()
    teachers = {
        str(teacher_id): teacher
        for teacher_id, teacher in teachers.items()
        if not faculty_status.get(str(teacher_id), {}).get("is_admin", False)
        and not is_fy_department(teacher.get("department"))
    }
    db_duty_counts = {}
    fairness_map = {}
    for teacher_id, status in list(faculty_status.items()):
        history = _history_in_period(status.get("history") or [], start, end)
        if history or status.get("history"):
            counts = _counts_from_history(history)
            high_role = any(
                normalize_role(entry.get("role_assigned") or entry.get("role"))
                in {"Senior", "Squad"}
                for entry in history
            )
        else:
            counts = status.get("duty_counts")
            if not counts:
                flat = int(status.get("duty_count") or 0)
                role = normalize_role(status.get("role") or "Junior").lower()
                counts = {
                    "squad": flat if role == "squad" else 0,
                    "junior": flat if role == "junior" else 0,
                    "senior": flat if role == "senior" else 0,
                }
            high_role = bool(status.get("has_served_high_role", False))
        db_duty_counts[str(teacher_id)] = counts
        fairness_map[str(teacher_id)] = high_role
        faculty_status[str(teacher_id)] = {
            **status,
            "history": history,
            "duty_counts": counts,
            "has_served_high_role": high_role,
        }
    return teachers, faculty_status, db_duty_counts, fairness_map


def _validate_assignments(assignments, duties, teachers):
    """Return a list of human-readable invariant violations."""
    problems = []
    seen_teacher_date = set()
    duty_by_id = {str(duty["duty_id"]): duty for duty in duties}
    teacher_departments = {
        str(teacher_id): canonical_department(teacher.get("department"))
        for teacher_id, teacher in teachers.items()
    }
    for assignment in assignments:
        teacher_id = str(assignment.get("teacher_id"))
        date_value = str(assignment.get("date"))
        key = (teacher_id, date_value)
        if key in seen_teacher_date:
            problems.append(f"{teacher_id} has more than one duty on {date_value}")
        seen_teacher_date.add(key)

        duty = duty_by_id.get(str(assignment.get("duty_id")))
        if not duty:
            problems.append(f"Duty {assignment.get('duty_id')} is not part of the run")
            continue
        teacher = teachers.get(teacher_id)
        if not teacher:
            problems.append(f"Unknown teacher {teacher_id}")
            continue
        if is_fy_department(teacher.get("department")):
            problems.append(f"FY faculty {teacher_id} was allocated to a department paper")
        exam_department = canonical_department(duty.get("exam_department") or duty.get("department"))
        if teacher_departments.get(teacher_id) == exam_department:
            problems.append(f"{teacher_id} was allocated to their own department ({exam_department})")
        if not teacher_can_invigilate(teacher, duty):
            problems.append(f"{teacher_id} is not eligible for duty {duty.get('duty_id')}")
        if normalize_role(teacher.get("role")) != normalize_role(duty.get("role_required")):
            problems.append(f"{teacher_id} has the wrong role for duty {duty.get('duty_id')}")
    return problems


def _validate_student_sessions(timetable, student_courses):
    """No student may have two papers in the same date/session."""
    seen = {}
    violations = []
    for row in timetable:
        key = (str(row.get("date")), str(row.get("session")))
        for student_id, courses in (student_courses or {}).items():
            if str(row.get("course_id")) in {str(course) for course in courses}:
                existing = seen.get((student_id, key))
                if existing and existing != str(row.get("course_id")):
                    violations.append(
                        f"{student_id}: {existing} and {row.get('course_id')} "
                        f"on {row.get('date')} {row.get('session')}"
                    )
                else:
                    seen[(student_id, key)] = str(row.get("course_id"))
    return violations


def _teacher_view(teacher_id, faculty_status):
    status = (faculty_status or {}).get(str(teacher_id), {})
    return {
        "teacher_department": _text(status.get("department"), ""),
        "teacher_role": normalize_role(status.get("role") or ""),
        "last_role": status.get("last_role") or "N/A",
        "is_priority": "No" if status.get("has_served_high_role", True) else "Yes",
    }


def generate_multi_department_run(
    filepath,
    year_filter="TY",
    exam_type="Insem",
    start_date=None,
    end_date=None,
    departments=None,
    selected_slots=None,
):
    """Generate one multi-department examination run.

    Returns a JSON-friendly dictionary containing public timetable rows, room
    allocation, all duty seats, assignments and per-department PDF groups.
    """
    if not start_date or not end_date:
        raise ValueError("Both the examination start date and end date are required.")
    start = _as_date(start_date)
    end = _as_date(end_date)
    if end < start:
        raise ValueError("The examination end date must be on or after the start date.")

    selected_departments = _normalize_departments(departments)
    if not selected_departments:
        raise ValueError("Select at least one department.")

    student_courses, _loaded_slots, course_meta, enrolled_counts = load_data(filepath)
    rooms = load_room_data(filepath)
    year_key = normalize_year(year_filter)
    if year_key == "ALL":
        year_key = "ALL"

    courses_by_department = defaultdict(dict)
    for course_id, info in course_meta.items():
        dept = canonical_department(info.get("department"))
        if dept not in selected_departments:
            continue
        course_year = normalize_year(info.get("year"))
        if course_year == "FY":
            # Department examinations are 2nd-4th year papers only.
            continue
        if year_key != "ALL" and course_year != year_key:
            continue
        courses_by_department[dept][course_id] = info

    missing_departments = [
        dept for dept in selected_departments
        if not courses_by_department.get(dept)
    ]
    if missing_departments:
        raise ValueError(
            "No courses were found for: "
            + ", ".join(missing_departments)
            + f" (year {year_filter}). Check the Courses sheet."
        )

    selected_course_ids = {
        course_id
        for courses in courses_by_department.values()
        for course_id in courses
    }
    filtered_students = {}
    for student_id, courses in student_courses.items():
        relevant = [course_id for course_id in courses if course_id in selected_course_ids]
        if relevant:
            filtered_students[student_id] = relevant

    groups_by_department = {}
    for dept in selected_departments:
        groups_by_department[dept] = _group_courses(
            courses_by_department[dept], filtered_students, exam_type, year_key
        )

    max_groups = max((len(groups) for groups in groups_by_department.values()), default=0)
    dates = _calendar_for_run(_loaded_slots, start, end, max_groups)
    sessions_by_department = _build_sessions(selected_departments, groups_by_department, dates)

    room_allocations = []
    timetable = []
    departments_payload = []
    all_duties = []
    duty_id = 1

    for dept in selected_departments:
        blocks = _block_definitions(dept, rooms)
        sessions_payload = []
        dept_timetable = []
        for session in sessions_by_department.get(dept, []):
            usage = _allocate_blocks(session["courses"], blocks)
            unique_students = _unique_students_for_courses(
                [course["course_id"] for course in session["courses"]], filtered_students
            )
            room_ids = [block["room_id"] for block in usage]
            total_capacity = sum(block["capacity"] for block in usage)
            student_count = len(unique_students)
            status = "Allocated" if total_capacity >= student_count else "Partial"
            session_payload = {
                **session,
                "blocks": usage,
                "room_ids": room_ids,
                "total_capacity": total_capacity,
                "student_count": student_count,
                "status": status,
            }
            sessions_payload.append(session_payload)

            room_allocations.append({
                "course_id": ", ".join(
                    str(course.get("course_code") or course.get("course_id"))
                    for course in session["courses"]
                ),
                "slot": session["slot"],
                "date": session["date_label"],
                "session": session["session"],
                "rooms_assigned": ", ".join(room_ids),
                "total_capacity": total_capacity,
                "status": status,
                "department": dept,
                "blocks": ", ".join(block["block"] for block in usage),
            })

            for course in session["courses"]:
                enrolled = enrolled_counts.get(course["course_id"], 0)
                timetable.append({
                    "course_id": course["course_id"],
                    "course_code": course.get("course_code") or course["course_id"],
                    "course_name": course.get("course_name") or "Course",
                    "year": course.get("year") or normalize_year(year_filter),
                    "department": dept,
                    "exam_department": dept,
                    "category": session.get("category") or "",
                    "slot": session["slot"],
                    "date": session["date_label"],
                    "session": session["session"],
                    "declared_students": int(course.get("students_count") or 0),
                    "enrolled_students": int(enrolled),
                    "room_assigned": ", ".join(room_ids),
                    "blocks": room_ids,
                })

        department_duties, squad_count = _build_department_duties(
            dept, sessions_payload, blocks, year_key, duty_id
        )
        duty_id += len(department_duties)
        all_duties.extend(department_duties)
        departments_payload.append({
            "department": dept,
            "title": department_title(dept),
            "short_title": department_short_title(dept),
            "block_count": len(blocks),
            "blocks": [
                {
                    "block": block["block"],
                    "room_id": block["room_id"],
                    "capacity": block["capacity"],
                    "courses": [],
                }
                for block in blocks
            ],
            "squad_count": squad_count,
            "sessions": sessions_payload,
            "timetable": dept_timetable,
        })

    # Fill the per-department timetable and block course labels after allocation.
    department_by_key = {item["department"]: item for item in departments_payload}
    for row in timetable:
        department_by_key[row["department"]]["timetable"].append(row)
    for department in departments_payload:
        block_courses = defaultdict(set)
        for session in department["sessions"]:
            for block in session["blocks"]:
                for course_id in block["courses"]:
                    block_courses[block["block"]].add(course_id)
        for block in department["blocks"]:
            block["courses"] = sorted(block_courses.get(block["block"], set()))

    try:
        teachers = load_teacher_data(filepath)
        teachers, faculty_status, db_duty_counts, fairness_map = _faculty_inputs(
            teachers, start=start, end=end
        )
    except Exception as exc:
        raise RuntimeError(f"Unable to initialise faculty records: {exc}") from exc

    assignments, unassigned, teacher_duty_count = assign_teachers(
        teachers,
        all_duties,
        fairness_map=fairness_map,
        db_duty_counts=db_duty_counts,
        faculty_status=faculty_status,
        MAX_DUTIES=int(os.getenv("MAX_DUTIES", "5") or 5),
    )

    assignment_by_duty = {str(item.get("duty_id")): item for item in assignments}
    duty_rows = []
    for duty in sorted(all_duties, key=lambda item: (
        str(item.get("date")),
        int(item.get("slot") or 0),
        str(item.get("exam_department")),
        str(item.get("role_required")),
    )):
        row = dict(duty)
        assignment = assignment_by_duty.get(str(duty["duty_id"]))
        if assignment:
            row.update({
                "teacher_id": assignment.get("teacher_id"),
                "teacher_name": assignment.get("teacher_name"),
                "pinned": bool(assignment.get("pinned")),
                **_teacher_view(assignment.get("teacher_id"), faculty_status),
            })
        else:
            row.update({
                "teacher_id": None,
                "teacher_name": None,
                "pinned": False,
                "teacher_department": "",
                "teacher_role": "",
                "last_role": "N/A",
                "is_priority": "No",
            })
        duty_rows.append(row)

    teacher_duties = []
    for assignment in assignments:
        row = dict(assignment)
        row["role"] = row.get("role_required")
        row["room_assigned"] = row.get("room") or row.get("room_assigned")
        row.update(_teacher_view(row.get("teacher_id"), faculty_status))
        teacher_duties.append(row)

    violations = _validate_student_sessions(timetable, filtered_students)
    assignment_problems = _validate_assignments(assignments, all_duties, teachers)

    role_counts = defaultdict(int)
    assigned_role_counts = defaultdict(int)
    for duty in all_duties:
        role_counts[duty["role_required"]] += 1
    for assignment in assignments:
        assigned_role_counts[assignment["role_required"]] += 1
    generated_loads = [int(value) for value in (teacher_duty_count or {}).values()]
    preferred = sum(1 for item in assignments if float(item.get("cost") or 9999) <= 1)

    all_student_ids = set()
    for student_id, courses in filtered_students.items():
        if set(courses) & selected_course_ids:
            all_student_ids.add(student_id)

    summary = {
        "total_courses": len(timetable),
        "total_students": len(all_student_ids),
        "slots_used": len({(row["date"], row["session"]) for row in timetable}),
        "conflict_edges": 0,
        "conflicts_found": len(violations),
        "rooms_allocated": len(room_allocations),
        "rooms_failed": sum(1 for row in room_allocations if row["status"] != "Allocated"),
        "duties_required": len(all_duties),
        "duties_assigned": len(assignments),
        "duties_failed": len(unassigned),
        "duties_by_role": dict(role_counts),
        "assigned_by_role": dict(assigned_role_counts),
        "max_generated_load": max(generated_loads) if generated_loads else 0,
        "min_generated_load": min(generated_loads) if generated_loads else 0,
        "pref_satisfaction": round((preferred / len(assignments)) * 100, 1) if assignments else 0,
        "assignment_problems": assignment_problems,
        "student_conflicts": violations,
    }

    return {
        "mode": "multi_department",
        "year": year_key,
        "exam_type": exam_type,
        "start_date": _date_label(start),
        "end_date": _date_label(end),
        "departments_requested": selected_departments,
        "blocks_by_department": {dept: DEPARTMENT_BLOCKS.get(dept, 0) for dept in selected_departments},
        "calendar_dates": [_date_label(value) for value in dates],
        "summary": summary,
        "departments": departments_payload,
        "timetable": timetable,
        "room_allocation": room_allocations,
        "teacher_duties": sorted(teacher_duties, key=lambda item: (
            str(item.get("date")), int(item.get("slot") or 0), str(item.get("role_required"))
        )),
        "duties": duty_rows,
        "unassigned_duties": [
            {
                **duty,
                "teacher_id": None,
                "teacher_name": None,
                "teacher_department": "",
                "teacher_role": "",
            }
            for duty in unassigned
        ],
        "warnings": {
            "unallocated_rooms": [row["course_id"] for row in room_allocations if row["status"] != "Allocated"],
            "unassigned_duties": [duty["duty_id"] for duty in unassigned],
            "assignment_problems": assignment_problems,
            "student_conflicts": violations,
        },
    }
