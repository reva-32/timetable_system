"""
main.py — Core algorithm logic for the ExamSched pipeline.

This module is imported by app.py (Flask server) and contains the full
production pipeline:
  - build_conflict_graph()  : builds course conflict graph from student enrollments
  - dsatur_coloring()       : DSATUR graph coloring for conflict-free slot assignment
  - load_data()             : reads Student_Courses, Courses, Slots sheets
  - load_room_data()        : reads Rooms sheet
  - allocate_rooms()        : greedy room allocation (Stage 2)
  - load_teacher_data()     : reads Teachers + Preferences sheets
  - build_duties()          : generates supervision duties from timetable
  - assign_teachers()       : Hungarian algorithm teacher assignment (Stage 3)
  - adjust_exam_dates()     : post-process to spread same-day exams across days

NOTE: colouring.py and graph.py contain standalone versions of the graph
building and coloring functions. They are intentionally NOT imported here
because main.py needs slightly extended versions (e.g., all_courses filter
in build_conflict_graph). Those files are kept for isolated testing only.
"""
import pandas as pd
import os
import csv
import sys
from datetime import datetime
from itertools import combinations
from tabulate import tabulate
import numpy as np
from scipy.optimize import linear_sum_assignment
from eligibility import (
    normalize_year,
    normalize_role,
    effective_teaching_years,
    teacher_can_invigilate as _teacher_can_invigilate,
    department_matches,
)

try:
    MAX_DUTIES = max(1, int(os.getenv("MAX_DUTIES", "5")))
except (TypeError, ValueError):
    MAX_DUTIES = 5

# ================================================
# CORE ALGORITHM LOGIC
# ================================================

def build_conflict_graph(student_courses, all_courses=None):
    """
    Build conflict graph: two courses conflict if ANY student is enrolled in both.
    """
    graph = {}
    if all_courses:
        # Initialize graph only for the specified subset of courses
        for c in all_courses:
            graph[c] = set()

        # Only consider student-enrollments that include two or more courses
        # from the provided all_courses set; this prevents cross-group edges
        for courses in student_courses.values():
            # keep only courses that are in our group
            relevant = [c for c in courses if c in graph]
            for c1, c2 in combinations(set(relevant), 2):
                graph[c1].add(c2)
                graph[c2].add(c1)
        return graph

    # Fallback: build graph from all encountered student courses
    for courses in student_courses.values():
        for course in courses:
            if course not in graph:
                graph[course] = set()
    for courses in student_courses.values():
        for c1, c2 in combinations(set(courses), 2):
            if c1 in graph and c2 in graph:
                graph[c1].add(c2)
                graph[c2].add(c1)
    return graph


def dsatur_coloring(graph):
    """
    DSATUR Algorithm for conflict-free slot assignment.
    """
    result = {}
    saturation = {node: 0 for node in graph}
    degree = {node: len(graph[node]) for node in graph}
    uncolored = set(graph.keys())

    while uncolored:
        node = max(uncolored, key=lambda x: (saturation[x], degree[x]))
        
        used_colors = {result[nb] for nb in graph[node] if nb in result}
        color = 0
        while color in used_colors:
            color += 1
        
        result[node] = color
        uncolored.remove(node)

        # Update saturation of neighbors
        for neighbor in graph[node]:
            if neighbor in uncolored:
                neighbor_colors = {result[nb] for nb in graph[neighbor] if nb in result}
                saturation[neighbor] = len(neighbor_colors)
    return result


def normalize_teacher_years(value):
    """Backward-compatible helper used by the API layer."""
    return {normalize_year(v) for v in (value or []) if safe_str(v, default="")}



def validate_schedule(final_data, student_courses):
    course_slot = {str(r.get("course_id")): (str(r.get("date", "")), str(r.get("slot", ""))) for r in final_data}
    violations = []
    for student, courses in student_courses.items():
        seen = {}
        for course in courses:
            key = course_slot.get(str(course))
            if key is None:
                continue
            if key in seen:
                violations.append(f"{student}: {course} and {seen[key]} on {key[0]} Slot {key[1]}")
            else:
                seen[key] = course
    return violations

def adjust_exam_dates(final_data, student_courses, slot_meta):
    """Safely spread same-day exams; every accepted move is revalidated."""
    candidates = []
    for key, meta in slot_meta.items():
        if meta.get("date"):
            candidates.append({
                "slot": meta.get("display_id", key + 1),
                "date": meta["date"],
                "session": meta.get("session", "TBD")
            })
    candidates.sort(key=lambda x: (str(x["date"]), str(x["slot"])))
    course_to_students = {}
    for student, courses in student_courses.items():
        for course in set(courses):
            course_to_students.setdefault(str(course), set()).add(str(student))
    row_by_course = {str(r["course_id"]): r for r in final_data}
    moved = []
    seen_states = set()
    max_moves = max(10, len(final_data) * 3)
    moves_made = 0
    changed = True
    while changed and moves_made < max_moves:
        state = tuple(sorted(
            (str(row.get("course_id")), str(row.get("date")), str(row.get("slot")))
            for row in final_data
        ))
        if state in seen_states:
            break
        seen_states.add(state)
        changed = False
        changed = False
        for student, courses in student_courses.items():
            by_date = {}
            for course in courses:
                row = row_by_course.get(str(course))
                if row:
                    by_date.setdefault(str(row.get("date", "")), []).append(row)
            duplicate = next((rows for rows in by_date.values() if len(rows) > 1), None)
            if not duplicate:
                continue
            row = duplicate[-1]
            original = (str(row.get("date")), str(row.get("slot")))
            ordered = sorted(candidates, key=lambda c: (str(c["date"]) == original[0], str(c["date"]), str(c["slot"])))
            for cand in ordered:
                target = (str(cand["date"]), str(cand["slot"]))
                if target == original:
                    continue
                conflict = False
                for other in final_data:
                    if other is row or (str(other.get("date")), str(other.get("slot"))) != target:
                        continue
                    if course_to_students.get(str(row["course_id"]), set()) & course_to_students.get(str(other["course_id"]), set()):
                        conflict = True
                        break
                if conflict:
                    continue
                old = dict(row)
                row.update({"slot": cand["slot"], "date": cand["date"], "session": cand["session"]})
                if validate_schedule(final_data, student_courses):
                    row.clear()
                    row.update(old)
                    continue
                moved.append({"course_id": row["course_id"], "from": original, "to": target})
                moves_made += 1
                changed = True
                break
            if changed:
                break
    return moved


# ================================================
# DATA LOADING HELPERS
# ================================================

def safe_str(val, default="N/A"):
    if val is None:
        return default
    if pd.isna(val) if not isinstance(val, str) else False:
        return default
    s = str(val).strip()
    return s if s and s.lower() not in ("nan", "none", "") else default


def safe_id(val, default="N/A"):
    """
    Safely convert IDs to clean strings while preserving leading zeros (e.g., '001').
    Strips trailing '.0' from float conversions.
    """
    if val is None:
        return default
    if pd.isna(val) if not isinstance(val, str) else False:
        return default
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return default
    if s.endswith(".0"):
        s = s[:-2]
    return s


def find_sheet(excel_file, possible_names):
    """Find a sheet name ignoring case and underscores/spaces."""
    try:
        xl = pd.ExcelFile(excel_file)
        names = xl.sheet_names
        # Exact match first
        for target in possible_names:
            for actual in names:
                if actual.strip().lower() == target.strip().lower():
                    return actual
        # Normalized match
        for target in possible_names:
            norm_target = target.lower().replace("_", "").replace(" ", "")
            for actual in names:
                norm_actual = actual.lower().replace("_", "").replace(" ", "")
                if norm_actual == norm_target:
                    return actual
    except Exception:
        pass
    return None


def get_col(row, possible_keys, default=None):
    """Retrieve value from a pandas Series/dict using case-insensitive key lookup."""
    keys_map = {str(k).strip().lower().replace("_", "").replace(" ", ""): k for k in row.index}
    for candidate in possible_keys:
        norm = str(candidate).lower().replace("_", "").replace(" ", "")
        if norm in keys_map:
            val = row.get(keys_map[norm])
            if pd.notna(val):
                return val
    return default


def load_data(filepath):
    from db import format_date_to_standard
    
    student_sheet = find_sheet(filepath, ["Student_Courses", "StudentCourses", "Students", "Enrollments"])
    if not student_sheet:
        raise ValueError("Could not find 'Student_Courses' sheet in uploaded Excel workbook.")
    df_students = pd.read_excel(filepath, sheet_name=student_sheet)

    course_sheet = find_sheet(filepath, ["Courses", "Course", "Subjects", "Course_Details"])
    if not course_sheet:
        raise ValueError("Could not find 'Courses' sheet in uploaded Excel workbook.")
    df_courses = pd.read_excel(filepath, sheet_name=course_sheet)

    slots_sheet = find_sheet(filepath, ["Slots", "Slot", "Exam_Slots", "ExamSlots"])
    df_slots = pd.read_excel(filepath, sheet_name=slots_sheet) if slots_sheet else None

    student_courses = {}
    for _, row in df_students.iterrows():
        s_id = safe_id(get_col(row, ["student_id", "studentid", "roll_no", "rollno", "id"]))
        c_id = safe_id(get_col(row, ["course_id", "courseid", "subject_code", "course_code", "course"]))
        if s_id != "N/A" and c_id != "N/A":
            student_courses.setdefault(s_id, []).append(c_id)

    enrolled_counts = {}
    for courses in student_courses.values():
        for c in courses:
            enrolled_counts[c] = enrolled_counts.get(c, 0) + 1

    slot_metadata = {}
    if df_slots is not None and not df_slots.empty:
        for i, row in df_slots.iterrows():
            slot_id_val = get_col(row, ["slot_id", "slotid", "slot"])
            try:
                slot_id = int(slot_id_val) if slot_id_val is not None else (i + 1)
            except (ValueError, TypeError):
                slot_id = i + 1

            idx = slot_id - 1
            raw_date = get_col(row, ["date", "exam_date", "examdate"])
            date_str = format_date_to_standard(raw_date)
            session_str = safe_str(get_col(row, ["session", "slot_time", "time"]), default="Morning")

            slot_metadata[idx] = {
                "display_id": slot_id,
                "date": date_str,
                "session": session_str
            }

    course_metadata = {}
    for _, row in df_courses.iterrows():
        c_id = safe_id(get_col(row, ["course_id", "courseid", "subject_code", "course_code", "course"]))
        if c_id == "N/A":
            continue

        students_count_val = get_col(row, ["students_count", "studentscount", "count", "declared_students", "capacity"])
        try:
            declared = int(students_count_val) if students_count_val is not None else 0
        except (ValueError, TypeError):
            declared = 0

        raw_date = get_col(row, ["exam_date", "examdate", "date", "Date"])
        exam_date_str = format_date_to_standard(raw_date) if raw_date is not None else None
        session = safe_str(get_col(row, ["session", "slot_time", "time"]), default="General")

        course_metadata[c_id] = {
            "course_name"    : safe_str(get_col(row, ["course_name", "coursename", "name", "subject", "title"]), default="Unknown"),
            "course_code"    : safe_str(get_col(row, ["course_code", "coursecode", "subject_code", "subjectcode", "code"]), default=c_id),
            "year"           : safe_id(get_col(row, ["year", "Year", "academic_year"]), default="N/A"),
            "students_count" : declared,
            "department"     : safe_str(get_col(row, ["department", "dept", "branch"]), default="General"),
            "exam_date"      : exam_date_str,
            "session"        : session,
            "exam_category"  : safe_str(get_col(row, ["exam_category", "examcategory", "category", "paper_type", "papertype", "component"]), default=""),
            "exam_order"     : safe_id(get_col(row, ["exam_order", "examorder", "sequence", "order", "sort_order"]), default=""),
        }

    return student_courses, slot_metadata, course_metadata, enrolled_counts


# ================================================
# STAGE 2 — ROOM ALLOCATION
# ================================================

def load_room_data(filepath):
    room_sheet = find_sheet(filepath, ["Rooms", "Room", "Classrooms", "Halls"])
    if not room_sheet:
        return []
    df_rooms = pd.read_excel(filepath, sheet_name=room_sheet)
    rooms = []
    for _, row in df_rooms.iterrows():
        r_id = safe_str(get_col(row, ["room_id", "roomid", "room", "hall_no", "hall", "name"]))
        if r_id == "N/A":
            continue
        try:
            capacity = int(get_col(row, ["capacity", "seats", "size", "room_capacity"], default=0))
            dept = safe_str(get_col(row, ["department", "dept", "branch"]), default="General")
            rooms.append({"room_id": r_id, "capacity": capacity, "department": dept})
        except Exception:
            continue
    rooms.sort(key=lambda x: x["capacity"])
    return rooms


def allocate_rooms(final_data, rooms):
    room_assignments = []
    unallocated = []
    usage_counter = {r["room_id"]: 0 for r in rooms}
    slot_used_rooms = {}

    sorted_data = sorted(final_data, key=lambda x: (x.get("date", ""), x.get("slot", 0), -x.get("enrolled_students", 0)))

    for row in sorted_data:
        students = row.get("enrolled_students", 0)
        slot = row.get("slot")
        date = row.get("date")
        if students == 0:
            students = row.get("declared_students", 0)
        slot_key = (date, slot)
        used_in_slot = slot_used_rooms.setdefault(slot_key, set())
        available = [r for r in rooms if r["room_id"] not in used_in_slot]
        available.sort(key=lambda r: (r["capacity"], usage_counter[r["room_id"]]))

        assigned = []
        total_cap = 0
        temp_students = students

        single_fit = [r for r in available if r["capacity"] >= temp_students]
        if single_fit:
            best = min(single_fit, key=lambda r: (r["capacity"], usage_counter[r["room_id"]]))
            assigned = [best["room_id"]]
            total_cap = best["capacity"]
            used_in_slot.add(best["room_id"])
            usage_counter[best["room_id"]] += 1
        else:
            while temp_students > 0 and available:
                best_room = max(available, key=lambda r: (r["capacity"], -usage_counter[r["room_id"]]))
                assigned.append(best_room["room_id"])
                total_cap += best_room["capacity"]
                temp_students -= best_room["capacity"]
                used_in_slot.add(best_room["room_id"])
                usage_counter[best_room["room_id"]] += 1
                available = [r for r in available if r["room_id"] not in used_in_slot]

        res = {
            **row,
            "rooms_assigned": ", ".join(assigned) if assigned else "None",
            "total_capacity": total_cap,
            "status": "Allocated" if total_cap >= students else "Partial"
        }
        room_assignments.append(res)
        if total_cap < students:
            unallocated.append(res)

    return room_assignments, unallocated


# ================================================
# STAGE 3 — TEACHER ASSIGNMENT
# ================================================

def load_teacher_data(filepath, preferences_filepath=None):
    """
    Loads teacher profiles and slot preferences from:
    1. The 'Teachers' sheet (if present in workbook)
    2. Fallback to MongoDB 'teachers' collection if sheet not present
    3. An optional 'Preferences' sheet or external file
    """
    teachers = {}
    teacher_sheet = find_sheet(filepath, ["Teachers", "Teacher", "Faculty", "Supervisors"])

    if teacher_sheet:
        df_teachers = pd.read_excel(filepath, sheet_name=teacher_sheet)
        
        # Detect if header is missing (e.g. columns are T001, Prof. IT_1, Squad)
        col0_str = str(df_teachers.columns[0]).strip().upper()
        if col0_str.startswith("T") and (len(col0_str) <= 6 or col0_str[1:].isdigit()):
            df_teachers = pd.read_excel(filepath, sheet_name=teacher_sheet, header=None)
            col_names = ["teacher_id", "name", "role", "department", "preferred_slots"]
            df_teachers.columns = col_names[:len(df_teachers.columns)]

        for _, row in df_teachers.iterrows():
            t_id = safe_id(get_col(row, ["teacher_id", "teacherid", "id", "faculty_id", 0]))
            if t_id == "N/A":
                continue

            raw_prefs = get_col(row, ["preferred_slots", "preferredslots", "preferences", "preference", 4], default="")
            preferred_slots = set()
            if pd.notna(raw_prefs):
                try:
                    for p in str(raw_prefs).split(","):
                        p_str = p.strip()
                        if p_str:
                            try:
                                preferred_slots.add(int(p_str))
                            except ValueError:
                                preferred_slots.add(p_str)
                except Exception:
                    pass

            t_name = safe_str(get_col(row, ["name", "teacher_name", "faculty_name", 1]), default=f"Prof. {t_id}")
            t_role = normalize_role(safe_str(get_col(row, ["role", "designation", 2]), default="Junior"))
            t_dept = safe_str(get_col(row, ["department", "dept", "branch", 3]), default="IT")

            teachers[t_id] = {
                "id": t_id,
                "name": t_name,
                "role": t_role,
                "department": t_dept,
                "preferred_slots": preferred_slots
            }

    # Fallback to MongoDB teachers if no teachers loaded from Excel
    if not teachers:
        try:
            from db import get_all_teachers_from_db
            teachers = get_all_teachers_from_db()
            print(f"Loaded {len(teachers)} teachers from MongoDB database.")
        except Exception as db_err:
            print(f"Warning: Could not fetch fallback teachers from DB: {db_err}")

    # Check for 'Preferences' sheet in the main Excel workbook
    pref_sheet = find_sheet(filepath, ["Preferences", "Preference", "Teacher_Preferences"])
    if pref_sheet:
        try:
            df_prefs = pd.read_excel(filepath, sheet_name=pref_sheet)
            for _, row in df_prefs.iterrows():
                t_id = safe_id(get_col(row, ["teacher_id", "teacherid", "id"]))
                if t_id in teachers:
                    slot_val = get_col(row, ["slot_id", "slot", "preferred_slots", "preferences"])
                    if pd.notna(slot_val):
                        for p in str(slot_val).split(","):
                            p_str = p.strip()
                            if p_str:
                                try:
                                    teachers[t_id]["preferred_slots"].add(int(p_str))
                                except ValueError:
                                    teachers[t_id]["preferred_slots"].add(p_str)
        except Exception:
            pass

    # Check for separate uploaded preferences CSV / Excel file
    if preferences_filepath and os.path.exists(preferences_filepath):
        try:
            if preferences_filepath.endswith(".csv"):
                df_pref_ext = pd.read_csv(preferences_filepath)
            else:
                df_pref_ext = pd.read_excel(preferences_filepath)

            for _, row in df_pref_ext.iterrows():
                t_id = safe_id(get_col(row, ["teacher_id", "teacherid", "id"]))
                if t_id in teachers:
                    slot_val = get_col(row, ["preferred_slots", "slot_id", "slot", "preferences"])
                    if pd.notna(slot_val):
                        for p in str(slot_val).split(","):
                            p_str = p.strip()
                            if p_str:
                                try:
                                    teachers[t_id]["preferred_slots"].add(int(p_str))
                                except ValueError:
                                    teachers[t_id]["preferred_slots"].add(p_str)
        except Exception as pref_err:
            print(f"Warning: Could not load external preferences file: {pref_err}")

    # Academic-year eligibility is authoritative in MongoDB. The Excel workbook
    # intentionally does not contain teaching_years/teaching_year.
    try:
        from db import get_db
        mongo_docs = list(get_db()["teachers"].find({}))
        mongo_by_id = {}
        for doc in mongo_docs:
            tid = doc.get("teacher_id") or str(doc.get("email", "")).split("@")[0]
            if tid:
                mongo_by_id[str(tid)] = doc
        for tid, teacher in teachers.items():
            doc = mongo_by_id.get(str(tid), {})
            teacher["department"] = safe_str(doc.get("department"), default=teacher.get("department", "General"))
            teacher["role"] = normalize_role(doc.get("role") or teacher.get("role", "Junior"))
            # Department policy is authoritative; teaching_years only narrows it.
            teacher["teaching_years"] = effective_teaching_years({
                "department": teacher.get("department"),
                "teaching_years": doc.get("teaching_years"),
            })
            try:
                teacher["max_duties"] = int(doc.get("max_duties") or MAX_DUTIES)
            except (TypeError, ValueError):
                teacher["max_duties"] = MAX_DUTIES
            teacher["pinned_duties"] = clean_pinned_duties(doc.get("pinned_duties"))
            teacher["preferred_slots"] = set(teacher.get("preferred_slots", set())) | {str(x) for x in (doc.get("preferred_slots") or [])}
    except Exception as exc:
        raise RuntimeError("Unable to load teacher eligibility from MongoDB") from exc

    return teachers


def clean_pinned_duties(raw):
    """Normalize coordinator-pinned duties read from MongoDB."""
    pinned = []
    if not isinstance(raw, list):
        return pinned
    for item in raw:
        if not isinstance(item, dict) or not item.get("date"):
            continue
        slot = item.get("slot")
        if slot is None:
            continue
        pinned.append({
            "date": str(item.get("date")).strip(),
            "slot": str(slot).strip(),
            "role": normalize_role(item.get("role") or "Junior"),
            "course_id": str(item.get("course_id") or "ALL"),
        })
    return pinned



def _date_sort_key(value):
    try:
        parsed = pd.to_datetime(str(value), dayfirst=True, errors="coerce")
        if not pd.isna(parsed):
            return parsed.to_pydatetime()
    except Exception:
        pass
    return datetime.max


def _as_non_negative_int(value, default=0):
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return default


def normalize_duty_config(config=None):
    """Normalize the per-session duty demand configuration.

    ``js_per_session`` is a minimum number of Junior Supervisor duties per
    session (room duties are always kept; extra roaming duties are added).
    ``squad_per_session`` is the exact Internal Squad demand per session.
    ``last_day_squad`` optionally overrides the Squad demand on the final day.
    ``overrides`` accepts ``{"date|session": {"js": 3, "squad": 1}}`` or
    ``{"date": {"js": 3, "squad": 1}}``.
    """
    config = config or {}
    if not isinstance(config, dict):
        config = {}

    def read_int(container, names, default):
        for name in names:
            if name in container and container.get(name) is not None:
                return _as_non_negative_int(container.get(name), default)
        return default

    overrides = {}
    raw_overrides = config.get("overrides") or {}
    if isinstance(raw_overrides, str):
        try:
            import json
            raw_overrides = json.loads(raw_overrides)
        except (TypeError, ValueError):
            raw_overrides = {}
    if isinstance(raw_overrides, dict):
        for key, value in raw_overrides.items():
            if not isinstance(value, dict):
                continue
            overrides[str(key).strip()] = {
                "js": read_int(value, ["js", "js_per_session", "junior", "junior_supervisors"], 0),
                "squad": read_int(value, ["squad", "squad_per_session", "internal_squad"], 0),
            }

    return {
        "js_per_session": read_int(config, ["js_per_session", "js", "junior_per_session"], 0),
        "squad_per_session": read_int(config, ["squad_per_session", "squad"], 9),
        "last_day_squad": read_int(config, ["last_day_squad", "last_day_squad_per_session"], 2),
        "overrides": overrides,
    }


def build_duties(final_data, room_assignments=None, duty_config=None):
    """Generate all invigilation duties, including configurable JS/Squad demand.

    Junior (JS) duties are one per allocated room.  A per-session minimum can
    add roaming JS seats.  Senior demand is one per session; Internal Squad
    demand is configured per session (with a separate last-day value).
    """
    config = normalize_duty_config(duty_config)
    duties = []
    duty_id = 1
    course_rooms = {}
    if room_assignments:
        for a in room_assignments:
            # use course_id + slot + date to uniquely identify exam instance
            key = (a["course_id"], a["slot"], a.get("date"))
            course_rooms[key] = [r.strip() for r in str(a["rooms_assigned"]).split(",") if r.strip() != "None"]

    session_meta = {}
    for row in final_data:
        date_value = row.get("date")
        session_value = row.get("session") or "TBD"
        session_key = (date_value, session_value)
        meta = session_meta.setdefault(session_key, {
            "slot": row.get("slot"),
            "years": set(),
            "departments": set(),
            "junior_count": 0,
        })
        year = normalize_year(row.get("year", "ALL"))
        if year in {"FY", "SY", "TY", "BTECH"}:
            meta["years"].add(year)
        department = safe_str(row.get("department"), default="General")
        if department.upper() not in {"GENERAL", "COMMON", "SHARED", ""}:
            meta["departments"].add(department)

        key = (row["course_id"], row["slot"], date_value)
        rooms = course_rooms.get(key, ["TBD"])
        for room in rooms:
            duties.append({
                "duty_id": duty_id,
                "slot": row["slot"],
                "date": date_value,
                "session": session_value,
                "course_id": row["course_id"],
                "year": year,
                "department": department,
                "room": room,
                "role_required": "Junior",
                "allowed_years": [year] if year != "ALL" else [],
            })
            meta["junior_count"] += 1
            duty_id += 1

    if not session_meta:
        return duties

    dated_sessions = sorted(session_meta.keys(), key=lambda key: (_date_sort_key(key[0]), str(key[0]), str(key[1])))
    sorted_dates = sorted({key[0] for key in dated_sessions}, key=_date_sort_key)
    last_date = sorted_dates[-1] if sorted_dates else None

    for (date_value, session_value) in dated_sessions:
        meta = session_meta[(date_value, session_value)]
        scope_years = sorted(meta["years"])
        if not scope_years:
            scope_years = ["FY", "SY", "TY", "BTECH"]
        override = config["overrides"].get(f"{date_value}|{session_value}")
        if override is None:
            override = config["overrides"].get(str(date_value))
        is_last_day = str(date_value) == str(last_date)

        if override is not None:
            js_target = override.get("js", 0)
            squad_target = override.get("squad", 0)
        else:
            js_target = config["js_per_session"]
            squad_target = config["last_day_squad"] if is_last_day else config["squad_per_session"]

        extra_js = max(0, js_target - meta["junior_count"])
        for _ in range(extra_js):
            duties.append({
                "duty_id": duty_id,
                "slot": meta["slot"],
                "date": date_value,
                "session": session_value,
                "course_id": "ALL",
                "year": "ALL",
                "allowed_years": scope_years,
                "department": "General",
                "room": "Roaming",
                "role_required": "Junior",
            })
            duty_id += 1

        duties.append({
            "duty_id": duty_id,
            "slot": meta["slot"],
            "date": date_value,
            "session": session_value,
            "course_id": "ALL",
            "year": "ALL",
            "allowed_years": scope_years,
            "department": "General",
            "room": "Control",
            "role_required": "Senior",
        })
        duty_id += 1

        for _ in range(squad_target):
            duties.append({
                "duty_id": duty_id,
                "slot": meta["slot"],
                "date": date_value,
                "session": session_value,
                "course_id": "ALL",
                "year": "ALL",
                "allowed_years": scope_years,
                "department": "General",
                "room": "Roaming",
                "role_required": "Squad",
            })
            duty_id += 1

    return duties


def teacher_can_invigilate(teacher, duty):
    """Thin wrapper so callers importing from main keep working."""
    return _teacher_can_invigilate(teacher, duty)


def compute_cost(teacher, duty, teacher_duty_count, fairness_map=None, db_duty_counts=None, MAX_DUTIES=5):
    INF = 10_000
    if teacher["role"] != duty["role_required"] or not teacher_can_invigilate(teacher, duty):
        return INF

    # Past duties from database (used to spread load across examination
    # cycles).  The per-teacher cap applies to duties assigned in this run,
    # not to the lifetime count, otherwise a teacher would become permanently
    # ineligible after a few cycles.
    past_duties_dict = db_duty_counts.get(teacher["id"], {}) if db_duty_counts else {}
    past_total = sum(past_duties_dict.values())

    inst = teacher.get("instance", 0)
    total_duty_count = past_total + inst
    max_duties = teacher.get("max_duties", MAX_DUTIES)

    if inst >= max_duties:
        return INF

    cost = 10
    if duty["role_required"] in ["Senior", "Squad"] and fairness_map:
        if not fairness_map.get(teacher["id"], True):
            cost = 0.5
        else:
            cost = 5

    preferred = {str(p) for p in teacher.get("preferred_slots", set())}
    if str(duty["slot"]) in preferred:
        cost = min(cost, 1)

    # Penalty of 100 * total_duty_count to load-balance
    cost += total_duty_count * 100
    return cost


def _duty_key(date_value, slot_value):
    return (str(date_value).strip(), str(slot_value).strip())


def _history_dates_by_teacher(faculty_status):
    """Dates on which each teacher already has a duty (prevents two per day)."""
    busy = {}
    for tid, doc in (faculty_status or {}).items():
        dates = set()
        for entry in (doc.get("history") or []):
            if entry.get("exam_date"):
                dates.add(str(entry.get("exam_date")))
        busy[str(tid)] = dates
    return busy


def assign_teachers(teachers, duties, fairness_map=None, db_duty_counts=None,
                    MAX_DUTIES=5, faculty_status=None, admin_keys=None):
    """Assign duties with one-duty-per-day and per-teacher caps.

    The old implementation built a square Hungarian matrix with several
    "instances" per teacher and then dropped assignments that clashed on the
    same day.  This version assigns each date independently, so a teacher can
    never receive two duties on one date and load is balanced against a target
    of ``ceil(total_duties / eligible_teachers)``.
    """
    INF = 10_000
    assignments = []
    assigned_duty_ids = set()
    teacher_duty_count = {}

    if admin_keys is None:
        admin_keys = set()
        try:
            from db import get_db
            db_conn = get_db()
            for doc in db_conn["teachers"].find({"is_admin": True}):
                if doc.get("teacher_id"):
                    admin_keys.add(str(doc["teacher_id"]))
                if doc.get("email"):
                    admin_keys.add(str(doc["email"]))
        except Exception as exc:
            print(f"Error querying admins for exclusion: {exc}")
    admin_keys = {str(key) for key in admin_keys}

    busy_dates = _history_dates_by_teacher(faculty_status)
    base_counts = {}
    for teacher in teachers.values():
        tid = str(teacher["id"])
        counts = (db_duty_counts or {}).get(tid, {})
        if isinstance(counts, dict):
            base_counts[tid] = sum(_as_non_negative_int(value, 0) for value in counts.values())
        else:
            base_counts[tid] = _as_non_negative_int(counts, 0)
        teacher_duty_count[tid] = 0

    def append_assignment(duty, teacher, cost=10, pinned=False):
        tid = str(teacher["id"])
        assignments.append({
            "duty_id": duty["duty_id"], "slot": duty["slot"], "date": duty["date"],
            "session": duty["session"], "course_id": duty["course_id"],
            "room": duty["room"],
            "role_required": duty["role_required"], "teacher_id": tid,
            "teacher_name": teacher["name"], "year": duty.get("year", "ALL"),
            "department": duty.get("department", "General"),
            "allowed_years": duty.get("allowed_years", []),
            "pinned": pinned,
            "cost": int(cost) if cost >= 1 else cost,
        })
        assigned_duty_ids.add(duty["duty_id"])
        teacher_duty_count[tid] += 1
        busy_dates.setdefault(tid, set()).add(str(duty.get("date")))

    # ---- 1. Pinned duties are assigned first and may exceed the normal cap.
    pinned_by_duty = {}
    for teacher in teachers.values():
        tid = str(teacher["id"])
        for pin in teacher.get("pinned_duties") or []:
            key = (_duty_key(pin.get("date"), pin.get("slot")), normalize_role(pin.get("role") or "Junior"))
            pinned_by_duty.setdefault(key, []).append(tid)

    for duty in sorted(duties, key=lambda item: (_date_sort_key(item.get("date")), str(item.get("slot")))):
        key = (_duty_key(duty.get("date"), duty.get("slot")), normalize_role(duty.get("role_required")))
        candidates = pinned_by_duty.get(key) or []
        for tid in candidates:
            teacher = teachers.get(tid)
            if not teacher:
                continue
            if tid in admin_keys:
                continue
            if normalize_role(teacher.get("role")) != normalize_role(duty.get("role_required")):
                continue
            if not teacher_can_invigilate(teacher, duty):
                # A pin cannot override the year/department rule.
                continue
            if duty["duty_id"] in assigned_duty_ids:
                continue
            if str(duty.get("date")) in busy_dates.get(tid, set()):
                # A pinned duty cannot break the one-duty-per-day rule.  The
                # coordinator must fix the pin; do not silently drop another
                # duty.
                continue
            append_assignment(duty, teacher, cost=0, pinned=True)
            break

    # ---- 2. Assign the remaining duties, one date at a time.
    remaining = [duty for duty in duties if duty["duty_id"] not in assigned_duty_ids]
    duties_by_role = {}
    for duty in remaining:
        duties_by_role.setdefault(duty["role_required"], []).append(duty)

    for role, role_duties in duties_by_role.items():
        eligible_teachers = [
            teacher for teacher in teachers.values()
            if normalize_role(teacher.get("role")) == role and str(teacher["id"]) not in admin_keys
        ]
        if not eligible_teachers:
            continue

        # Target total load for this role across the whole roster + history.
        history_total = sum(base_counts.get(str(t["id"]), 0) for t in eligible_teachers)
        target = int(np.ceil((history_total + len(role_duties)) / max(1, len(eligible_teachers))))

        duties_by_date = {}
        for duty in role_duties:
            duties_by_date.setdefault(str(duty.get("date")), []).append(duty)

        for date_value in sorted(duties_by_date, key=_date_sort_key):
            date_duties = sorted(duties_by_date[date_value], key=lambda item: str(item.get("slot")))
            available = []
            for teacher in eligible_teachers:
                tid = str(teacher["id"])
                if date_value in busy_dates.get(tid, set()):
                    continue
                if teacher_duty_count.get(tid, 0) >= teacher.get("max_duties", MAX_DUTIES):
                    continue
                # Drop teachers who cannot invigilate any duty on this date.
                if not any(teacher_can_invigilate(teacher, duty) for duty in date_duties):
                    continue
                available.append(teacher)

            if not available:
                continue

            n_teachers = len(available)
            n_duties = len(date_duties)
            cost_matrix = np.full((n_teachers, n_duties), INF, dtype=float)
            for row_idx, teacher in enumerate(available):
                tid = str(teacher["id"])
                current_load = base_counts.get(tid, 0) + teacher_duty_count.get(tid, 0)
                for col_idx, duty in enumerate(date_duties):
                    if normalize_role(teacher.get("role")) != duty["role_required"]:
                        continue
                    if not teacher_can_invigilate(teacher, duty):
                        continue
                    cost = 10
                    if duty["role_required"] in {"Senior", "Squad"} and fairness_map:
                        cost = 0.5 if not fairness_map.get(tid, True) else 5
                    preferred = {str(p) for p in teacher.get("preferred_slots", set())}
                    if str(duty["slot"]) in preferred:
                        cost = min(cost, 1)
                    cost += current_load * 100
                    if current_load >= target:
                        cost += (current_load - target + 1) * 250
                    cost_matrix[row_idx][col_idx] = cost

            row_ind, col_ind = linear_sum_assignment(cost_matrix)
            for row_idx, col_idx in zip(row_ind, col_ind):
                if cost_matrix[row_idx][col_idx] >= INF:
                    continue
                teacher = available[row_idx]
                duty = date_duties[col_idx]
                if duty["duty_id"] in assigned_duty_ids:
                    continue
                append_assignment(duty, teacher, cost=cost_matrix[row_idx][col_idx])
            # Safety: a teacher can only appear once per date.
            seen_teacher_ids = {str(a["teacher_id"]) for a in assignments if str(a.get("date")) == date_value}
            for duty in date_duties:
                if duty["duty_id"] in assigned_duty_ids:
                    continue
                fallback = [
                    teacher for teacher in eligible_teachers
                    if str(teacher["id"]) not in seen_teacher_ids
                    and str(teacher["id"]) not in admin_keys
                    and date_value not in busy_dates.get(str(teacher["id"]), set())
                    and teacher_can_invigilate(teacher, duty)
                    and teacher_duty_count.get(str(teacher["id"]), 0)
                        < teacher.get("max_duties", MAX_DUTIES)
                ]
                if fallback:
                    best = min(
                        fallback,
                        key=lambda t: (base_counts.get(str(t["id"]), 0) + teacher_duty_count.get(str(t["id"]), 0)),
                    )
                    append_assignment(duty, best, cost=10)
                    seen_teacher_ids.add(str(best["id"]))

    unassigned = [duty for duty in duties if duty["duty_id"] not in assigned_duty_ids]
    return assignments, unassigned, teacher_duty_count
