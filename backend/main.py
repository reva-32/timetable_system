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
from itertools import combinations
from tabulate import tabulate
import numpy as np
from scipy.optimize import linear_sum_assignment

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


def adjust_exam_dates(final_data, student_courses, slot_meta):
    """
    Ensure students with multiple exams have them on separate days if possible.
    """
    next_day_slots = {1: [], 2: []}
    for slot, meta in slot_meta.items():
        if meta["session"] == "Morning":
            next_day_slots[1].append(slot)
        elif meta["session"] == "Afternoon":
            next_day_slots[2].append(slot)

    for student, courses in student_courses.items():
        assigned_slots = {course: None for course in courses}
        for row in final_data:
            if row["course_id"] in assigned_slots:
                assigned_slots[row["course_id"]] = row["date"]

        for course1, course2 in combinations(courses, 2):
            if assigned_slots[course1] == assigned_slots[course2]:
                if next_day_slots[1]:
                    new_slot = next_day_slots[1].pop(0)
                    for idx, row in enumerate(final_data):
                        if row["course_id"] == course2:
                            final_data[idx]["slot"] = slot_meta[new_slot]["display_id"]
                            final_data[idx]["date"] = slot_meta[new_slot]["date"]
                            final_data[idx]["session"] = slot_meta[new_slot]["session"]
                            break
                elif next_day_slots[2]:
                    new_slot = next_day_slots[2].pop(0)
                    for idx, row in enumerate(final_data):
                        if row["course_id"] == course2:
                            final_data[idx]["slot"] = slot_meta[new_slot]["display_id"]
                            final_data[idx]["date"] = slot_meta[new_slot]["date"]
                            final_data[idx]["session"] = slot_meta[new_slot]["session"]
                            break


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
            "year"           : safe_id(get_col(row, ["year", "Year", "academic_year"]), default="N/A"),
            "students_count" : declared,
            "department"     : safe_str(get_col(row, ["department", "dept", "branch"]), default="General"),
            "exam_date"      : exam_date_str,
            "session"        : session
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
            t_role = safe_str(get_col(row, ["role", "designation", 2]), default="Junior")
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

    return teachers



def build_duties(final_data, room_assignments=None):
    duties = []
    duty_id = 1
    course_rooms = {}
    if room_assignments:
        for a in room_assignments:
            # use course_id + slot + date to uniquely identify exam instance
            key = (a["course_id"], a["slot"], a.get("date"))
            course_rooms[key] = [r.strip() for r in str(a["rooms_assigned"]).split(",") if r.strip() != "None"]

    for row in final_data:
        key = (row["course_id"], row["slot"], row.get("date"))
        rooms = course_rooms.get(key, ["TBD"])
        for room in rooms:
            duties.append({
                "duty_id": duty_id, "slot": row["slot"], "date": row["date"],
                "session": row["session"], "course_id": row["course_id"],
                "room": room, "role_required": "Junior"
            })
            duty_id += 1

    # Add Senior and Squad duties per unique (date, slot) pair
    seen_slots = set()
    for row in final_data:
        slot_key = (row.get("date"), row.get("slot"))
        if slot_key not in seen_slots:
            for role in ["Senior", "Squad"]:
                duties.append({
                    "duty_id": duty_id, "slot": row["slot"], "date": row["date"],
                    "session": row["session"], "course_id": "ALL",
                    "room": "Control" if role=="Senior" else "Roaming",
                    "role_required": role
                })
                duty_id += 1
            seen_slots.add(slot_key)
    return duties


def compute_cost(teacher, duty, teacher_duty_count, fairness_map=None, db_duty_counts=None, MAX_DUTIES=5):
    INF = 10_000
    if teacher["role"] != duty["role_required"]: return INF
    
    # Past duties from database
    past_duties_dict = db_duty_counts.get(teacher["id"], {}) if db_duty_counts else {}
    past_total = sum(past_duties_dict.values())
    
    # Replicated instance number indicates how many duties are being assigned in this pass.
    # Total duty count = past database duties + instance number
    inst = teacher.get("instance", 0)
    total_duty_count = past_total + inst
    
    if total_duty_count >= MAX_DUTIES:
        return INF
    
    cost = 10
    if duty["role_required"] in ["Senior", "Squad"] and fairness_map:
        if not fairness_map.get(teacher["id"], True): cost = 0.5
        else: cost = 5
    
    if duty["slot"] in teacher.get("preferred_slots", set()) or str(duty["slot"]) in [str(p) for p in teacher.get("preferred_slots", set())]:
        cost = min(cost, 1)
    
    # Penalty of 100 * total_duty_count to load-balance
    cost += total_duty_count * 100
    return cost


def assign_teachers(teachers, duties, fairness_map=None, db_duty_counts=None, MAX_DUTIES=5):
    """
    Improved teacher assignment using role-batching and Hungarian matching with greedy fill fallback.
    """
    INF = 10_000
    assignments = []
    assigned_duty_ids = set()
    teacher_duty_count = {tid: 0 for tid in teachers}
    
    # Exclude admins before building the matrix
    admin_keys = set()
    try:
        from db import get_db
        db_conn = get_db()
        admin_docs = list(db_conn["teachers"].find({"is_admin": True}))
        for doc in admin_docs:
            if doc.get("teacher_id"):
                admin_keys.add(doc["teacher_id"])
            if doc.get("email"):
                admin_keys.add(doc["email"])
    except Exception as e:
        print(f"Error querying admins for exclusion: {e}")

    from collections import defaultdict
    duties_by_role = defaultdict(list)
    for d in duties: duties_by_role[d["role_required"]].append(d)

    for role, role_duties in duties_by_role.items():
        eligible_teachers = [
            t for t in teachers.values() 
            if t["role"] == role and t["id"] not in admin_keys
        ]
        if not eligible_teachers: continue

        # Replicate teachers to allow multiple duties in one Hungarian pass
        expanded_teachers = []
        for t in eligible_teachers:
            for inst in range(MAX_DUTIES):
                expanded_teachers.append({**t, "instance": inst})

        n_t = len(expanded_teachers)
        n_d = len(role_duties)
        size = max(n_t, n_d)
        cost_matrix = np.full((size, size), INF, dtype=float)

        for i, t_inst in enumerate(expanded_teachers):
            for j, duty in enumerate(role_duties):
                cost_matrix[i][j] = compute_cost(
                    t_inst, duty, teacher_duty_count, 
                    fairness_map=fairness_map, 
                    db_duty_counts=db_duty_counts,
                    MAX_DUTIES=MAX_DUTIES
                )

        row_ind, col_ind = linear_sum_assignment(cost_matrix)

        for r, c in zip(row_ind, col_ind):
            if r >= n_t or c >= n_d: continue
            cost = cost_matrix[r][c]
            if cost >= INF: continue

            teacher = expanded_teachers[r]
            duty = role_duties[c]

            # Collision check: same teacher, same slot & date
            if any(a["teacher_id"] == teacher["id"] and a["slot"] == duty["slot"] and a.get("date") == duty.get("date") for a in assignments):
                continue

            teacher_duty_count[teacher["id"]] += 1
            assigned_duty_ids.add(duty["duty_id"])
            assignments.append({
                "duty_id": duty["duty_id"], "slot": duty["slot"], "date": duty["date"],
                "session": duty["session"], "course_id": duty["course_id"],
                "room": duty["room"],
                "role_required": duty["role_required"], "teacher_id": teacher["id"],
                "teacher_name": teacher["name"], "cost": int(cost) if cost >= 1 else cost
            })

        # Fallback pass for any unassigned duties in this role
        unassigned_in_role = [d for d in role_duties if d["duty_id"] not in assigned_duty_ids]
        for duty in unassigned_in_role:
            # Find an eligible teacher free in that (date, slot) with least current duties
            free_candidates = [
                t for t in eligible_teachers
                if not any(a["teacher_id"] == t["id"] and a["slot"] == duty["slot"] and a.get("date") == duty.get("date") for a in assignments)
                and (teacher_duty_count[t["id"]] + (sum(db_duty_counts.get(t["id"], {}).values()) if db_duty_counts else 0)) < MAX_DUTIES
            ]
            if free_candidates:
                # Pick teacher with lowest duties so far
                best_t = min(free_candidates, key=lambda t: teacher_duty_count[t["id"]])
                teacher_duty_count[best_t["id"]] += 1
                assigned_duty_ids.add(duty["duty_id"])
                assignments.append({
                    "duty_id": duty["duty_id"], "slot": duty["slot"], "date": duty["date"],
                    "session": duty["session"], "course_id": duty["course_id"],
                    "room": duty["room"],
                    "role_required": duty["role_required"], "teacher_id": best_t["id"],
                    "teacher_name": best_t["name"], "cost": 10
                })

    unassigned = [d for d in duties if d["duty_id"] not in assigned_duty_ids]
    return assignments, unassigned, teacher_duty_count