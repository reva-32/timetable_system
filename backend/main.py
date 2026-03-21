import pandas as pd
import os
import csv
from itertools import combinations
from tabulate import tabulate
import numpy as np
from scipy.optimize import linear_sum_assignment

# ================================================
# CORE ALGORITHM LOGIC
# ================================================

def build_conflict_graph(student_courses):
    """
    Build conflict graph: two courses conflict if ANY student is enrolled in both.
    These conflicting courses CANNOT be in the same slot.
    """
    graph = {}
    for courses in student_courses.values():
        for course in courses:
            if course not in graph:
                graph[course] = set()

    for courses in student_courses.values():
        for c1, c2 in combinations(set(courses), 2):
            graph[c1].add(c2)
            graph[c2].add(c1)

    return graph


def dsatur_coloring(graph):
    """
    DSATUR Algorithm:
    - Saturation degree = number of distinct colors in a node's neighbors
    - Always color the node with the highest saturation (ties broken by degree)
    - Guarantees no two conflicting courses share the same slot (color)
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

        for nb in graph[node]:
            if nb in uncolored:
                nb_colors = {result[n] for n in graph[nb] if n in result}
                saturation[nb] = len(nb_colors)

    return result


# ================================================
# DATA LOADING
# ================================================

def safe_str(val, default="N/A"):
    """Convert value to string, replacing NaN/None with default."""
    if pd.isna(val) if not isinstance(val, str) else False:
        return default
    s = str(val).strip()
    return s if s and s.lower() not in ("nan", "none", "") else default


def load_data(filepath):
    """
    Loads all required sheets from the Excel file.

    Returns:
        student_courses  : dict { student_id -> [course_id, ...] }
        slot_metadata    : dict { 0-based-index -> {display_id, date, session} }
        course_metadata  : dict { course_id -> {course_name, year, students_count, department} }
        enrolled_counts  : dict { course_id -> int (actual enrolled from Student_Courses) }
    """

    # --- Load Sheets ---
    df_students = pd.read_excel(filepath, sheet_name="Student_Courses")
    df_slots    = pd.read_excel(filepath, sheet_name="Slots")
    df_courses  = pd.read_excel(filepath, sheet_name="Courses")

    # -----------------------------------------------
    # 1. Student -> Courses Mapping
    # -----------------------------------------------
    student_courses = {}
    for _, row in df_students.iterrows():
        s_id = safe_str(row["student_id"])
        c_id = safe_str(row["course_id"])
        if s_id != "N/A" and c_id != "N/A":
            student_courses.setdefault(s_id, []).append(c_id)

    # -----------------------------------------------
    # 2. Actual Enrollment Count per Course
    #    (from Student_Courses sheet — this is the TRUE count)
    # -----------------------------------------------
    enrolled_counts = {}
    for courses in student_courses.values():
        for c in courses:
            enrolled_counts[c] = enrolled_counts.get(c, 0) + 1

    # -----------------------------------------------
    # 3. Slot Metadata (0-based index used internally)
    # -----------------------------------------------
    slot_metadata = {}
    for _, row in df_slots.iterrows():
        try:
            slot_id = int(row["slot_id"])
        except (ValueError, TypeError):
            continue

        idx = slot_id - 1  # Convert to 0-based for DSATUR mapping

        # Normalize date: handle both "27-Nov" strings and datetime objects
        raw_date = row["date"]
        if isinstance(raw_date, pd.Timestamp):
            date_str = raw_date.strftime("%d-%b")     # e.g. "27-Nov"
        else:
            date_str = str(raw_date).strip().split(" ")[0]  # clean any time suffix

        slot_metadata[idx] = {
            "display_id" : slot_id,
            "date"       : date_str,
            "session"    : safe_str(row["session"])
        }

    # -----------------------------------------------
    # 4. Course Metadata (declared info from Courses sheet)
    # -----------------------------------------------
    course_metadata = {}
    for _, row in df_courses.iterrows():
        c_id = safe_str(row["course_id"])
        if c_id == "N/A":
            continue

        # students_count: declared count (official class size)
        try:
            declared = int(row["students_count"])
        except (ValueError, TypeError):
            declared = 0

        course_metadata[c_id] = {
            "course_name"    : safe_str(row.get("course_name", ""), default="Unknown"),
            "year"           : safe_str(row.get("year", ""), default="N/A"),
            "students_count" : declared,
            "department"     : safe_str(row.get("department", ""), default="General")
        }

    return student_courses, slot_metadata, course_metadata, enrolled_counts


# ================================================
# VALIDATION HELPERS
# ================================================

def validate_no_conflicts(student_courses, coloring):
    """
    Post-scheduling validation:
    Check that no student has two exams in the SAME slot.
    Prints any violations found.
    """
    violations = []
    for student, courses in student_courses.items():
        slots_taken = {}
        for course in courses:
            slot = coloring.get(course)
            if slot is None:
                continue
            if slot in slots_taken:
                violations.append(
                    f"  ⚠️  {student}: {course} and {slots_taken[slot]} both in slot {slot + 1}"
                )
            else:
                slots_taken[slot] = course

    if violations:
        print("\n🚨 CONFLICT VIOLATIONS FOUND:")
        for v in violations:
            print(v)
    else:
        print("\n✅ Zero student conflicts — every student has at most 1 exam per slot.")
    return len(violations) == 0


def print_slot_summary(final_data, total_slots):
    """Print a per-slot summary table showing which courses are grouped together."""
    slot_groups = {}
    for row in final_data:
        key = (row["slot"], row["date"], row["session"])
        slot_groups.setdefault(key, []).append(row["course_id"])

    summary = []
    for (slot, date, session), courses in sorted(slot_groups.items()):
        summary.append({
            "Slot"    : slot,
            "Date"    : date,
            "Session" : session,
            "Courses" : ", ".join(sorted(courses)),
            "Count"   : len(courses)
        })

    print("\n--- SLOT UTILIZATION SUMMARY ---")
    print(tabulate(summary, headers="keys", tablefmt="fancy_grid"))
    print(f"\n📅 Total Slots Used   : {total_slots}")
    print(f"📚 Total Courses      : {sum(s['Count'] for s in summary)}")


# ================================================
# STAGE 3 — TEACHER ASSIGNMENT (BIPARTITE MATCHING)
# ================================================

def load_teacher_data(filepath):
    """
    Load Teachers and Preferences sheets from Excel.

    Returns:
        teachers : dict { teacher_id -> {name, role, preferred_slots: set} }
    """
    df_teachers    = pd.read_excel(filepath, sheet_name="Teachers")
    df_preferences = pd.read_excel(filepath, sheet_name="Preferences")

    # Build preference map: teacher_id -> set of preferred slot display IDs
    pref_map = {}
    for _, row in df_preferences.iterrows():
        t_id = str(row["teacher_id"]).strip()
        raw  = str(row["preferred_slots"]).strip()
        # Stored as "1,2,3" strings
        try:
            slots = {int(s.strip()) for s in raw.split(",") if s.strip().isdigit()}
        except Exception:
            slots = set()
        pref_map[t_id] = slots

    teachers = {}
    for _, row in df_teachers.iterrows():
        t_id   = str(row["teacher_id"]).strip()
        t_name = str(row["name"]).strip()
        t_role = str(row["role"]).strip()   # "Senior" | "Squad" | "Junior"
        teachers[t_id] = {
            "name"            : t_name,
            "role"            : t_role,
            "preferred_slots" : pref_map.get(t_id, set())
        }

    return teachers


def build_duties(final_data):
    """
    Generate one duty per (slot, course) combination from the exam schedule.
    Each duty requires exactly ONE Junior supervisor inside the room.
    Senior and Squad duties are generated per unique slot.

    Returns:
        duties : list of dicts
            { duty_id, slot, date, session, course_id, role_required }
    """
    duties = []
    duty_id = 1

    # Junior duty — one per course per slot (inside the room)
    for row in final_data:
        duties.append({
            "duty_id"      : duty_id,
            "slot"         : row["slot"],
            "date"         : row["date"],
            "session"      : row["session"],
            "course_id"    : row["course_id"],
            "role_required": "Junior"
        })
        duty_id += 1

    # Senior duty — one per unique slot (overall control)
    seen_slots = set()
    for row in final_data:
        if row["slot"] not in seen_slots:
            duties.append({
                "duty_id"      : duty_id,
                "slot"         : row["slot"],
                "date"         : row["date"],
                "session"      : row["session"],
                "course_id"    : "ALL",
                "role_required": "Senior"
            })
            duty_id += 1
            seen_slots.add(row["slot"])

    # Squad duty — one per unique slot (roving across rooms)
    seen_slots = set()
    for row in final_data:
        if row["slot"] not in seen_slots:
            duties.append({
                "duty_id"      : duty_id,
                "slot"         : row["slot"],
                "date"         : row["date"],
                "session"      : row["session"],
                "course_id"    : "ALL",
                "role_required": "Squad"
            })
            duty_id += 1
            seen_slots.add(row["slot"])

    return duties


def compute_cost(teacher, duty, teacher_duty_count, MAX_DUTIES=5):
    """
    Compute assignment cost for a (teacher, duty) pair.

    Cost logic:
    - Role mismatch          → INF (hard constraint, never assign)
    - Already at max duties  → INF (hard constraint)
    - Preferred slot         → cost = 1  (strongly preferred)
    - Non-preferred slot     → cost = 10 (allowed but penalised)

    Lower cost = better assignment.
    """
    INF = 10_000

    # Hard constraint: role must match
    if teacher["role"] != duty["role_required"]:
        return INF

    # Hard constraint: max 5 duties per teacher
    if teacher_duty_count.get(teacher["id"], 0) >= MAX_DUTIES:
        return INF

    # Soft constraint: prefer slots the teacher listed
    if duty["slot"] in teacher["preferred_slots"]:
        return 1
    else:
        return 10


def assign_teachers(teachers, duties):
    """
    Solve the teacher-duty assignment using Minimum Cost Bipartite Matching.

    Uses scipy.optimize.linear_sum_assignment (Hungarian Algorithm) on the
    cost matrix built from compute_cost().

    Returns:
        assignments : list of dicts
            { duty_id, slot, date, session, course_id,
              role_required, teacher_id, teacher_name, cost }
        unassigned  : list of duty dicts that could not be filled
    """
    INF = 10_000

    # Flatten teacher dict into a list for matrix indexing
    teacher_list = [{"id": tid, **tdata} for tid, tdata in teachers.items()]
    n_teachers   = len(teacher_list)
    n_duties     = len(duties)

    # Pad to square matrix (Hungarian algorithm needs square)
    size = max(n_teachers, n_duties)

    # Track how many duties each teacher has been given (updated iteratively)
    teacher_duty_count = {t["id"]: 0 for t in teacher_list}

    # We solve in ONE pass using the full cost matrix.
    # Build cost matrix: rows = teachers, cols = duties
    cost_matrix = np.full((size, size), INF, dtype=float)

    for i, teacher in enumerate(teacher_list):
        for j, duty in enumerate(duties):
            cost_matrix[i][j] = compute_cost(
                teacher, duty, teacher_duty_count
            )

    # Run Hungarian algorithm — finds optimal min-cost assignment
    row_ind, col_ind = linear_sum_assignment(cost_matrix)

    assignments = []
    unassigned  = []
    assigned_duty_ids = set()

    for r, c in zip(row_ind, col_ind):
        # Skip padding rows/cols
        if r >= n_teachers or c >= n_duties:
            continue

        cost = cost_matrix[r][c]
        if cost >= INF:
            # Could not assign — hard constraint violated
            continue

        teacher = teacher_list[r]
        duty    = duties[c]

        teacher_duty_count[teacher["id"]] = teacher_duty_count.get(teacher["id"], 0) + 1
        assigned_duty_ids.add(duty["duty_id"])

        assignments.append({
            "duty_id"      : duty["duty_id"],
            "slot"         : duty["slot"],
            "date"         : duty["date"],
            "session"      : duty["session"],
            "course_id"    : duty["course_id"],
            "role_required": duty["role_required"],
            "teacher_id"   : teacher["id"],
            "teacher_name" : teacher["name"],
            "cost"         : int(cost)
        })

    # Collect unassigned duties
    for duty in duties:
        if duty["duty_id"] not in assigned_duty_ids:
            unassigned.append(duty)

    return assignments, unassigned, teacher_duty_count


def validate_teacher_assignment(assignments, teacher_duty_count, MAX_DUTIES=5):
    """
    Post-assignment validation:
    1. No teacher assigned two duties in the same slot
    2. No teacher exceeds max duty count
    3. Role integrity — teacher role matches duty role
    """
    violations = []

    # Check: same teacher in same slot twice
    slot_map = {}
    for a in assignments:
        key = (a["teacher_id"], a["slot"])
        if key in slot_map:
            violations.append(
                f"  ⚠️  {a['teacher_name']} has TWO duties in slot {a['slot']}"
            )
        else:
            slot_map[key] = True

    # Check: max duties
    for t_id, count in teacher_duty_count.items():
        if count > MAX_DUTIES:
            violations.append(
                f"  ⚠️  Teacher {t_id} has {count} duties — exceeds max of {MAX_DUTIES}"
            )

    if violations:
        print("\n🚨 TEACHER ASSIGNMENT VIOLATIONS:")
        for v in violations:
            print(v)
    else:
        print("\n✅ Teacher assignment valid — no constraint violations.")

    return len(violations) == 0


def print_teacher_summary(assignments, teacher_duty_count):
    """Print per-teacher workload summary."""
    workload = {}
    for a in assignments:
        t_id   = a["teacher_id"]
        t_name = a["teacher_name"]
        role   = a["role_required"]
        workload.setdefault(t_id, {"name": t_name, "role": role, "duties": 0})
        workload[t_id]["duties"] += 1

    summary = [
        {
            "Teacher ID"  : t_id,
            "Name"        : data["name"],
            "Role"        : data["role"],
            "Duties Assigned": data["duties"]
        }
        for t_id, data in workload.items()
    ]
    summary.sort(key=lambda x: x["Role"])

    print("\n--- TEACHER WORKLOAD SUMMARY ---")
    print(tabulate(summary, headers="keys", tablefmt="fancy_grid"))

    # Preference satisfaction rate
    preferred = sum(1 for a in assignments if a["cost"] == 1)
    total     = len(assignments)
    pct       = (preferred / total * 100) if total else 0
    print(f"\n📊 Preference satisfaction: {preferred}/{total} duties ({pct:.1f}%) assigned to preferred slots")


# ================================================
# MAIN
# ================================================

def main():
    input_file  = os.path.join("uploads", "exam_data.xlsx")
    output_file = os.path.join("uploads", "final_exam_schedule.csv")

    if not os.path.exists(input_file):
        print(f"❌ Error: '{input_file}' not found.")
        print("   → Make sure 'exam_data.xlsx' is inside the 'uploads/' folder.")
        return

    print("📂 Loading data from Excel...")
    student_courses, slot_meta, course_meta, enrolled_counts = load_data(input_file)

    print(f"   ✔ {len(course_meta)} courses loaded")
    print(f"   ✔ {len(student_courses)} students loaded")
    print(f"   ✔ {len(slot_meta)} slots available")

    # Build conflict graph and color it
    print("\n🔗 Building conflict graph...")
    graph = build_conflict_graph(student_courses)
    print(f"   ✔ {len(graph)} nodes, "
          f"{sum(len(v) for v in graph.values()) // 2} conflict edges")

    print("\n🎨 Running DSATUR coloring algorithm...")
    coloring = dsatur_coloring(graph)
    total_slots_used = max(coloring.values()) + 1
    print(f"   ✔ Coloring complete — {total_slots_used} slot(s) needed")

    # Warn if more slots needed than available
    if total_slots_used > len(slot_meta):
        print(f"\n⚠️  WARNING: Algorithm needs {total_slots_used} slots "
              f"but only {len(slot_meta)} are defined in the Excel sheet.")
        print("   → Extra courses will be assigned a fallback label.")

    # -----------------------------------------------
    # Build Final Output Rows
    # -----------------------------------------------
    final_data = []

    for course_id, slot_idx in coloring.items():
        # Slot info (fallback if out of range)
        if slot_idx in slot_meta:
            slot_info = slot_meta[slot_idx]
        else:
            slot_info = {
                "display_id" : slot_idx + 1,
                "date"       : "TBD",
                "session"    : "TBD"
            }

        # Course info
        c_info = course_meta.get(course_id, {
            "course_name"    : "Unknown",
            "year"           : "N/A",
            "students_count" : 0,
            "department"     : "General"
        })

        # Enrolled count from actual Student_Courses data
        enrolled = enrolled_counts.get(course_id, 0)

        final_data.append({
            "course_id"        : course_id,
            "course_name"      : c_info["course_name"],
            "year"             : c_info["year"],
            "department"       : c_info["department"],
            "slot"             : slot_info["display_id"],
            "date"             : slot_info["date"],
            "session"          : slot_info["session"],
            "declared_students": c_info["students_count"],  # from Courses sheet
            "enrolled_students": enrolled,                  # from Student_Courses sheet
        })

    # Sort: primary by slot, secondary by department, then course_id
    final_data.sort(key=lambda x: (x["slot"], x["department"], x["course_id"]))

    # -----------------------------------------------
    # Terminal Output — Full Schedule
    # -----------------------------------------------
    print("\n" + "=" * 70)
    print("            📋  FINAL EXAM TIMETABLE  📋")
    print("=" * 70)
    print(tabulate(final_data, headers="keys", tablefmt="fancy_grid"))

    # -----------------------------------------------
    # Slot Utilization Summary
    # -----------------------------------------------
    print_slot_summary(final_data, total_slots_used)

    # -----------------------------------------------
    # Conflict Validation
    # -----------------------------------------------
    validate_no_conflicts(student_courses, coloring)

    # -----------------------------------------------
    # CSV Export
    # -----------------------------------------------
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    if not final_data:
        print("\n⚠️  No data to export.")
        return

    keys = list(final_data[0].keys())
    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(final_data)

    print(f"\n💾 CSV exported → {output_file}")
    print(
        "\n📌 Columns for your teammates:\n"
        "   course_id, course_name, year, department → course identity\n"
        "   slot, date, session                      → time assignment\n"
        "   declared_students                        → official class size (Courses sheet)\n"
        "   enrolled_students                        → actual enrolled (Student_Courses sheet)\n"
        "   Room allocation and teacher assignment can join on course_id + slot.\n"
    )

    # -----------------------------------------------
    # STAGE 3 — Teacher Assignment
    # -----------------------------------------------
    print("\n" + "=" * 70)
    print("       👩‍🏫  STAGE 3: TEACHER ASSIGNMENT (BIPARTITE MATCHING)  👨‍🏫")
    print("=" * 70)

    print("\n📂 Loading teacher data...")
    teachers = load_teacher_data(input_file)
    print(f"   ✔ {len(teachers)} teachers loaded")

    # Role breakdown
    role_counts = {}
    for t in teachers.values():
        role_counts[t["role"]] = role_counts.get(t["role"], 0) + 1
    for role, count in role_counts.items():
        print(f"   ✔ {count} {role} supervisor(s)")

    print("\n📋 Generating duties from exam schedule...")
    duties = build_duties(final_data)
    duty_role_counts = {}
    for d in duties:
        duty_role_counts[d["role_required"]] = duty_role_counts.get(d["role_required"], 0) + 1
    print(f"   ✔ {len(duties)} total duties generated")
    for role, count in duty_role_counts.items():
        print(f"      → {count} {role} duties")

    print("\n🔗 Running minimum cost bipartite matching...")
    assignments, unassigned, teacher_duty_count = assign_teachers(teachers, duties)
    print(f"   ✔ {len(assignments)} duties successfully assigned")

    if unassigned:
        print(f"\n⚠️  {len(unassigned)} duties could NOT be assigned:")
        for d in unassigned:
            print(f"   → Duty {d['duty_id']} | Slot {d['slot']} | {d['date']} "
                  f"| {d['session']} | Role: {d['role_required']}")
        print("   Tip: Add more teachers of the required role to the Excel sheet.")

    # Print full assignment table
    print("\n" + "=" * 70)
    print("         📋  FINAL TEACHER DUTY SCHEDULE  📋")
    print("=" * 70)

    display_assignments = [
        {
            "Slot"        : a["slot"],
            "Date"        : a["date"],
            "Session"     : a["session"],
            "Course"      : a["course_id"],
            "Role"        : a["role_required"],
            "Teacher ID"  : a["teacher_id"],
            "Teacher Name": a["teacher_name"],
            "Preferred?"  : "✅ Yes" if a["cost"] == 1 else "🔸 No"
        }
        for a in sorted(assignments, key=lambda x: (x["slot"], x["role_required"]))
    ]
    print(tabulate(display_assignments, headers="keys", tablefmt="fancy_grid"))

    # Workload and preference summary
    print_teacher_summary(assignments, teacher_duty_count)

    # Validation
    validate_teacher_assignment(assignments, teacher_duty_count)

    # Export teacher assignment to CSV
    teacher_output_file = os.path.join("uploads", "teacher_duty_schedule.csv")
    if assignments:
        keys = list(assignments[0].keys())
        with open(teacher_output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=keys)
            writer.writeheader()
            writer.writerows(assignments)
        print(f"\n💾 Teacher duty schedule exported → {teacher_output_file}")


if __name__ == "__main__":
    main()