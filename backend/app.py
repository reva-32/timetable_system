from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import io
from datetime import date, timedelta

# Import all functions from main.py
from main import (
    load_data,
    build_conflict_graph,
    dsatur_coloring,
    allocate_rooms,
    load_room_data,
    build_duties,
    assign_teachers,
    load_teacher_data
)
# Import MongoDB functions
from db import (
    init_faculty,
    get_all_faculty_status,
    update_faculty_duty,
    check_reset_fairness
)
from flask import send_file
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(
    __name__,
    static_folder=os.path.join("..", "frontend"),
)
CORS(app)


# ================================================
# ROUTES
# ================================================

@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")

@app.route("/index.css")
def css():
    return send_from_directory(app.static_folder, "index.css")


@app.route("/login", methods=["POST"])
def login():
    data = request.get_json()

    email    = data.get("email", "").strip()    if data else ""
    password = data.get("password", "").strip() if data else ""

    if not email:
        return jsonify({"error": "Email is required"}), 400
    if not password:
        return jsonify({"error": "Password is required"}), 400

    try:
        from db import get_db, check_password_hash
        database = get_db()

        user = database["teachers"].find_one({
            "$or": [
                {"email": email},
                {"teacher_id": email}
            ]
        })

        if user is None:
            return jsonify({"error": "Invalid credentials"}), 401

        # --- @pict.edu domain check ---
        if not email.lower().endswith("@pict.edu"):
            return jsonify({"error": "Only @pict.edu email addresses are allowed"}), 403

        # --- Password check ---
        stored_hash = user.get("password_hash")
        teacher_id  = user.get("teacher_id", user.get("email", "").split("@")[0])

        if stored_hash:
            # Normal path: validate against stored hash
            if not check_password_hash(stored_hash, password):
                return jsonify({"error": "Invalid credentials"}), 401
        else:
            # Legacy path: no password set yet → default password = teacher_id (e.g. "T1")
            if password != teacher_id:
                return jsonify({"error": "Invalid credentials"}), 401

        user_email = user.get("email", f"{teacher_id}@pict.edu")
        return jsonify({
            "name"      : user.get("name"),
            "role"      : user.get("role"),
            "is_admin"  : user.get("is_admin", False),
            "history"   : user.get("history", []),
            "identifier": user_email
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
        

@app.route("/change_password", methods=["POST"])
def change_password():
    data             = request.get_json()
    identifier       = (data.get("identifier", "") or "").strip()
    current_password = (data.get("current_password", "") or "").strip()
    new_password     = (data.get("new_password", "") or "").strip()

    if not identifier or not current_password or not new_password:
        return jsonify({"error": "All fields are required"}), 400
    if len(new_password) < 6:
        return jsonify({"error": "New password must be at least 6 characters"}), 400

    try:
        from db import get_db, check_password_hash, update_password
        database = get_db()

        user = database["teachers"].find_one({
            "$or": [{"email": identifier}, {"teacher_id": identifier}]
        })
        if user is None:
            return jsonify({"error": "User not found"}), 404

        stored_hash = user.get("password_hash")
        teacher_id  = user.get("teacher_id", user.get("email", "").split("@")[0])

        # Validate current password (same logic as login)
        if stored_hash:
            if not check_password_hash(stored_hash, current_password):
                return jsonify({"error": "Current password is incorrect"}), 401
        else:
            # Legacy default = teacher_id
            if current_password != teacher_id:
                return jsonify({"error": "Current password is incorrect"}), 401

        ok = update_password(identifier, new_password)
        if ok:
            return jsonify({"success": True})
        else:
            return jsonify({"error": "Failed to update password"}), 500

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/generate", methods=["POST"])
def generate():
    """
    Main pipeline: Filter -> DSATUR -> Rooms -> Teachers
    """
    # -----------------------------------------------
    # 0. Get Config from request
    # -----------------------------------------------
    year_filter    = request.form.get("year", "ALL")
    branch_filter  = request.form.get("branch", "ALL")
    exam_type      = request.form.get("exam_type", "Endsem")
    selected_slots_raw = request.form.get("selected_slots", "").strip()
    frontend_slots     = [s.strip() for s in selected_slots_raw.split(",") if s.strip()] if selected_slots_raw else []
    # NOTE: Date selection removed from UI. Exam dates are read from the Courses sheet

    # Default slot labels per exam type (used when user picks no slots in UI)
    EXAM_TYPE_DEFAULT_SLOTS = {
        "Insem"     : ["10:00 – 11:00", "14:00 – 15:00", "16:00 – 17:00"],
        "Endsem"    : ["10:00 – 12:30", "14:00 – 16:30"],
        "Practical" : ["12:00 – 14:00", "14:30 – 16:30", "17:00 – 19:00"],
    }
    if not frontend_slots:
        frontend_slots = EXAM_TYPE_DEFAULT_SLOTS.get(exam_type, ["10:00 – 12:30", "14:00 – 16:30"])

    # -----------------------------------------------
    # 1. Receive and validate uploaded file
    # -----------------------------------------------
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    file_content = file.read()
    temp_path = "temp_data.xlsx"
    with open(temp_path, "wb") as f:
        f.write(file_content)

    # -----------------------------------------------
    # 1b. Validate required Excel sheets are present
    # -----------------------------------------------
    REQUIRED_SHEETS = ["Student_Courses", "Courses", "Rooms", "Teachers", "Preferences"]
    try:
        import openpyxl as _openpyxl
        _wb = _openpyxl.load_workbook(temp_path, read_only=True)
        missing = [s for s in REQUIRED_SHEETS if s not in _wb.sheetnames]
        _wb.close()
        if missing:
            return jsonify({
                "error": (
                    f"Missing required sheet(s) in uploaded Excel: {', '.join(missing)}. "
                    f"Expected sheets: {', '.join(REQUIRED_SHEETS)}."
                )
            }), 400
    except Exception as val_err:
        return jsonify({"error": f"Could not open Excel file: {str(val_err)}"}), 400

    # -----------------------------------------------
    # 2. Load and Filter Data
    # -----------------------------------------------
    try:
        student_courses, loaded_slots, course_meta, enrolled_counts = load_data(temp_path)
    except Exception as e:
        return jsonify({"error": f"Error loading data: {str(e)}"}), 500

    # 2a. Filter data based on year and branch
    if year_filter != "ALL" or branch_filter != "ALL":
        year_map = {
            "FY"     : ["1", "FY", "FIRST"],
            "SY"     : ["2", "SY", "SECOND"],
            "TY"     : ["3", "TY", "THIRD"],
            "B.TECH" : ["4", "BT", "BTECH", "FINAL", "4TH"]
        }
        
        # DEBUG PRINTS
        if course_meta:
            first_c = next(iter(course_meta.values()))
            print(f"DEBUG: Excel Year is {first_c['year']} (type: {type(first_c['year'])}) and Frontend Year is {year_filter} (type: {type(year_filter)})")


        def is_match(val, target, category):
            v = str(val).strip().upper()
            t = str(target).strip().upper()
            if t == "ALL": return True

            if category == "year":
                allowed_vals = year_map.get(t, [t])
                return v in allowed_vals

            if category == "branch":
                branch_keywords = {
                    "CE"   : ["COMPUTER", "CE", "COMP"],
                    "ENTC" : ["ENTC", "E&TC", "ELECTRONICS"],
                    "AIDS" : ["AIDS", "AI&DS", "AI", "DS"],
                    "IT"   : ["IT", "INFORMATION TECHNOLOGY", "INFO TECH"]
                }
                allowed_branches = branch_keywords.get(t, [t])
                return any(m in v for m in allowed_branches) or t in v

            return t in v

        filtered_courses = {
            cid: info for cid, info in course_meta.items()
            if is_match(info["year"], year_filter, "year") and
               is_match(info["department"], branch_filter, "branch")
        }

        if not filtered_courses:
            return jsonify({
                "error": (
                    f"Filter Mismatch: No courses found for Branch '{branch_filter}' "
                    f"and Year '{year_filter}'. "
                    "Check if your Excel 'year' column uses 1, 2, 3, 4 instead of FY, SY, TY."
                )
            }), 400

        filtered_student_courses = {}
        for sid, courses in student_courses.items():
            valid_courses = [c for c in courses if c in filtered_courses]
            if valid_courses:
                filtered_student_courses[sid] = valid_courses

        course_meta     = filtered_courses
        student_courses = filtered_student_courses
        
        enrolled_counts = {c: count for c, count in enrolled_counts.items() if c in course_meta}
        print(f"Filter Applied: Year={year_filter}, Branch={branch_filter}. Remaining Courses: {len(course_meta)}")

    # 2b. Build conflict graphs and assign slots
    use_dynamic_dates = any(not info.get("exam_date") for info in course_meta.values())

    slot_meta = {}
    final_data = []

    if use_dynamic_dates:
        print("Using dynamic slot scheduling from Slots sheet (or generated).")
        # Build global conflict graph for all courses
        subgraph = build_conflict_graph(student_courses, all_courses=list(course_meta.keys()))
        coloring = dsatur_coloring(subgraph)

        orig_colors = sorted(set(coloring.values())) if coloring else []
        remap = {orig: idx for idx, orig in enumerate(orig_colors)}

        has_loaded_slots = bool(loaded_slots)
        start_date = date.today()
        L = len(frontend_slots)
        if L == 0:
            frontend_slots = ["Morning", "Afternoon"]
            L = 2

        color_to_global = {}
        for orig_color in orig_colors:
            idx = remap[orig_color]
            color_to_global[orig_color] = idx
            
            if has_loaded_slots and idx in loaded_slots:
                slot_info = loaded_slots[idx]
                slot_meta[idx] = {
                    "display_id": slot_info.get("display_id", idx + 1),
                    "date": slot_info.get("date"),
                    "session": slot_info.get("session")
                }
            else:
                day_offset = idx // L
                slot_in_day = idx % L
                slot_date = start_date + timedelta(days=day_offset)
                date_str = slot_date.strftime("%d-%b-%Y")
                slot_meta[idx] = {
                    "display_id": slot_in_day + 1,
                    "date": date_str,
                    "session": frontend_slots[slot_in_day]
                }

        for course_id in course_meta.keys():
            c_info = course_meta[course_id]
            enrolled = enrolled_counts.get(course_id, 0)
            orig_color = coloring.get(course_id, 0)
            new_color = remap.get(orig_color, 0)
            global_slot = color_to_global.get(orig_color, None)
            
            slot_info = slot_meta.get(new_color, {})
            display_slot = slot_info.get("display_id", 1)
            exam_date = slot_info.get("date", "TBD")
            session = slot_info.get("session", "TBD")

            final_data.append({
                "course_id": course_id,
                "course_name": c_info.get("course_name", "Unknown"),
                "year": c_info.get("year", "N/A"),
                "department": c_info.get("department", "General"),
                "slot": display_slot,
                "date": exam_date,
                "session": session,
                "declared_students": c_info.get("students_count", 0),
                "enrolled_students": enrolled,
                "_global_slot": global_slot
            })
        print(f"DEBUG: Dynamic final_data length: {len(final_data)}")
    else:
        # Pre-assigned dates mode: group by exam date from Courses sheet
        print("Using pre-assigned exam dates from Courses sheet.")
        from collections import defaultdict
        date_groups = defaultdict(list)
        for cid, info in course_meta.items():
            date_groups[info["exam_date"]].append(cid)

        global_slot_counter = 0
        for exam_date, courses_on_date in sorted(date_groups.items()):
            subgraph = build_conflict_graph(student_courses, all_courses=list(courses_on_date))
            coloring = dsatur_coloring(subgraph)

            orig_colors = sorted(set(coloring.values())) if coloring else []
            remap = {orig: idx for idx, orig in enumerate(orig_colors)}

            color_to_global = {}
            L = len(frontend_slots)
            for orig_color in orig_colors:
                new_color = remap[orig_color]
                color_to_global[orig_color] = global_slot_counter
                session_val = frontend_slots[new_color % L] if L > 0 else None
                slot_meta[global_slot_counter] = {
                    "display_id": new_color + 1,
                    "date": exam_date,
                    "session": session_val
                }
                global_slot_counter += 1

            for course_id in courses_on_date:
                c_info = course_meta.get(course_id, {})
                enrolled = enrolled_counts.get(course_id, 0)
                orig_color = coloring.get(course_id, 0)
                new_color = remap.get(orig_color, 0)
                global_slot = color_to_global.get(orig_color, None)
                display_slot = new_color + 1
                session = None
                if global_slot is not None and global_slot in slot_meta:
                    session = slot_meta[global_slot]["session"]
                if not session:
                    session = c_info.get("session") or (frontend_slots[0] if L>0 else "TBD")

                final_data.append({
                    "course_id": course_id,
                    "course_name": c_info.get("course_name", "Unknown"),
                    "year": c_info.get("year", "N/A"),
                    "department": c_info.get("department", "General"),
                    "slot": display_slot,
                    "date": exam_date,
                    "session": session,
                    "declared_students": c_info.get("students_count", 0),
                    "enrolled_students": enrolled,
                    "_global_slot": global_slot
                })
        print(f"DEBUG: Pre-assigned final_data length: {len(final_data)}")

    # -----------------------------------------------
    # 2b. Adjust exam dates (move same-day exams to different days where possible)
    # -----------------------------------------------
    try:
        from main import adjust_exam_dates
        adjust_exam_dates(final_data, student_courses, slot_meta)
    except Exception as adj_err:
        print(f"adjust_exam_dates warning (non-fatal): {adj_err}")

    # -----------------------------------------------
    # 3. Run Stage 2 — Room Allocation
    # -----------------------------------------------
    # Build global conflict graph (used in summary)
    try:
        graph = build_conflict_graph(student_courses, all_courses=list(course_meta.keys()))
    except Exception:
        graph = {}
    try:
        rooms = load_room_data(temp_path)
        if branch_filter != "ALL":
            # Include rooms matching the branch OR rooms marked as 'General' (shared across branches)
            rooms = [
                r for r in rooms
                if is_match(r.get("department", "General"), branch_filter, "branch")
                or str(r.get("department", "")).strip().upper() in ("GENERAL", "COMMON", "SHARED", "")
            ]
        room_assignments, unallocated = allocate_rooms(final_data, rooms)
    except Exception as e:
        return jsonify({"error": f"Error in room allocation: {str(e)}"}), 500

    # Build a mapping of assigned rooms for quick lookup: (course_id, slot, date) -> rooms_assigned
    room_map = {}
    for r in room_assignments:
        key = (str(r.get("course_id")), str(r.get("slot")), str(r.get("date")))
        room_map[key] = r.get("rooms_assigned", "TBD")

    # Attach room assignment to timetable rows
    for row in final_data:
        key = (str(row.get("course_id")), str(row.get("slot")), str(row.get("date")))
        row["room_assigned"] = room_map.get(key, "TBD")

    # -----------------------------------------------
    # 4. Run Stage 3 — Teacher Assignment
    # -----------------------------------------------
    faculty_status = {}
    try:
        teachers = load_teacher_data(temp_path)
        if branch_filter != "ALL":
            # Include teachers from the selected branch OR those marked as 'General' (cross-department)
            teachers = {
                tid: t for tid, t in teachers.items()
                if is_match(t.get("department", "General"), branch_filter, "branch")
                or str(t.get("department", "")).strip().upper() in ("GENERAL", "COMMON", "SHARED", "")
            }
            
        try:
            init_faculty(teachers)
            faculty_status = get_all_faculty_status()
            
            # Filter out admin teachers from the teachers pool
            teachers = {
                tid: t for tid, t in teachers.items()
                if not faculty_status.get(tid, {}).get("is_admin", False)
            }
            
            # Extract db_duty_counts
            db_duty_counts = {}
            for tid, f in faculty_status.items():
                dc = f.get("duty_counts")
                if not dc:
                    flat = f.get("duty_count", 0)
                    role = f.get("role", "Junior")
                    role_key = role.lower() if role.lower() in ["junior", "senior", "squad"] else "junior"
                    dc = {
                        "squad": flat if role_key == "squad" else 0,
                        "junior": flat if role_key == "junior" else 0,
                        "senior": flat if role_key == "senior" else 0
                    }
                db_duty_counts[tid] = dc

            fairness_map   = {tid: f["has_served_high_role"] for tid, f in faculty_status.items()}
        except Exception as db_err:
            print(f"DB Warning: {db_err}")
            fairness_map   = None
            db_duty_counts = None

        duties = build_duties(final_data, room_assignments=room_assignments)
        assignments, unassigned, teacher_duty_count = assign_teachers(
            teachers, duties, fairness_map=fairness_map, db_duty_counts=db_duty_counts
        )
    except Exception as e:
        return jsonify({"error": f"Error in teacher assignment: {str(e)}"}), 500
    # -----------------------
    # Sanitize outputs: convert numpy types to native Python types and ensure strings
    # -----------------------
    def _clean_row(r):
        out = {}
        for k, v in r.items():
            # drop internal-only keys
            if k.startswith("_"): continue
            try:
                if isinstance(v, (int, float, str, bool)) or v is None:
                    out[k] = v
                else:
                    out[k] = int(v) if (hasattr(v, 'item') and isinstance(v.item(), (int,))) else str(v)
            except Exception:
                out[k] = str(v)
        return out

    final_data = [_clean_row(r) for r in final_data]
    room_assignments = [_clean_row(r) for r in room_assignments]
    # -----------------------------------------------
    # 5. Build and return JSON response
    # -----------------------------------------------
    # Build mapping course_id -> (date, slot_display)
    course_slot_map = {r["course_id"]: (r["date"], r["slot"]) for r in final_data}

    violations = []
    for student, courses in student_courses.items():
        slots_taken = {}
        for course in courses:
            cs = course_slot_map.get(course)
            if not cs: continue
            date_taken, slot_taken = cs
            key = (date_taken, slot_taken)
            if key in slots_taken:
                violations.append(f"{student}: {course} and {slots_taken[key]} on {date_taken} Slot {slot_taken}")
            else:
                slots_taken[key] = course

    preferred_count = sum(1 for a in assignments if a.get("cost", 9999) <= 1)
    pref_pct = round(preferred_count / len(assignments) * 100, 1) if assignments else 0

    response = {
        "summary": {
            "total_courses"    : len(final_data),
            "total_students"   : len(student_courses),
            # Number of actual exam sessions (unique date + display slot pairs)
            "slots_used"       : len({(r["date"], r["slot"]) for r in final_data}),
            "conflict_edges"   : sum(len(v) for v in graph.values()) // 2,
            "conflicts_found"  : len(violations),
            "rooms_allocated"  : len(room_assignments),
            "rooms_failed"     : len(unallocated),
            "duties_assigned"  : len(assignments),
            "duties_failed"    : len(unassigned),
            "pref_satisfaction": pref_pct,
        },
        "timetable": final_data,
        "room_allocation": room_assignments,
        "teacher_duties": [
            {
                "duty_id"      : a["duty_id"],
                "slot"         : a["slot"],
                "date"         : a["date"],
                "session"      : a["session"],
                "course_id"    : a["course_id"],
                "role"         : a["role_required"],
                "teacher_id"   : a["teacher_id"],
                "teacher_name" : a["teacher_name"],
                "last_role"    : faculty_status.get(a["teacher_id"], {}).get("last_role", "N/A"),
                "is_priority"  : "Yes" if not faculty_status.get(a["teacher_id"], {}).get("has_served_high_role", True) else "No",
                "room_assigned" : (
                    room_map.get((str(a.get("course_id")), str(a.get("slot")), str(a.get("date"))))
                    if a.get("course_id") != "ALL" else ("Control" if a.get("role_required") == "Senior" else "Roaming")
                )
            }
            for a in sorted(assignments, key=lambda x: (x["date"], x["slot"], x["role_required"]))
        ],
        "warnings": {
            "unallocated_rooms" : [u["course_id"] for u in unallocated],
            "unassigned_duties" : [u["duty_id"] for u in unassigned],
            "student_conflicts" : violations
        }
    }

    return jsonify(response)


@app.route("/export_excel", methods=["POST"])
def export_excel():
    """Accepts JSON with keys 'timetable', 'room_allocation', 'teacher_duties' and returns a styled xlsx."""
    data = request.get_json()
    if not data:
        return jsonify({"error": "JSON body required"}), 400

    tables = {
        'Timetable': data.get('timetable', []),
        'Room Allocation': data.get('room_allocation', []),
        'Teacher Duties': data.get('teacher_duties', [])
    }

    # Build room_map from Room Allocation for merging into other sheets
    room_map = {}
    for r in tables.get('Room Allocation', []):
        key = (str(r.get('course_id')), str(r.get('slot')), str(r.get('date')))
        room_map[key] = r.get('rooms_assigned', '')

    # Merge Room Assigned into Timetable
    tt = []
    for row in tables.get('Timetable', []):
        key = (str(row.get('course_id')), str(row.get('slot')), str(row.get('date')))
        new = dict(row)
        new['Room Assigned'] = room_map.get(key, '')
        tt.append(new)
    tables['Timetable'] = tt

    # Prepare Teacher Duties: drop 'preferred' and add 'Room Assigned'
    td = []
    for row in tables.get('Teacher Duties', []):
        new = {k: v for k, v in row.items() if k != 'preferred'}
        key = (str(row.get('course_id')), str(row.get('slot')), str(row.get('date')))
        if row.get('course_id') == 'ALL':
            # infer control/roaming for senior/squad duties
            role = row.get('role') or row.get('role_required')
            new['Room Assigned'] = 'Control' if role == 'Senior' else 'Roaming'
        else:
            new['Room Assigned'] = room_map.get(key, '')
        td.append(new)
    tables['Teacher Duties'] = td

    # Build workbook in memory
    out = BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for name, rows in tables.items():
            df = pd.DataFrame(rows)
            if df.empty:
                df = pd.DataFrame(columns=["Empty"])
            df.to_excel(writer, sheet_name=name[:31], index=False)
        # Note: writer.save() is deprecated in pandas >= 2.0.
        # The context manager (with block) handles saving automatically on __exit__.
    out.seek(0)

    # Post-process with openpyxl for styling
    wb = load_workbook(out)
    thin = Side(border_style="thin", color="000000")
    for ws in wb.worksheets:
        # Header style
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Column widths: approximate by max length
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is None: continue
                v = str(cell.value)
                if len(v) > max_len: max_len = len(v)
            adjusted_width = (max_len + 2)
            ws.column_dimensions[col_letter].width = adjusted_width

        # Borders
        for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
            for cell in row:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Freeze header
        ws.freeze_panes = 'A2'

    # Save workbook back to bytes
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    return send_file(bio, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='timetable_export.xlsx')


@app.route("/confirm", methods=["POST"])
def confirm():
    data = request.json
    assignments = data.get("assignments", [])
    dept = data.get("department", "IT")

    try:
        for a in assignments:
            is_high = a["role"] in ["Senior", "Squad"]
            update_faculty_duty(
                a["teacher_id"],
                a["role"],
                a["date"],
                a["slot"],
                is_high_role=is_high
            )

        reset_triggered = check_reset_fairness(department=dept)
        return jsonify({"success": True, "reset_triggered": reset_triggered})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("Running Flask server...")
    print("   -> Open http://127.0.0.1:5000 in your browser")
    app.run(debug=False, port=5000)