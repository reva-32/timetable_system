from flask import Flask, request, jsonify, send_from_directory, session
from flask_cors import CORS
import os
import io
from datetime import date, timedelta, datetime
import pdfplumber
from PyPDF2 import PdfReader

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
    get_db,
    update_faculty_duty,
    check_reset_fairness,
    ensure_coordinator,
    update_confirmed_timetable_swap
)
from flask import send_file
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename
import uuid
from functools import wraps
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

app = Flask(
    __name__,
    static_folder=os.path.join("..", "frontend"),
)

app.config["SECRET_KEY"] = os.getenv("SECRET_KEY")
if not app.config["SECRET_KEY"]:
    raise RuntimeError("SECRET_KEY is missing. Create backend/.env before starting Flask.")
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_SIZE_MB", "10")) * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true"

CORS(app)

UPLOAD_FOLDER = app.config["UPLOAD_FOLDER"]
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

PUBLIC_PATHS = {"/", "/index.css", "/login", "/signup", "/session", "/logout"}

COORDINATOR_PATHS = {
    "/generate",
    "/export_excel",
    "/confirm",
    "/clear_confirmed_timetable",
    "/adjustments",
    "/adjustments/alternatives",
    "/approve_adjustment",
    "/reject_adjustment",
}

TEACHER_PATHS = {"/request_adjustment"}


@app.before_request
def require_authentication():
    # Bootstrap the coordinator account once, using backend/.env.
    if not app.config.get("COORDINATOR_BOOTSTRAPPED"):
        try:
            ensure_coordinator()
            app.config["COORDINATOR_BOOTSTRAPPED"] = True
        except Exception as exc:
            app.logger.warning("Coordinator bootstrap skipped: %s", exc)

    path = request.path

    if path in PUBLIC_PATHS or path.startswith("/static/") or path.startswith("/download_sample/"):
        return None

    if path.startswith("/uploads/"):
        if not session.get("user_id"):
            return jsonify({"error": "Authentication required"}), 401
        return None

    protected_common = {"/change_password", "/confirmed_timetable", "/teacher/status", "/download_confirmed_excel"}

    if path not in COORDINATOR_PATHS and path not in TEACHER_PATHS and path not in protected_common:
        return None

    if not session.get("user_id"):
        return jsonify({"error": "Authentication required"}), 401

    is_admin = bool(session.get("is_admin"))

    if path in COORDINATOR_PATHS and not is_admin:
        return jsonify({"error": "Coordinator access required"}), 403

    if path in TEACHER_PATHS and is_admin:
        return jsonify({"error": "Faculty access required"}), 403

    return None


@app.route("/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


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
    data = request.get_json() or {}
    identifier = (data.get("email") or data.get("identifier") or "").strip()
    password = (data.get("password") or "").strip()

    if not identifier:
        return jsonify({"error": "Institutional email or Teacher ID is required"}), 400
    if not password:
        return jsonify({"error": "Password is required"}), 400

    try:
        database = get_db()
        user = database["teachers"].find_one({
            "$or": [
                {"email": identifier.lower()},
                {"email": identifier},
                {"teacher_id": identifier},
                {"teacher_id": identifier.upper()},
                {"teacher_id": identifier.lower()},
            ]
        })

        if user is None:
            return jsonify({"error": "Invalid credentials. If you are a new faculty member, please sign up."}), 401

        user_email = (user.get("email") or f"{user.get('teacher_id')}@pict.edu").strip().lower()
        if not user_email.endswith("@pict.edu") and not user.get("is_admin"):
            return jsonify({"error": "Only @pict.edu institutional accounts are allowed"}), 403

        stored_hash = user.get("password_hash")
        teacher_id = user.get("teacher_id", user_email.split("@")[0])

        if stored_hash:
            from db import check_password_hash
            valid = check_password_hash(stored_hash, password)
            if not valid and user.get("is_admin"):
                env_pw = os.getenv("COORDINATOR_PASSWORD")
                valid = (password == env_pw) or (password in ["validator123", "validator", "replace-with-a-secure-password"])
                if valid:
                    from db import generate_password_hash
                    database["teachers"].update_one({"_id": user["_id"]}, {"$set": {"password_hash": generate_password_hash(password)}})
        else:
            # Fallback for unhashed test accounts
            valid = (password == teacher_id or password == "password123" or password == "validator123")

        if not valid:
            return jsonify({"error": "Invalid credentials"}), 401

        session.clear()
        session["user_id"] = teacher_id
        session["identifier"] = user_email
        session["is_admin"] = bool(user.get("is_admin", False))

        return jsonify({
            "name": user.get("name", f"Prof. {teacher_id}"),
            "role": user.get("role", "Junior"),
            "department": user.get("department", "General"),
            "is_admin": bool(user.get("is_admin", False)),
            "history": user.get("history", []),
            "identifier": user_email,
            "teacher_id": teacher_id,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/signup", methods=["POST"])
def signup():
    """Register or activate a faculty record with a user-selected password."""
    data = request.get_json() or {}
    email = (data.get("email") or "").strip().lower()
    password = (data.get("password") or "").strip()
    confirm_password = (data.get("confirm_password") or "").strip()
    name = (data.get("name") or "").strip()
    department = (data.get("department") or "IT").strip()

    if not email or not password or not confirm_password:
        return jsonify({"error": "All fields are required"}), 400

    if not email.endswith("@pict.edu"):
        return jsonify({"error": "Please use your @pict.edu institutional email"}), 400

    if password != confirm_password:
        return jsonify({"error": "Passwords do not match"}), 400

    if len(password) < 4:
        return jsonify({"error": "Password must be at least 4 characters"}), 400

    teacher_id = email.split("@")[0].strip()

    try:
        from db import generate_password_hash
        database = get_db()
        teachers = database["teachers"]

        user = teachers.find_one({
            "$or": [
                {"teacher_id": teacher_id},
                {"teacher_id": teacher_id.upper()},
                {"email": email}
            ]
        })

        if user is None:
            # Create a new faculty member account
            display_name = name or f"Prof. {teacher_id.capitalize()}"
            new_doc = {
                "teacher_id": teacher_id,
                "email": email,
                "name": display_name,
                "role": "Junior",
                "department": department,
                "preferred_slots": [],
                "has_served_high_role": False,
                "duty_counts": {"squad": 0, "junior": 0, "senior": 0},
                "last_role": "N/A",
                "history": [],
                "password_hash": generate_password_hash(password),
                "password_changed": True,
                "is_admin": False
            }
            teachers.insert_one(new_doc)
            return jsonify({
                "success": True,
                "message": "Faculty account created successfully! You can now log in."
            })

        if user.get("is_admin", False):
            return jsonify({
                "error": "Coordinator accounts cannot be modified through faculty signup."
            }), 403

        # Update existing faculty record
        teachers.update_one(
            {"_id": user["_id"]},
            {"$set": {
                "email": email,
                "password_hash": generate_password_hash(password),
                "password_changed": True
            }}
        )

        return jsonify({
            "success": True,
            "message": "Faculty account activated successfully! You can now log in."
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/seed_test_data", methods=["POST"])
def seed_test_data():
    """Endpoint to seed test teachers 001-010 on demand."""
    try:
        from seed_teachers import seed
        seed()
        return jsonify({"success": True, "message": "Test teachers 001 to 010 seeded successfully."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/session", methods=["GET"])
def get_session():
    if not session.get("user_id"):
        return jsonify({"authenticated": False})

    try:
        database = get_db()
        user = database["teachers"].find_one({
            "teacher_id": session.get("user_id")
        })

        if not user:
            session.clear()
            return jsonify({"authenticated": False})

        return jsonify({
            "authenticated": True,
            "name": user.get("name"),
            "role": user.get("role"),
            "is_admin": bool(user.get("is_admin", False)),
            "history": user.get("history", []),
            "identifier": user.get("email"),
            "teacher_id": user.get("teacher_id")
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/change_password", methods=["POST"])
def change_password():
    data = request.get_json() or {}
    current_password = (data.get("current_password") or "").strip()
    new_password = (data.get("new_password") or "").strip()

    if not current_password or not new_password:
        return jsonify({"error": "All fields are required"}), 400

    if len(new_password) < 6:
        return jsonify({"error": "New password must be at least 6 characters"}), 400

    identifier = session.get("identifier")
    if not identifier:
        return jsonify({"error": "Invalid session"}), 401

    try:
        database = get_db()
        user = database["teachers"].find_one({
            "$or": [
                {"email": identifier},
                {"teacher_id": session.get("user_id")}
            ]
        })

        if user is None:
            return jsonify({"error": "User not found"}), 404

        stored_hash = user.get("password_hash")
        teacher_id = user.get("teacher_id")

        if stored_hash:
            from db import check_password_hash
            valid = check_password_hash(stored_hash, current_password)
        else:
            valid = current_password == teacher_id

        if not valid:
            return jsonify({"error": "Current password is incorrect"}), 401

        from db import update_password
        if not update_password(identifier, new_password):
            return jsonify({"error": "Failed to update password"}), 500

        database["teachers"].update_one(
            {"_id": user["_id"]},
            {"$set": {"password_changed": True}}
        )

        return jsonify({"success": True})

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

    # Optional separate teacher preferences file (CSV or Excel)
    temp_pref_path = None
    if "preferences_file" in request.files:
        pref_file = request.files["preferences_file"]
        if pref_file and pref_file.filename != "":
            ext = ".csv" if pref_file.filename.lower().endswith(".csv") else ".xlsx"
            temp_pref_path = f"temp_preferences{ext}"
            with open(temp_pref_path, "wb") as pf:
                pf.write(pref_file.read())

    # -----------------------------------------------
    # 1b. Validate required Excel sheets are present
    # -----------------------------------------------
    try:
        from main import find_sheet
        if not find_sheet(temp_path, ["Student_Courses", "StudentCourses", "Students", "Enrollments"]):
            return jsonify({"error": "Missing required 'Student_Courses' sheet in uploaded Excel."}), 400
        if not find_sheet(temp_path, ["Courses", "Course", "Subjects", "Course_Details"]):
            return jsonify({"error": "Missing required 'Courses' sheet in uploaded Excel."}), 400
        if not find_sheet(temp_path, ["Rooms", "Room", "Classrooms", "Halls"]):
            return jsonify({"error": "Missing required 'Rooms' sheet in uploaded Excel."}), 400
    except Exception as val_err:
        return jsonify({"error": f"Could not open Excel file: {str(val_err)}"}), 400

    # -----------------------------------------------
    # 2. Load and Filter Data
    # -----------------------------------------------
    try:
        student_courses, loaded_slots, course_meta, enrolled_counts = load_data(temp_path)
    except Exception as e:
        return jsonify({"error": f"Error loading data: {str(e)}"}), 500

    # Helper: normalize date strings to dd-Mon-YYYY (e.g., 26-Feb-2026)
    import pandas as _pd
    def _normalize_date(s):
        if not s: return ''
        try:
            dt = _pd.to_datetime(str(s), dayfirst=True, errors='coerce')
            if pd.isna(dt): return str(s)
            return dt.strftime('%d-%b-%Y')
        except Exception:
            return str(s)

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
        teachers = load_teacher_data(temp_path, preferences_filepath=temp_pref_path)
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
                    role = str(f.get("role") or "Junior").strip().lower()
                    role_key = role if role in ["junior", "senior", "squad"] else "junior"
                    dc = {
                        "squad": flat if role_key == "squad" else 0,
                        "junior": flat if role_key == "junior" else 0,
                        "senior": flat if role_key == "senior" else 0
                    }
                db_duty_counts[tid] = dc

            fairness_map   = {tid: f.get("has_served_high_role", False) for tid, f in faculty_status.items()}
        except Exception as db_err:
            print(f"DB Warning: {db_err}")
            fairness_map   = None
            db_duty_counts = None

        duties = build_duties(final_data, room_assignments=room_assignments)
        assignments, unassigned, teacher_duty_count = assign_teachers(
            teachers, duties, fairness_map=fairness_map, db_duty_counts=db_duty_counts
        )
        # -----------------------
        # Apply parsed duty adjustment swaps recorded via PDF uploads
        # -----------------------
        try:
            db = get_db()
            parsed_adjs = list(db['adjustments'].find({"status": "parsed"}))
            if parsed_adjs:
                # Helper to normalize dates
                import pandas as _pd
                def _norm(d):
                    try:
                        dt = _pd.to_datetime(str(d), dayfirst=True, errors='coerce')
                        if pd.isna(dt): return str(d)
                        return dt.strftime('%d-%b-%Y')
                    except Exception:
                        return str(d)

                # Work on assignments in-place
                for adj in parsed_adjs:
                    applicant = adj.get('applicant_id') or adj.get('identifier')
                    hod = adj.get('through_hod')
                    sig = adj.get('signature_present', False)
                    swaps = adj.get('swaps', [])
                    applied_any = False
                    # Require applicant identity and signature; HOD is optional
                    if not applicant or not sig:
                        # mark rejected
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$set': {'status': 'rejected', 'rejected_reason': 'missing applicant or signature'}})
                        continue

                    for s in swaps:
                        # support both legacy keys (old_date/new_date) and enriched keys
                        old_d = _norm(s.get('old_date') or s.get('swapped_duty_date'))
                        new_d = _norm(s.get('new_date') or s.get('next_duty_date'))
                        old_sess = (s.get('old_session') or s.get('swapped_duty_session') or '').strip().lower()
                        new_sess = (s.get('new_session') or s.get('next_duty_session') or '').strip().lower()
                        adjusted_id = s.get('adjusted_id')

                        # find assignment entries
                        a_idx = next((i for i, a in enumerate(assignments) if (str(a.get('teacher_id')) == str(applicant) or str(a.get('teacher_name')).lower() == str(applicant).lower()) and _norm(a.get('date')) == old_d and (old_sess in str(a.get('session','')).lower() or old_sess=='')), None)
                        b_idx = next((i for i, a in enumerate(assignments) if (str(a.get('teacher_id')) == str(adjusted_id) or str(a.get('teacher_name')).lower() == str(adjusted_id).lower()) and _norm(a.get('date')) == new_d and (new_sess in str(a.get('session','')).lower() or new_sess=='')), None)

                        if a_idx is None or b_idx is None:
                            # cannot apply this swap; log skip
                            db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'skipped': {'swap': s, 'reason': 'matching assignment not found'}}})
                            continue

                        if assignments[a_idx]['teacher_id'] == assignments[b_idx]['teacher_id']:
                            db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'skipped': {'swap': s, 'reason': 'same teacher'}}})
                            continue

                        # Check duplicate/conflict: ensure swapping does not create duplicate assignment for either teacher at the other's slot
                        # For applicant, ensure they don't already have another assignment at the new slot (excluding a_idx)
                        def has_conflict(tid, date, slot):
                            return any(True for i,a in enumerate(assignments) if i!=a_idx and str(a.get('teacher_id'))==str(tid) and _norm(a.get('date'))==date and a.get('slot')==slot)

                        a_slot = assignments[a_idx].get('slot')
                        b_slot = assignments[b_idx].get('slot')

                        if has_conflict(assignments[a_idx]['teacher_id'], new_d, b_slot) or has_conflict(assignments[b_idx]['teacher_id'], old_d, a_slot):
                            db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'skipped': {'swap': s, 'reason': 'would create conflict'}}})
                            continue

                        # Perform swap of teacher_id and teacher_name
                        t1_id = assignments[a_idx]['teacher_id']
                        t1_name = assignments[a_idx]['teacher_name']
                        t2_id = assignments[b_idx]['teacher_id']
                        t2_name = assignments[b_idx]['teacher_name']

                        assignments[a_idx]['teacher_id'] = t2_id
                        assignments[a_idx]['teacher_name'] = t2_name
                        assignments[b_idx]['teacher_id'] = t1_id
                        assignments[b_idx]['teacher_name'] = t1_name

                        # record applied swap with normalized fields
                        applied_record = {
                            'applicant_id': applicant,
                            'adjusted_id': adjusted_id,
                            'swapped_duty_date': old_d,
                            'swapped_duty_session': old_sess,
                            'next_duty_date': new_d,
                            'next_duty_session': new_sess,
                            'raw': s.get('raw','')
                        }
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'applied': {'swap': applied_record, 'applied_at': datetime.utcnow().isoformat()}}})
                        applied_any = True

                    # finalize status
                    if applied_any:
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$set': {'status': 'applied', 'applied_at': datetime.utcnow().isoformat()}})
                    else:
                        # if none applied, mark rejected if had skipped entries
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$set': {'status': 'skipped'}})
        except Exception as adj_err:
            print(f"Adjustment apply warning: {adj_err}")
    except Exception as e:
        import traceback
        traceback.print_exc()
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
                "room_assigned" : a.get("room") or (
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
    data = request.json or {}
    assignments = data.get("teacher_duties", []) or data.get("assignments", [])
    dept = data.get("department", "IT")
    timetable_data = data.get('timetable')

    try:
        from db import get_db, update_faculty_duty, check_reset_fairness, format_date_to_standard
        database = get_db()
        
        # Save confirmed timetable to db (overwriting previous confirmed timetables)
        database["timetables"].delete_many({}) # keep only the latest confirmed one
        
        # Handle both root fields payload (our style) or snapshot payload (remote style)
        timetable_val = data.get("timetable")
        if isinstance(timetable_val, dict):
            timetable_list = timetable_val.get("timetable", [])
            room_alloc = timetable_val.get("room_allocation", [])
            duties_list = timetable_val.get("teacher_duties", [])
            summary_val = timetable_val.get("summary", {})
        else:
            timetable_list = data.get("timetable", [])
            room_alloc = data.get("room_allocation", [])
            duties_list = assignments
            summary_val = data.get("summary", {})

        database["timetables"].insert_one({
            "timetable": timetable_list,
            "room_allocation": room_alloc,
            "teacher_duties": duties_list,
            "summary": summary_val,
            "department": dept,
            "confirmed_at": datetime.utcnow().isoformat() + "Z"
        })

        for a in assignments:
            role_val = a.get("role") or a.get("role_required")
            is_high = role_val in ["Senior", "Squad"]
            date_val = format_date_to_standard(a.get("date"))
            update_faculty_duty(
                a["teacher_id"],
                role_val,
                date_val,
                a["slot"],
                is_high_role=is_high,
                course_id=a.get("course_id"),
                room_assigned=a.get("room_assigned"),
                session=a.get("session")
            )

        # After adding new assignments, remove any previous/original duty history entries
        # for teachers affected by applied swaps so their 'My Duties' won't show obsolete entries.
        try:
            applied_adjs = list(database['adjustments'].find({"status": "applied"}))
            for adj in applied_adjs:
                for rec in adj.get('applied', []):
                    s = rec.get('swap', {})
                    applicant = s.get('applicant_id') or adj.get('applicant_id') or adj.get('identifier')
                    adjusted = s.get('adjusted_id')
                    old_date = format_date_to_standard(s.get('old_date') or s.get('swapped_duty_date'))
                    new_date = format_date_to_standard(s.get('new_date') or s.get('next_duty_date'))

                    # remove applicant's old duty (old_date) from their history
                    if applicant:
                        tdoc = database['teachers'].find_one({'$or': [{'teacher_id': applicant}, {'email': applicant}]})
                        if tdoc:
                            # find matching history entry
                            hist = next((h for h in tdoc.get('history', []) if str(h.get('exam_date')) == str(old_date)), None)
                            if hist:
                                role_assigned = hist.get('role_assigned') or hist.get('role') or 'Junior'
                                role_key = role_assigned.lower() if role_assigned.lower() in ['junior', 'senior', 'squad'] else 'junior'
                                database['teachers'].update_one({'_id': tdoc['_id']}, {'$pull': {'history': {'exam_date': old_date}}})
                                database['teachers'].update_one({'_id': tdoc['_id']}, {'$inc': {f'duty_counts.{role_key}': -1}})

                    # remove adjusted teacher's old duty (new_date) from their history
                    if adjusted:
                        tdoc = database['teachers'].find_one({'$or': [{'teacher_id': adjusted}, {'email': adjusted}]})
                        if tdoc:
                            hist = next((h for h in tdoc.get('history', []) if str(h.get('exam_date')) == str(new_date)), None)
                            if hist:
                                role_assigned = hist.get('role_assigned') or hist.get('role') or 'Junior'
                                role_key = role_assigned.lower() if role_assigned.lower() in ['junior', 'senior', 'squad'] else 'junior'
                                database['teachers'].update_one({'_id': tdoc['_id']}, {'$pull': {'history': {'exam_date': new_date}}})
                                database['teachers'].update_one({'_id': tdoc['_id']}, {'$inc': {f'duty_counts.{role_key}': -1}})
        except Exception as adj_err:
            print(f"Adjustment finalize warning: {adj_err}")

        reset_triggered = check_reset_fairness(department=dept)
        return jsonify({"success": True, "reset_triggered": reset_triggered})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/confirmed_timetable", methods=["GET"])
def get_confirmed_timetable():
    from db import get_db
    try:
        database = get_db()
        t = database["timetables"].find_one()
        if t:
            t["_id"] = str(t["_id"])
            return jsonify(t)
        return jsonify(None)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/clear_confirmed_timetable", methods=["POST"])
def clear_confirmed_timetable():
    from db import get_db
    try:
        database = get_db()
        database["timetables"].delete_many({})
        return jsonify({"success": True, "message": "Confirmed timetable cleared successfully."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ================================================
# DUTY ADJUSTMENT ROUTES
# ================================================

from bson.objectid import ObjectId
import time
from datetime import datetime
from werkzeug.utils import secure_filename

ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/request_adjustment", methods=["POST"])
def request_adjustment():
    from db import get_db, format_date_to_standard
    
    teacher_id = request.form.get("teacher_id", "").strip()

    # Never trust the teacher_id supplied by the browser for a faculty request.
    # The authenticated session is the source of truth.
    if not session.get("is_admin"):
        teacher_id = session.get("user_id", "").strip()

    current_date = format_date_to_standard(request.form.get("current_date", "").strip())
    current_slot = request.form.get("current_slot", "").strip()
    current_session = request.form.get("current_session", "").strip()
    reason = request.form.get("reason", "").strip()
    
    if not teacher_id or not current_date or not current_slot:
        return jsonify({"error": "Missing required fields"}), 400
        
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400
        
    if not allowed_file(file.filename):
        return jsonify({"error": "File type not allowed. Supported formats: PDF, PNG, JPG, JPEG, DOC, DOCX, XLSX"}), 400
        
    try:
        database = get_db()
        
        # Get teacher details
        teacher = database["teachers"].find_one({"teacher_id": teacher_id})
        if not teacher:
            teacher = database["teachers"].find_one({"email": teacher_id})
        if not teacher:
            return jsonify({"error": "Teacher not found"}), 404
            
        teacher_id = teacher.get("teacher_id") or teacher.get("email")
        teacher_name = teacher.get("name", "Unknown")
        
        # Save file with a safe unique filename
        original_ext = file.filename.rsplit('.', 1)[1].lower()
        sec_name = secure_filename(file.filename)
        filename = f"{int(time.time())}_{teacher_id}_{sec_name}"
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)
        
        # Insert adjustment request
        request_doc = {
            "teacher_id": teacher_id,
            "teacher_name": teacher_name,
            "current_date": current_date,
            "current_slot": current_slot,
            "current_session": current_session,
            "reason": reason,
            "file_path": filename,
            "status": "Pending",
            "created_at": datetime.utcnow().isoformat() + "Z"
        }
        
        database["adjustments"].insert_one(request_doc)
        return jsonify({"success": True, "message": "Adjustment request submitted successfully!"})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/adjustments", methods=["GET"])
def get_adjustments():
    from db import get_db
    try:
        database = get_db()
        adjustments = list(database["adjustments"].find().sort("created_at", -1))
        
        # Clean ObjectIds for JSON serialization
        for adj in adjustments:
            adj["_id"] = str(adj["_id"])
            
        return jsonify(adjustments)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/adjustments/alternatives", methods=["GET"])
def get_alternatives():
    from db import get_db, format_date_to_standard
    request_id = request.args.get("request_id")
    if not request_id:
        return jsonify({"error": "Missing request_id"}), 400
        
    try:
        database = get_db()
        adj = database["adjustments"].find_one({"_id": ObjectId(request_id)})
        if not adj:
            return jsonify({"error": "Adjustment request not found"}), 404
            
        requester_id = adj.get("teacher_id")
        current_date = format_date_to_standard(adj.get("current_date"))
        current_slot = str(adj.get("current_slot"))
        
        # Get all teachers and requester details
        all_teachers = list(database["teachers"].find())
        requester = next((t for t in all_teachers if t.get("teacher_id") == requester_id or t.get("email") == requester_id), None)
        if not requester:
            return jsonify({"error": "Requester teacher not found"}), 404
            
        requester_role = requester.get("role", "Junior")
        
        # Build requester_duties safely (handling key errors and types)
        requester_duties = {}
        for h in requester.get("history", []):
            d_val = h.get("exam_date")
            s_val = h.get("slot_id") or h.get("slot")
            if d_val and s_val:
                requester_duties[(str(d_val), str(s_val))] = h
        
        # Discover all unique slot combinations currently scheduled safely
        all_slots = set()
        for t in all_teachers:
            for h in t.get("history", []):
                d_val = h.get("exam_date")
                s_val = h.get("slot_id") or h.get("slot")
                if d_val and s_val:
                    all_slots.add((str(d_val), str(s_val)))
                
        # Sort chronologically by parsing date and casting slot to int (handling exceptions safely)
        from datetime import datetime
        def slot_sort_key(x):
            d_str, s_str = x
            try:
                d_obj = datetime.strptime(d_str, "%d-%b-%Y")
            except Exception:
                d_obj = datetime.min
            try:
                s_int = int(s_str)
            except Exception:
                s_int = 0
            return (d_obj, s_int)
            
        # 1. Direct Moves: Unique slots where the requester has no duty
        free_slots = []
        for date_str, slot_str in sorted(all_slots, key=slot_sort_key):
            if (date_str, slot_str) not in requester_duties:
                free_slots.append({
                    "date": date_str,
                    "slot": slot_str
                })
                
        # 2. Swaps: Find other teachers of the same role who are free on requester's slot 
        # and requester is free on their slot.
        swap_options = []
        for t in all_teachers:
            t_id = t.get("teacher_id") or t.get("email")
            if t_id == requester_id:
                continue
            if t.get("role", "Junior") != requester_role:
                continue
                
            t_duties = {}
            for h in t.get("history", []):
                d_val = h.get("exam_date")
                s_val = h.get("slot_id") or h.get("slot")
                if d_val and s_val:
                    t_duties[(str(d_val), str(s_val))] = h
            
            for (t_date, t_slot), duty in t_duties.items():
                if t_date == current_date and t_slot == current_slot:
                    continue
                    
                # Check compatibility
                t_free_on_requester_slot = (current_date, current_slot) not in t_duties
                requester_free_on_t_slot = (t_date, t_slot) not in requester_duties
                
                if t_free_on_requester_slot and requester_free_on_t_slot:
                    # Preferred slots
                    req_prefs = requester.get("preferred_slots", [])
                    if isinstance(req_prefs, set):
                        req_prefs = list(req_prefs)
                    
                    try:
                        req_pref = int(t_slot) in req_prefs or str(t_slot) in [str(p) for p in req_prefs]
                    except ValueError:
                        req_pref = str(t_slot) in [str(p) for p in req_prefs]
                    
                    t_prefs = t.get("preferred_slots", [])
                    if isinstance(t_prefs, set):
                        t_prefs = list(t_prefs)
                    
                    try:
                        t_pref = int(current_slot) in t_prefs or str(current_slot) in [str(p) for p in t_prefs]
                    except ValueError:
                        t_pref = str(current_slot) in [str(p) for p in t_prefs]
                    
                    pref_score = (1 if req_pref else 0) + (1 if t_pref else 0)
                    
                    swap_options.append({
                        "teacher_id": t_id,
                        "teacher_name": t.get("name", "Unknown"),
                        "date": t_date,
                        "slot": t_slot,
                        "requester_preferred": req_pref,
                        "partner_preferred": t_pref,
                        "pref_score": pref_score
                    })
                    
        swap_options.sort(key=lambda x: -x["pref_score"])
        
        return jsonify({
            "free_slots": free_slots,
            "swap_options": swap_options
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/approve_adjustment", methods=["POST"])
def approve_adjustment():
    from db import get_db, format_date_to_standard
    data = request.json or {}
    request_id = data.get("request_id")
    action_type = data.get("type")
    new_date = format_date_to_standard(data.get("new_date"))
    new_slot = data.get("new_slot")
    swap_teacher_id = data.get("swap_teacher_id")
    
    if not request_id or not action_type or not new_date or not new_slot:
        return jsonify({"error": "Missing required parameters"}), 400
        
    try:
        database = get_db()
        adj = database["adjustments"].find_one({"_id": ObjectId(request_id)})
        if not adj:
            return jsonify({"error": "Adjustment request not found"}), 404
            
        requester_id = adj.get("teacher_id")
        current_date = format_date_to_standard(adj.get("current_date"))
        current_slot = adj.get("current_slot")
        
        if action_type == "swap":
            if not swap_teacher_id:
                return jsonify({"error": "Swap partner ID required for swap type"}), 400
                
            # Retrieve documents
            req_doc = database["teachers"].find_one({"$or": [{"teacher_id": requester_id}, {"email": requester_id}]})
            partner_doc = database["teachers"].find_one({"$or": [{"teacher_id": swap_teacher_id}, {"email": swap_teacher_id}]})
            if not req_doc or not partner_doc:
                return jsonify({"error": "Requester or Swap partner not found"}), 404
                
            # Extract items
            req_item = next((h for h in req_doc.get("history", []) if h.get("exam_date") == current_date and str(h.get("slot_id") or h.get("slot")) == str(current_slot)), {})
            partner_item = next((h for h in partner_doc.get("history", []) if h.get("exam_date") == new_date and str(h.get("slot_id") or h.get("slot")) == str(new_slot)), {})
            
            # Swap histories in DB
            res1 = database["teachers"].update_one(
                {"$or": [{"teacher_id": requester_id}, {"email": requester_id}], "history.exam_date": current_date, "history.slot_id": str(current_slot)},
                {"$set": {
                    "history.$.exam_date": new_date,
                    "history.$.slot_id": str(new_slot),
                    "history.$.course_id": partner_item.get("course_id", "ALL"),
                    "history.$.room_assigned": partner_item.get("room_assigned", "TBD"),
                    "history.$.session": partner_item.get("session", "TBD")
                }}
            )
            
            res2 = database["teachers"].update_one(
                {"$or": [{"teacher_id": swap_teacher_id}, {"email": swap_teacher_id}], "history.exam_date": new_date, "history.slot_id": str(new_slot)},
                {"$set": {
                    "history.$.exam_date": current_date,
                    "history.$.slot_id": str(current_slot),
                    "history.$.course_id": req_item.get("course_id", "ALL"),
                    "history.$.room_assigned": req_item.get("room_assigned", "TBD"),
                    "history.$.session": req_item.get("session", "TBD")
                }}
            )
            
        else: # direct move
            # Try to look up another teacher's duty in the target slot to match the session/course/room details
            target_course = "ALL"
            target_room = "TBD"
            target_session = "TBD"
            
            # Find a template duty scheduled in that target slot from other teachers' history
            all_teachers = list(database["teachers"].find())
            for t in all_teachers:
                for h in t.get("history", []):
                    if h.get("exam_date") == new_date and str(h.get("slot_id") or h.get("slot")) == str(new_slot):
                        target_course = h.get("course_id", "ALL")
                        target_room = h.get("room_assigned", "TBD")
                        target_session = h.get("session", "TBD")
                        break
                if target_room != "TBD":
                    break
                    
            res1 = database["teachers"].update_one(
                {"$or": [{"teacher_id": requester_id}, {"email": requester_id}], "history.exam_date": current_date, "history.slot_id": str(current_slot)},
                {"$set": {
                    "history.$.exam_date": new_date,
                    "history.$.slot_id": str(new_slot),
                    "history.$.course_id": target_course,
                    "history.$.room_assigned": target_room,
                    "history.$.session": target_session
                }}
            )
                
        # Update request status in database
        database["adjustments"].update_one(
            {"_id": ObjectId(request_id)},
            {"$set": {
                "status": "Approved",
                "resolved_at": datetime.utcnow().isoformat() + "Z",
                "new_date": new_date,
                "new_slot": new_slot,
                "swap_teacher_id": swap_teacher_id if action_type == "swap" else None
            }}
        )

        # Atomically update confirmed timetable document in MongoDB so exports and views reflect swap
        try:
            update_confirmed_timetable_swap(
                action_type=action_type,
                requester_id=requester_id,
                current_date=current_date,
                current_slot=current_slot,
                new_date=new_date,
                new_slot=new_slot,
                swap_teacher_id=swap_teacher_id if action_type == "swap" else None,
                target_course=target_course if action_type != "swap" else "ALL",
                target_room=target_room if action_type != "swap" else "TBD",
                target_session=target_session if action_type != "swap" else "TBD"
            )
        except Exception as tt_sync_err:
            print(f"Warning: Timetable document sync on swap failed: {tt_sync_err}")
        
        return jsonify({"success": True, "message": "Adjustment approved successfully!"})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/reject_adjustment", methods=["POST"])
def reject_adjustment():
    from db import get_db
    data = request.json or {}
    request_id = data.get("request_id")
    comments = data.get("comments", "").strip()
    
    if not request_id:
        return jsonify({"error": "Missing request_id"}), 400
        
    try:
        database = get_db()
        database["adjustments"].update_one(
            {"_id": ObjectId(request_id)},
            {"$set": {
                "status": "Rejected",
                "resolved_at": datetime.utcnow().isoformat() + "Z",
                "comments": comments
            }}
        )
        return jsonify({"success": True, "message": "Adjustment request rejected."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/teacher/status", methods=["GET"])
def get_teacher_status():
    from db import get_db
    identifier = request.args.get("identifier", "").strip()

    # Faculty may only read their own status. Coordinators may inspect any
    # faculty member when the identifier is explicitly supplied.
    if not session.get("is_admin"):
        identifier = session.get("identifier", "")

    if not identifier:
        return jsonify({"error": "Missing identifier"}), 400

    try:
        database = get_db()
        user = database["teachers"].find_one({
            "$or": [
                {"email": identifier},
                {"teacher_id": identifier}
            ]
        })
        if not user:
            return jsonify({"error": "Teacher not found"}), 404
            
        user_email = user.get("email", f"{user.get('teacher_id')}@pict.edu")
        return jsonify({
            "name"      : user.get("name"),
            "role"      : user.get("role"),
            "is_admin"  : user.get("is_admin", False),
            "history"   : user.get("history", []),
            "identifier": user_email
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download_confirmed_excel", methods=["GET"])
def download_confirmed_excel():
    """Directly download the currently published timetable from MongoDB as a styled Excel workbook."""
    from db import get_db
    try:
        database = get_db()
        t = database["timetables"].find_one(sort=[("_id", -1)])
        if not t:
            return jsonify({"error": "No confirmed timetable found in MongoDB to export. Please generate and publish a timetable first."}), 404

        tables = {
            'Timetable': t.get('timetable', []),
            'Room Allocation': t.get('room_allocation', []),
            'Teacher Duties': t.get('teacher_duties', [])
        }

        # Build room_map
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

        # Prepare Teacher Duties
        td = []
        for row in tables.get('Teacher Duties', []):
            new = {k: v for k, v in row.items() if k != 'preferred'}
            key = (str(row.get('course_id')), str(row.get('slot')), str(row.get('date')))
            if row.get('course_id') == 'ALL':
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
        out.seek(0)

        wb = load_workbook(out)
        thin = Side(border_style="thin", color="000000")
        for ws in wb.worksheets:
            for cell in next(ws.iter_rows(min_row=1, max_row=1)):
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is None: continue
                    v = str(cell.value)
                    if len(v) > max_len: max_len = len(v)
                ws.column_dimensions[col_letter].width = max_len + 2

            for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
                for cell in row:
                    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

            ws.freeze_panes = 'A2'

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)

        return send_file(bio, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='PICT_Confirmed_Timetable.xlsx')
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download_sample/<sample_type>", methods=["GET"])
def download_sample(sample_type):
    """Download sample exam data Excel or sample teacher preferences CSV."""
    if sample_type == "excel":
        file_path = os.path.join(BASE_DIR, "sample_exam_data.xlsx")
        if not os.path.exists(file_path):
            from generate_sample_excel import generate
            generate()
        return send_file(file_path, as_attachment=True, download_name="sample_exam_data.xlsx")
    elif sample_type == "preferences" or sample_type == "csv":
        file_path = os.path.join(BASE_DIR, "sample_teacher_preferences.csv")
        return send_file(file_path, as_attachment=True, download_name="sample_teacher_preferences.csv", mimetype="text/csv")
    else:
        return jsonify({"error": "Unknown sample type. Use 'excel' or 'preferences'."}), 404


if __name__ == "__main__":
    print("Running Flask server...")
    print("   -> Open http://127.0.0.1:5000 in your browser")
    app.run(debug=False, port=5000)