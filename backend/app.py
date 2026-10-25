from flask import Flask, request, jsonify, send_from_directory, session
from flask_cors import CORS
import os
import io
import re
import hmac
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
    validate_schedule,
    assign_teachers,
    load_teacher_data,
    normalize_duty_config,
    teacher_can_invigilate,
    normalize_year,
    normalize_role,
    _date_sort_key,
)
# Import MongoDB functions
from db import (
    init_faculty,
    get_all_faculty_status,
    get_db,
    update_faculty_duty,
    check_reset_fairness,
    ensure_coordinator,
    update_confirmed_timetable_swap,
    set_teacher_eligibility
)
from assignments import (
    AssignmentConflict,
    build_all_teacher_duties_payload,
    count_approved_swaps,
    derive_history,
    get_active_timetable,
    migrate_legacy_timetables,
    move_assignment,
    rebuild_teacher_history,
    select_best_timetable_doc,
    swap_assignments,
    sync_timetable_assignments,
)
from eligibility import effective_teaching_years, session_code
from department_config import DEPARTMENT_ORDER
from department_pdf import build_detailed_timetable_pdf, build_public_timetable_pdf
from multi_department import generate_multi_department_run
from rosters import (
    build_roster_context,
    normalize_roster_type,
    render_excel as render_roster_excel,
    render_pdf as render_roster_pdf,
    roster_filename,
)
from reports import (
    DOCX_MIME,
    XLSX_MIME,
    generate_daily_change_reports,
    generate_faculty_reports,
)
from mailer import (
    get_mail_job,
    get_mail_status,
    mail_config,
    queue_daily_change_emails,
    queue_timetable_emails,
    send_timetable_emails,
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

app.config["SECRET_KEY"] = os.getenv("SECRET_KEY") or os.getenv("FLASK_SECRET_KEY")
if not app.config["SECRET_KEY"]:
    raise RuntimeError("SECRET_KEY is missing. Create backend/.env before starting Flask.")
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_SIZE_MB", "10")) * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = os.getenv("SESSION_COOKIE_SAMESITE", "Lax")
app.config["SESSION_COOKIE_SECURE"] = os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true"

ALLOWED_ORIGIN = os.getenv("FRONTEND_ORIGIN", "").strip()
if ALLOWED_ORIGIN:
    CORS(app, origins=[ALLOWED_ORIGIN], supports_credentials=True)

UPLOAD_FOLDER = app.config["UPLOAD_FOLDER"]
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DEFAULT_JS_PER_SESSION = int(os.getenv("JS_PER_SESSION", "0") or 0)
DEFAULT_SQUAD_PER_SESSION = int(os.getenv("SQUAD_PER_SESSION", "9") or 9)
DEFAULT_LAST_DAY_SQUAD = int(os.getenv("LAST_DAY_SQUAD", "2") or 2)
ALLOW_DIRECT_MOVE = os.getenv("ALLOW_DIRECT_MOVE", "false").lower() == "true"
try:
    MAX_DUTIES = max(1, int(os.getenv("MAX_DUTIES", "5")))
except (TypeError, ValueError):
    MAX_DUTIES = 5

PUBLIC_PATHS = {"/", "/index.css", "/login", "/signup", "/session", "/logout", "/csrf-token"}

COORDINATOR_PATHS = {
    "/generate",
    "/generate_departments",
    "/confirm",
    "/confirm_departments",
    "/export_excel",
    "/clear_confirmed_timetable",
    "/adjustments",
    "/adjustments/alternatives",
    "/approve_adjustment",
    "/reject_adjustment",
    "/admin/teachers",
    "/timetable/versions",
    "/timetable/rollback",
    "/api/timetable/duties",
}

TEACHER_PATHS = {"/request_adjustment", "/teacher/requests"}


def json_safe_value(value):
    """Convert numpy/pandas scalars while preserving JSON containers.

    The old sanitizer stringified lists (for example ``allowed_years``), which
    silently broke year eligibility during the generate -> confirm round trip.
    """
    if value is None or isinstance(value, (int, float, str, bool)):
        return value
    if isinstance(value, dict):
        return {key: json_safe_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [json_safe_value(item) for item in value]
    if hasattr(value, "item"):
        try:
            return json_safe_value(value.item())
        except Exception:
            pass
    return str(value)


def json_safe_row(row):
    return {
        key: json_safe_value(value)
        for key, value in row.items()
        if not key.startswith("_")
    }


def _csrf_token():
    token = session.get("csrf_token")
    if not token:
        token = uuid.uuid4().hex
        session["csrf_token"] = token
    return token

def require_admin(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            return jsonify({"error": "Authentication required"}), 401
        if not session.get("is_admin"):
            return jsonify({"error": "Coordinator access required"}), 403
        return fn(*args, **kwargs)
    return wrapper

def require_csrf(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        expected = session.get("csrf_token")
        supplied = request.headers.get("X-CSRF-Token", "")
        if not expected or not supplied or not hmac.compare_digest(expected, supplied):
            return jsonify({"error": "Invalid CSRF token"}), 403
        return fn(*args, **kwargs)
    return wrapper

def safe_server_error(message="Internal server error"):
    app.logger.exception(message)
    return jsonify({"error": message}), 500

@app.errorhandler(Exception)
def handle_unexpected_error(exc):
    app.logger.exception("Unhandled application error")
    return jsonify({"error": "Internal server error"}), 500

@app.route("/csrf-token", methods=["GET"])
def csrf_token():
    return jsonify({"csrf_token": _csrf_token()})

@app.before_request
def require_authentication():
    # Bootstrap the coordinator account once, using backend/.env.
    if not app.config.get("TESTING") and not app.config.get("COORDINATOR_BOOTSTRAPPED"):
        try:
            ensure_coordinator()
            app.config["COORDINATOR_BOOTSTRAPPED"] = True
        except Exception as exc:
            app.logger.warning("Coordinator bootstrap skipped: %s", exc)

    # Import legacy timetable/history rows into the assignments collection once.
    if not app.config.get("TESTING") and not app.config.get("ASSIGNMENTS_BOOTSTRAPPED"):
        try:
            migrated = migrate_legacy_timetables()
            if migrated:
                rebuild_teacher_history()
                app.logger.info("Migrated %s duty rows into assignments", migrated)
            app.config["ASSIGNMENTS_BOOTSTRAPPED"] = True
        except Exception as exc:
            app.logger.warning("Assignment bootstrap skipped: %s", exc)

    path = request.path

    if path in PUBLIC_PATHS or path.startswith("/static/") or path.startswith("/download_sample/"):
        return None

    if path.startswith("/uploads/"):
        if not session.get("user_id"):
            return jsonify({"error": "Authentication required"}), 401
        return None

    if path.startswith("/roster/"):
        if not session.get("user_id"):
            return jsonify({"error": "Authentication required"}), 401
        if not session.get("is_admin"):
            return jsonify({"error": "Coordinator access required"}), 403
        return None

    if path.startswith("/generation_runs/"):
        if not session.get("user_id"):
            return jsonify({"error": "Authentication required"}), 401
        if path.endswith("/detailed.pdf") and not session.get("is_admin"):
            return jsonify({"error": "Coordinator access required"}), 403
        return None

    protected_common = {"/change_password", "/confirmed_timetable", "/teacher/status", "/download_confirmed_excel", "/api/teacher/duties"}

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
    email = (data.get("email") or "").strip().lower()
    password = (data.get("password") or "").strip()

    if not email:
        return jsonify({"error": "Institutional email is required"}), 400
    if "@" not in email:
        return jsonify({"error": "Sign in with your full institutional email, e.g. t001@pict.edu"}), 400
    if not password:
        return jsonify({"error": "Password is required"}), 400

    try:
        database = get_db()
        user = database["teachers"].find_one({"email": email})

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
        else:
            return jsonify({"error": "Account password is not initialized. Ask the coordinator to run the user migration."}), 403

        if not valid:
            return jsonify({"error": "Invalid credentials"}), 401

        session.clear()
        session["user_id"] = teacher_id
        session["identifier"] = user_email
        session["authenticated"] = True
        session["teacher_id"] = teacher_id
        session["email"] = user_email
        session["is_admin"] = bool(user.get("is_admin", False))

        try:
            history = derive_history(teacher_id, database=database)
        except Exception:
            history = user.get("history", [])

        return jsonify({
            "name": user.get("name", f"Prof. {teacher_id}"),
            "role": user.get("role", "Junior"),
            "department": user.get("department", "General"),
            "is_admin": bool(user.get("is_admin", False)),
            "history": history,
            "identifier": user_email,
            "teacher_id": teacher_id,
        })

    except Exception:
        return safe_server_error()


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

    if len(password) < 8 or not re.search(r"[A-Z]", password) or not re.search(r"[a-z]", password) or not re.search(r"\d", password) or not re.search(r"[^A-Za-z0-9]", password):
        return jsonify({"error": "Password must be at least 8 characters and include upper, lower, digit and special character"}), 400

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
                "teaching_years": sorted(effective_teaching_years({"department": department})),
                "max_duties": MAX_DUTIES,
                "pinned_duties": [],
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

    except Exception:
        return safe_server_error()


@app.route("/seed_test_data", methods=["POST"])
@require_csrf
@require_admin
def seed_test_data():
    """Endpoint to seed test teachers 001-010 on demand."""
    try:
        from seed_teachers import seed
        seed()
        return jsonify({"success": True, "message": "Test teachers 001 to 010 seeded successfully."})
    except Exception:
        return safe_server_error()


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

        try:
            history = derive_history(user.get("teacher_id"), database=database)
        except Exception:
            history = user.get("history", [])

        return jsonify({
            "authenticated": True,
            "name": user.get("name"),
            "role": user.get("role"),
            "department": user.get("department"),
            "is_admin": bool(user.get("is_admin", False)),
            "history": history,
            "identifier": user.get("email"),
            "teacher_id": user.get("teacher_id")
        })
    except Exception as e:
        return safe_server_error()


@app.route("/logout", methods=["POST"])
@require_csrf
def logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/change_password", methods=["POST"])
@require_csrf
def change_password():
    data = request.get_json() or {}
    current_password = (data.get("current_password") or "").strip()
    new_password = (data.get("new_password") or "").strip()

    if not current_password or not new_password:
        return jsonify({"error": "All fields are required"}), 400

    if len(new_password) < 8 or not re.search(r"[A-Z]", new_password) or not re.search(r"[a-z]", new_password) or not re.search(r"\d", new_password) or not re.search(r"[^A-Za-z0-9]", new_password):
        return jsonify({"error": "Password must be at least 8 characters and include upper, lower, digit and special character"}), 400

    identifier = session.get("identifier")
    if not session.get("authenticated") or not identifier:
        return jsonify({"error": "Authentication required"}), 401
    if identifier not in {session.get("identifier"), session.get("teacher_id"), session.get("email")} and not session.get("is_admin"):
        return jsonify({"error": "You can only change your own password"}), 403

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
            return jsonify({"error": "Account password is not initialized. Ask the coordinator to run the user migration."}), 403

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
        return safe_server_error()


@app.route("/admin/teachers", methods=["GET", "PUT"])
@require_admin
def admin_teachers():
    if request.method == "GET":
        docs = list(get_db()["teachers"].find({}, {"_id": 0, "password_hash": 0}))
        return jsonify({"teachers": docs})
    expected = session.get("csrf_token")
    supplied = request.headers.get("X-CSRF-Token", "")
    if not expected or not supplied or not hmac.compare_digest(expected, supplied):
        return jsonify({"error": "Invalid CSRF token"}), 403
    data = request.get_json() or {}
    teacher_id = str(data.get("teacher_id", "")).strip()
    years = data.get("teaching_years", [])
    if not teacher_id:
        return jsonify({"error": "teacher_id is required"}), 400
    try:
        ok = set_teacher_eligibility(
            teacher_id,
            years,
            data.get("department"),
            max_duties=data.get("max_duties"),
            pinned_duties=data.get("pinned_duties"),
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    if not ok:
        return jsonify({"error": "Teacher not found"}), 404
    updated = get_db()["teachers"].find_one(
        {"teacher_id": teacher_id},
        {"_id": 0, "password_hash": 0},
    )
    return jsonify({"success": True, "teacher": updated})

@app.route("/generate", methods=["POST"])
@require_csrf
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
    # Dates from the Courses sheet are authoritative.  The start date is used
    # only when a workbook has no pre-assigned exam dates.
    start_date_raw = (request.form.get("start_date") or "").strip()
    try:
        start_date = datetime.strptime(start_date_raw, "%Y-%m-%d").date() if start_date_raw else date.today()
    except ValueError:
        start_date = date.today()

    duty_config = normalize_duty_config({
        "js_per_session": request.form.get("js_per_session", DEFAULT_JS_PER_SESSION),
        "squad_per_session": request.form.get("squad_per_session", DEFAULT_SQUAD_PER_SESSION),
        "last_day_squad": request.form.get("last_day_squad", DEFAULT_LAST_DAY_SQUAD),
        "overrides": request.form.get("duty_overrides", "{}"),
    })
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

        duties = build_duties(final_data, room_assignments=room_assignments, duty_config=duty_config)
        assignments, unassigned, teacher_duty_count = assign_teachers(
            teachers, duties, fairness_map=fairness_map, db_duty_counts=db_duty_counts,
            MAX_DUTIES=MAX_DUTIES, faculty_status=faculty_status,
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
    # Sanitize outputs: convert numpy types to native Python types while
    # preserving lists/dicts (year eligibility depends on allowed_years).
    # -----------------------
    _clean_row = json_safe_row

    final_data = [_clean_row(r) for r in final_data]
    room_assignments = [_clean_row(r) for r in room_assignments]
    # -----------------------------------------------
    # 5. Build and return JSON response
    # -----------------------------------------------
    # Build mapping course_id -> (date, slot_display)
    course_slot_map = {r["course_id"]: (r["date"], r["slot"]) for r in final_data}

    violations = validate_schedule(final_data, student_courses)
    if violations:
        app.logger.error("Scheduler invariant violation: %s", violations)
        return jsonify({"error": "Generated timetable violates a student scheduling constraint", "violations": violations}), 422

    preferred_count = sum(1 for a in assignments if a.get("cost", 9999) <= 1)
    pref_pct = round(preferred_count / len(assignments) * 100, 1) if assignments else 0
    generated_loads = [int(v) for v in (teacher_duty_count or {}).values()]
    role_counts = {}
    for duty in duties:
        role_counts[duty["role_required"]] = role_counts.get(duty["role_required"], 0) + 1
    assigned_role_counts = {}
    for assignment in assignments:
        role = assignment["role_required"]
        assigned_role_counts[role] = assigned_role_counts.get(role, 0) + 1

    assignment_by_duty = {str(a.get("duty_id")): a for a in assignments}
    all_duties = []
    for duty in sorted(duties, key=lambda item: (_date_sort_key(item.get("date")), str(item.get("slot")), item.get("role_required"))):
        row = {
            key: value for key, value in _clean_row(duty).items()
            if key not in ("preferred",)
        }
        assigned = assignment_by_duty.get(str(duty.get("duty_id")))
        if assigned:
            row.update({
                "teacher_id": assigned.get("teacher_id"),
                "teacher_name": assigned.get("teacher_name"),
                "pinned": bool(assigned.get("pinned")),
            })
        all_duties.append(row)
    unassigned_rows = [
        _clean_row(duty) for duty in sorted(
            unassigned, key=lambda item: (_date_sort_key(item.get("date")), str(item.get("slot")), item.get("role_required"))
        )
    ]

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
            "duties_required"  : len(duties),
            "duties_assigned"  : len(assignments),
            "duties_failed"    : len(unassigned),
            "duties_by_role"   : role_counts,
            "assigned_by_role" : assigned_role_counts,
            "max_generated_load": max(generated_loads) if generated_loads else 0,
            "min_generated_load": min(generated_loads) if generated_loads else 0,
            "pref_satisfaction": pref_pct,
        },
        "duty_config": duty_config,
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
                "year"         : a.get("year", "ALL"),
                "department"   : a.get("department", "General"),
                "pinned"       : bool(a.get("pinned")),
                "last_role"    : faculty_status.get(a["teacher_id"], {}).get("last_role", "N/A"),
                "is_priority"  : "Yes" if not faculty_status.get(a["teacher_id"], {}).get("has_served_high_role", True) else "No",
                "room_assigned" : a.get("room") or (
                    room_map.get((str(a.get("course_id")), str(a.get("slot")), str(a.get("date"))))
                    if a.get("course_id") != "ALL" else ("Control" if a.get("role_required") == "Senior" else "Roaming")
                )
            }
            for a in sorted(assignments, key=lambda x: (x["date"], x["slot"], x["role_required"]))
        ],
        # Every seat, including unfilled ones, is published so the roster can
        # track vacancies and headcount precisely.
        "duties": all_duties,
        "unassigned_duties": unassigned_rows,
        "warnings": {
            "unallocated_rooms" : [u["course_id"] for u in unallocated],
            "unassigned_duties" : [u["duty_id"] for u in unassigned],
            "student_conflicts" : violations
        }
    }

    return jsonify(response)


def _run_pdf_filename(run, detailed=False):
    year = str(run.get("year") or "TY").replace(".", "")
    exam = str(run.get("exam_type") or "Insem")
    stem = f"PICT_{year}_{exam}_{'Detailed_' if detailed else ''}Timetable"
    return re.sub(r"[^A-Za-z0-9_-]+", "_", stem) + ".pdf"


def _department_summary(run, department):
    department = str(department)
    rows = [row for row in run.get("timetable") or [] if str(row.get("department")) == department]
    sessions = [
        session for item in run.get("departments") or []
        if str(item.get("department")) == department
        for session in item.get("sessions") or []
    ]
    room_rows = [
        row for row in run.get("room_allocation") or []
        if str(row.get("department")) == department
    ]
    duties = [
        duty for duty in run.get("duties") or []
        if str(duty.get("exam_department") or duty.get("department")) == department
    ]
    assignments = [
        row for row in run.get("teacher_duties") or []
        if str(row.get("exam_department") or row.get("department")) == department
    ]
    role_counts = {}
    assigned_role_counts = {}
    for duty in duties:
        role = duty.get("role_required") or "Junior"
        role_counts[role] = role_counts.get(role, 0) + 1
    for assignment in assignments:
        role = assignment.get("role_required") or assignment.get("role") or "Junior"
        assigned_role_counts[role] = assigned_role_counts.get(role, 0) + 1
    return {
        "total_courses": len(rows),
        "total_students": sum(int(session.get("student_count") or 0) for session in sessions),
        "slots_used": len({(row.get("date"), row.get("session")) for row in rows}),
        "conflict_edges": 0,
        "conflicts_found": 0,
        "rooms_allocated": len(room_rows),
        "rooms_failed": sum(1 for row in room_rows if row.get("status") != "Allocated"),
        "duties_required": len(duties),
        "duties_assigned": len(assignments),
        "duties_failed": max(0, len(duties) - len(assignments)),
        "duties_by_role": role_counts,
        "assigned_by_role": assigned_role_counts,
        "pref_satisfaction": 0,
        "max_generated_load": 0,
        "min_generated_load": 0,
    }


@app.route("/generate_departments", methods=["POST"])
@require_csrf
def generate_departments():
    """Generate a coordinated timetable for every selected department."""
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file uploaded"}), 400
    if not str(file.filename).lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Upload an Excel workbook (.xlsx or .xls)."}), 400

    start_date_raw = (request.form.get("start_date") or "").strip()
    end_date_raw = (request.form.get("end_date") or "").strip()
    if not start_date_raw or not end_date_raw:
        return jsonify({"error": "Select both the examination start date and end date."}), 400

    year_filter = (request.form.get("year") or "TY").strip() or "TY"
    exam_type = (request.form.get("exam_type") or "Insem").strip() or "Insem"
    departments_raw = (request.form.get("departments") or "").strip()
    departments = [
        value.strip() for value in departments_raw.split(",") if value.strip()
    ] or list(DEPARTMENT_ORDER)
    selected_slots = [
        value.strip() for value in (request.form.get("selected_slots") or "").split(",")
        if value.strip()
    ]

    # Keep the uploaded source workbook instead of overwriting temp_data.xlsx.
    filename = f"run_{uuid.uuid4().hex}_{secure_filename(file.filename)}"
    source_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(source_path)

    try:
        run = generate_multi_department_run(
            source_path,
            year_filter=year_filter,
            exam_type=exam_type,
            start_date=start_date_raw,
            end_date=end_date_raw,
            departments=departments,
            selected_slots=selected_slots,
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Multi-department generation failed")
        return jsonify({"error": "Could not generate the department timetables. Check the workbook and try again."}), 500

    run_id = uuid.uuid4().hex
    created_at = datetime.utcnow().isoformat() + "Z"
    run.update({
        "run_id": run_id,
        "status": "draft",
        "created_at": created_at,
        "created_by": session.get("user_id"),
        "source_filename": secure_filename(file.filename),
        "source_path": filename,
        "timetable_pdf_url": f"/generation_runs/{run_id}/timetable.pdf",
        "detailed_pdf_url": f"/generation_runs/{run_id}/detailed.pdf",
    })
    try:
        get_db()["generation_runs"].insert_one({
            "run_id": run_id,
            "status": "draft",
            "created_at": created_at,
            "created_by": session.get("user_id"),
            "payload": json_safe_value(run),
        })
    except Exception:
        app.logger.exception("Could not save generation run")
        return jsonify({"error": "The schedule was generated but could not be saved as a draft."}), 500

    return jsonify(json_safe_value(run))


@app.route("/generation_runs/<run_id>/timetable.pdf", methods=["GET"])
def download_generation_timetable(run_id):
    try:
        document = get_db()["generation_runs"].find_one({"run_id": run_id})
        if not document:
            return jsonify({"error": "Generation run not found"}), 404
        payload = document.get("payload") or {}
        if not payload.get("departments"):
            return jsonify({"error": "This run has no department timetables"}), 404
        pdf_bytes = build_public_timetable_pdf(payload)
        return send_file(
            BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=False,
            download_name=_run_pdf_filename(payload),
        )
    except Exception:
        return safe_server_error("Could not render the timetable PDF")


@app.route("/generation_runs/<run_id>/detailed.pdf", methods=["GET"])
@require_admin
def download_generation_detailed(run_id):
    try:
        document = get_db()["generation_runs"].find_one({"run_id": run_id})
        if not document:
            return jsonify({"error": "Generation run not found"}), 404
        payload = document.get("payload") or {}
        if not payload.get("departments"):
            return jsonify({"error": "This run has no department timetables"}), 404
        pdf_bytes = build_detailed_timetable_pdf(payload)
        return send_file(
            BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=_run_pdf_filename(payload, detailed=True),
        )
    except Exception:
        return safe_server_error("Could not render the detailed timetable PDF")


@app.route("/confirm_departments", methods=["POST"])
@require_csrf
@require_admin
def confirm_departments():
    """Publish every department timetable in one coordinated run."""
    data = request.get_json(silent=True) or {}
    run_id = str(data.get("run_id") or "").strip()
    if not run_id:
        return jsonify({"error": "Generation run id is required"}), 400

    database = get_db()
    run_document = database["generation_runs"].find_one({"run_id": run_id})
    if not run_document:
        return jsonify({"error": "Generation run not found"}), 404
    if run_document.get("status") == "published":
        return jsonify({
            "success": True,
            "already_published": True,
            "run_id": run_id,
            "timetable_ids": run_document.get("timetable_ids") or {},
        })
    run = run_document.get("payload") or {}
    departments = [item for item in run.get("departments") or [] if item.get("department")]
    if not departments:
        return jsonify({"error": "Generation run has no departments"}), 400

    confirmed_at = datetime.utcnow().isoformat() + "Z"
    combined_at = (datetime.utcnow() + timedelta(microseconds=1)).isoformat() + "Z"
    teacher_lookup = {
        str(doc.get("teacher_id")): doc
        for doc in database["teachers"].find({})
        if doc.get("teacher_id")
    }

    published = {}
    inserted_timetable_ids = []
    period_dates = run.get("calendar_dates") or sorted({
        str(row.get("date")) for row in run.get("timetable") or [] if row.get("date")
    })
    try:
        for department in departments:
            dept = str(department.get("department"))
            dept_rows = [row for row in run.get("timetable") or [] if str(row.get("department")) == dept]
            dept_rooms = [
                row for row in run.get("room_allocation") or []
                if str(row.get("department")) == dept
            ]
            dept_duties = [
                duty for duty in run.get("duties") or []
                if str(duty.get("exam_department") or duty.get("department")) == dept
            ]
            dept_assignments = [
                row for row in run.get("teacher_duties") or []
                if str(row.get("exam_department") or row.get("department")) == dept
            ]
            latest = database["timetables"].find_one({"department": dept}, sort=[("version", -1)])
            next_version = int(latest.get("version", 0)) + 1 if latest else 1
            summary = _department_summary(run, dept)
            document = {
                "timetable": dept_rows,
                "room_allocation": dept_rooms,
                "teacher_duties": dept_assignments,
                "all_duties": dept_duties,
                "summary": summary,
                "duty_config": {},
                "department": dept,
                "version": next_version,
                "revision": 1,
                "cleared": False,
                "confirmed_at": confirmed_at,
                "exam_period": {
                    "start_date": run.get("start_date"),
                    "end_date": run.get("end_date"),
                },
                "generation_run_id": run_id,
                "multi_department": False,
            }
            inserted = database["timetables"].insert_one(document)
            inserted_timetable_ids.append(inserted.inserted_id)
            sync_timetable_assignments(
                str(inserted.inserted_id),
                dept,
                dept_duties,
                teachers=teacher_lookup,
                confirmed_at=confirmed_at,
                database=database,
            )
            # The old roster for these same dates has just been replaced.
            # Remove its archived seats so fairness history does not count the
            # superseded duties twice.
            try:
                database["assignments"].delete_many({
                    "department": dept,
                    "date": {"$in": period_dates},
                    "active": False,
                })
            except Exception:
                app.logger.exception("Could not remove superseded assignments for %s", dept)
            published[dept] = str(inserted.inserted_id)

        # Combined read-only view used by the dashboard and public PDF flow.
        combined = {
            "timetable": run.get("timetable") or [],
            "room_allocation": run.get("room_allocation") or [],
            "teacher_duties": run.get("teacher_duties") or [],
            "all_duties": run.get("duties") or [],
            "summary": run.get("summary") or {},
            "duty_config": {},
            "department": "ALL",
            "departments": [item.get("department") for item in departments],
            "version": 1,
            "revision": 1,
            "cleared": False,
            "confirmed_at": combined_at,
            "generation_run_id": run_id,
            "multi_department": True,
            "exam_period": {
                "start_date": run.get("start_date"),
                "end_date": run.get("end_date"),
            },
        }
        combined_insert = database["timetables"].insert_one(combined)
        inserted_timetable_ids.append(combined_insert.inserted_id)
        database["timetables"].update_many(
            {
                "multi_department": True,
                "cleared": {"$ne": True},
                "_id": {"$ne": combined_insert.inserted_id},
            },
            {"$set": {"cleared": True, "cleared_at": confirmed_at}},
        )

        teacher_ids = {
            str(duty.get("teacher_id"))
            for duty in run.get("duties") or []
            if duty.get("teacher_id")
        }
        try:
            rebuild_teacher_history(teacher_ids=teacher_ids, database=database)
        except Exception:
            app.logger.exception("Teacher history refresh failed after multi-department publish")
    except Exception:
        # Remove any partially-inserted view documents.  Assignment writes are
        # themselves transactional in sync_timetable_assignments.
        for timetable_id in inserted_timetable_ids:
            try:
                database["timetables"].delete_one({"_id": timetable_id})
            except Exception:
                pass
        app.logger.exception("Multi-department publish failed")
        return jsonify({"error": "Could not publish the department timetables."}), 500

    database["generation_runs"].update_one(
        {"run_id": run_id},
        {"$set": {
            "status": "published",
            "published_at": combined_at,
            "published_by": session.get("user_id"),
            "timetable_ids": published,
            "combined_timetable_id": str(combined_insert.inserted_id),
        }},
    )
    return jsonify({
        "success": True,
        "run_id": run_id,
        "timetable_ids": published,
        "combined_timetable_id": str(combined_insert.inserted_id),
        "assignments_created": len(run.get("teacher_duties") or []),
        "summary": run.get("summary") or {},
        "mail": {"queued": False, "reason": "Use the detailed PDF for room and duty distribution."},
    })


@app.route("/export_excel", methods=["POST"])
@require_csrf
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
@require_csrf
def confirm():
    data = request.json or {}
    assignments = data.get("teacher_duties", []) or data.get("assignments", [])
    dept = data.get("department", "IT") or "IT"
    summary_val = data.get("summary", {})
    duty_config = data.get("duty_config") or {}

    try:
        database = get_db()

        # Handle both root fields payload (our style) or snapshot payload (remote style)
        timetable_val = data.get("timetable")
        if isinstance(timetable_val, dict):
            timetable_list = timetable_val.get("timetable", [])
            room_alloc = timetable_val.get("room_allocation", [])
            duties_list = timetable_val.get("teacher_duties", []) or assignments
            summary_val = timetable_val.get("summary", summary_val)
            duty_config = timetable_val.get("duty_config", duty_config)
        else:
            timetable_list = data.get("timetable", []) or []
            room_alloc = data.get("room_allocation", []) or []
            duties_list = assignments

        # Every seat (including vacancies) is published.  Older clients only
        # send assigned duties; convert them to seats.
        seats = data.get("duties") or data.get("all_duties") or []
        seat_map = {str(seat.get("duty_id")): seat for seat in seats if seat.get("duty_id") is not None}
        if not seats:
            for assignment in duties_list:
                seat = dict(assignment)
                seat["role_required"] = seat.get("role") or seat.get("role_required")
                seats.append(seat)
        else:
            # Make sure every assigned duty is present as a seat and that the
            # assigned teacher is attached even if the client sent a stale
            # all-duties list (older clients did not include teachers there).
            for assignment in duties_list:
                key = str(assignment.get("duty_id"))
                seat = seat_map.get(key)
                if seat is None:
                    seat = {
                        **assignment,
                        "role_required": assignment.get("role") or assignment.get("role_required"),
                    }
                    seat_map[key] = seat
                else:
                    seat["teacher_id"] = assignment.get("teacher_id")
                    seat["teacher_name"] = assignment.get("teacher_name")
                    seat["pinned"] = bool(assignment.get("pinned"))
            seats = list(seat_map.values())

        # Publish as a new immutable timetable version instead of deleting prior history.
        latest = database["timetables"].find_one({"department": dept}, sort=[("version", -1)])
        next_version = int(latest.get("version", 0)) + 1 if latest else 1
        confirmed_at = datetime.utcnow().isoformat() + "Z"

        insert_result = database["timetables"].insert_one({
            "timetable": timetable_list,
            "room_allocation": room_alloc,
            "teacher_duties": duties_list,
            "all_duties": seats,
            "summary": summary_val,
            "duty_config": duty_config,
            "department": dept,
            "version": next_version,
            "revision": 1,
            "cleared": False,
            "confirmed_at": confirmed_at,
        })
        timetable_id = insert_result.inserted_id

        # Authoritative seat documents.  If this fails the publish is rolled back
        # so the coordinator never sees a "success" with a stale roster.
        try:
            teachers = {
                str(doc.get("teacher_id")): doc
                for doc in database["teachers"].find({})
                if doc.get("teacher_id")
            }
            created = sync_timetable_assignments(
                str(timetable_id), dept, seats, teachers=teachers,
                confirmed_at=confirmed_at, database=database,
            )
        except Exception as assignment_error:
            database["timetables"].delete_one({"_id": timetable_id})
            app.logger.exception("Failed to publish assignment seats: %s", assignment_error)
            return jsonify({"error": "Could not publish duty assignments. The previous roster was kept."}), 500

        teacher_ids = sorted({
            str(seat.get("teacher_id"))
            for seat in seats
            if seat.get("teacher_id")
        })
        try:
            rebuild_teacher_history(teacher_ids=teacher_ids, database=database)
        except Exception as history_error:
            # Assignments are authoritative and teacher endpoints read them;
            # a stale history cache must not make publication fail.
            app.logger.exception("Teacher history refresh failed: %s", history_error)

        reset_triggered = check_reset_fairness(department=dept)
        try:
            mail_status = queue_timetable_emails(
                str(timetable_id), database=database, triggered_by=session.get("user_id")
            )
        except Exception as mail_error:
            app.logger.exception("Could not queue timetable emails: %s", mail_error)
            mail_status = {
                "queued": False,
                "reason": "Duty report emails could not be queued.",
            }
        return jsonify({
            "success": True,
            "timetable_id": str(timetable_id),
            "version": next_version,
            "assignments_created": created,
            "reset_triggered": reset_triggered,
            "mail": mail_status,
            "mail_queued": bool(mail_status.get("queued")),
            "mail_job_id": mail_status.get("job_id"),
        })
    except Exception:
        return safe_server_error()


@app.route("/confirmed_timetable", methods=["GET"])
def get_confirmed_timetable():
    try:
        database = get_db()
        combined = database["timetables"].find_one(
            {"multi_department": True, "cleared": {"$ne": True}},
            sort=[("confirmed_at", -1)],
        )
        single = get_active_timetable(database=database)
        # A newer single-department publication must still win over an older
        # multi-department view.
        if combined and (
            not single
            or str(combined.get("confirmed_at") or "") >= str(single.get("confirmed_at") or "")
        ):
            payload = dict(combined)
            payload["_id"] = str(payload["_id"])
            return jsonify(payload)
        if single:
            payload = dict(single)
            payload["_id"] = str(payload["_id"])
            return jsonify(payload)
        return jsonify(None)
    except Exception:
        return safe_server_error()


@app.route("/timetable/versions", methods=["GET"])
@require_admin
def timetable_versions():
    docs = list(get_db()["timetables"].find({}, {"_id": 1, "department": 1, "version": 1, "confirmed_at": 1, "summary": 1, "cleared": 1}).sort([("department", 1), ("version", -1)]))
    for d in docs:
        d["_id"] = str(d["_id"])
    return jsonify(docs)


@app.route("/timetable/rollback", methods=["POST"])
@require_csrf
@require_admin
def timetable_rollback():
    data = request.get_json() or {}
    department = str(data.get("department", "IT"))
    try:
        version = int(data.get("version"))
    except Exception:
        return jsonify({"error": "Version must be an integer"}), 400
    database = get_db()
    source = database["timetables"].find_one({"department": department, "version": version})
    if not source:
        return jsonify({"error": "Timetable version not found"}), 404
    latest = database["timetables"].find_one({"department": department}, sort=[("version", -1)])
    next_version = int(latest.get("version", 0)) + 1 if latest else 1
    clone = {k: v for k, v in source.items() if k != "_id"}
    clone.update({"version": next_version, "rollback_of": version, "cleared": False, "revision": 1, "confirmed_at": datetime.utcnow().isoformat() + "Z"})
    result = database["timetables"].insert_one(clone)
    try:
        teachers = {str(doc.get("teacher_id")): doc for doc in database["teachers"].find({}) if doc.get("teacher_id")}
        seats = clone.get("all_duties") or clone.get("teacher_duties") or []
        sync_timetable_assignments(str(result.inserted_id), department, seats, teachers=teachers,
                                   confirmed_at=clone["confirmed_at"], database=database)
        rebuild_teacher_history(database=database)
    except Exception as rollback_error:
        database["timetables"].delete_one({"_id": result.inserted_id})
        app.logger.exception("Rollback assignment sync failed: %s", rollback_error)
        return jsonify({"error": "Could not restore duty assignments for this version"}), 500
    return jsonify({"success": True, "version": next_version, "rollback_of": version})


@app.route("/clear_confirmed_timetable", methods=["POST"])
@require_csrf
def clear_confirmed_timetable():
    try:
        database = get_db()
        active_ids = [doc["_id"] for doc in database["timetables"].find({"cleared": {"$ne": True}}, {"_id": 1})]
        database["timetables"].update_many({"cleared": {"$ne": True}}, {"$set": {"cleared": True, "cleared_at": datetime.utcnow().isoformat() + "Z"}})
        if active_ids:
            database["assignments"].update_many(
                {"timetable_id": {"$in": [str(doc_id) for doc_id in active_ids]}, "active": True},
                {"$set": {"active": False, "archived_at": datetime.utcnow().isoformat() + "Z"}},
            )
        return jsonify({"success": True, "message": "Confirmed timetable cleared successfully."})
    except Exception:
        return safe_server_error()


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


ADJUSTMENT_STATUS_META = {
    "pending": {"label": "Under Scrutiny", "tone": "warning"},
    "parsed": {"label": "Under Scrutiny", "tone": "warning"},
    "processing": {"label": "In Review", "tone": "info"},
    "approved": {"label": "Approved", "tone": "success"},
    "applied": {"label": "Approved", "tone": "success"},
    "rejected": {"label": "Rejected", "tone": "danger"},
    "skipped": {"label": "Not Applied", "tone": "muted"},
}


def adjustment_status_meta(raw_status):
    key = str(raw_status or "Pending").strip().lower()
    return key, ADJUSTMENT_STATUS_META.get(
        key, {"label": str(raw_status or "Pending"), "tone": "muted"}
    )


@app.route("/request_adjustment", methods=["POST"])
@require_csrf
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
        return safe_server_error()


@app.route("/adjustments", methods=["GET"])
def get_adjustments():
    from db import get_db
    try:
        database = get_db()
        adjustments = list(database["adjustments"].find().sort("created_at", -1))
        
        # Clean ObjectIds for JSON serialization and expose stable status labels.
        for adj in adjustments:
            adj["_id"] = str(adj["_id"])
            status_key, status_meta = adjustment_status_meta(adj.get("status"))
            adj["status_key"] = status_key
            adj["status_label"] = status_meta["label"]
            
        return jsonify(adjustments)
    except Exception as e:
        return safe_server_error()


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

        timetable = get_active_timetable(database=database)
        if not timetable:
            return jsonify({"error": "No published timetable. Publish a timetable before approving changes."}), 409
        timetable_id = str(timetable["_id"])

        requester_query = adj.get("teacher_id")
        requester = database["teachers"].find_one({
            "$or": [
                {"teacher_id": requester_query},
                {"email": requester_query},
                {"email": str(requester_query).lower()},
            ]
        })
        if not requester:
            return jsonify({"error": "Requester teacher not found"}), 404
        requester_id = str(requester.get("teacher_id"))

        current_date = format_date_to_standard(adj.get("current_date"))
        current_slot = str(adj.get("current_slot"))

        rows = list(database["assignments"].find({
            "timetable_id": timetable_id,
            "active": True,
        }))
        requester_rows = [row for row in rows if str(row.get("teacher_id")) == requester_id]
        requester_duty = next(
            (row for row in requester_rows
             if format_date_to_standard(row.get("date")) == current_date and str(row.get("slot")) == current_slot),
            None,
        )
        if not requester_duty:
            return jsonify({"error": "The requested duty is no longer on the active roster"}), 409

        # The published seat is authoritative. A teacher's profile role or
        # eligibility may have changed since publication; the swap must still
        # be based on the role of the duty that is actually being exchanged.
        requester_role = normalize_role(requester_duty.get("role") or requester.get("role", "Junior"))
        requester_dates = {format_date_to_standard(row.get("date")) for row in requester_rows}

        # A direct move would leave the source day short-staffed.  It is
        # disabled by default; the college uses mutual swaps.
        free_slots = []
        if ALLOW_DIRECT_MOVE:
            for row in rows:
                if row.get("teacher_id"):
                    continue
                if normalize_role(row.get("role")) != requester_role:
                    continue
                target_date = format_date_to_standard(row.get("date"))
                if target_date in requester_dates:
                    continue
                if not teacher_can_invigilate(requester, {
                    "year": row.get("year"),
                    "allowed_years": row.get("allowed_years"),
                    "department": row.get("department_required") or row.get("department"),
                }):
                    continue
                free_slots.append({
                    "duty_id": row.get("duty_id"),
                    "date": target_date,
                    "slot": str(row.get("slot")),
                    "session": row.get("session"),
                    "role": normalize_role(row.get("role")),
                })

        def slot_sort_key(item):
            try:
                d_obj = datetime.strptime(item["date"], "%d-%b-%Y")
            except Exception:
                d_obj = datetime.min
            try:
                s_int = int(item["slot"])
            except Exception:
                s_int = 0
            return (d_obj, s_int)

        free_slots.sort(key=slot_sort_key)

        # Mutual swaps with the same role, year rule and free dates.
        swap_options = []
        seen = set()
        diagnostics = {
            "same_role_duties": 0,
            "date_clashes": 0,
            "requester_ineligible": 0,
            "partner_ineligible": 0,
            "missing_partner": 0,
            "eligible_options": 0,
        }
        partners = {
            str(doc.get("teacher_id")): doc
            for doc in database["teachers"].find({})
            if doc.get("teacher_id")
        }
        requester_prefs = requester.get("preferred_slots", []) or []
        if isinstance(requester_prefs, str):
            requester_prefs = [p.strip() for p in requester_prefs.split(",") if p.strip()]

        for row in rows:
            if normalize_role(row.get("role")) != requester_role:
                continue
            partner_id = str(row.get("teacher_id") or "")
            if not partner_id or partner_id == requester_id:
                continue
            diagnostics["same_role_duties"] += 1

            partner = partners.get(partner_id)
            if not partner:
                diagnostics["missing_partner"] += 1
                continue
            if partner.get("is_admin"):
                continue

            partner_dates = {
                format_date_to_standard(other.get("date"))
                for other in rows
                if str(other.get("teacher_id")) == partner_id
            }
            target_date = format_date_to_standard(row.get("date"))
            target_slot = str(row.get("slot"))
            if target_date == current_date and target_slot == current_slot:
                continue
            if target_date in requester_dates or current_date in partner_dates:
                diagnostics["date_clashes"] += 1
                continue

            partner_duty = {
                "date": target_date,
                "slot": target_slot,
                "year": row.get("year"),
                "allowed_years": row.get("allowed_years"),
                "department": row.get("department_required") or row.get("department"),
            }
            requester_duty_for_partner = {
                "date": current_date,
                "slot": current_slot,
                "year": requester_duty.get("year"),
                "allowed_years": requester_duty.get("allowed_years"),
                "department": requester_duty.get("department_required") or requester_duty.get("department"),
            }
            if not teacher_can_invigilate(requester, partner_duty):
                diagnostics["requester_ineligible"] += 1
                continue
            if not teacher_can_invigilate(partner, requester_duty_for_partner):
                diagnostics["partner_ineligible"] += 1
                continue

            option_key = (partner_id, target_date, target_slot, requester_role)
            if option_key in seen:
                continue
            seen.add(option_key)

            # Preferences describe the slots each teacher would *receive* after
            # the swap, not the duty they already hold.
            req_pref = str(target_slot) in {str(p) for p in requester_prefs}
            partner_prefs = partner.get("preferred_slots", []) or []
            if isinstance(partner_prefs, str):
                partner_prefs = [p.strip() for p in partner_prefs.split(",") if p.strip()]
            partner_pref = str(current_slot) in {str(p) for p in partner_prefs}
            swap_options.append({
                "duty_id": row.get("duty_id"),
                "teacher_id": partner_id,
                "teacher_name": partner.get("name", "Unknown"),
                "date": target_date,
                "slot": target_slot,
                "session": row.get("session") or "TBD",
                "course_id": row.get("course_id") or "ALL",
                "room": row.get("room") or "TBD",
                "role": requester_role,
                "requester_preferred": req_pref,
                "partner_preferred": partner_pref,
                "pref_score": (1 if req_pref else 0) + (1 if partner_pref else 0),
            })

        diagnostics["eligible_options"] = len(swap_options)
        swap_options.sort(key=lambda item: (-item["pref_score"], item["date"], str(item["slot"])))

        if swap_options:
            no_options_message = ""
        elif not effective_teaching_years(requester):
            no_options_message = (
                f"{requester.get('name', requester_id)} has no eligible invigilation years "
                f"under the current department/year rules "
                f"(department {requester.get('department', 'General')}, "
                f"allow-list {requester.get('teaching_years') or []}). "
                "Fix that teacher's eligibility before choosing a swap."
            )
        elif diagnostics["same_role_duties"] == 0:
            no_options_message = (
                f"No other {requester_role} duty exists on the published roster, "
                "so a same-role mutual swap is impossible."
            )
        elif diagnostics["date_clashes"]:
            no_options_message = (
                f"No compatible {requester_role} swap was found. "
                f"{diagnostics['same_role_duties']} same-role duty seat(s) exist, but "
                f"{diagnostics['date_clashes']} are blocked by a same-day clash for one of the teachers."
            )
        else:
            no_options_message = (
                f"No compatible {requester_role} swap was found. "
                "The remaining same-role duties fail the year or department eligibility rules."
            )

        return jsonify({
            "timetable_id": timetable_id,
            "revision": int(timetable.get("revision") or 1),
            "request": {
                "date": current_date,
                "slot": current_slot,
                "role": requester_role,
            },
            "direct_move_allowed": ALLOW_DIRECT_MOVE,
            "free_slots": free_slots,
            "swap_options": swap_options,
            "diagnostics": diagnostics,
            "message": no_options_message,
        })

    except Exception:
        app.logger.exception("Failed to load adjustment alternatives")
        return safe_server_error()


@app.route("/approve_adjustment", methods=["POST"])
@require_csrf
def approve_adjustment():
    from db import format_date_to_standard

    data = request.json or {}
    request_id = data.get("request_id")
    action_type = str(data.get("type") or "").strip().lower()
    new_date = format_date_to_standard(data.get("new_date"))
    new_slot = data.get("new_slot")
    swap_teacher_id = data.get("swap_teacher_id")

    if not request_id or action_type not in {"swap", "move"} or not new_date or new_slot is None:
        return jsonify({"error": "Missing required parameters"}), 400

    try:
        database = get_db()
        try:
            adj = database["adjustments"].find_one({"_id": ObjectId(request_id)})
        except Exception:
            return jsonify({"error": "Adjustment request not found"}), 404
        if not adj:
            return jsonify({"error": "Adjustment request not found"}), 404

        status_key = str(adj.get("status") or "").strip().lower()
        if status_key not in {"pending", "parsed"}:
            return jsonify({"error": "This adjustment request has already been resolved"}), 409

        requester_id = adj.get("teacher_id")
        current_date = format_date_to_standard(adj.get("current_date"))
        current_slot = adj.get("current_slot")

        timetable = get_active_timetable(database=database)
        if not timetable:
            return jsonify({"error": "No published timetable. Publish a timetable before approving changes."}), 409
        timetable_id = str(timetable["_id"])

        # Claim the request first so two coordinators cannot approve it twice.
        claim = database["adjustments"].update_one(
            {"_id": adj["_id"], "status": adj.get("status")},
            {"$set": {
                "status": "Processing",
                "processing_at": datetime.utcnow().isoformat() + "Z",
                "processing_by": session.get("user_id"),
            }},
        )
        if claim.modified_count != 1:
            return jsonify({"error": "This adjustment request is already being processed"}), 409

        def restore_status():
            database["adjustments"].update_one(
                {"_id": adj["_id"], "status": "Processing"},
                {"$set": {"status": adj.get("status")}},
            )

        try:
            if action_type == "swap":
                if not swap_teacher_id:
                    restore_status()
                    return jsonify({"error": "Swap partner ID required for swap type"}), 400
                result = swap_assignments(
                    timetable_id,
                    requester_id,
                    swap_teacher_id,
                    current_date,
                    current_slot,
                    new_date,
                    new_slot,
                    actor_id=session.get("user_id"),
                    reason=adj.get("reason", ""),
                    database=database,
                )
                message = "Duty swap approved and the roster was updated."
            else:
                if not ALLOW_DIRECT_MOVE:
                    restore_status()
                    return jsonify({
                        "error": "Direct moves are disabled. Use a mutual swap so both days stay fully staffed."
                    }), 409
                result = move_assignment(
                    timetable_id,
                    requester_id,
                    new_date,
                    new_slot,
                    actor_id=session.get("user_id"),
                    reason=adj.get("reason", ""),
                    database=database,
                )
                message = "Duty moved to the vacant seat and the roster was updated."
        except AssignmentConflict as conflict:
            restore_status()
            return jsonify({"error": str(conflict)}), 409
        except Exception:
            restore_status()
            raise

        try:
            rebuild_teacher_history(database=database)
        except Exception:
            app.logger.exception("Teacher history refresh failed after adjustment approval")

        database["adjustments"].update_one(
            {"_id": adj["_id"], "status": "Processing"},
            {"$set": {
                "status": "Approved",
                "resolved_at": datetime.utcnow().isoformat() + "Z",
                "resolved_by": session.get("user_id"),
                "action": action_type,
                "new_date": new_date,
                "new_slot": new_slot,
                "swap_teacher_id": swap_teacher_id if action_type == "swap" else None,
                "timetable_id": timetable_id,
                "revision": result.get("revision"),
            }},
        )

        daily_reports = {"generated": 0, "report_dates": [], "files": []}
        daily_mail = {"queued": False, "reason": "Daily reports were not generated."}
        try:
            timetable_after = get_active_timetable(database=database) or timetable
            generated = generate_daily_change_reports(
                database,
                timetable_after,
                [current_date, new_date],
                persist=True,
                generated_by=session.get("user_id"),
            )
            daily_reports = {
                "generated": len(generated),
                "report_dates": [item["context"]["report_date"] for item in generated],
                "files": [
                    {
                        "report_date": item["context"]["report_date"],
                        "docx": item["filenames"]["docx"],
                        "xlsx": item["filenames"]["xlsx"],
                    }
                    for item in generated
                ],
            }
            if generated:
                daily_mail = queue_daily_change_emails(
                    timetable_id,
                    [item["context"]["report_date"] for item in generated],
                    database=database,
                    triggered_by=session.get("user_id"),
                )
        except Exception as report_error:
            app.logger.exception("Daily change report generation failed: %s", report_error)
            daily_reports["error"] = "Daily change reports could not be generated."

        return jsonify({
            "success": True,
            "message": message,
            "revision": result.get("revision"),
            "timetable_id": timetable_id,
            "daily_reports": daily_reports,
            "mail": daily_mail,
            "mail_queued": bool(daily_mail.get("queued")),
            "mail_job_id": daily_mail.get("job_id"),
        })
    except Exception:
        return safe_server_error()


@app.route("/reject_adjustment", methods=["POST"])
@require_csrf
def reject_adjustment():
    data = request.json or {}
    request_id = data.get("request_id")
    comments = (data.get("comments") or "").strip()

    if not request_id:
        return jsonify({"error": "Missing request_id"}), 400

    try:
        database = get_db()
        try:
            adj = database["adjustments"].find_one({"_id": ObjectId(request_id)})
        except Exception:
            return jsonify({"error": "Adjustment request not found"}), 404
        if not adj:
            return jsonify({"error": "Adjustment request not found"}), 404

        status_key = str(adj.get("status") or "").strip().lower()
        if status_key not in {"pending", "parsed"}:
            return jsonify({"error": "This adjustment request has already been resolved"}), 409

        result = database["adjustments"].update_one(
            {"_id": adj["_id"], "status": adj.get("status")},
            {"$set": {
                "status": "Rejected",
                "resolved_at": datetime.utcnow().isoformat() + "Z",
                "resolved_by": session.get("user_id"),
                "comments": comments,
            }},
        )
        if result.modified_count != 1:
            return jsonify({"error": "This adjustment request was already resolved by another coordinator"}), 409
        return jsonify({"success": True, "message": "Adjustment request rejected."})
    except Exception:
        return safe_server_error()


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

        try:
            history = derive_history(user.get("teacher_id"), database=database)
        except Exception:
            history = user.get("history", [])

        user_email = user.get("email", f"{user.get('teacher_id')}@pict.edu")
        return jsonify({
            "name"      : user.get("name"),
            "role"      : user.get("role"),
            "department": user.get("department"),
            "is_admin"  : user.get("is_admin", False),
            "history"   : history,
            "identifier": user_email,
            "teacher_id": user.get("teacher_id"),
            "last_role" : user.get("last_role", "N/A"),
            "has_served_high_role": user.get("has_served_high_role"),
            "max_duties": user.get("max_duties", MAX_DUTIES),
        })
    except Exception as e:
        return safe_server_error()


@app.route("/teacher/requests", methods=["GET"])
def get_teacher_requests():
    """Return only the signed-in teacher's change requests and their decisions."""
    teacher_id = str(session.get("user_id") or "").strip()
    if not teacher_id:
        return jsonify({"error": "Authentication required"}), 401

    try:
        database = get_db()
        teacher = database["teachers"].find_one({
            "$or": [
                {"teacher_id": teacher_id},
                {"email": str(session.get("identifier") or "").lower()},
            ]
        })
        if not teacher:
            return jsonify({"error": "Teacher not found"}), 404

        teacher_id = str(teacher.get("teacher_id") or teacher_id)
        requests = list(database["adjustments"].find({
            "$or": [{"teacher_id": teacher_id}, {"teacher_id": str(teacher.get("email") or "").lower()}]
        }).sort("created_at", -1))

        partner_ids = {
            str(row.get("swap_teacher_id"))
            for row in requests
            if row.get("swap_teacher_id")
        }
        partners = {
            str(doc.get("teacher_id")): doc.get("name", str(doc.get("teacher_id")))
            for doc in database["teachers"].find(
                {"teacher_id": {"$in": list(partner_ids)}},
                {"teacher_id": 1, "name": 1},
            )
        }

        payload = []
        for row in requests:
            status_key, status_meta = adjustment_status_meta(row.get("status"))
            partner_id = str(row.get("swap_teacher_id") or "")
            partner_name = partners.get(partner_id) if partner_id else None
            new_date = row.get("new_date")
            new_slot = row.get("new_slot")

            if status_key in {"approved", "applied"}:
                if (row.get("action") == "swap" or partner_id) and new_date and new_slot is not None:
                    resolution = (
                        f"Approved — swapped with {partner_name or partner_id or 'another teacher'}. "
                        f"New duty: {new_date}, slot {new_slot}."
                    )
                elif new_date and new_slot is not None:
                    resolution = f"Approved — moved to {new_date}, slot {new_slot}."
                else:
                    resolution = "Approved by the examination coordinator."
            elif status_key in {"rejected", "skipped"}:
                resolution = (row.get("comments") or "").strip() or "Rejected by the examination coordinator."
            elif status_key == "processing":
                resolution = "The coordinator is reviewing this request."
            else:
                resolution = "Awaiting coordinator review."

            payload.append({
                "request_id": str(row["_id"]),
                "status": row.get("status"),
                "status_key": status_key,
                "status_label": status_meta["label"],
                "status_tone": status_meta["tone"],
                "current_date": row.get("current_date"),
                "current_slot": str(row.get("current_slot") or ""),
                "current_session": row.get("current_session"),
                "reason": row.get("reason") or "",
                "comments": row.get("comments") or "",
                "resolution": resolution,
                "action": row.get("action"),
                "new_date": new_date,
                "new_slot": str(new_slot) if new_slot is not None else None,
                "swap_teacher_id": partner_id or None,
                "swap_teacher_name": partner_name,
                "created_at": row.get("created_at"),
                "resolved_at": row.get("resolved_at"),
                "file_url": f"/uploads/{row['file_path']}" if row.get("file_path") else None,
            })

        return jsonify({"teacher_id": teacher_id, "requests": payload})
    except Exception:
        return safe_server_error()


def _active_roster_data():
    database = get_db()
    timetable = get_active_timetable(database=database)
    if not timetable:
        return None, None, None, None
    timetable_id = str(timetable["_id"])
    assignments = list(database["assignments"].find({
        "timetable_id": timetable_id,
        "active": True,
    }).sort([("date", 1), ("slot", 1), ("role", 1)]))
    teachers = list(database["teachers"].find({}, {"password_hash": 0}))
    return database, timetable, assignments, teachers


def _duty_payload(row):
    return {
        "duty_id": row.get("duty_id"),
        "date": row.get("date"),
        "slot": str(row.get("slot")),
        "session": row.get("session") or "TBD",
        "course_id": row.get("course_id") or "ALL",
        "role": normalize_role(row.get("role")),
        "room_assigned": row.get("room") or "TBD",
        "year": row.get("year") or "ALL",
        "department": row.get("department_required") or row.get("department") or "General",
        "pinned": bool(row.get("pinned")),
    }


@app.route("/api/teacher/duties", methods=["GET"])
def teacher_duties_api():
    identifier = (request.args.get("teacher_id") or request.args.get("identifier") or "").strip()
    if not session.get("is_admin"):
        identifier = session.get("user_id") or session.get("identifier") or ""
    if not identifier:
        return jsonify({"error": "Missing teacher identifier"}), 400

    try:
        database, timetable, assignments, teachers = _active_roster_data()
        user = database["teachers"].find_one({
            "$or": [
                {"teacher_id": identifier},
                {"teacher_id": identifier.upper()},
                {"email": identifier.lower()},
                {"email": identifier},
            ]
        })
        if not user:
            return jsonify({"error": "Teacher not found"}), 404
        teacher_id = str(user.get("teacher_id") or identifier)

        rows = [
            row for row in (assignments or [])
            if str(row.get("teacher_id")) == teacher_id
        ]
        duties = [_duty_payload(row) for row in rows]
        history = [
            {
                "slot_id": duty["slot"],
                "slot": duty["slot"],
                "exam_date": duty["date"],
                "session": duty["session"],
                "course_id": duty["course_id"],
                "year": duty["year"],
                "department": duty["department"],
                "role_assigned": duty["role"],
                "room_assigned": duty["room_assigned"],
                "assigned_at": timetable.get("confirmed_at") if timetable else None,
                "pinned": duty["pinned"],
            }
            for duty in duties
        ]
        return jsonify({
            "teacher_id": teacher_id,
            "name": user.get("name"),
            "role": user.get("role"),
            "department": user.get("department"),
            "timetable_id": str(timetable["_id"]) if timetable else None,
            "revision": int(timetable.get("revision") or 1) if timetable else None,
            "duties": duties,
            "history": history,
        })
    except Exception:
        return safe_server_error()


@app.route("/api/timetable/duties", methods=["GET"])
@require_admin
def timetable_duties_api():
    try:
        database, timetable, assignments, teachers = _active_roster_data()
        if not timetable:
            return jsonify({"error": "No published timetable found"}), 404

        teachers_by_id = {
            str(teacher.get("teacher_id")): teacher
            for teacher in teachers or []
            if teacher.get("teacher_id")
        }
        assigned = [row for row in assignments or [] if row.get("teacher_id")]
        duties = []
        for row in assigned:
            payload = _duty_payload(row)
            teacher = teachers_by_id.get(str(row.get("teacher_id")), {})
            payload.update({
                "teacher_id": str(row.get("teacher_id")),
                "teacher_name": row.get("teacher_name") or teacher.get("name", ""),
                "last_role": teacher.get("last_role", "N/A"),
                "is_priority": "Yes" if not teacher.get("has_served_high_role", True) else "No",
            })
            duties.append(payload)

        seats = []
        for row in assignments or []:
            payload = _duty_payload(row)
            payload.update({
                "teacher_id": str(row.get("teacher_id")) if row.get("teacher_id") else None,
                "teacher_name": row.get("teacher_name"),
            })
            seats.append(payload)

        return jsonify({
            "timetable_id": str(timetable["_id"]),
            "department": timetable.get("department"),
            "revision": int(timetable.get("revision") or 1),
            "duties": duties,
            "seats": seats,
        })
    except Exception:
        return safe_server_error()


@app.route("/roster/<roster_type>.pdf", methods=["GET"])
@require_admin
def download_roster_pdf(roster_type):
    try:
        roster_type = normalize_roster_type(roster_type)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404

    try:
        database, timetable, assignments, teachers = _active_roster_data()
        if not timetable:
            return jsonify({"error": "No published timetable found. Publish a timetable first."}), 404
        timetable_id = str(timetable["_id"])
        revision = count_approved_swaps(timetable_id, database=database) + 1
        context = build_roster_context(
            timetable, assignments, teachers, roster_type,
            database=database, revision=revision,
        )
        buffer = render_roster_pdf(context)
        return send_file(
            buffer,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=roster_filename(context, "pdf"),
        )
    except Exception:
        return safe_server_error()


@app.route("/roster/<roster_type>.xlsx", methods=["GET"])
@require_admin
def download_roster_excel(roster_type):
    try:
        roster_type = normalize_roster_type(roster_type)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404

    try:
        database, timetable, assignments, teachers = _active_roster_data()
        if not timetable:
            return jsonify({"error": "No published timetable found. Publish a timetable first."}), 404
        timetable_id = str(timetable["_id"])
        revision = count_approved_swaps(timetable_id, database=database) + 1
        context = build_roster_context(
            timetable, assignments, teachers, roster_type,
            database=database, revision=revision,
        )
        buffer = render_roster_excel(context)
        return send_file(
            buffer,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=roster_filename(context, "xlsx"),
        )
    except Exception:
        return safe_server_error()


@app.route("/download_confirmed_excel", methods=["GET"])
def download_confirmed_excel():
    """Directly download the currently published timetable from MongoDB as a styled Excel workbook."""
    from db import get_db
    try:
        database = get_db()
        t = get_active_timetable(database=database)
        if not t:
            return jsonify({"error": "No confirmed timetable found in MongoDB to export. Please generate and publish a timetable first."}), 404

        teacher_duties = list(database["assignments"].find({
            "timetable_id": str(t["_id"]),
            "active": True,
            "teacher_id": {"$ne": None},
        }))
        export_duties = []
        if teacher_duties:
            for row in teacher_duties:
                export_duties.append({
                    "duty_id": row.get("duty_id"),
                    "slot": row.get("slot"),
                    "date": row.get("date"),
                    "session": row.get("session"),
                    "course_id": row.get("course_id"),
                    "role": row.get("role"),
                    "teacher_id": row.get("teacher_id"),
                    "teacher_name": row.get("teacher_name"),
                    "year": row.get("year"),
                    "department": row.get("department_required"),
                    "room_assigned": row.get("room"),
                })
        else:
            export_duties = t.get("teacher_duties", [])

        tables = {
            'Timetable': t.get('timetable', []),
            'Room Allocation': t.get('room_allocation', []),
            'Teacher Duties': export_duties,
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
        return safe_server_error()


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


@app.route("/reports/faculty/<teacher_id>.<fmt>", methods=["GET"])
def download_faculty_duty_report(teacher_id, fmt):
    """Download the signed Word/Excel duty report for a faculty member."""
    fmt = str(fmt or "").lower()
    if fmt not in {"docx", "xlsx"}:
        return jsonify({"error": "Use .docx or .xlsx"}), 404
    if not session.get("user_id"):
        return jsonify({"error": "Authentication required"}), 401
    if not session.get("is_admin") and str(session.get("user_id") or "").upper() != str(teacher_id).upper():
        return jsonify({"error": "You can only download your own duty report."}), 403

    try:
        database = get_db()
        timetable = get_active_timetable(database=database)
        if not timetable:
            return jsonify({"error": "No published timetable found."}), 404
        teacher = database["teachers"].find_one({"teacher_id": teacher_id})
        if not teacher:
            teacher = database["teachers"].find_one({"teacher_id": str(teacher_id).upper()})
        if not teacher:
            return jsonify({"error": "Faculty member not found."}), 404
        report = generate_faculty_reports(database, timetable, teacher, persist=False)
        if not report:
            return jsonify({"error": "No active duties found for this faculty member."}), 404
        filename = report["filenames"][fmt]
        mime = DOCX_MIME if fmt == "docx" else XLSX_MIME
        return send_file(
            BytesIO(report["buffers"][fmt]),
            mimetype=mime,
            as_attachment=True,
            download_name=filename,
        )
    except Exception:
        return safe_server_error()


@app.route("/admin/reports/daily-change/<report_date>.<fmt>", methods=["GET"])
@require_admin
def download_daily_change_report(report_date, fmt):
    """Download the signed daily duty-change report for one examination date."""
    fmt = str(fmt or "").lower()
    if fmt not in {"docx", "xlsx"}:
        return jsonify({"error": "Use .docx or .xlsx"}), 404
    try:
        database = get_db()
        timetable = get_active_timetable(database=database)
        if not timetable:
            return jsonify({"error": "No published timetable found."}), 404
        generated = generate_daily_change_reports(
            database, timetable, [report_date], persist=True,
            generated_by=session.get("user_id"),
        )
        if not generated:
            return jsonify({"error": "No duty changes were found for that date."}), 404
        item = generated[0]
        mime = DOCX_MIME if fmt == "docx" else XLSX_MIME
        return send_file(
            BytesIO(item["buffers"][fmt]),
            mimetype=mime,
            as_attachment=True,
            download_name=item["filenames"][fmt],
        )
    except Exception:
        return safe_server_error()


@app.route("/admin/mail/timetable", methods=["POST"])
@require_admin
@require_csrf
def email_active_timetable():
    """Queue (or synchronously send) the faculty duty report emails."""
    try:
        database = get_db()
        timetable = get_active_timetable(database=database)
        if not timetable:
            return jsonify({"error": "No published timetable found."}), 404
        result = queue_timetable_emails(
            str(timetable["_id"]), database=database, triggered_by=session.get("user_id")
        )
        if not result.get("queued") and not result.get("summary"):
            return jsonify({
                "success": False,
                "error": result.get("reason") or "Email delivery could not be started.",
                "mail": result,
            }), 400
        return jsonify({"success": True, "mail": result, **result})
    except Exception:
        return safe_server_error()


@app.route("/admin/mail/test", methods=["POST"])
@require_admin
@require_csrf
def email_test_faculty_report():
    """Send one sample faculty report to the configured test recipient."""
    config = mail_config()
    if not config.get("test_recipient"):
        return jsonify({"error": "Set MAIL_TEST_RECIPIENT in backend/.env before sending a test mail."}), 400
    if not config.get("enabled") and not config.get("dry_run"):
        return jsonify({
            "error": config.get("reason") or "Mail is not configured.",
            "missing": config.get("missing") or [],
        }), 400

    try:
        database = get_db()
        timetable = get_active_timetable(database=database)
        if not timetable:
            return jsonify({"error": "No published timetable found."}), 404
        data = request.json or {}
        teacher_id = str(data.get("teacher_id") or "").strip()
        if not teacher_id:
            row = database["assignments"].find_one(
                {
                    "timetable_id": str(timetable["_id"]),
                    "active": True,
                    "teacher_id": {"$ne": None},
                },
                sort=[("teacher_id", 1)],
            )
            teacher_id = str(row.get("teacher_id")) if row else ""
        if not teacher_id:
            return jsonify({"error": "No assigned faculty member available for a test mail."}), 404
        summary = send_timetable_emails(
            str(timetable["_id"]),
            database=database,
            only_teacher_id=teacher_id,
            force_test=True,
            triggered_by=session.get("user_id"),
        )
        delivered = summary.get("sent", 0) + summary.get("dry_run", 0)
        return jsonify({
            "success": delivered > 0 and summary.get("failed", 0) == 0,
            "teacher_id": teacher_id,
            "test_recipient": config.get("test_recipient"),
            "dry_run": bool(config.get("dry_run")),
            "summary": summary,
        })
    except Exception:
        return safe_server_error()


@app.route("/admin/mail/jobs/<job_id>", methods=["GET"])
@require_admin
def mail_job_status(job_id):
    job = get_mail_job(job_id)
    if not job:
        return jsonify({"error": "Mail job not found."}), 404
    return jsonify(job)


@app.route("/admin/mail/status", methods=["GET"])
@require_admin
def mail_delivery_status():
    try:
        limit = request.args.get("limit", 25)
        return jsonify(get_mail_status(limit=limit))
    except Exception:
        return safe_server_error()


if __name__ == "__main__":
    print("Running Flask server...")
    print("   -> Open http://127.0.0.1:5000 in your browser")
    app.run(debug=False, port=5000)