import os
from pymongo import MongoClient
from dotenv import load_dotenv
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME     = os.getenv("DB_NAME", "TimeTable")



client = None
db     = None


def format_date_to_standard(raw_date):
    """
    Format any date string, datetime object, or raw value to '%d-%b-%Y' (e.g., '26-Nov-2026').
    Returns 'TBD' if the date is invalid or empty.
    """
    if raw_date is None:
        return "TBD"
    
    if hasattr(raw_date, "strftime"):
        return raw_date.strftime("%d-%b-%Y")
        
    raw_str = str(raw_date).strip()
    if not raw_str or raw_str.upper() in ("NAN", "NONE", "TBD", ""):
        return "TBD"
        
    # If it's already in the correct format, return it
    try:
        datetime.strptime(raw_str, "%d-%b-%Y")
        return raw_str
    except ValueError:
        pass
        
    # Try parsing via pandas to_datetime
    try:
        import pandas as pd
        parsed = pd.to_datetime(raw_str, dayfirst=True, errors='coerce')
        if not pd.isna(parsed):
            return parsed.strftime("%d-%b-%Y")
    except Exception:
        pass
        
    return raw_str


def migrate_db_schema():
    """Migrate legacy flat 'duty_count' to structured 'duty_counts', and ensure other schema defaults exist."""
    try:
        database = get_db()
        faculty_col = database["teachers"]
        for f in faculty_col.find():
            update_fields = {}
            unset_fields = {}

            if "duty_counts" not in f:
                flat_count = f.get("duty_count", 0)
                role = f.get("role", "Junior")
                role_key = role.lower() if role.lower() in ["junior", "senior", "squad"] else "junior"
                update_fields["duty_counts"] = {
                    "squad": flat_count if role_key == "squad" else 0,
                    "junior": flat_count if role_key == "junior" else 0,
                    "senior": flat_count if role_key == "senior" else 0
                }
                if "duty_count" in f:
                    unset_fields["duty_count"] = ""

            if "has_served_high_role" not in f:
                update_fields["has_served_high_role"] = False

            if "last_role" not in f:
                update_fields["last_role"] = "N/A"

            # Normalize history dates to %d-%b-%Y
            history = f.get("history", [])
            history_changed = False
            new_history = []
            for h in history:
                exam_date = h.get("exam_date")
                if exam_date:
                    normalized = format_date_to_standard(exam_date)
                    if normalized != exam_date:
                        h["exam_date"] = normalized
                        history_changed = True
                new_history.append(h)
                
            if "history" not in f:
                update_fields["history"] = new_history
            elif history_changed:
                update_fields["history"] = new_history

            if update_fields or unset_fields:
                update_op = {}
                if update_fields:
                    update_op["$set"] = update_fields
                if unset_fields:
                    update_op["$unset"] = unset_fields
                faculty_col.update_one({"_id": f["_id"]}, update_op)

        # Normalize adjustments dates to %d-%b-%Y
        adjustments_col = database["adjustments"]
        for adj in adjustments_col.find():
            adj_changed = False
            adj_updates = {}
            for field in ["current_date", "new_date"]:
                val = adj.get(field)
                if val:
                    normalized = format_date_to_standard(val)
                    if normalized != val:
                        adj_updates[field] = normalized
                        adj_changed = True
            if adj_changed:
                adjustments_col.update_one({"_id": adj["_id"]}, {"$set": adj_updates})

        print("Schema migration complete: all legacy fields migrated and defaults set.")
    except Exception as e:
        print(f"Error during schema migration: {e}")

def get_db():
    global client, db

    if db is None:
        if not MONGODB_URI:
            raise RuntimeError(
                "MONGODB_URI is not set. "
                "Add it to your .env file: MONGODB_URI=mongodb+srv://..."
            )
        try:
            client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
            new_db = client[DB_NAME]
            new_db.command("ping")
            db = new_db
            migrate_db_schema()
        except Exception:
            client = None
            db = None
            raise

    return db

def init_faculty(teachers_data):
    """
    Initialize or update faculty records in MongoDB.
    Called once per /generate request so new teachers from the Excel
    are always upserted before duty assignment.

    teachers_data: dict { teacher_id -> {name, role, preferred_slots, ...} }

    FIX: The original code used teacher_id as 'email' (a misnomer —
    the Teachers sheet uses IDs like T1, T2, not emails). We now store
    and look up by 'teacher_id' consistently. The 'email' field is kept
    for backward-compatibility with any existing documents.
    """
    database      = get_db()
    faculty_col   = database["teachers"]

    for t_id, t_info in teachers_data.items():
        existing = faculty_col.find_one({"teacher_id": t_id})
        if not existing:
            # Also check legacy documents keyed by email == t_id
            existing = faculty_col.find_one({"email": t_id})

        if not existing:
            teacher_email = f"{t_id}@pict.edu"  # canonical email format
            faculty_col.insert_one({
                "teacher_id"          : t_id,
                "email"               : teacher_email,
                "name"                : t_info["name"],
                "role"                : t_info.get("role", "Junior"),
                "department"          : t_info.get("department", "IT"),
                "has_served_high_role": False,
                "duty_counts"         : { "squad": 0, "junior": 0, "senior": 0 },
                "last_role"           : "N/A",
                "history"             : [],
                # Default password = teacher_id (e.g. "T1"), not the full email
                "password_hash"       : generate_password_hash(t_id)
            })
        else:
            # Keep name/role in sync with Excel in case it was updated
            update_fields = {
                "teacher_id": t_id,
                "name"      : t_info["name"],
                "role"      : t_info.get("role", existing.get("role", "Junior")),
            }
            if "duty_counts" not in existing:
                flat_count = existing.get("duty_count", 0)
                role_key = update_fields["role"].lower() if update_fields["role"].lower() in ["junior", "senior", "squad"] else "junior"
                update_fields["duty_counts"] = {
                    "squad": flat_count if role_key == "squad" else 0,
                    "junior": flat_count if role_key == "junior" else 0,
                    "senior": flat_count if role_key == "senior" else 0
                }
            if "has_served_high_role" not in existing:
                update_fields["has_served_high_role"] = False
            faculty_col.update_one(
                {"_id": existing["_id"]},
                {"$set": update_fields}
            )


def get_priority_faculty():
    """Fetch faculty where has_served_high_role == False."""
    database = get_db()
    return list(database["teachers"].find({"has_served_high_role": False}))


def update_faculty_duty(teacher_id, role, date, slot_id, is_high_role=False, course_id=None, room_assigned=None, session=None):
    """
    Update faculty duty status after confirmation.

    FIX: Query by teacher_id OR email to handle both old and new documents.
    """
    database    = get_db()
    date = format_date_to_standard(date)
    role_key = role.lower() if role.lower() in ["junior", "senior", "squad"] else "junior"
    
    # Sensible defaults for room based on role if not provided
    if not room_assigned:
        if role == "Senior":
            room_assigned = "Control"
        elif role == "Squad":
            room_assigned = "Roaming"
        else:
            room_assigned = "TBD"

    update_data = {
        "$inc" : {f"duty_counts.{role_key}": 1},
        "$set" : {"last_role": role},
        "$push": {
            "history": {
                "slot_id": str(slot_id),
                "exam_date": date,
                "role_assigned": role,
                "course_id": course_id or "ALL",
                "room_assigned": room_assigned,
                "session": session or "TBD",
                "assigned_at": datetime.utcnow().isoformat() + "Z"
            }
        }
    }
    if is_high_role:
        update_data["$set"]["has_served_high_role"] = True

    # Query by teacher_id, numeric formats, and email
    query_filters = [
        {"teacher_id": str(teacher_id)},
        {"teacher_id": str(teacher_id).upper()},
        {"email": str(teacher_id).lower()},
        {"email": f"{teacher_id}@pict.edu".lower()}
    ]
    if str(teacher_id).isdigit():
        query_filters.append({"teacher_id": f"{int(teacher_id):03d}"})
        query_filters.append({"teacher_id": str(int(teacher_id))})

    database["teachers"].update_one({"$or": query_filters}, update_data)


def check_reset_fairness(department="IT"):
    """
    If everyone in the department has now served a high role,
    reset all flags so the fairness cycle restarts.
    Returns True if a reset was triggered.
    """
    database = get_db()

    count_false = database["teachers"].count_documents({
        "department"          : department,
        "has_served_high_role": False
    })

    if count_false == 0:
        database["teachers"].update_many(
            {"department": department},
            {"$set": {"has_served_high_role": False}}
        )
        return True
    return False


def get_all_faculty_status():
    """
    Fetch all faculty records as a dict keyed by teacher_id.
    Falls back to 'email' for documents that predate the teacher_id field.

    FIX: The original code always keyed by f["email"]. Since teacher IDs
    in the Teachers sheet are like T1/T2 (not emails), the lookup in
    app.py — faculty_status.get(a["teacher_id"], {}) — would always miss,
    causing every teacher to show last_role='N/A' and is_priority=False.
    We now key by teacher_id so the lookup works correctly.
    """
    database = get_db()
    result   = {}
    for f in database["teachers"].find():
        key          = f.get("teacher_id") or f.get("email", "")
        result[key]  = f
    return result


def update_password(identifier: str, new_password: str) -> bool:
    """
    Hash and store a new password for the teacher identified by email or teacher_id.
    Returns True if a matching document was found and updated, False otherwise.
    """
    database = get_db()
    new_hash = generate_password_hash(new_password)
    result = database["teachers"].update_one(
        {"$or": [{"email": identifier}, {"teacher_id": identifier}]},
        {"$set": {"password_hash": new_hash}}
    )
    return result.matched_count > 0


def ensure_coordinator():
    """
    Create the coordinator account from environment variables if it does not
    already exist. Existing coordinator passwords are not overwritten on every
    server restart.
    """
    coordinator_email = (os.getenv("COORDINATOR_EMAIL") or "").strip().lower()
    coordinator_password = os.getenv("COORDINATOR_PASSWORD") or ""

    if not coordinator_email or not coordinator_password:
        return False

    database = get_db()
    teachers = database["teachers"]
    existing = teachers.find_one({"email": coordinator_email})

    if existing:
        update = {"is_admin": True}
        if not existing.get("role"):
            update["role"] = "Coordinator"
        teachers.update_one({"_id": existing["_id"]}, {"$set": update})
        return True

    teacher_id = coordinator_email.split("@")[0].upper()
    teachers.insert_one({
        "teacher_id": teacher_id,
        "email": coordinator_email,
        "name": os.getenv("COORDINATOR_NAME", "Examination Coordinator"),
        "role": "Coordinator",
        "department": os.getenv("COORDINATOR_DEPARTMENT", "IT"),
        "is_admin": True,
        "has_served_high_role": True,
        "duty_counts": {"squad": 0, "junior": 0, "senior": 0},
        "last_role": "N/A",
        "history": [],
        "password_hash": generate_password_hash(coordinator_password),
        "password_changed": True,
    })
    return True


if __name__ == "__main__":
    print("Testing get_db()...")

    try:
        database = get_db()
        print("Connected successfully!")
        print("Collections:", database.list_collection_names())
    except Exception as e:
        import traceback
        traceback.print_exc()   