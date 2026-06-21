import os
from pymongo import MongoClient
from dotenv import load_dotenv
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME     = os.getenv("DB_NAME", "TimeTable")



client = None
db     = None


def migrate_db_schema():
    """Migrate legacy flat 'duty_count' to structured 'duty_counts'."""
    try:
        database = get_db()
        faculty_col = database["teachers"]
        for f in faculty_col.find():
            if "duty_counts" not in f:
                flat_count = f.get("duty_count", 0)
                role = f.get("role", "Junior")
                role_key = role.lower() if role.lower() in ["junior", "senior", "squad"] else "junior"
                
                duty_counts = {
                    "squad": flat_count if role_key == "squad" else 0,
                    "junior": flat_count if role_key == "junior" else 0,
                    "senior": flat_count if role_key == "senior" else 0
                }
                
                faculty_col.update_one(
                    {"_id": f["_id"]},
                    {
                        "$set": {"duty_counts": duty_counts},
                        "$unset": {"duty_count": ""}
                    }
                )
        print("Schema migration complete: all legacy duty_count fields migrated to duty_counts.")
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
            faculty_col.update_one(
                {"_id": existing["_id"]},
                {"$set": update_fields}
            )


def get_priority_faculty():
    """Fetch faculty where has_served_high_role == False."""
    database = get_db()
    return list(database["teachers"].find({"has_served_high_role": False}))


def update_faculty_duty(teacher_id, role, date, slot_id, is_high_role=False):
    """
    Update faculty duty status after confirmation.

    FIX: Query by teacher_id OR email to handle both old and new documents.
    """
    database    = get_db()
    role_key = role.lower() if role.lower() in ["junior", "senior", "squad"] else "junior"
    
    update_data = {
        "$inc" : {f"duty_counts.{role_key}": 1},
        "$set" : {"last_role": role},
        "$push": {
            "history": {
                "slot_id": str(slot_id),
                "exam_date": date,
                "role_assigned": role,
                "assigned_at": datetime.utcnow().isoformat()
            }
        }
    }
    if is_high_role:
        update_data["$set"]["has_served_high_role"] = True

    # Try the new field first, fall back to email for legacy docs
    result = database["teachers"].update_one({"teacher_id": teacher_id}, update_data)
    if result.matched_count == 0:
        database["teachers"].update_one({"email": teacher_id}, update_data)


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


if __name__ == "__main__":
    print("Testing get_db()...")

    try:
        database = get_db()
        print("Connected successfully!")
        print("Collections:", database.list_collection_names())
    except Exception as e:
        import traceback
        traceback.print_exc()   