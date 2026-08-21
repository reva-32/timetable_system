"""
clean_and_ingest.py — Clean temporary data and cleanly ingest 30 teachers (T001 to T030) + coordinator into MongoDB.

Usage:
    python clean_and_ingest.py
"""
import os
import glob
from pymongo import MongoClient
from dotenv import load_dotenv
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "dhairya_jotwani")

# Teacher roster matching the user's dataset
TEACHERS_ROSTER = [
    {"teacher_id": "T001", "name": "Prof. IT_1", "role": "Squad"},
    {"teacher_id": "T002", "name": "Prof. IT_2", "role": "Squad"},
    {"teacher_id": "T003", "name": "Prof. IT_3", "role": "Junior"},
    {"teacher_id": "T004", "name": "Prof. IT_4", "role": "Junior"},
    {"teacher_id": "T005", "name": "Prof. IT_5", "role": "Senior"},
    {"teacher_id": "T006", "name": "Prof. IT_6", "role": "Squad"},
    {"teacher_id": "T007", "name": "Prof. IT_7", "role": "Senior"},
    {"teacher_id": "T008", "name": "Prof. IT_8", "role": "Junior"},
    {"teacher_id": "T009", "name": "Prof. IT_9", "role": "Junior"},
    {"teacher_id": "T010", "name": "Prof. IT_10", "role": "Squad"},
    {"teacher_id": "T011", "name": "Prof. IT_11", "role": "Squad"},
    {"teacher_id": "T012", "name": "Prof. IT_12", "role": "Squad"},
    {"teacher_id": "T013", "name": "Prof. IT_13", "role": "Squad"},
    {"teacher_id": "T014", "name": "Prof. IT_14", "role": "Junior"},
    {"teacher_id": "T015", "name": "Prof. IT_15", "role": "Senior"},
    {"teacher_id": "T016", "name": "Prof. IT_16", "role": "Squad"},
    {"teacher_id": "T017", "name": "Prof. IT_17", "role": "Junior"},
    {"teacher_id": "T018", "name": "Prof. IT_18", "role": "Junior"},
    {"teacher_id": "T019", "name": "Prof. IT_19", "role": "Senior"},
    {"teacher_id": "T020", "name": "Prof. IT_20", "role": "Senior"},
    {"teacher_id": "T021", "name": "Prof. IT_21", "role": "Senior"},
    {"teacher_id": "T022", "name": "Prof. IT_22", "role": "Junior"},
    {"teacher_id": "T023", "name": "Prof. IT_23", "role": "Squad"},
    {"teacher_id": "T024", "name": "Prof. IT_24", "role": "Squad"},
    {"teacher_id": "T025", "name": "Prof. IT_25", "role": "Squad"},
    {"teacher_id": "T026", "name": "Prof. IT_26", "role": "Junior"},
    {"teacher_id": "T027", "name": "Prof. IT_27", "role": "Junior"},
    {"teacher_id": "T028", "name": "Prof. IT_28", "role": "Squad"},
    {"teacher_id": "T029", "name": "Prof. IT_29", "role": "Squad"},
    {"teacher_id": "T030", "name": "Prof. IT_30", "role": "Junior"},
]

def clean_and_ingest():
    print("=" * 60)
    print("Starting Database Cleanup & Ingestion...")
    print("=" * 60)

    if not MONGODB_URI:
        print("Error: MONGODB_URI not found in .env")
        return False

    client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
    db = client[DB_NAME]

    # 1. Clean stale collections
    print("Clearing timetables and adjustments collections...")
    db["timetables"].delete_many({})
    db["adjustments"].delete_many({})

    # 2. Clean temporary files
    print("Cleaning temporary upload files...")
    temp_patterns = [
        os.path.join(BASE_DIR, "temp_data.xlsx"),
        os.path.join(BASE_DIR, "temp_preferences*"),
        os.path.join(BASE_DIR, "uploads", "*")
    ]
    for pattern in temp_patterns:
        for f in glob.glob(pattern):
            try:
                if os.path.isfile(f):
                    os.remove(f)
                    print(f"   Deleted temp file: {os.path.basename(f)}")
            except Exception as e:
                print(f"   Could not delete {f}: {e}")

    # 3. Ensure Coordinator exists
    coord_email = (os.getenv("COORDINATOR_EMAIL") or "validator@pict.edu").strip().lower()
    coord_pw = os.getenv("COORDINATOR_PASSWORD") or "validator123"
    coord_name = os.getenv("COORDINATOR_NAME", "Examination Coordinator")
    coord_dept = os.getenv("COORDINATOR_DEPARTMENT", "IT")
    coord_id = coord_email.split("@")[0].upper()

    print(f"Ingesting Coordinator: {coord_name} ({coord_email})...")
    db["teachers"].update_one(
        {"email": coord_email},
        {"$set": {
            "teacher_id": coord_id,
            "email": coord_email,
            "name": coord_name,
            "role": "Coordinator",
            "department": coord_dept,
            "is_admin": True,
            "has_served_high_role": True,
            "duty_counts": {"squad": 0, "junior": 0, "senior": 0},
            "last_role": "N/A",
            "history": [],
            "password_hash": generate_password_hash(coord_pw),
            "password_changed": True
        }},
        upsert=True
    )

    # 4. Ingest 30 IT faculty members
    print(f"Ingesting {len(TEACHERS_ROSTER)} Faculty members (T001 to T030)...")
    for f in TEACHERS_ROSTER:
        t_id = f["teacher_id"]
        email = f"{t_id.lower()}@pict.edu"
        db["teachers"].update_one(
            {"$or": [{"teacher_id": t_id}, {"email": email}, {"teacher_id": t_id.lower()}]},
            {"$set": {
                "teacher_id": t_id,
                "email": email,
                "name": f["name"],
                "department": "IT",
                "role": f["role"],
                "preferred_slots": [1, 2, 3],
                "has_served_high_role": False,
                "duty_counts": {"squad": 0, "junior": 0, "senior": 0},
                "last_role": "N/A",
                "history": [],
                "is_admin": False,
                "password_hash": generate_password_hash(t_id),
                "password_changed": True
            }},
            upsert=True
        )

    # Clean old non-matching teachers
    allowed_emails = {f"{f['teacher_id'].lower()}@pict.edu" for f in TEACHERS_ROSTER} | {coord_email}
    allowed_ids = {f["teacher_id"] for f in TEACHERS_ROSTER} | {coord_id}
    removed = db["teachers"].delete_many({
        "email": {"$nin": list(allowed_emails)},
        "teacher_id": {"$nin": list(allowed_ids)}
    })
    if removed.deleted_count > 0:
        print(f"   Removed {removed.deleted_count} obsolete teacher records.")

    print("=" * 60)
    print("SUCCESS: Ingested exactly 30 Faculty members + Coordinator!")
    print(f"   Total Faculty in DB: {db['teachers'].count_documents({'is_admin': {'$ne': True}})}")
    print(f"   Total Coordinators in DB: {db['teachers'].count_documents({'is_admin': True})}")
    print(f"   Active Timetables in DB: {db['timetables'].count_documents({})}")
    print(f"   Active Adjustments in DB: {db['adjustments'].count_documents({})}")
    print("=" * 60)
    return True

if __name__ == "__main__":
    clean_and_ingest()
