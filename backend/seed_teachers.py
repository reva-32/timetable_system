"""
seed_teachers.py — Populate MongoDB with test faculty records (IDs 001 to 010).

Usage:
    cd backend
    python seed_teachers.py
"""

import os
from dotenv import load_dotenv
from pymongo import MongoClient
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME     = os.getenv("DB_NAME", "dhairya_jotwani")

if not MONGODB_URI:
    print("ERROR: MONGODB_URI not found in .env")
    exit(1)

client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
db = client[DB_NAME]
teachers_col = db["teachers"]

# Define 10 test teachers with balanced roles and departments
test_teachers = [
    {
        "teacher_id": "001",
        "name": "Prof. Aarav Sharma",
        "email": "001@pict.edu",
        "department": "IT",
        "role": "Junior",
        "preferred_slots": [1, 2],
    },
    {
        "teacher_id": "002",
        "name": "Dr. Bhavna Patil",
        "email": "002@pict.edu",
        "department": "IT",
        "role": "Junior",
        "preferred_slots": [2, 3],
    },
    {
        "teacher_id": "003",
        "name": "Prof. Chetan Kulkarni",
        "email": "003@pict.edu",
        "department": "CE",
        "role": "Junior",
        "preferred_slots": [1, 3],
    },
    {
        "teacher_id": "004",
        "name": "Dr. Deepa Deshmukh",
        "email": "004@pict.edu",
        "department": "CE",
        "role": "Junior",
        "preferred_slots": [2],
    },
    {
        "teacher_id": "005",
        "name": "Prof. Eshan Joshi",
        "email": "005@pict.edu",
        "department": "ENTC",
        "role": "Junior",
        "preferred_slots": [1],
    },
    {
        "teacher_id": "006",
        "name": "Dr. Fatima Shaikh",
        "email": "006@pict.edu",
        "department": "AIDS",
        "role": "Junior",
        "preferred_slots": [2, 3],
    },
    {
        "teacher_id": "007",
        "name": "Prof. Girish Mehta",
        "email": "007@pict.edu",
        "department": "IT",
        "role": "Senior",
        "preferred_slots": [1, 2],
    },
    {
        "teacher_id": "008",
        "name": "Dr. Hemlata Rao",
        "email": "008@pict.edu",
        "department": "CE",
        "role": "Senior",
        "preferred_slots": [2],
    },
    {
        "teacher_id": "009",
        "name": "Prof. Ishaan Verma",
        "email": "009@pict.edu",
        "department": "ENTC",
        "role": "Squad",
        "preferred_slots": [1],
    },
    {
        "teacher_id": "010",
        "name": "Dr. Jayant Nambiar",
        "email": "010@pict.edu",
        "department": "AIDS",
        "role": "Squad",
        "preferred_slots": [1, 2],
    },
]

def seed():
    print(f"Connecting to database: {DB_NAME}...")
    seeded_count = 0
    updated_count = 0

    for t in test_teachers:
        t_id = t["teacher_id"]
        email = t["email"].strip().lower()
        default_password = t_id.strip().lower()
        doc = {
            "teacher_id": t_id,
            "email": email,
            "name": t["name"],
            "role": t["role"],
            "department": t["department"],
            "preferred_slots": t.get("preferred_slots", []),
            "teaching_years": [],
            "has_served_high_role": False,
            "duty_counts": {"squad": 0, "junior": 0, "senior": 0},
            "last_role": "N/A",
            "history": [],
            "password_hash": generate_password_hash(default_password),
            "password_changed": False,
            "is_admin": False
        }

        existing = teachers_col.find_one({"$or": [{"teacher_id": t_id}, {"email": email}]})
        if not existing:
            teachers_col.insert_one(doc)
            print(f" [+] Inserted Teacher {t_id}: {t['name']} ({t['role']} - {t['department']})")
            seeded_count += 1
        else:
            # Update fields while ensuring proper defaults
            update_data = {
                "teacher_id": t_id,
                "email": email,
                "name": t["name"],
                "role": t["role"],
                "department": t["department"],
                "preferred_slots": t.get("preferred_slots", []),
            }
            if "duty_counts" not in existing:
                update_data["duty_counts"] = {"squad": 0, "junior": 0, "senior": 0}
            if "history" not in existing:
                update_data["history"] = []
            if not existing.get("password_hash"):
                update_data["password_hash"] = generate_password_hash(default_password)
                update_data["password_changed"] = False

            teachers_col.update_one({"_id": existing["_id"]}, {"$set": update_data})
            print(f" [~] Updated Teacher {t_id}: {t['name']}")
            updated_count += 1

    print(f"\nSeeding Complete! Seeded: {seeded_count}, Updated: {updated_count}")
    print(f"Total teachers in DB: {teachers_col.count_documents({})}")
    print("\nTeacher login email: <lowercase teacher id>@pict.edu; default password: <lowercase teacher id>.")

if __name__ == "__main__":
    seed()
