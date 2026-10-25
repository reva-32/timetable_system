"""
reset_teacher_passwords.py — Set faculty login passwords to their Teacher ID.

For every non-admin faculty record:
  - the login email is stored in lowercase (e.g. t001@pict.edu)
  - the password is the lowercase Teacher ID (e.g. t001)

Coordinator/admin accounts are left untouched.

Usage:
    cd backend
    python reset_teacher_passwords.py
"""

import os

from dotenv import load_dotenv
from pymongo import MongoClient
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "TimeTable")

if not MONGODB_URI:
    print("ERROR: MONGODB_URI not found in .env")
    raise SystemExit(1)

client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
teachers = client[DB_NAME]["teachers"]


def main():
    updated = 0
    skipped = 0

    for teacher in teachers.find({"is_admin": {"$ne": True}}):
        teacher_id = str(teacher.get("teacher_id") or "").strip()
        if not teacher_id:
            email = str(teacher.get("email") or "").strip()
            teacher_id = email.split("@", 1)[0].strip()
        if not teacher_id:
            skipped += 1
            print(f"  SKIP   {teacher.get('_id')} (no teacher_id or email)")
            continue

        email = str(teacher.get("email") or f"{teacher_id}@pict.edu").strip().lower()
        default_password = teacher_id.lower()

        teachers.update_one(
            {"_id": teacher["_id"]},
            {
                "$set": {
                    "teacher_id": teacher_id,
                    "email": email,
                    "password_hash": generate_password_hash(default_password),
                    "password_changed": False,
                }
            },
        )
        updated += 1
        print(f"  RESET  {email:<28} password={default_password}")

    print(f"\nDone. Updated: {updated}, skipped: {skipped}")
    print("Admin/coordinator accounts were not changed.")


if __name__ == "__main__":
    main()
