"""
migrate_users.py — One-time migration script.

Run this ONCE to update existing teacher documents in MongoDB:
  1. Converts email from bare teacher_id (e.g. "T1") to "T1@pict.edu"
  2. Sets password_hash = hash(teacher_id) for any teacher missing a password

Usage:
    cd backend
    python migrate_users.py

Safe to run multiple times — already-migrated records are skipped.
"""

import os
from dotenv import load_dotenv
from pymongo import MongoClient
from werkzeug.security import generate_password_hash

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME     = os.getenv("DB_NAME", "TimeTable")

if not MONGODB_URI:
    print("ERROR: MONGODB_URI not set in .env")
    exit(1)

client   = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
database = client[DB_NAME]
col      = database["teachers"]

print(f"Connected to: {DB_NAME}")
print(f"Total teacher documents: {col.count_documents({})}\n")

updated_email    = 0
updated_password = 0
skipped          = 0

for doc in col.find():
    doc_id     = doc["_id"]
    teacher_id = doc.get("teacher_id", "")
    email      = doc.get("email", "")
    has_hash   = bool(doc.get("password_hash"))

    # Decide the canonical email
    if email.endswith("@pict.edu"):
        new_email = email           # already correct
        email_changed = False
    elif teacher_id:
        new_email     = f"{teacher_id}@pict.edu"
        email_changed = (new_email != email)
    else:
        # No teacher_id — derive teacher_id from email as fallback
        teacher_id    = email.split("@")[0]
        new_email     = f"{teacher_id}@pict.edu"
        email_changed = (new_email != email)

    # Decide the default password hash
    new_hash      = generate_password_hash(teacher_id) if not has_hash else None
    password_note = "(password already set)" if has_hash else f"(set default pw = '{teacher_id}')"

    if not email_changed and has_hash:
        print(f"  SKIP   {doc.get('name', teacher_id)!r:<30} email={email!r}  {password_note}")
        skipped += 1
        continue

    update = {}
    if email_changed:
        update["email"] = new_email

    if new_hash:
        update["password_hash"] = new_hash

    if update:
        col.update_one({"_id": doc_id}, {"$set": update})
        print(f"  UPDATE {doc.get('name', teacher_id)!r:<30} email: {email!r} -> {new_email!r}  {password_note}")

        if email_changed:
            updated_email += 1
        if new_hash:
            updated_password += 1

print(f"\nDone.")
print(f"  Emails updated   : {updated_email}")
print(f"  Passwords set    : {updated_password}")
print(f"  Already OK       : {skipped}")
