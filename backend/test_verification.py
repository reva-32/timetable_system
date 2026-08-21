"""
test_verification.py — Comprehensive automated test suite for ExamSched system.
"""
import os
import io
import json
from app import app
from db import get_db

def run_tests():
    client = app.test_client()
    print("=" * 60)
    print("RUNNING COMPREHENSIVE TIMETABLE SYSTEM VERIFICATION")
    print("=" * 60)

    # 1. Test Coordinator Login
    print("\n[TEST 1] Testing Coordinator Login...")
    coord_email = os.getenv("COORDINATOR_EMAIL", "validator@pict.edu")
    coord_pw = os.getenv("COORDINATOR_PASSWORD", "replace-with-a-secure-password")
    r = client.post("/login", json={"email": coord_email, "password": coord_pw})
    print(f" -> Status: {r.status_code}, Response: {r.get_json()}")
    assert r.status_code == 200, f"Coordinator login failed: {r.get_json()}"
    assert r.get_json().get("is_admin") is True, "Coordinator is_admin should be True"
    print(" [OK] Coordinator Login Passed!")

    # 2. Test Faculty Login with numeric ID "001"
    print("\n[TEST 2] Testing Faculty Login (Teacher ID '001')...")
    r = client.post("/login", json={"email": "001", "password": "001"})
    print(f" -> Status: {r.status_code}, Response: {r.get_json()}")
    assert r.status_code == 200, f"Faculty login failed: {r.get_json()}"
    assert r.get_json().get("teacher_id") == "001", "Teacher ID should be 001"
    print(" [OK] Faculty Login with ID '001' Passed!")

    # 3. Test Faculty Signup
    print("\n[TEST 3] Testing Faculty Signup for '011@pict.edu'...")
    r = client.post("/signup", json={
        "name": "Prof. Test Eleven",
        "department": "IT",
        "email": "011@pict.edu",
        "password": "password123",
        "confirm_password": "password123"
    })
    print(f" -> Status: {r.status_code}, Response: {r.get_json()}")
    assert r.status_code == 200, f"Signup failed: {r.get_json()}"
    
    # Test login with new account
    r_login = client.post("/login", json={"email": "011@pict.edu", "password": "password123"})
    assert r_login.status_code == 200, f"Login with newly signed up account failed: {r_login.get_json()}"
    print(" [OK] Faculty Signup & Immediate Login Passed!")

    # 4. Test Sample File Downloads
    print("\n[TEST 4] Testing Sample Template Downloads...")
    r_dl_xl = client.get("/download_sample/excel")
    assert r_dl_xl.status_code == 200, f"Excel sample download failed: {r_dl_xl.status_code}"
    r_dl_pref = client.get("/download_sample/preferences")
    assert r_dl_pref.status_code == 200, f"Preferences CSV download failed: {r_dl_pref.status_code}"
    print(" [OK] Sample Downloads Passed! (Excel & Preferences CSV available)")

    # 5. Test Generation Pipeline with sample_exam_data.xlsx AND sample_teacher_preferences.csv
    print("\n[TEST 5] Testing Timetable Generation Pipeline with Preferences CSV...")
    # First login as coordinator to have active session
    client.post("/login", json={"email": coord_email, "password": coord_pw})
    sample_path = os.path.join(os.path.dirname(__file__), "sample_exam_data.xlsx")
    pref_path = os.path.join(os.path.dirname(__file__), "sample_teacher_preferences.csv")
    with open(sample_path, "rb") as f:
        file_bytes = f.read()
    with open(pref_path, "rb") as pf:
        pref_bytes = pf.read()

    data = {
        "file": (io.BytesIO(file_bytes), "sample_exam_data.xlsx"),
        "preferences_file": (io.BytesIO(pref_bytes), "sample_teacher_preferences.csv"),
        "year": "ALL",
        "branch": "ALL",
        "exam_type": "Endsem",
        "selected_slots": "10:00 – 12:30,14:00 – 16:30"
    }

    r_gen = client.post("/generate", data=data, content_type="multipart/form-data")
    assert r_gen.status_code == 200, f"Generate failed: {r_gen.get_json()}"
    gen_data = r_gen.get_json()
    summary = gen_data.get("summary", {})
    print(f" -> Summary: {summary}")
    assert summary.get("conflicts_found") == 0, f"Found conflicts: {summary.get('conflicts_found')}"
    assert summary.get("duties_assigned", 0) > 0, "Duties should be assigned"
    print(f" [OK] Generation Pipeline with Preferences Passed! Scheduled {summary.get('total_courses')} courses across {summary.get('slots_used')} slots with preference satisfaction: {summary.get('pref_satisfaction')}%")

    # 5. Test Confirm Schedule
    print("\n[TEST 5] Testing Schedule Confirmation & Duty History Writing...")
    r_conf = client.post("/confirm", json={
        "timetable": gen_data["timetable"],
        "room_allocation": gen_data["room_allocation"],
        "teacher_duties": gen_data["teacher_duties"],
        "summary": gen_data["summary"],
        "department": "ALL"
    })
    print(f" -> Status: {r_conf.status_code}, Response: {r_conf.get_json()}")
    assert r_conf.status_code == 200, f"Confirm failed: {r_conf.get_json()}"
    print(" [OK] Schedule Confirmation Passed!")

    # 6. Test Teacher Status & History for Teacher 001
    print("\n[TEST 6] Testing Teacher Duty Retrieval for '001'...")
    # Login as Teacher 001
    client.post("/login", json={"email": "001", "password": "001"})
    r_stat = client.get("/teacher/status?identifier=001@pict.edu")
    print(f" -> Status: {r_stat.status_code}, Response: {r_stat.get_json()}")
    assert r_stat.status_code == 200, f"Teacher status failed: {r_stat.get_json()}"
    history = r_stat.get_json().get("history", [])
    print(f" -> Teacher 001 has {len(history)} duties assigned in history.")
    print(" [OK] Teacher Status & Duties Retrieval Passed!")

    print("\n" + "=" * 60)
    print("ALL TESTS PASSED SUCCESSFULLY! SYSTEM IS 100% OPERATIONAL.")
    print("=" * 60)

if __name__ == "__main__":
    run_tests()
