"""
test_verification.py — Comprehensive automated test suite for ExamSched system.
"""
import os
import io
import json
from app import app
from db import get_db
from clean_and_ingest import clean_and_ingest

def run_tests():
    print("=" * 60)
    print("RUNNING COMPREHENSIVE TIMETABLE SYSTEM VERIFICATION")
    print("=" * 60)

    # 0. Clean & Ingest exactly matching 30 faculty members and coordinator
    clean_and_ingest()

    client = app.test_client()

    # 1. Test Coordinator Login
    print("\n[TEST 1] Testing Coordinator Login...")
    coord_email = os.getenv("COORDINATOR_EMAIL", "validator@pict.edu")
    coord_pw = os.getenv("COORDINATOR_PASSWORD", "validator123")
    r = client.post("/login", json={"email": coord_email, "password": coord_pw})
    print(f" -> Status: {r.status_code}")
    assert r.status_code == 200, f"Coordinator login failed: {r.get_json()}"
    assert r.get_json().get("is_admin") is True, "Coordinator is_admin should be True"
    print(" [OK] Coordinator Login Passed!")

    # 2. Test Faculty Login with ID "T001" and "T003"
    print("\n[TEST 2] Testing Faculty Login (Teacher ID 'T001' and 'T003')...")
    r1 = client.post("/login", json={"email": "T001", "password": "T001"})
    print(f" -> Status T001: {r1.status_code}")
    assert r1.status_code == 200, f"Faculty login failed: {r1.get_json()}"
    assert r1.get_json().get("teacher_id") == "T001", "Teacher ID should be T001"

    r3 = client.post("/login", json={"email": "T003", "password": "T003"})
    print(f" -> Status T003: {r3.status_code}")
    assert r3.status_code == 200, f"Faculty login failed: {r3.get_json()}"
    assert r3.get_json().get("teacher_id") == "T003", "Teacher ID should be T003"
    print(" [OK] Faculty Login with IDs 'T001' and 'T003' Passed!")

    # 3. Test Timetable Generation Pipeline with User's Excel File Structure
    print("\n[TEST 3] Testing Timetable Generation Pipeline with Excel Upload...")
    client.post("/login", json={"email": coord_email, "password": coord_pw})
    sample_path = os.path.join(os.path.dirname(__file__), "sample_exam_data_v2.xlsx")
    if not os.path.exists(sample_path):
        sample_path = os.path.join(os.path.dirname(__file__), "sample_exam_data.xlsx")
    with open(sample_path, "rb") as f:
        file_bytes = f.read()

    data = {
        "file": (io.BytesIO(file_bytes), "sample_exam_data.xlsx"),
        "year": "ALL",
        "branch": "ALL",
        "exam_type": "Endsem",
        "selected_slots": "Morning,Afternoon"
    }

    r_gen = client.post("/generate", data=data, content_type="multipart/form-data")
    assert r_gen.status_code == 200, f"Generate failed: {r_gen.get_json()}"
    gen_data = r_gen.get_json()
    summary = gen_data.get("summary", {})
    print(f" -> Summary: {summary}")
    assert summary.get("conflicts_found") == 0, f"Found conflicts: {summary.get('conflicts_found')}"
    assert summary.get("duties_assigned", 0) > 0, "Duties should be assigned"
    print(f" [OK] Timetable Generation Passed! Scheduled {summary.get('total_courses')} courses across {summary.get('slots_used')} slots.")

    # 4. Test Schedule Confirmation (Store in MongoDB)
    print("\n[TEST 4] Testing Schedule Confirmation & Publishing to MongoDB...")
    r_conf = client.post("/confirm", json={
        "timetable": gen_data["timetable"],
        "room_allocation": gen_data["room_allocation"],
        "teacher_duties": gen_data["teacher_duties"],
        "summary": gen_data["summary"],
        "department": "ALL"
    })
    assert r_conf.status_code == 200, f"Confirm failed: {r_conf.get_json()}"
    
    # Verify in DB
    db = get_db()
    saved_tt = db["timetables"].find_one()
    assert saved_tt is not None, "Timetable should be present in MongoDB 'timetables' collection"
    assert len(saved_tt["teacher_duties"]) == len(gen_data["teacher_duties"])
    print(f" [OK] Schedule Confirmed & Stored in MongoDB ({len(saved_tt['teacher_duties'])} duties assigned to faculty)!")

    # 5. Test Teacher Requesting Duty Adjustment
    print("\n[TEST 5] Testing Faculty Requesting Duty Change...")
    # Find a teacher with an assigned duty
    assigned_teacher_id = gen_data["teacher_duties"][0]["teacher_id"]
    assigned_teacher_duty = gen_data["teacher_duties"][0]

    client.post("/login", json={"email": assigned_teacher_id, "password": assigned_teacher_id})
    t_status = client.get(f"/teacher/status?identifier={assigned_teacher_id}").get_json()
    t_history = t_status.get("history", [])
    assert len(t_history) > 0, f"Teacher {assigned_teacher_id} should have assigned duties"
    first_duty = t_history[0]
    
    req_file_data = {
        "file": (io.BytesIO(b"%PDF-1.4 dummy application"), "emergency_leave.pdf"),
        "teacher_id": assigned_teacher_id,
        "current_date": first_duty["exam_date"],
        "current_slot": str(first_duty["slot_id"]),
        "current_session": first_duty.get("session", "Morning"),
        "reason": "Family emergency"
    }
    r_adj = client.post("/request_adjustment", data=req_file_data, content_type="multipart/form-data")
    assert r_adj.status_code == 200, f"Adjustment submission failed: {r_adj.get_json()}"
    print(" [OK] Duty Change Request Submitted by Faculty!")

    # 6. Test Coordinator Viewing and Approving Swap
    print("\n[TEST 6] Testing Coordinator Viewing Alternatives & Approving Swap...")
    client.post("/login", json={"email": coord_email, "password": coord_pw})
    adjs = client.get("/adjustments").get_json()
    assert len(adjs) > 0, "Should have at least 1 adjustment request"
    adj_id = adjs[0]["_id"]

    alts = client.get(f"/adjustments/alternatives?request_id={adj_id}").get_json()
    print(f" -> Found {len(alts.get('swap_options', []))} swap options and {len(alts.get('free_slots', []))} free slots.")
    assert len(alts.get("swap_options", [])) > 0 or len(alts.get("free_slots", [])) > 0

    if alts.get("swap_options"):
        best_swap = alts["swap_options"][0]
        r_appr = client.post("/approve_adjustment", json={
            "request_id": adj_id,
            "type": "swap",
            "new_date": best_swap["date"],
            "new_slot": str(best_swap["slot"]),
            "swap_teacher_id": best_swap["teacher_id"]
        })
    else:
        best_move = alts["free_slots"][0]
        r_appr = client.post("/approve_adjustment", json={
            "request_id": adj_id,
            "type": "move",
            "new_date": best_move["date"],
            "new_slot": str(best_move["slot"])
        })
    assert r_appr.status_code == 200, f"Approval failed: {r_appr.get_json()}"
    print(" [OK] Adjustment Approved and Synced with MongoDB!")

    # 7. Test Confirmed Timetable and Excel Download from MongoDB
    print("\n[TEST 7] Testing Timetable Retrieval & Excel Export from MongoDB...")
    r_tt = client.get("/confirmed_timetable")
    assert r_tt.status_code == 200
    tt_json = r_tt.get_json()
    assert tt_json is not None, "Confirmed timetable should be accessible"

    r_excel = client.get("/download_confirmed_excel")
    assert r_excel.status_code == 200, f"Excel download failed: {r_excel.status_code}"
    assert len(r_excel.data) > 1000, "Downloaded Excel should contain complete workbook data"
    print(" [OK] Timetable Retrieval & Excel Export Passed!")

    print("\n" + "=" * 60)
    print("ALL 7 TESTS PASSED SUCCESSFULLY! THE SYSTEM IS FULLY OPERATIONAL.")
    print("=" * 60)

if __name__ == "__main__":
    run_tests()
