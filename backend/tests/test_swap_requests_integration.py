"""End-to-end tests for faculty change requests and mutual duty swaps.

The tests use a temporary MongoDB database on the configured cluster so the
published roster in the application database is never modified.  They are
skipped when MongoDB is unavailable.
"""

import os
import shutil
import sys
import tempfile
import unittest
import uuid
from datetime import datetime
from io import BytesIO
from unittest.mock import patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from bson.objectid import ObjectId
from pymongo import MongoClient
from werkzeug.security import generate_password_hash

import app as app_module
import db as db_module
from app import app as flask_app


@unittest.skipUnless(os.getenv("MONGODB_URI") or db_module.MONGODB_URI, "MongoDB is not configured")
class SwapRequestIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        uri = os.getenv("MONGODB_URI") or db_module.MONGODB_URI
        cls.client = MongoClient(uri, serverSelectionTimeoutMS=5000)
        try:
            cls.client.admin.command("ping")
        except Exception as exc:
            raise unittest.SkipTest(f"MongoDB unavailable: {exc}")

        cls.db_name = f"TimeTable_swap_test_{uuid.uuid4().hex[:10]}"
        cls.db = cls.client[cls.db_name]

    @classmethod
    def tearDownClass(cls):
        try:
            # Atlas application users may create collections but are not
            # always allowed to drop the database itself. Remove all test
            # documents explicitly instead.
            for collection_name in cls.db.list_collection_names():
                cls.db[collection_name].delete_many({})
        finally:
            cls.client.close()

    def setUp(self):
        flask_app.config.update(TESTING=True, SESSION_COOKIE_SECURE=False)
        self.app_patch = patch.object(app_module, "get_db", return_value=self.db)
        self.db_patch = patch.object(db_module, "get_db", return_value=self.db)
        self.app_patch.start()
        self.db_patch.start()
        self.addCleanup(self.app_patch.stop)
        self.addCleanup(self.db_patch.stop)

        self.upload_dir = tempfile.mkdtemp(prefix="swap-test-")
        self.old_upload_dir = flask_app.config["UPLOAD_FOLDER"]
        flask_app.config["UPLOAD_FOLDER"] = self.upload_dir
        self.addCleanup(self._restore_upload_dir)

        self.report_dir = tempfile.mkdtemp(prefix="report-test-")
        self.report_env_patch = patch.dict(
            os.environ,
            {
                "REPORT_OUTPUT_DIR": self.report_dir,
                "MAIL_ENABLED": "false",
                "MAIL_DRY_RUN": "false",
                "MAIL_TEST_MODE": "false",
                "SMTP_HOST": "",
                "SMTP_FROM_EMAIL": "",
            },
            clear=False,
        )
        self.report_env_patch.start()
        self.addCleanup(lambda: shutil.rmtree(self.report_dir, ignore_errors=True))
        self.addCleanup(self.report_env_patch.stop)

        self.http = flask_app.test_client()
        self.db["teachers"].delete_many({})
        self.db["assignments"].delete_many({})
        self.db["timetables"].delete_many({})
        self.db["adjustments"].delete_many({})
        self.db["swap_log"].delete_many({})
        self._seed_roster()

    def _restore_upload_dir(self):
        flask_app.config["UPLOAD_FOLDER"] = self.old_upload_dir
        shutil.rmtree(self.upload_dir, ignore_errors=True)

    def _seed_roster(self):
        self.db["teachers"].insert_many([
            {
                "teacher_id": "ADMIN",
                "email": "admin@pict.edu",
                "name": "Test Coordinator",
                "role": "Coordinator",
                "department": "IT",
                "is_admin": True,
                "password_hash": generate_password_hash("Admin@123!"),
            },
            {
                "teacher_id": "T1",
                "email": "t1@pict.edu",
                "name": "Teacher One",
                "role": "Junior",
                "department": "IT",
                "teaching_years": [],
                "preferred_slots": [],
                "password_hash": generate_password_hash("t1"),
            },
            {
                "teacher_id": "T2",
                "email": "t2@pict.edu",
                "name": "Teacher Two",
                "role": "Junior",
                "department": "IT",
                "teaching_years": [],
                "preferred_slots": [],
                "password_hash": generate_password_hash("t2"),
            },
        ])

        timetable_id = ObjectId()
        duties = [
            {
                "duty_id": "d1", "date": "01-Mar-2026", "slot": "1",
                "session": "Morning", "course_id": "C1", "role": "Junior",
                "teacher_id": "T1", "teacher_name": "Teacher One",
                "year": "SY", "department": "IT", "room_assigned": "R1",
            },
            {
                "duty_id": "d2", "date": "02-Mar-2026", "slot": "2",
                "session": "Afternoon", "course_id": "C2", "role": "Junior",
                "teacher_id": "T2", "teacher_name": "Teacher Two",
                "year": "TY", "department": "IT", "room_assigned": "R2",
            },
        ]
        self.timetable_id = timetable_id
        self.db["timetables"].insert_one({
            "_id": timetable_id,
            "department": "IT",
            "revision": 1,
            "confirmed_at": "2026-02-20T10:00:00Z",
            "teacher_duties": duties,
            "all_duties": [dict(duty) for duty in duties],
        })
        self.db["assignments"].insert_many([
            {
                "timetable_id": str(timetable_id),
                "duty_id": duty["duty_id"],
                "date": duty["date"],
                "slot": duty["slot"],
                "session": duty["session"],
                "course_id": duty["course_id"],
                "role": duty["role"],
                "teacher_id": duty["teacher_id"],
                "teacher_name": duty["teacher_name"],
                "year": duty["year"],
                "allowed_years": [duty["year"]],
                "department_required": duty["department"],
                "department": duty["department"],
                "room": duty["room_assigned"],
                "active": True,
                "version": 1,
                "confirmed_at": "2026-02-20T10:00:00Z",
                "created_at": "2026-02-20T10:00:00Z",
                "updated_at": "2026-02-20T10:00:00Z",
            }
            for duty in duties
        ])

    def _login(self, email, password):
        response = self.http.post("/login", json={"email": email, "password": password})
        self.assertEqual(response.status_code, 200, response.get_json())
        return response

    def _csrf(self):
        response = self.http.get("/csrf-token")
        self.assertEqual(response.status_code, 200)
        return response.get_json()["csrf_token"]

    def _submit_request(self):
        self._login("t1@pict.edu", "t1")
        token = self._csrf()
        response = self.http.post(
            "/request_adjustment",
            headers={"X-CSRF-Token": token},
            data={
                "teacher_id": "T1",
                "current_date": "01-Mar-2026",
                "current_slot": "1",
                "current_session": "Morning",
                "reason": "Medical appointment",
                "file": (BytesIO(b"%PDF-1.4 test proof"), "proof.pdf"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        request_doc = self.db["adjustments"].find_one({"teacher_id": "T1"})
        self.assertIsNotNone(request_doc)
        self.assertEqual(request_doc["status"], "Pending")
        return request_doc

    def _login_admin(self):
        self._login("admin@pict.edu", "Admin@123!")
        return self._csrf()

    def test_valid_swap_is_listed_approved_and_visible_to_teacher(self):
        request_doc = self._submit_request()
        token = self._login_admin()

        response = self.http.get(f"/adjustments/alternatives?request_id={request_doc['_id']}")
        self.assertEqual(response.status_code, 200, response.get_json())
        alternatives = response.get_json()
        self.assertEqual(len(alternatives["swap_options"]), 1)
        option = alternatives["swap_options"][0]
        self.assertEqual(option["teacher_id"], "T2")
        self.assertEqual(option["date"], "02-Mar-2026")
        self.assertEqual(option["slot"], "2")

        response = self.http.post(
            "/approve_adjustment",
            headers={"X-CSRF-Token": token},
            json={
                "request_id": str(request_doc["_id"]),
                "type": "swap",
                "swap_teacher_id": "T2",
                "new_date": "02-Mar-2026",
                "new_slot": "2",
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        approval = response.get_json()
        self.assertTrue(approval["success"])
        self.assertGreaterEqual(approval["daily_reports"]["generated"], 2)
        self.assertFalse(approval["mail_queued"])

        first = self.db["assignments"].find_one({"duty_id": "d1", "active": True})
        second = self.db["assignments"].find_one({"duty_id": "d2", "active": True})
        self.assertEqual(first["teacher_id"], "T2")
        self.assertEqual(second["teacher_id"], "T1")

        saved = self.db["adjustments"].find_one({"_id": request_doc["_id"]})
        self.assertEqual(saved["status"], "Approved")
        self.assertEqual(saved["action"], "swap")
        self.assertEqual(saved["new_date"], "02-Mar-2026")
        self.assertEqual(saved["new_slot"], "2")
        self.assertEqual(saved["swap_teacher_id"], "T2")

        daily_docx = self.http.get("/admin/reports/daily-change/01-Mar-2026.docx")
        self.assertEqual(daily_docx.status_code, 200, daily_docx.data[:200])
        self.assertTrue(daily_docx.data.startswith(b"PK"))
        daily_xlsx = self.http.get("/admin/reports/daily-change/02-Mar-2026.xlsx")
        self.assertEqual(daily_xlsx.status_code, 200, daily_xlsx.data[:200])
        self.assertTrue(daily_xlsx.data.startswith(b"PK"))

        self._login("t1@pict.edu", "t1")
        response = self.http.get("/teacher/requests")
        self.assertEqual(response.status_code, 200, response.get_json())
        requests = response.get_json()["requests"]
        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0]["status_label"], "Approved")
        self.assertIn("Teacher Two", requests[0]["resolution"])
        self.assertIn("02-Mar-2026", requests[0]["resolution"])

        faculty_docx = self.http.get("/reports/faculty/T1.docx")
        self.assertEqual(faculty_docx.status_code, 200, faculty_docx.data[:200])
        self.assertTrue(faculty_docx.data.startswith(b"PK"))
        faculty_xlsx = self.http.get("/reports/faculty/T1.xlsx")
        self.assertEqual(faculty_xlsx.status_code, 200, faculty_xlsx.data[:200])
        self.assertTrue(faculty_xlsx.data.startswith(b"PK"))

        self._login("t2@pict.edu", "t2")
        forbidden = self.http.get("/reports/faculty/T1.docx")
        self.assertEqual(forbidden.status_code, 403, forbidden.get_json())

    def test_rejected_request_is_visible_with_coordinator_note(self):
        request_doc = self._submit_request()
        token = self._login_admin()

        response = self.http.post(
            "/reject_adjustment",
            headers={"X-CSRF-Token": token},
            json={
                "request_id": str(request_doc["_id"]),
                "comments": "No eligible replacement is available.",
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())

        self._login("t1@pict.edu", "t1")
        response = self.http.get("/teacher/requests")
        self.assertEqual(response.status_code, 200, response.get_json())
        request = response.get_json()["requests"][0]
        self.assertEqual(request["status_label"], "Rejected")
        self.assertIn("No eligible replacement", request["resolution"])

    def test_pending_request_is_under_scrutiny_and_not_visible_to_other_teacher(self):
        request_doc = self._submit_request()
        self.assertEqual(request_doc["status"], "Pending")

        self._login("t1@pict.edu", "t1")
        response = self.http.get("/teacher/requests")
        self.assertEqual(response.status_code, 200, response.get_json())
        request = response.get_json()["requests"][0]
        self.assertEqual(request["status_label"], "Under Scrutiny")

        self._login("t2@pict.edu", "t2")
        response = self.http.get("/teacher/requests")
        self.assertEqual(response.status_code, 200, response.get_json())
        self.assertEqual(response.get_json()["requests"], [])

    def test_no_option_response_explains_the_reason(self):
        request_doc = self._submit_request()
        self.db["assignments"].delete_one({"duty_id": "d2"})
        self._login_admin()

        response = self.http.get(f"/adjustments/alternatives?request_id={request_doc['_id']}")
        self.assertEqual(response.status_code, 200, response.get_json())
        alternatives = response.get_json()
        self.assertEqual(alternatives["swap_options"], [])
        self.assertEqual(alternatives["diagnostics"]["same_role_duties"], 0)
        self.assertIn("No other Junior duty exists", alternatives["message"])


    def test_mail_queue_reports_missing_smtp_configuration(self):
        token = self._login_admin()
        response = self.http.post("/admin/mail/timetable", headers={"X-CSRF-Token": token})
        self.assertEqual(response.status_code, 400, response.get_json())
        self.assertIn("Mail is disabled", response.get_json()["error"])

        status = self.http.get("/admin/mail/status")
        self.assertEqual(status.status_code, 200, status.get_json())
        self.assertFalse(status.get_json()["config"]["enabled"])

        self._login("t1@pict.edu", "t1")
        forbidden = self.http.get("/admin/mail/status")
        self.assertEqual(forbidden.status_code, 403, forbidden.get_json())


if __name__ == "__main__":
    unittest.main()
