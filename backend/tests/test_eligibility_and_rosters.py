import os
import sys
import unittest
from unittest.mock import patch

import numpy as np

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import db as db_module
from app import json_safe_row
from eligibility import effective_teaching_years, session_code, teacher_can_invigilate
from main import assign_teachers, build_duties, normalize_duty_config
from rosters import build_roster_context, format_session_label, render_excel, render_pdf


def teacher(teacher_id, role="Junior", department="IT", years=None, **extra):
    doc = {
        "id": teacher_id,
        "teacher_id": teacher_id,
        "name": f"Prof. {teacher_id}",
        "role": role,
        "department": department,
        "preferred_slots": set(),
        "max_duties": 5,
        "pinned_duties": [],
    }
    if years is not None:
        doc["teaching_years"] = years
    doc.update(extra)
    return doc


class EligibilityTests(unittest.TestCase):
    def test_department_policy(self):
        self.assertEqual(effective_teaching_years({"department": "FY"}), {"FY"})
        self.assertEqual(
            effective_teaching_years({"department": "IT"}),
            {"SY", "TY", "BTECH"},
        )
        self.assertEqual(
            effective_teaching_years({"department": "IT", "teaching_years": []}),
            {"SY", "TY", "BTECH"},
        )

    def test_explicit_years_only_narrow(self):
        self.assertEqual(
            effective_teaching_years({"department": "IT", "teaching_years": ["SY"]}),
            {"SY"},
        )
        self.assertEqual(
            effective_teaching_years({"department": "IT", "teaching_years": ["FY"]}),
            set(),
        )

    def test_stringified_legacy_year_lists_still_narrow(self):
        self.assertEqual(
            effective_teaching_years({"department": "IT", "teaching_years": "['SY']"}),
            {"SY"},
        )
        self.assertEqual(
            effective_teaching_years({"department": "FY", "teaching_years": "['FY']"}),
            {"FY"},
        )

    def test_fy_teacher_cannot_take_sy(self):
        fy = teacher("F1", department="FY", years=["FY"])
        self.assertTrue(teacher_can_invigilate(
            fy, {"year": "FY", "department": "General"}))
        self.assertFalse(teacher_can_invigilate(
            fy, {"year": "SY", "department": "IT"}))

    def test_session_codes(self):
        self.assertEqual(session_code("Morning"), "M")
        self.assertEqual(session_code("10:00 - 12:30"), "M")
        self.assertEqual(session_code("14:00 - 16:30"), "A")
        self.assertEqual(session_code("17:00 - 19:00"), "E")
        self.assertEqual(session_code("2:00 pm to 3:00 pm"), "A")
        self.assertEqual(session_code("Evening: 2:00 pm to 3:00 pm"), "E")
        self.assertEqual(
            format_session_label("Evening: 2:00 pm to 3:00 pm"),
            "Evening: 2:00 pm to 3:00 pm",
        )


class DutyDemandTests(unittest.TestCase):
    def test_json_sanitizer_preserves_year_lists(self):
        row = {
            "duty_id": 1,
            "allowed_years": ["FY"],
            "_internal": "drop me",
            "nested": {"years": ("SY", "TY")},
            "score": np.int64(7),
        }
        cleaned = json_safe_row(row)
        self.assertEqual(cleaned["allowed_years"], ["FY"])
        self.assertEqual(cleaned["nested"], {"years": ["SY", "TY"]})
        self.assertEqual(cleaned["score"], 7)
        self.assertNotIn("_internal", cleaned)

    def test_js_and_squad_demand(self):
        final_data = [
            {"course_id": "C1", "slot": 1, "date": "01-Mar-2026", "session": "Morning",
             "year": "SY", "department": "IT"},
            {"course_id": "C2", "slot": 2, "date": "01-Mar-2026", "session": "Afternoon",
             "year": "TY", "department": "IT"},
        ]
        rooms = [
            {"course_id": "C1", "slot": 1, "date": "01-Mar-2026", "rooms_assigned": "R1"},
            {"course_id": "C2", "slot": 2, "date": "01-Mar-2026", "rooms_assigned": "R2"},
        ]
        config = normalize_duty_config({
            "js_per_session": 3,
            "squad_per_session": 2,
            "last_day_squad": 1,
        })
        duties = build_duties(final_data, room_assignments=rooms, duty_config=config)
        counts = {}
        for duty in duties:
            counts[duty["role_required"]] = counts.get(duty["role_required"], 0) + 1
        self.assertEqual(len(duties), 10)
        self.assertEqual(counts["Junior"], 6)
        self.assertEqual(counts["Senior"], 2)
        self.assertEqual(counts["Squad"], 2)

    def test_pinned_duty_cannot_override_year_rule(self):
        duties = [
            {"duty_id": 1, "slot": 1, "date": "01-Mar-2026", "session": "Morning",
             "course_id": "C1", "year": "SY", "department": "IT", "room": "R1",
             "role_required": "Junior", "allowed_years": ["SY"]},
        ]
        teachers = {
            "F1": teacher(
                "F1", department="FY", years=["FY"],
                pinned_duties=[{"date": "01-Mar-2026", "slot": "1", "role": "Junior"}],
            ),
        }
        assignments, unassigned, _ = assign_teachers(
            teachers, duties, faculty_status={}, admin_keys=set(), MAX_DUTIES=5,
        )
        self.assertEqual(assignments, [])
        self.assertEqual([duty["duty_id"] for duty in unassigned], [1])

    def test_one_duty_per_day_and_load(self):
        duties = [
            {"duty_id": 1, "slot": 1, "date": "01-Mar-2026", "session": "Morning",
             "course_id": "C1", "year": "SY", "department": "IT", "room": "R1",
             "role_required": "Junior", "allowed_years": ["SY"]},
            {"duty_id": 2, "slot": 2, "date": "01-Mar-2026", "session": "Afternoon",
             "course_id": "C2", "year": "SY", "department": "IT", "room": "R2",
             "role_required": "Junior", "allowed_years": ["SY"]},
            {"duty_id": 3, "slot": 1, "date": "02-Mar-2026", "session": "Morning",
             "course_id": "C3", "year": "SY", "department": "IT", "room": "R1",
             "role_required": "Junior", "allowed_years": ["SY"]},
            {"duty_id": 4, "slot": 2, "date": "02-Mar-2026", "session": "Afternoon",
             "course_id": "C4", "year": "SY", "department": "IT", "room": "R2",
             "role_required": "Junior", "allowed_years": ["SY"]},
        ]
        teachers = {
            "T1": teacher("T1"),
            "T2": teacher("T2"),
        }
        assignments, unassigned, counts = assign_teachers(
            teachers, duties, faculty_status={}, admin_keys=set(), MAX_DUTIES=5,
        )
        self.assertEqual(unassigned, [])
        self.assertEqual(len(assignments), 4)
        per_day = {}
        for row in assignments:
            key = (row["teacher_id"], row["date"])
            per_day[key] = per_day.get(key, 0) + 1
        self.assertTrue(all(value == 1 for value in per_day.values()))
        self.assertEqual(sorted(counts.values()), [2, 2])


class FairnessResetTests(unittest.TestCase):
    class FakeTeachers:
        def __init__(self, documents):
            self.documents = [dict(document) for document in documents]
            self.updated = []

        def find(self, query):
            return [doc for doc in self.documents if doc.get("is_admin") is not True]

        def update_many(self, query, update):
            ids = set(query.get("_id", {}).get("$in", []))
            for doc in self.documents:
                if doc.get("_id") in ids:
                    doc.update(update.get("$set", {}))
                    self.updated.append(doc["_id"])

    class FakeDB:
        def __init__(self, documents):
            self.teachers = FairnessResetTests.FakeTeachers(documents)

        def __getitem__(self, key):
            if key != "teachers":
                raise KeyError(key)
            return self.teachers

    def test_all_department_ignores_juniors_and_admins(self):
        fake = self.FakeDB([
            {"_id": 1, "teacher_id": "S1", "role": "Senior", "department": "IT", "has_served_high_role": True},
            {"_id": 2, "teacher_id": "J1", "role": "Junior", "department": "IT", "has_served_high_role": False},
            {"_id": 3, "teacher_id": "A1", "role": "Squad", "department": "IT", "is_admin": True, "has_served_high_role": False},
        ])
        with patch.object(db_module, "get_db", return_value=fake):
            self.assertTrue(db_module.check_reset_fairness("ALL"))
        self.assertFalse(fake.teachers.documents[0]["has_served_high_role"])
        self.assertIn("fairness_cycle_started_at", fake.teachers.documents[0])
        self.assertFalse(fake.teachers.documents[1]["has_served_high_role"])

    def test_department_without_supervisors_does_not_reset(self):
        fake = self.FakeDB([
            {"_id": 1, "teacher_id": "F1", "role": "Junior", "department": "FY", "has_served_high_role": False},
        ])
        with patch.object(db_module, "get_db", return_value=fake):
            self.assertFalse(db_module.check_reset_fairness("FY"))
        self.assertFalse(fake.teachers.documents[0]["has_served_high_role"])

    def test_pending_supervisor_blocks_reset(self):
        fake = self.FakeDB([
            {"_id": 1, "teacher_id": "S1", "role": "Senior", "department": "IT", "has_served_high_role": True},
            {"_id": 2, "teacher_id": "S2", "role": "Squad", "department": "IT", "has_served_high_role": False},
        ])
        with patch.object(db_module, "get_db", return_value=fake):
            self.assertFalse(db_module.check_reset_fairness("IT"))


class RosterTests(unittest.TestCase):
    def setUp(self):
        self.timetable = {
            "_id": "tt1",
            "department": "IT",
            "revision": 1,
            "timetable": [
                {"date": "01-Mar-2026", "session": "Morning"},
                {"date": "02-Mar-2026", "session": "Afternoon"},
            ],
        }
        self.teachers = [teacher("T1"), teacher("T2")]
        self.assignments = [
            {"duty_id": "d1", "teacher_id": "T1", "teacher_name": "Prof. T1",
             "date": "01-Mar-2026", "slot": "1", "session": "Morning",
             "course_id": "C1", "role": "Junior", "room": "R1", "year": "SY"},
            {"duty_id": "d2", "teacher_id": "T2", "teacher_name": "Prof. T2",
             "date": "02-Mar-2026", "slot": "2", "session": "Afternoon",
             "course_id": "C2", "role": "Junior", "room": "R2", "year": "SY"},
        ]

    def test_matrix(self):
        context = build_roster_context(
            self.timetable, self.assignments, self.teachers, "junior",
        )
        self.assertEqual(context["dates"], ["01-Mar-2026", "02-Mar-2026"])
        self.assertEqual(context["codes_by_date"]["01-Mar-2026"], ["M"])
        self.assertEqual(context["codes_by_date"]["02-Mar-2026"], ["A"])
        rows = {row["teacher_id"]: row for row in context["rows"]}
        self.assertEqual(rows["T1"]["cells"]["01-Mar-2026"], "M")
        self.assertEqual(rows["T2"]["cells"]["02-Mar-2026"], "A")
        self.assertEqual(rows["T1"]["cells"].get("02-Mar-2026", ""), "")

    def test_renderers(self):
        context = build_roster_context(
            self.timetable, self.assignments, self.teachers, "junior",
        )
        pdf = render_pdf(context).getvalue()
        excel = render_excel(context).getvalue()
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertTrue(excel.startswith(b"PK"))


if __name__ == "__main__":
    unittest.main()
