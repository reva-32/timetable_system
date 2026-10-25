import os
import sys
import unittest
from datetime import date

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from department_config import (
    DEPARTMENT_BLOCKS,
    category_session,
    squads_for_blocks,
)
from eligibility import cross_department_allowed, teacher_can_invigilate
from multi_department import (
    _build_department_duties,
    _calendar_for_run,
    _select_calendar_dates,
    _workbook_exam_dates,
)


def teacher(teacher_id, department, years=None, role="Junior"):
    return {
        "teacher_id": teacher_id,
        "name": teacher_id,
        "role": role,
        "department": department,
        "teaching_years": years or [],
    }


class DepartmentConfigTests(unittest.TestCase):
    def test_fixed_block_counts(self):
        self.assertEqual(
            DEPARTMENT_BLOCKS,
            {"CE": 10, "ENTC": 10, "IT": 6, "AIDS": 2, "ECE": 2},
        )

    def test_squad_rule(self):
        self.assertEqual(squads_for_blocks(0), 0)
        self.assertEqual(squads_for_blocks(1), 1)
        self.assertEqual(squads_for_blocks(2), 2)
        self.assertEqual(squads_for_blocks(5), 2)
        self.assertEqual(squads_for_blocks(6), 3)
        self.assertEqual(squads_for_blocks(10), 3)

    def test_category_sessions(self):
        self.assertEqual(category_session("MD", "Insem"), "11:00 AM - 12:00 PM")
        self.assertEqual(category_session("PC", "Insem"), "3:30 PM - 4:30 PM")
        self.assertEqual(category_session("HC", "Insem"), "3:30 PM - 4:30 PM")


class CrossDepartmentEligibilityTests(unittest.TestCase):
    def test_fy_faculty_never_invigilates_department_papers(self):
        fy = teacher("F001", "FY")
        for department in ("CE", "ENTC", "IT", "AIDS", "ECE"):
            duty = {
                "exam_department": department,
                "cross_department": True,
                "allowed_years": ["TY"],
                "department": department,
            }
            self.assertFalse(teacher_can_invigilate(fy, duty))
            self.assertFalse(cross_department_allowed(fy, department))

    def test_own_department_is_rejected(self):
        it = teacher("ITJ01", "IT")
        self.assertFalse(cross_department_allowed(it, "IT"))
        self.assertTrue(cross_department_allowed(it, "CE"))
        duty = {
            "exam_department": "IT",
            "cross_department": True,
            "allowed_years": ["TY"],
            "department": "IT",
        }
        self.assertFalse(teacher_can_invigilate(it, duty))

    def test_cross_department_duty_is_allowed(self):
        it = teacher("ITJ01", "IT")
        duty = {
            "exam_department": "CE",
            "cross_department": True,
            "allowed_years": ["TY"],
            "department": "CE",
        }
        self.assertTrue(teacher_can_invigilate(it, duty))

    def test_legacy_duty_keeps_same_department_rule(self):
        it = teacher("ITJ01", "IT")
        legacy = {"department": "IT", "allowed_years": ["TY"]}
        self.assertTrue(teacher_can_invigilate(it, legacy))
        other = {"department": "CE", "allowed_years": ["TY"]}
        self.assertFalse(teacher_can_invigilate(it, other))


class CalendarTests(unittest.TestCase):
    def test_calendar_includes_both_ends_and_skips_sunday(self):
        selected = _select_calendar_dates(
            date(2026, 8, 17), date(2026, 8, 24), 6
        )
        self.assertEqual(
            [value.isoformat() for value in selected],
            [
                "2026-08-17",
                "2026-08-18",
                "2026-08-19",
                "2026-08-21",
                "2026-08-22",
                "2026-08-24",
            ],
        )

    def test_calendar_rejects_a_period_that_is_too_short(self):
        with self.assertRaises(ValueError):
            _select_calendar_dates(date(2026, 8, 17), date(2026, 8, 18), 3)


class WorkbookDateTests(unittest.TestCase):
    """The workbook's Slots sheet must win over the even-spread fallback."""

    @staticmethod
    def slots(*dates):
        return {
            index + 1: {"display_id": index + 1, "date": value}
            for index, value in enumerate(dates)
        }

    def test_workbook_dates_are_used_in_slot_order(self):
        selected = _workbook_exam_dates(
            self.slots(
                "2026-08-17",
                "2026-08-18",
                "2026-08-19",
                "2026-08-20",
                "2026-08-22",
                "2026-08-24",
            ),
            date(2026, 8, 17),
            date(2026, 8, 24),
        )
        self.assertEqual(
            [value.isoformat() for value in selected],
            [
                "2026-08-17",
                "2026-08-18",
                "2026-08-19",
                "2026-08-20",
                "2026-08-22",
                "2026-08-24",
            ],
        )

    def test_workbook_dates_skip_sundays_duplicates_and_out_of_range_values(self):
        selected = _workbook_exam_dates(
            self.slots(
                "2026-08-17",
                "2026-08-17",
                "2026-08-23",  # Sunday
                "2026-09-01",  # after the exam period
                "2026-08-18",
                "not-a-date",
            ),
            date(2026, 8, 17),
            date(2026, 8, 24),
        )
        self.assertEqual(
            [value.isoformat() for value in selected],
            ["2026-08-17", "2026-08-18"],
        )

    def test_workbook_dates_accept_common_date_formats_and_lists(self):
        slots = [
            {"display_id": 1, "date": "17-Aug-2026"},
            {"display_id": 2, "date": "18/08/2026"},
            {"display_id": 3, "date": date(2026, 8, 19)},
        ]
        selected = _workbook_exam_dates(
            slots, date(2026, 8, 17), date(2026, 8, 24)
        )
        self.assertEqual(
            [value.isoformat() for value in selected],
            ["2026-08-17", "2026-08-18", "2026-08-19"],
        )


class DynamicCalendarTests(unittest.TestCase):
    """The coordinator's selected range controls the generated PDF dates."""

    SLOTS = {
        1: {"display_id": 1, "date": "2026-08-17"},
        2: {"display_id": 2, "date": "2026-08-18"},
        3: {"display_id": 3, "date": "2026-08-19"},
        4: {"display_id": 4, "date": "2026-08-20"},
        5: {"display_id": 5, "date": "2026-08-22"},
        6: {"display_id": 6, "date": "2026-08-24"},
    }

    @staticmethod
    def iso(dates):
        return [value.isoformat() for value in dates]

    def test_workbook_dates_are_used_when_they_fit_the_selected_range(self):
        selected = _calendar_for_run(
            self.SLOTS, date(2026, 8, 17), date(2026, 8, 24), 6
        )
        self.assertEqual(
            self.iso(selected),
            [
                "2026-08-17",
                "2026-08-18",
                "2026-08-19",
                "2026-08-20",
                "2026-08-22",
                "2026-08-24",
            ],
        )

    def test_custom_range_outside_the_workbook_uses_the_selected_range(self):
        selected = _calendar_for_run(
            self.SLOTS, date(2026, 9, 1), date(2026, 9, 10), 6
        )
        self.assertEqual(
            self.iso(selected),
            [
                "2026-09-01",
                "2026-09-03",
                "2026-09-04",
                "2026-09-07",
                "2026-09-08",
                "2026-09-10",
            ],
        )
        self.assertFalse(any(value.month == 8 for value in selected))

    def test_narrowed_range_uses_the_selected_range_when_workbook_dates_are_short(self):
        selected = _calendar_for_run(
            self.SLOTS, date(2026, 8, 18), date(2026, 8, 24), 6
        )
        self.assertEqual(
            self.iso(selected),
            [
                "2026-08-18",
                "2026-08-19",
                "2026-08-20",
                "2026-08-21",
                "2026-08-22",
                "2026-08-24",
            ],
        )

    def test_generated_dates_never_leave_the_selected_range(self):
        for start, end in [
            (date(2026, 8, 18), date(2026, 8, 24)),
            (date(2026, 9, 1), date(2026, 9, 10)),
            (date(2027, 1, 4), date(2027, 1, 20)),
        ]:
            with self.subTest(start=start, end=end):
                selected = _calendar_for_run(self.SLOTS, start, end, 6)
                self.assertTrue(selected)
                self.assertTrue(all(start <= value <= end for value in selected))
                self.assertTrue(all(value.weekday() != 6 for value in selected))


class SampleWorkbookAlignmentTests(unittest.TestCase):
    """The bundled sample workbook mirrors the August 2026 circular."""

    def test_slots_match_the_circular(self):
        from generate_sample_excel import build_slots

        slots = build_slots()
        self.assertEqual(
            [slot["date"] for slot in slots],
            [
                "2026-08-17",
                "2026-08-18",
                "2026-08-19",
                "2026-08-20",
                "2026-08-22",
                "2026-08-24",
            ],
        )
        self.assertEqual(
            [slot["session"] for slot in slots],
            [
                "3:30 PM - 4:30 PM",
                "3:30 PM - 4:30 PM",
                "3:30 PM - 4:30 PM",
                "3:30 PM - 4:30 PM",
                "11:00 AM - 12:00 PM",
                "3:30 PM - 4:30 PM",
            ],
        )

    def test_shared_md_and_hc_codes_match_the_circular(self):
        from generate_sample_excel import COURSE_CATALOGUE

        expected_md = {
            "CE": ["2605MD11", "2605MD12", "2605MD15"],
            "ENTC": ["2605MD12", "2605MD15", "2605MD16"],
            "IT": ["2605MD12", "2605MD15"],
            "AIDS": ["2605MD12", "2605MD15"],
            "ECE": ["2605MD15", "2605MD19"],
        }
        expected_hc = ["2625HC11", "2605HC11", "2635HC11"]
        for department, codes in expected_md.items():
            with self.subTest(department=department, category="MD"):
                self.assertEqual(
                    [code for code, _name in COURSE_CATALOGUE[department]["md"]],
                    codes,
                )
            with self.subTest(department=department, category="HC"):
                self.assertEqual(
                    [code for code, _name in COURSE_CATALOGUE[department]["hc"]],
                    expected_hc,
                )

    def test_internal_course_ids_stay_unique_for_shared_codes(self):
        from collections import defaultdict

        from generate_sample_excel import build_courses_and_students

        courses, students = build_courses_and_students()
        course_ids = [course["course_id"] for course in courses]
        self.assertEqual(len(course_ids), len(set(course_ids)))
        self.assertTrue(students)

        internal_by_display = defaultdict(set)
        for course in courses:
            internal_by_display[course["course_code"]].add(course["course_id"])
        # Shared circular codes still point at one distinct paper per department.
        self.assertEqual(len(internal_by_display["2605MD12"]), 4)
        self.assertEqual(len(internal_by_display["2625HC11"]), 5)

    def test_exam_order_follows_the_circular_category_sequence(self):
        from generate_sample_excel import (
            COURSE_CATALOGUE,
            build_courses_and_students,
        )

        courses, _students = build_courses_and_students()
        by_department = {}
        for course in courses:
            by_department.setdefault(course["department"], []).append(course)

        for department, department_courses in by_department.items():
            with self.subTest(department=department):
                self.assertEqual(
                    [course["course_code"] for course in department_courses],
                    [
                        code
                        for category in ("pc", "pe", "md", "hc")
                        for code, _name in COURSE_CATALOGUE[department][category]
                    ],
                )
                for category in ("PC", "PE", "MD", "HC"):
                    category_courses = [
                        course for course in department_courses
                        if course["exam_category"] == category
                    ]
                    with self.subTest(department=department, category=category):
                        self.assertEqual(
                            [course["exam_order"] for course in category_courses],
                            list(range(1, len(category_courses) + 1)),
                        )


class DutyShapeTests(unittest.TestCase):
    def test_one_junior_per_block_plus_squad_and_senior(self):
        blocks = [
            {"block": f"CE-B{index:02d}", "room_id": f"CE-{100 + index}", "capacity": 60}
            for index in range(1, 11)
        ]
        sessions = [{
            "slot": 1,
            "date_label": "17-Aug-2026",
            "session": "3:30 PM - 4:30 PM",
            "category": "PC",
            "position": 0,
            "courses": [{
                "course_id": "2615PC11",
                "course_code": "2615PC11",
                "year": "TY",
            }],
        }]
        duties, squad_count = _build_department_duties(
            "CE", sessions, blocks, "TY", 1
        )
        self.assertEqual(squad_count, 3)
        roles = [duty["role_required"] for duty in duties]
        self.assertEqual(roles.count("Junior"), 10)
        self.assertEqual(roles.count("Squad"), 3)
        self.assertEqual(roles.count("Senior"), 1)
        self.assertEqual(len(duties), 14)
        for duty in duties:
            self.assertEqual(duty["exam_department"], "CE")
            self.assertTrue(duty["cross_department"])
            self.assertEqual(duty["allowed_years"], ["TY"])
            self.assertEqual(duty["session"], "3:30 PM - 4:30 PM")
            self.assertEqual(duty["date"], "17-Aug-2026")


if __name__ == "__main__":
    unittest.main()
