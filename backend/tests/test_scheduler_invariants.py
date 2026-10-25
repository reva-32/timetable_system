import os, sys, unittest, tempfile, types
from pathlib import Path
from unittest.mock import patch
import openpyxl

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from main import dsatur_coloring, validate_schedule, adjust_exam_dates, allocate_rooms, teacher_can_invigilate, compute_cost, load_teacher_data


def make_workbook():
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Teachers"
    ws.append(["teacher_id", "name", "role", "department", "preferred_slots"])
    ws.append(["T001", "Prof. One", "Junior", "IT", "1"])
    wb.save(path)
    wb.close()
    return path


class Tests(unittest.TestCase):
    def test_dsatur(self):
        g = {"A": {"B", "C"}, "B": {"A", "C"}, "C": {"A", "B"}}
        c = dsatur_coloring(g)
        for n, ns in g.items():
            for x in ns:
                self.assertNotEqual(c[n], c[x])

    def test_date_adjustment(self):
        rows = [
            {"course_id": "C1", "date": "01-Jul-2026", "slot": 1, "session": "Morning"},
            {"course_id": "C2", "date": "01-Jul-2026", "slot": 2, "session": "Afternoon"},
        ]
        students = {"S1": ["C1", "C2"]}
        meta = {
            0: {"display_id": 1, "date": "01-Jul-2026", "session": "Morning"},
            1: {"display_id": 2, "date": "01-Jul-2026", "session": "Afternoon"},
            2: {"display_id": 1, "date": "02-Jul-2026", "session": "Morning"},
        }
        adjust_exam_dates(rows, students, meta)
        self.assertEqual(validate_schedule(rows, students), [])

    def test_year_rule(self):
        t = {"role": "Junior", "department": "IT", "teaching_years": {"FY"}}
        d = {"role_required": "Junior", "department": "IT", "year": "SY", "slot": 1}
        self.assertFalse(teacher_can_invigilate(t, d))
        self.assertEqual(compute_cost(t, d, {}, db_duty_counts={}), 10000)

    def test_same_year(self):
        t = {"role": "Junior", "department": "IT", "teaching_years": {"SY"}}
        d = {"role_required": "Junior", "department": "IT", "year": "2", "slot": 1}
        self.assertTrue(teacher_can_invigilate(t, d))

    def test_excel_has_no_teacher_year_field(self):
        path = make_workbook()
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            headers = {str(c.value).strip().lower() for c in wb["Teachers"][1]}
            self.assertFalse(headers & {"teaching_year", "teaching_years", "years_taught"})
            wb.close()
        finally:
            Path(path).unlink(missing_ok=True)

    def test_teacher_years_come_from_mongodb(self):
        class FakeCollection:
            def find(self, *args, **kwargs):
                return [{"teacher_id": "T001", "department": "IT", "teaching_years": [], "preferred_slots": ["1"]}]

        class FakeDB:
            def __getitem__(self, key):
                return FakeCollection()

        path = make_workbook()
        fake_db = types.ModuleType("db")
        fake_db.get_db = lambda: FakeDB()
        try:
            with patch.dict(sys.modules, {"db": fake_db}):
                teachers = load_teacher_data(path)
            # An empty teaching_years list means "use the department policy",
            # never "no invigilation".
            self.assertEqual(teachers["T001"]["teaching_years"], {"SY", "TY", "BTECH"})
            self.assertNotIn("teaching_year", teachers["T001"])
        finally:
            import gc
            gc.collect()
            Path(path).unlink(missing_ok=True)

    def test_rooms(self):
        rows = [
            {"course_id": "A", "date": "01-Jul-2026", "slot": 1, "enrolled_students": 30},
            {"course_id": "B", "date": "01-Jul-2026", "slot": 1, "enrolled_students": 30},
        ]
        rooms = [{"room_id": "R1", "capacity": 30}, {"room_id": "R2", "capacity": 30}]
        allocated, unallocated = allocate_rooms(rows, rooms)
        self.assertFalse(unallocated)
        self.assertNotEqual(allocated[0]["rooms_assigned"], allocated[1]["rooms_assigned"])


if __name__ == "__main__":
    unittest.main()
