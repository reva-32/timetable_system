"""
generate_sample_excel.py — build a complete multi-department sample workbook.

The workbook mirrors the structure of a real examination dataset:

* Courses for CE, ENTC, IT, AIDS, ECE (third year / TY)
* PC, PE, MD and HC paper categories so the generated PDF follows the
  uploaded college circular (three core days, one elective day, one
  multidisciplinary day and one honours/minor day)
* Faculty for every department, split into Junior, Squad and Senior roles
* Exactly the configured classroom blocks for every department
* Student enrolments that create genuinely conflict-free paper groups

The workbook is written to ``sample_exam_data.xlsx``. Regenerating overwrites
that file, so close it in Excel first.

Usage:
    cd backend
    python generate_sample_excel.py
"""

import os
from collections import defaultdict

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "sample_exam_data.xlsx")

DEPARTMENT_ORDER = ["CE", "ENTC", "IT", "AIDS", "ECE"]
DEPARTMENT_BLOCKS = {"CE": 10, "ENTC": 10, "IT": 6, "AIDS": 2, "ECE": 2}
DEPARTMENT_ROOM_PREFIX = {
    "CE": "CE",
    "ENTC": "ENTC",
    "IT": "IT",
    "AIDS": "AIDS",
    "ECE": "ECE",
}
CATEGORY_SESSION = {
    "PC": "3:30 PM - 4:30 PM",
    "PE": "3:30 PM - 4:30 PM",
    "MD": "11:00 AM - 12:00 PM",
    "HC": "3:30 PM - 4:30 PM",
}

TEACHER_ROLES = [
    ("Junior", 8, "J"),
    ("Squad", 4, "SQ"),
    ("Senior", 2, "SS"),
]
FY_TEACHERS = [
    ("F001", "Prof. FY_1"),
    ("F002", "Prof. FY_2"),
    ("F003", "Prof. FY_3"),
    ("F004", "Prof. FY_4"),
    ("F005", "Prof. FY_5"),
]

# Paper catalogues.  Codes are unique across the whole workbook so that a
# student/course id can never belong to two departments.
COURSE_CATALOGUE = {
    "CE": {
        "pc": [
            ("2615PC11", "Computer Networks (CN)"),
            ("2615PC12", "Theory of Computation (TOC)"),
            ("2615PC13", "Artificial Intelligence (AI)"),
        ],
        "pe": [
            ("2615PE11", "Data Warehousing and Mining (DWM)"),
            ("2615PE12", "File Systems Design and Rust Programming (FSRP)"),
            ("2615PE13", "Computer Graphics (CG)"),
            ("2615PE14", "Client-Side Technologies and Frameworks (CSTF)"),
            ("2615PE15", "Cryptography (CGR)"),
        ],
        "md": [
            ("2605MD11", "Data Analytics for Smart and Sustainable Systems (DASSS)"),
            ("2605MD12", "Fundamentals of Stock Market (FSM)"),
            ("2605MD15", "Graph Neural Networks with Python (GNNP)"),
        ],
        "hc": [
            ("2625HC11", "VLSI Design and Embedded Systems (VDES)"),
            ("2605HC11", "Fundamental of Blockchain"),
            ("2635HC11", "Fundamentals of Internet of Things (FIoT)"),
        ],
    },
    "ENTC": {
        "pc": [
            ("2625PC11", "Computer Networks (CN)"),
            ("2625PC12", "Embedded Processors (EP)"),
            ("2625PC13", "Digital Signal Processing (DSP)"),
        ],
        "pe": [
            ("2625PE11", "Software Defined Radio (SDR)"),
            ("2625PE12", "Mechatronics (MECHX)"),
            ("2625PE13", "Digital Image and Video Processing (DIVP)"),
            ("2625PE14", "Cloud Computing (CC)"),
            ("2625PE15", "CMOS Analog Circuit Design (CACD)"),
            ("2625PE16", "Microcontrollers and IoT (MI)"),
        ],
        "md": [
            ("2605MD12", "Fundamentals of Stock Market (FSM)"),
            ("2605MD15", "Graph Neural Networks with Python (GNNP)"),
            ("2605MD16", "Database Management Systems (DBMS)"),
        ],
        "hc": [
            ("2625HC11", "VLSI Design and Embedded Systems (VDES)"),
            ("2605HC11", "Fundamental of Blockchain"),
            ("2635HC11", "Fundamentals of Internet of Things (FIoT)"),
        ],
    },
    "IT": {
        "pc": [
            ("2635PC11", "Theory of Computation (TOC)"),
            ("2635PC12", "Operating Systems (OS)"),
            ("2635PC13", "Artificial Intelligence and Machine Learning (AIML)"),
        ],
        "pe": [
            ("2635PE11", "Information & Cyber Security (ICS)"),
            ("2635PE12", "Deep & Reinforcement Learning (DRL)"),
            ("2635PE13", "Data Warehouse & Data Mining (DWDM)"),
            ("2635PE14", "DevOps for Scalable Systems (DOSS)"),
            ("2635PE15", "Computer Graphics & Multimedia Techniques (CGMT)"),
        ],
        "md": [
            ("2605MD12", "Fundamentals of Stock Market (FSM)"),
            ("2605MD15", "Graph Neural Networks with Python (GNNP)"),
        ],
        "hc": [
            ("2625HC11", "VLSI Design and Embedded Systems (VDES)"),
            ("2605HC11", "Fundamental of Blockchain"),
            ("2635HC11", "Fundamentals of Internet of Things (FIoT)"),
        ],
    },
    "AIDS": {
        "pc": [
            ("2645PC11", "Deep Learning (DL)"),
            ("2645PC12", "Cloud Computing (CC)"),
            ("2645PC13", "Big Data Analytics (BDA)"),
        ],
        "pe": [
            ("2645PE11", "Computer Architecture and Organization (CAO)"),
            ("2645PE12", "Foundations of Game Design (FGD)"),
            ("2645PE13", "Foundations of Web Development (FWD)"),
            ("2645PE14", "Cyber Security Principles (CSP)"),
            ("2645PE15", "Fundamentals of Computer Vision (FCV)"),
        ],
        "md": [
            ("2605MD12", "Fundamentals of Stock Market (FSM)"),
            ("2605MD15", "Graph Neural Networks with Python (GNNP)"),
        ],
        "hc": [
            ("2625HC11", "VLSI Design and Embedded Systems (VDES)"),
            ("2605HC11", "Fundamental of Blockchain"),
            ("2635HC11", "Fundamentals of Internet of Things (FIoT)"),
        ],
    },
    "ECE": {
        "pc": [
            ("2655PC11", "Computer Networks (CN)"),
            ("2655PC12", "Database Management System (DBMS)"),
            ("2655PC13", "Embedded System (ES)"),
        ],
        "pe": [
            ("2655PE11", "Mechatronics (MECHX)"),
            ("2655PE12", "Web Development (WD)"),
            ("2655PE13", "Micro Electro Mechanical System (MEMS)"),
            ("2655PE14", "Big Data Analytics (BDA)"),
            ("2655PE15", "Security and Privacy (SP)"),
        ],
        "md": [
            ("2605MD15", "Graph Neural Networks with Python (GNNP)"),
            ("2605MD19", "Artificial Intelligence (AI)"),
        ],
        "hc": [
            ("2625HC11", "VLSI Design and Embedded Systems (VDES)"),
            ("2605HC11", "Fundamental of Blockchain"),
            ("2635HC11", "Fundamentals of Internet of Things (FIoT)"),
        ],
    },
}

STUDENT_COUNTS = {"CE": 80, "ENTC": 80, "IT": 60, "AIDS": 30, "ECE": 30}


def build_teachers():
    teachers = []
    for department in DEPARTMENT_ORDER:
        for role, count, prefix in TEACHER_ROLES:
            for index in range(1, count + 1):
                teacher_id = f"{department}{prefix}{index:02d}"
                teachers.append({
                    "teacher_id": teacher_id,
                    "name": f"Prof. {department} {role} {index}",
                    "role": role,
                    "department": department,
                })
    for teacher_id, name in FY_TEACHERS:
        teachers.append({
            "teacher_id": teacher_id,
            "name": name,
            "role": "Junior",
            "department": "FY",
        })
    return teachers


def course_key(department, category, display_code):
    """Return the internal key used by enrolments for a displayed paper code.

    The college circular reuses the same MD/HC codes across departments, while
    the scheduler needs a unique key per department paper.  The workbook keeps
    the displayed code in ``course_code`` and stores the unique key in
    ``course_id``.
    """
    if category.upper() in {"MD", "HC"}:
        return f"{department}-{display_code}"
    return display_code


def build_courses_and_students():
    courses = []
    students = []
    enrollments = defaultdict(int)

    for department in DEPARTMENT_ORDER:
        catalogue = COURSE_CATALOGUE[department]
        student_count = STUDENT_COUNTS[department]
        code_set = set()
        for category in ("pc", "pe", "md", "hc"):
            for order, (code, name) in enumerate(catalogue[category], start=1):
                if code in code_set:
                    raise ValueError(f"Duplicate course code {code} in {department}")
                code_set.add(code)
                courses.append({
                    "course_id": course_key(department, category, code),
                    "course_code": code,
                    "course_name": name,
                    "year": 3,
                    "students_count": 0,
                    "department": department,
                    "exam_category": category.upper(),
                    "session": CATEGORY_SESSION[category.upper()],
                    "exam_order": order,
                })

        for student_index in range(1, student_count + 1):
            student_id = f"{department}_TY_{student_index:03d}"
            selected = [
                (course_key(department, "PC", code), code)
                for code, _name in catalogue["pc"]
            ]
            selected.append((
                course_key(department, "PE", catalogue["pe"][(student_index - 1) % len(catalogue["pe"])][0]),
                catalogue["pe"][(student_index - 1) % len(catalogue["pe"])][0],
            ))
            selected.append((
                course_key(department, "MD", catalogue["md"][(student_index - 1) % len(catalogue["md"])][0]),
                catalogue["md"][(student_index - 1) % len(catalogue["md"])][0],
            ))
            selected.append((
                course_key(department, "HC", catalogue["hc"][(student_index - 1) % len(catalogue["hc"])][0]),
                catalogue["hc"][(student_index - 1) % len(catalogue["hc"])][0],
            ))
            for internal_id, _display_code in selected:
                students.append({"student_id": student_id, "course_id": internal_id})
                enrollments[internal_id] += 1

    for course in courses:
        course["students_count"] = enrollments.get(course["course_id"], 0)
    return courses, students


def build_rooms():
    rooms = []
    capacities = [45, 45, 50, 50, 55, 55, 60, 60, 65, 65]
    for department in DEPARTMENT_ORDER:
        prefix = DEPARTMENT_ROOM_PREFIX[department]
        for index in range(1, DEPARTMENT_BLOCKS[department] + 1):
            rooms.append({
                "room_id": f"{prefix}-{100 + index:03d}",
                "capacity": capacities[(index - 1) % len(capacities)],
                "department": department,
            })
    rooms.extend([
        {"room_id": "Audi-1", "capacity": 120, "department": "General"},
        {"room_id": "Audi-2", "capacity": 120, "department": "General"},
    ])
    return rooms


def build_slots():
    return [
        {"slot_id": 1, "date": "2026-08-17", "session": "3:30 PM - 4:30 PM"},
        {"slot_id": 2, "date": "2026-08-18", "session": "3:30 PM - 4:30 PM"},
        {"slot_id": 3, "date": "2026-08-19", "session": "3:30 PM - 4:30 PM"},
        {"slot_id": 4, "date": "2026-08-20", "session": "3:30 PM - 4:30 PM"},
        {"slot_id": 5, "date": "2026-08-22", "session": "11:00 AM - 12:00 PM"},
        {"slot_id": 6, "date": "2026-08-24", "session": "3:30 PM - 4:30 PM"},
    ]


def generate():
    teachers = build_teachers()
    courses, students = build_courses_and_students()
    rooms = build_rooms()
    slots = build_slots()

    sheets = {
        "Student_Courses": pd.DataFrame(students),
        "Courses": pd.DataFrame(courses),
        "Rooms": pd.DataFrame(rooms),
        "Teachers": pd.DataFrame(teachers),
        "Slots": pd.DataFrame(slots),
    }

    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            for sheet_name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Generated multi-department sample workbook: {OUTPUT_FILE}")
        print(
            f"  {len(courses)} courses, {len(students)} enrolments, "
            f"{len(teachers)} faculty, {len(rooms)} rooms"
        )
    except PermissionError as exc:
        raise RuntimeError(
            "sample_exam_data.xlsx is open in Excel. Close it and run the "
            "generator again."
        ) from exc


if __name__ == "__main__":
    try:
        generate()
    except RuntimeError as exc:
        raise SystemExit(str(exc))
