"""
generate_sample_excel.py — Creates a comprehensive test workbook (sample_exam_data.xlsx)
with 10 teachers (001 to 010), multi-department courses, rooms, and student enrollments.

Usage:
    cd backend
    python generate_sample_excel.py
"""

import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "sample_exam_data.xlsx")

# 1. Teachers Sheet (referencing 001 to 010)
teachers_data = [
    {"teacher_id": "001", "name": "Prof. Aarav Sharma", "role": "Junior", "department": "IT"},
    {"teacher_id": "002", "name": "Dr. Bhavna Patil", "role": "Junior", "department": "IT"},
    {"teacher_id": "003", "name": "Prof. Chetan Kulkarni", "role": "Junior", "department": "CE"},
    {"teacher_id": "004", "name": "Dr. Deepa Deshmukh", "role": "Junior", "department": "CE"},
    {"teacher_id": "005", "name": "Prof. Eshan Joshi", "role": "Junior", "department": "ENTC"},
    {"teacher_id": "006", "name": "Dr. Fatima Shaikh", "role": "Junior", "department": "AIDS"},
    {"teacher_id": "007", "name": "Prof. Girish Mehta", "role": "Senior", "department": "IT"},
    {"teacher_id": "008", "name": "Dr. Hemlata Rao", "role": "Senior", "department": "CE"},
    {"teacher_id": "009", "name": "Prof. Ishaan Verma", "role": "Squad", "department": "ENTC"},
    {"teacher_id": "010", "name": "Dr. Jayant Nambiar", "role": "Squad", "department": "AIDS"},
]

# 2. Courses Sheet
courses_data = [
    {"course_id": "IT301", "course_name": "Database Management Systems", "year": "3", "department": "IT", "students_count": 60, "exam_date": "10-Dec-2026", "session": "Morning"},
    {"course_id": "IT302", "course_name": "Operating Systems", "year": "3", "department": "IT", "students_count": 60, "exam_date": "12-Dec-2026", "session": "Morning"},
    {"course_id": "IT303", "course_name": "Computer Networks", "year": "3", "department": "IT", "students_count": 60, "exam_date": "14-Dec-2026", "session": "Morning"},
    {"course_id": "CE301", "course_name": "Theory of Computation", "year": "3", "department": "CE", "students_count": 60, "exam_date": "10-Dec-2026", "session": "Afternoon"},
    {"course_id": "CE302", "course_name": "Software Engineering", "year": "3", "department": "CE", "students_count": 60, "exam_date": "12-Dec-2026", "session": "Afternoon"},
    {"course_id": "ENTC301", "course_name": "Digital Signal Processing", "year": "3", "department": "ENTC", "students_count": 50, "exam_date": "10-Dec-2026", "session": "Morning"},
    {"course_id": "AIDS301", "course_name": "Artificial Intelligence", "year": "3", "department": "AIDS", "students_count": 55, "exam_date": "11-Dec-2026", "session": "Morning"},
    {"course_id": "AIDS302", "course_name": "Machine Learning", "year": "3", "department": "AIDS", "students_count": 55, "exam_date": "13-Dec-2026", "session": "Afternoon"},
]

# 3. Rooms Sheet
rooms_data = [
    {"room_id": "A101", "capacity": 60, "department": "IT"},
    {"room_id": "A102", "capacity": 60, "department": "IT"},
    {"room_id": "B201", "capacity": 70, "department": "CE"},
    {"room_id": "B202", "capacity": 60, "department": "CE"},
    {"room_id": "C301", "capacity": 55, "department": "ENTC"},
    {"room_id": "D401", "capacity": 60, "department": "AIDS"},
    {"room_id": "LH-01", "capacity": 120, "department": "General"},
    {"room_id": "LH-02", "capacity": 100, "department": "General"},
]

# 4. Student Enrollments Sheet
# Generate 60 IT students (IT_S1 to IT_S60) enrolled in IT301, IT302, IT303
# 60 CE students (CE_S1 to CE_S60) enrolled in CE301, CE302
# 50 ENTC students (ENTC_S1 to ENTC_S50) enrolled in ENTC301
# 55 AIDS students (AIDS_S1 to AIDS_S55) enrolled in AIDS301, AIDS302
students_data = []

for i in range(1, 61):
    sid = f"IT_S{i:02d}"
    students_data.append({"student_id": sid, "course_id": "IT301"})
    students_data.append({"student_id": sid, "course_id": "IT302"})
    students_data.append({"student_id": sid, "course_id": "IT303"})

for i in range(1, 61):
    sid = f"CE_S{i:02d}"
    students_data.append({"student_id": sid, "course_id": "CE301"})
    students_data.append({"student_id": sid, "course_id": "CE302"})

for i in range(1, 51):
    sid = f"ENTC_S{i:02d}"
    students_data.append({"student_id": sid, "course_id": "ENTC301"})

for i in range(1, 56):
    sid = f"AIDS_S{i:02d}"
    students_data.append({"student_id": sid, "course_id": "AIDS301"})
    students_data.append({"student_id": sid, "course_id": "AIDS302"})

# 5. Slots Sheet
slots_data = [
    {"slot_id": 1, "date": "10-Dec-2026", "session": "10:00 – 12:30"},
    {"slot_id": 2, "date": "10-Dec-2026", "session": "14:00 – 16:30"},
    {"slot_id": 3, "date": "11-Dec-2026", "session": "10:00 – 12:30"},
    {"slot_id": 4, "date": "12-Dec-2026", "session": "10:00 – 12:30"},
    {"slot_id": 5, "date": "12-Dec-2026", "session": "14:00 – 16:30"},
    {"slot_id": 6, "date": "13-Dec-2026", "session": "14:00 – 16:30"},
    {"slot_id": 7, "date": "14-Dec-2026", "session": "10:00 – 12:30"},
]

def generate():
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame(students_data).to_excel(writer, sheet_name="Student_Courses", index=False)
        pd.DataFrame(courses_data).to_excel(writer, sheet_name="Courses", index=False)
        pd.DataFrame(rooms_data).to_excel(writer, sheet_name="Rooms", index=False)
        pd.DataFrame(teachers_data).to_excel(writer, sheet_name="Teachers", index=False)
        pd.DataFrame(slots_data).to_excel(writer, sheet_name="Slots", index=False)

    print(f"Generated test dataset: {OUTPUT_FILE}")
    print(f" - Students enrollments: {len(students_data)}")
    print(f" - Courses: {len(courses_data)}")
    print(f" - Rooms: {len(rooms_data)}")
    print(f" - Teachers: {len(teachers_data)} (001 to 010)")
    print(f" - Slots: {len(slots_data)}")

if __name__ == "__main__":
    generate()
