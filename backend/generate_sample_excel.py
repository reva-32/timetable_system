"""
generate_sample_excel.py — Creates a complete test workbook (sample_exam_data.xlsx)
matching the user's exact dataset structure:
- Courses (20 IT subjects, numeric IDs 1 to 20, years 1 to 4)
- Teachers (30 faculty members T001 to T030 with exact roles)
- Slots (Slots 1 to 10 with dates and morning/afternoon sessions)
- Rooms (Capacity-aware exam rooms)
- Student_Courses (Student enrollments across IT subjects)

Usage:
    cd backend
    python generate_sample_excel.py
"""

import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "sample_exam_data.xlsx")

# 1. Teachers Sheet (T001 to T030 with exact roles)
teachers_data = [
    {"teacher_id": "T001", "name": "Prof. IT_1", "role": "Squad", "department": "IT"},
    {"teacher_id": "T002", "name": "Prof. IT_2", "role": "Squad", "department": "IT"},
    {"teacher_id": "T003", "name": "Prof. IT_3", "role": "Junior", "department": "IT"},
    {"teacher_id": "T004", "name": "Prof. IT_4", "role": "Junior", "department": "IT"},
    {"teacher_id": "T005", "name": "Prof. IT_5", "role": "Senior", "department": "IT"},
    {"teacher_id": "T006", "name": "Prof. IT_6", "role": "Squad", "department": "IT"},
    {"teacher_id": "T007", "name": "Prof. IT_7", "role": "Senior", "department": "IT"},
    {"teacher_id": "T008", "name": "Prof. IT_8", "role": "Junior", "department": "IT"},
    {"teacher_id": "T009", "name": "Prof. IT_9", "role": "Junior", "department": "IT"},
    {"teacher_id": "T010", "name": "Prof. IT_10", "role": "Squad", "department": "IT"},
    {"teacher_id": "T011", "name": "Prof. IT_11", "role": "Squad", "department": "IT"},
    {"teacher_id": "T012", "name": "Prof. IT_12", "role": "Squad", "department": "IT"},
    {"teacher_id": "T013", "name": "Prof. IT_13", "role": "Squad", "department": "IT"},
    {"teacher_id": "T014", "name": "Prof. IT_14", "role": "Junior", "department": "IT"},
    {"teacher_id": "T015", "name": "Prof. IT_15", "role": "Senior", "department": "IT"},
    {"teacher_id": "T016", "name": "Prof. IT_16", "role": "Squad", "department": "IT"},
    {"teacher_id": "T017", "name": "Prof. IT_17", "role": "Junior", "department": "IT"},
    {"teacher_id": "T018", "name": "Prof. IT_18", "role": "Junior", "department": "IT"},
    {"teacher_id": "T019", "name": "Prof. IT_19", "role": "Senior", "department": "IT"},
    {"teacher_id": "T020", "name": "Prof. IT_20", "role": "Senior", "department": "IT"},
    {"teacher_id": "T021", "name": "Prof. IT_21", "role": "Senior", "department": "IT"},
    {"teacher_id": "T022", "name": "Prof. IT_22", "role": "Junior", "department": "IT"},
    {"teacher_id": "T023", "name": "Prof. IT_23", "role": "Squad", "department": "IT"},
    {"teacher_id": "T024", "name": "Prof. IT_24", "role": "Squad", "department": "IT"},
    {"teacher_id": "T025", "name": "Prof. IT_25", "role": "Squad", "department": "IT"},
    {"teacher_id": "T026", "name": "Prof. IT_26", "role": "Junior", "department": "IT"},
    {"teacher_id": "T027", "name": "Prof. IT_27", "role": "Junior", "department": "IT"},
    {"teacher_id": "T028", "name": "Prof. IT_28", "role": "Squad", "department": "IT"},
    {"teacher_id": "T029", "name": "Prof. IT_29", "role": "Squad", "department": "IT"},
    {"teacher_id": "T030", "name": "Prof. IT_30", "role": "Junior", "department": "IT"},
]

# 2. Courses Sheet (Matching Image 1)
courses_data = [
    {"course_id": 1, "course_name": "IT_Subject_1", "year": 1, "students_count": 31, "department": "IT"},
    {"course_id": 2, "course_name": "IT_Subject_2", "year": 2, "students_count": 30, "department": "IT"},
    {"course_id": 3, "course_name": "IT_Subject_3", "year": 2, "students_count": 32, "department": "IT"},
    {"course_id": 4, "course_name": "IT_Subject_4", "year": 2, "students_count": 28, "department": "IT"},
    {"course_id": 5, "course_name": "IT_Subject_5", "year": 4, "students_count": 29, "department": "IT"},
    {"course_id": 6, "course_name": "IT_Subject_6", "year": 3, "students_count": 33, "department": "IT"},
    {"course_id": 7, "course_name": "IT_Subject_7", "year": 2, "students_count": 24, "department": "IT"},
    {"course_id": 8, "course_name": "IT_Subject_8", "year": 4, "students_count": 30, "department": "IT"},
    {"course_id": 9, "course_name": "IT_Subject_9", "year": 4, "students_count": 24, "department": "IT"},
    {"course_id": 10, "course_name": "IT_Subject_10", "year": 4, "students_count": 22, "department": "IT"},
    {"course_id": 11, "course_name": "IT_Subject_11", "year": 4, "students_count": 21, "department": "IT"},
    {"course_id": 12, "course_name": "IT_Subject_12", "year": 1, "students_count": 31, "department": "IT"},
    {"course_id": 13, "course_name": "IT_Subject_13", "year": 1, "students_count": 40, "department": "IT"},
    {"course_id": 14, "course_name": "IT_Subject_14", "year": 4, "students_count": 32, "department": "IT"},
    {"course_id": 15, "course_name": "IT_Subject_15", "year": 1, "students_count": 30, "department": "IT"},
    {"course_id": 16, "course_name": "IT_Subject_16", "year": 2, "students_count": 32, "department": "IT"},
    {"course_id": 17, "course_name": "IT_Subject_17", "year": 4, "students_count": 42, "department": "IT"},
    {"course_id": 18, "course_name": "IT_Subject_18", "year": 4, "students_count": 25, "department": "IT"},
    {"course_id": 19, "course_name": "IT_Subject_19", "year": 1, "students_count": 30, "department": "IT"},
    {"course_id": 20, "course_name": "IT_Subject_20", "year": 4, "students_count": 34, "department": "IT"},
]

# 3. Slots Sheet (Distinct exam dates & sessions)
slots_data = [
    {"slot_id": 1, "date": "2026-11-02", "session": "Morning"},
    {"slot_id": 2, "date": "2026-11-02", "session": "Afternoon"},
    {"slot_id": 3, "date": "2026-11-03", "session": "Morning"},
    {"slot_id": 4, "date": "2026-11-03", "session": "Afternoon"},
    {"slot_id": 5, "date": "2026-11-04", "session": "Morning"},
    {"slot_id": 6, "date": "2026-11-04", "session": "Afternoon"},
    {"slot_id": 7, "date": "2026-11-05", "session": "Morning"},
    {"slot_id": 8, "date": "2026-11-05", "session": "Afternoon"},
    {"slot_id": 9, "date": "2026-11-06", "session": "Morning"},
    {"slot_id": 10, "date": "2026-11-06", "session": "Afternoon"},
]

# 4. Rooms Sheet
rooms_data = [
    {"room_id": "IT-101", "capacity": 45, "department": "IT"},
    {"room_id": "IT-102", "capacity": 45, "department": "IT"},
    {"room_id": "IT-201", "capacity": 50, "department": "IT"},
    {"room_id": "IT-202", "capacity": 50, "department": "IT"},
    {"room_id": "IT-301", "capacity": 40, "department": "IT"},
    {"room_id": "IT-302", "capacity": 40, "department": "IT"},
    {"room_id": "Audi-1", "capacity": 100, "department": "General"},
    {"room_id": "Audi-2", "capacity": 100, "department": "General"},
]

# 5. Student Courses Sheet
# Generating enrollments for students according to their year and courses
students_data = []

# Year 1 courses: 1, 12, 13, 15, 19
for s in range(1, 41):
    sid = f"STU_Y1_{s:02d}"
    for c in [1, 12, 13, 15, 19]:
        students_data.append({"student_id": sid, "course_id": c})

# Year 2 courses: 2, 3, 4, 7, 16
for s in range(1, 33):
    sid = f"STU_Y2_{s:02d}"
    for c in [2, 3, 4, 7, 16]:
        students_data.append({"student_id": sid, "course_id": c})

# Year 3 course: 6
for s in range(1, 34):
    sid = f"STU_Y3_{s:02d}"
    students_data.append({"student_id": sid, "course_id": 6})

# Year 4 courses: 5, 8, 9, 10, 11, 14, 17, 18, 20
for s in range(1, 43):
    sid = f"STU_Y4_{s:02d}"
    for c in [5, 8, 9, 10, 11, 14, 17, 18, 20]:
        students_data.append({"student_id": sid, "course_id": c})

def generate():
    target = OUTPUT_FILE
    try:
        with pd.ExcelWriter(target, engine="openpyxl") as writer:
            pd.DataFrame(students_data).to_excel(writer, sheet_name="Student_Courses", index=False)
            pd.DataFrame(courses_data).to_excel(writer, sheet_name="Courses", index=False)
            pd.DataFrame(rooms_data).to_excel(writer, sheet_name="Rooms", index=False)
            pd.DataFrame(teachers_data).to_excel(writer, sheet_name="Teachers", index=False)
            pd.DataFrame(slots_data).to_excel(writer, sheet_name="Slots", index=False)
        print(f"Generated sample Excel dataset matching user's structure: {target}")
    except PermissionError:
        target = os.path.join(BASE_DIR, "sample_exam_data_v2.xlsx")
        with pd.ExcelWriter(target, engine="openpyxl") as writer:
            pd.DataFrame(students_data).to_excel(writer, sheet_name="Student_Courses", index=False)
            pd.DataFrame(courses_data).to_excel(writer, sheet_name="Courses", index=False)
            pd.DataFrame(rooms_data).to_excel(writer, sheet_name="Rooms", index=False)
            pd.DataFrame(teachers_data).to_excel(writer, sheet_name="Teachers", index=False)
            pd.DataFrame(slots_data).to_excel(writer, sheet_name="Slots", index=False)
        print(f"File locked by Excel. Generated copy at: {target}")

if __name__ == "__main__":
    generate()
