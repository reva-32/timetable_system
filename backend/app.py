from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import sys

# Import all functions from main.py
from main import (
    load_data,
    build_conflict_graph,
    dsatur_coloring,
    load_room_data,
    allocate_rooms,
    load_teacher_data,
    build_duties,
    assign_teachers
)

# ================================================
# FLASK APP SETUP
# ================================================

app = Flask(
    __name__,
    static_folder=os.path.join("..", "frontend"),  # serve frontend files
)
CORS(app)  # allow cross-origin requests from frontend

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ================================================
# ROUTES
# ================================================

@app.route("/")
def index():
    """Serve the frontend HTML page."""
    return send_from_directory(app.static_folder, "index.html")


@app.route("/index.css")
def css():
    return send_from_directory(app.static_folder, "index.css")

@app.route("/generate", methods=["POST"])
def generate():
    """
    Main endpoint — receives uploaded Excel file,
    runs all 3 stages, returns JSON with results.
    """

    # -----------------------------------------------
    # 1. Receive and save uploaded file
    # -----------------------------------------------
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported"}), 400

    filepath = os.path.join(UPLOAD_FOLDER, "exam_data.xlsx")
    file.save(filepath)

    # -----------------------------------------------
    # 2. Run Stage 1 — Course to Slot Assignment
    # -----------------------------------------------
    try:
        student_courses, slot_meta, course_meta, enrolled_counts = load_data(filepath)
    except Exception as e:
        return jsonify({"error": f"Error loading data: {str(e)}"}), 500

    graph    = build_conflict_graph(student_courses)
    coloring = dsatur_coloring(graph)

    # Build final exam schedule
    final_data = []
    for course_id, slot_idx in coloring.items():
        slot_info = slot_meta.get(slot_idx, {
            "display_id": slot_idx + 1,
            "date": "TBD",
            "session": "TBD"
        })
        c_info = course_meta.get(course_id, {
            "course_name": "Unknown",
            "year": "N/A",
            "students_count": 0,
            "department": "General"
        })
        enrolled = enrolled_counts.get(course_id, 0)

        final_data.append({
            "course_id"        : course_id,
            "course_name"      : c_info["course_name"],
            "year"             : c_info["year"],
            "department"       : c_info["department"],
            "slot"             : slot_info["display_id"],
            "date"             : slot_info["date"],
            "session"          : slot_info["session"],
            "declared_students": c_info["students_count"],
            "enrolled_students": enrolled,
        })

    final_data.sort(key=lambda x: (x["slot"], x["department"], x["course_id"]))

    # -----------------------------------------------
    # 3. Run Stage 2 — Room Allocation
    # -----------------------------------------------
    try:
        rooms = load_room_data(filepath)
        room_assignments, unallocated = allocate_rooms(final_data, rooms)
    except Exception as e:
        return jsonify({"error": f"Error in room allocation: {str(e)}"}), 500

    # -----------------------------------------------
    # 4. Run Stage 3 — Teacher Assignment
    # -----------------------------------------------
    try:
        teachers = load_teacher_data(filepath)
        duties   = build_duties(final_data)
        assignments, unassigned, teacher_duty_count = assign_teachers(teachers, duties)
    except Exception as e:
        return jsonify({"error": f"Error in teacher assignment: {str(e)}"}), 500

    # -----------------------------------------------
    # 5. Build and return JSON response
    # -----------------------------------------------

    # Conflict check
    violations = []
    for student, courses in student_courses.items():
        slots_taken = {}
        for course in courses:
            slot = coloring.get(course)
            if slot is None:
                continue
            if slot in slots_taken:
                violations.append(f"{student}: {course} clashes with {slots_taken[slot]}")
            else:
                slots_taken[slot] = course

    # Preference satisfaction
    preferred_count = sum(1 for a in assignments if a["cost"] == 1)
    pref_pct = round(preferred_count / len(assignments) * 100, 1) if assignments else 0

    response = {
        # Summary stats
        "summary": {
            "total_courses"  : len(final_data),
            "total_students" : len(student_courses),
            "slots_used"     : max(coloring.values()) + 1 if coloring else 0,
            "conflict_edges" : sum(len(v) for v in graph.values()) // 2,
            "conflicts_found": len(violations),
            "rooms_allocated": len(room_assignments),
            "rooms_failed"   : len(unallocated),
            "duties_assigned": len(assignments),
            "duties_failed"  : len(unassigned),
            "pref_satisfaction": pref_pct,
        },

        # Stage 1 output
        "timetable": final_data,

        # Stage 2 output
        "room_allocation": room_assignments,

        # Stage 3 output
        "teacher_duties": [
            {
                "duty_id"      : a["duty_id"],
                "slot"         : a["slot"],
                "date"         : a["date"],
                "session"      : a["session"],
                "course_id"    : a["course_id"],
                "role"         : a["role_required"],
                "teacher_id"   : a["teacher_id"],
                "teacher_name" : a["teacher_name"],
                "preferred"    : a["cost"] == 1
            }
            for a in sorted(assignments, key=lambda x: (x["slot"], x["role_required"]))
        ],

        # Unassigned warnings
        "warnings": {
            "unallocated_rooms"   : [u["course_id"] for u in unallocated],
            "unassigned_duties"   : [u["duty_id"] for u in unassigned],
            "student_conflicts"   : violations
        }
    }

    return jsonify(response)


# ================================================
# RUN
# ================================================

if __name__ == "__main__":
    print("🚀 Starting Flask server...")
    print("   → Open http://127.0.0.1:5000 in your browser")
    app.run(debug=True, port=5000)