from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import io
from datetime import date, timedelta, datetime
import pdfplumber
from PyPDF2 import PdfReader

# Import all functions from main.py
from main import (
    load_data,
    build_conflict_graph,
    dsatur_coloring,
    allocate_rooms,
    load_room_data,
    build_duties,
    assign_teachers,
    load_teacher_data
)
# Import MongoDB functions
from db import (
    init_faculty,
    get_all_faculty_status,
    get_db,
    update_faculty_duty,
    check_reset_fairness
)
from flask import send_file
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename
import uuid
import os

app = Flask(
    __name__,
    static_folder=os.path.join("..", "frontend"),
)
CORS(app)


# ================================================
# ROUTES
# ================================================

@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")

@app.route("/index.css")
def css():
    return send_from_directory(app.static_folder, "index.css")


@app.route("/login", methods=["POST"])
def login():
    data = request.get_json()

    email    = data.get("email", "").strip()    if data else ""
    password = data.get("password", "").strip() if data else ""

    if not email:
        return jsonify({"error": "Email is required"}), 400
    if not password:
        return jsonify({"error": "Password is required"}), 400

    try:
        from db import get_db, check_password_hash
        database = get_db()

        user = database["teachers"].find_one({
            "$or": [
                {"email": email},
                {"teacher_id": email}
            ]
        })

        if user is None:
            return jsonify({"error": "Invalid credentials"}), 401

        # --- @pict.edu domain check ---
        if not email.lower().endswith("@pict.edu"):
            return jsonify({"error": "Only @pict.edu email addresses are allowed"}), 403

        # --- Password check ---
        stored_hash = user.get("password_hash")
        teacher_id  = user.get("teacher_id", user.get("email", "").split("@")[0])

        if stored_hash:
            # Normal path: validate against stored hash
            if not check_password_hash(stored_hash, password):
                return jsonify({"error": "Invalid credentials"}), 401
        else:
            # Legacy path: no password set yet → default password = teacher_id (e.g. "T1")
            if password != teacher_id:
                return jsonify({"error": "Invalid credentials"}), 401

        user_email = user.get("email", f"{teacher_id}@pict.edu")
        return jsonify({
            "name"      : user.get("name"),
            "role"      : user.get("role"),
            "is_admin"  : user.get("is_admin", False),
            "history"   : user.get("history", []),
            "identifier": user_email
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
        

@app.route("/change_password", methods=["POST"])
def change_password():
    data             = request.get_json()
    identifier       = (data.get("identifier", "") or "").strip()
    current_password = (data.get("current_password", "") or "").strip()
    new_password     = (data.get("new_password", "") or "").strip()

    if not identifier or not current_password or not new_password:
        return jsonify({"error": "All fields are required"}), 400
    if len(new_password) < 6:
        return jsonify({"error": "New password must be at least 6 characters"}), 400

    try:
        from db import get_db, check_password_hash, update_password
        database = get_db()

        user = database["teachers"].find_one({
            "$or": [{"email": identifier}, {"teacher_id": identifier}]
        })
        if user is None:
            return jsonify({"error": "User not found"}), 404

        stored_hash = user.get("password_hash")
        teacher_id  = user.get("teacher_id", user.get("email", "").split("@")[0])

        # Validate current password (same logic as login)
        if stored_hash:
            if not check_password_hash(stored_hash, current_password):
                return jsonify({"error": "Current password is incorrect"}), 401
        else:
            # Legacy default = teacher_id
            if current_password != teacher_id:
                return jsonify({"error": "Current password is incorrect"}), 401

        ok = update_password(identifier, new_password)
        if ok:
            return jsonify({"success": True})
        else:
            return jsonify({"error": "Failed to update password"}), 500

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/request_adjustment", methods=["POST"])
def request_adjustment():
    data = request.get_json() or {}
    identifier = (data.get('identifier','') or '').strip()
    # New payload supports direct apply: old_* fields identify existing duty entry
    old_slot = data.get('old_slot_id','')
    old_date = data.get('old_exam_date','')
    new_slot = data.get('new_slot_id','')
    new_date = data.get('new_exam_date','')
    new_role = data.get('new_role','')
    reason = data.get('reason','')
    apply_direct = bool(data.get('apply_direct', False))

    if not identifier or not old_slot or not old_date:
        return jsonify({"error": "identifier, old_slot_id and old_exam_date are required"}), 400

    try:
        database = get_db()
        adj_doc = {
            'identifier': identifier,
            'old_slot_id': str(old_slot),
            'old_exam_date': old_date,
            'new_slot_id': str(new_slot),
            'new_exam_date': new_date,
            'new_role': new_role,
            'reason': reason,
            'requested_at': datetime.utcnow().isoformat(),
            'status': 'applied' if apply_direct else 'pending'
        }

        # If apply_direct, attempt to validate ownership and update teacher history
        if apply_direct:
            # Find applicant teacher and the target teacher
            applicant_q = {"$or": [{"email": identifier}, {"teacher_id": identifier}]}
            applicant = database['teachers'].find_one(applicant_q)
            if not applicant:
                adj_doc['status'] = 'rejected'
                adj_doc['reason_internal'] = 'Applicant not found'
                database['adjustments'].insert_one(adj_doc)
                return jsonify({"error": "Applicant not found"}), 400

            # Validation: applicant must have a history entry matching old_slot & old_date
            def owns_entry(doc, slot, date):
                for h in doc.get('history', []):
                    if str(h.get('slot_id')) == str(slot) and str(h.get('exam_date')) == str(date):
                        return True
                return False

            if not owns_entry(applicant, old_slot, old_date):
                adj_doc['status'] = 'rejected'
                adj_doc['reason_internal'] = 'Applicant does not own the specified original duty'
                database['adjustments'].insert_one(adj_doc)
                return jsonify({"error": "Applicant does not own the specified original duty"}), 400

            # Find adjusted faculty (the one to swap with)
            adjusted = database['teachers'].find_one({"$or": [{"teacher_id": new_slot}, {"email": new_slot}, {"teacher_id": adj_doc.get('new_slot_id')}, {"email": adj_doc.get('new_slot_id')}]})
            # Note: new_slot currently holds slot id; adjusted faculty will be looked up later during swap application in generate()

        database['adjustments'].insert_one(adj_doc)
        print(f"Adjustment request recorded: {identifier} old_slot={old_slot} old_date={old_date} -> new_slot={new_slot} new_date={new_date} role={new_role}")
        return jsonify({"success": True})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/upload_adjustment_pdf', methods=['POST'])
def upload_adjustment_pdf():
    # Accepts multipart/form-data: 'file' (PDF) and 'identifier' (user email or teacher_id)
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    identifier = (request.form.get('identifier','') or '').strip()
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    try:
        import io as _io
        # Read PDF bytes
        fp = file.stream.read()
        # Extract text
        text = ''
        sig_present = False
        try:
            import io as _io
            with pdfplumber.open(_io.BytesIO(fp)) as pdf:
                for p in pdf.pages:
                    t = p.extract_text()
                    if t: text += '\n' + t
                    if p.images and len(p.images) > 0:
                        sig_present = True
        except Exception:
            # fallback to PyPDF2 text extraction
            try:
                reader = PdfReader(_io.BytesIO(fp))
                for p in reader.pages:
                    try:
                        t = p.extract_text() or ''
                        text += '\n' + t
                    except: continue
                # images detection not available in PyPDF2 easily
            except Exception:
                text = ''

        ltext = (text or '').lower()
        if not sig_present and ('signature' in ltext or 'signed' in ltext):
            sig_present = True

        # Parse Through HOD and Applicant Faculty fields if present
        import re
        hod = None
        m = re.search(r'Through\s*HOD[:\-\s]*([A-Za-z .]+)', text, re.IGNORECASE)
        if m:
            hod = m.group(1).strip()

        applicant_name = None
        m2 = re.search(r'Applicant\s*Faculty[:\-\s]*([A-Za-z0-9 @.()\-]+)', text, re.IGNORECASE)
        if m2:
            applicant_name = m2.group(1).strip()

        # Attempt to resolve applicant to teacher_id via DB
        db = get_db()
        applicant_id = None
        if identifier:
            q = {'$or': [{'email': identifier}, {'teacher_id': identifier}]}
            user = db['teachers'].find_one(q)
            if user:
                applicant_id = user.get('teacher_id') or user.get('email')
        if not applicant_id and applicant_name:
            # try to match by name
            user = db['teachers'].find_one({'name': {'$regex': applicant_name, '$options': 'i'}})
            if user:
                applicant_id = user.get('teacher_id') or user.get('email')

        # Parse swaps: look for lines that mention two teacher ids and two dates
        swaps = []
        lines = [ln.strip() for ln in (text or '').splitlines() if ln.strip()]
        tid_re = re.compile(r'\bT\d+\b', re.IGNORECASE)
        date_re = re.compile(r'\b\d{1,2}[-/]\w{3,9}[-/]\d{2,4}\b|\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b')
        sess_re = re.compile(r'\b(Morning|Afternoon|Evening|Practical|\d{1,2}:\d{2})\b', re.IGNORECASE)
        for ln in lines:
            tids = tid_re.findall(ln)
            dates = date_re.findall(ln)
            sess = sess_re.findall(ln)
            if len(tids) >= 2 and len(dates) >= 2:
                # assume order: applicant_tid, adjusted_tid, old_date, new_date (or similar)
                # pick first two tids and first two dates
                s = {
                    'applicant_id': applicant_id or (tids[0] if tids else ''),
                    'adjusted_id': tids[1] if len(tids) > 1 else '',
                    'old_date': dates[0] if len(dates) > 0 else '',
                    'new_date': dates[1] if len(dates) > 1 else '',
                    'old_session': sess[0] if sess else '',
                    'new_session': sess[1] if len(sess) > 1 else '' ,
                    'raw': ln
                }
                swaps.append(s)

        # If no swaps found via lines, try pairing global ids/dates
        if not swaps:
            tids_all = tid_re.findall(text)
            dates_all = date_re.findall(text)
            for i in range(min(len(tids_all)-1, len(dates_all)-1)):
                swaps.append({
                    'applicant_id': applicant_id or tids_all[i],
                    'adjusted_id': tids_all[i+1],
                    'old_date': dates_all[i],
                    'new_date': dates_all[i+1],
                    'old_session': '', 'new_session': '', 'raw': ''
                })

        # Require signature present for valid adjustment PDFs
        if not sig_present:
            return jsonify({'error': 'Signature not detected on PDF. A handwritten signature is required for duty adjustments.'}), 400

        # Allow any HOD value: prefer parsed HOD, otherwise accept `hod` provided in form data
        if not hod:
            hod = (request.form.get('hod', '') or '').strip()

        # Applicant identity must still be resolvable
        if not applicant_id:
            return jsonify({'error': 'Applicant identity could not be resolved; ensure you are logged in or PDF contains Applicant Faculty field.'}), 400

        # Normalize and enrich swap entries: ensure explicit fields for next/swapped duty
        import pandas as _pd
        def _norm(d):
            try:
                dt = _pd.to_datetime(str(d), dayfirst=True, errors='coerce')
                if _pd.isna(dt): return str(d)
                return dt.strftime('%d-%b-%Y')
            except Exception:
                return str(d)

        enriched_swaps = []
        for s in swaps:
            a_id = s.get('applicant_id') or applicant_id
            adj_id = s.get('adjusted_id')
            old_date_raw = s.get('old_date') or ''
            new_date_raw = s.get('new_date') or ''
            old_sess = (s.get('old_session') or '').strip()
            new_sess = (s.get('new_session') or '').strip()
            old_date = _norm(old_date_raw)
            new_date = _norm(new_date_raw)

            enriched_swaps.append({
                'applicant_id': a_id,
                'adjusted_id': adj_id,
                'swapped_duty_date': old_date,
                'swapped_duty_session': old_sess,
                'next_duty_date': new_date,
                'next_duty_session': new_sess,
                'raw': s.get('raw','')
            })

        # Do NOT save PDF bytes to disk. Store only parsed metadata and enriched swap fields.
        created_at = datetime.utcnow().isoformat()
        created_docs = []
        for s in enriched_swaps:
            # Create document with legacy keys (keep existing behavior) plus explicit swap fields
            doc = {
                'teacher1_id': s.get('applicant_id'),
                'teacher2_id': s.get('adjusted_id'),
                'teacher1_date': s.get('swapped_duty_date'),
                'teacher1_session': s.get('swapped_duty_session'),
                'teacher2_date': s.get('next_duty_date'),
                'teacher2_session': s.get('next_duty_session'),
                'applicant_id': s.get('applicant_id'),
                'adjusted_id': s.get('adjusted_id'),
                'swapped_duty_date': s.get('swapped_duty_date'),
                'swapped_duty_session': s.get('swapped_duty_session'),
                'next_duty_date': s.get('next_duty_date'),
                'next_duty_session': s.get('next_duty_session'),
                'reason': request.form.get('reason',''),
                'status': 'Pending',
                'submitted_at': created_at,
                'approved_by': None,
                'approved_at': None,
                'raw': s.get('raw',''),
                'pdf_saved': False
            }
            res = db['duty_adjustments'].insert_one(doc)
            doc['_id'] = str(res.inserted_id)
            created_docs.append(doc)

            # Also add a clear formatted view required by validator dashboards
            # Keep keys explicit: who it's swapped with, session, and dates
            formatted_view = {
                'identifier': identifier or '',
                'applicant_id': s.get('applicant_id') or '',
                'through_hod': hod or '',
                'swapped_with': s.get('adjusted_id') or '',
                'swapped_session': s.get('swapped_duty_session') or s.get('old_session','') or '',
                'swapped_duty_date': s.get('swapped_duty_date') or '',
                'next_duty_date': s.get('next_duty_date') or '',
                'next_duty_session': s.get('next_duty_session') or s.get('new_session','') or '',
                'signature_present': bool(sig_present),
                'created_at': created_at,
                'status': 'pending'
            }
            try:
                db['duty_adjustments'].update_one({'_id': res.inserted_id}, {'$set': {'formatted_view': formatted_view}})
            except Exception:
                # best-effort: don't fail the whole request if update fails
                pass

        # Also save the parsed adjustments for auditing (without saving PDF bytes)
        audit = {
            'identifier': identifier,
            'applicant_id': applicant_id,
            'through_hod': hod,
            'applicant_name': applicant_name,
            'signature_present': True,
            'swaps': enriched_swaps,
            'created_at': created_at,
            'status': 'parsed',
            'pdf_saved': False
        }
        db['adjustments'].insert_one(audit)

        return jsonify({'success': True, 'created': created_docs})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/adjustments', methods=['GET'])
def get_adjustments():
    """Return parsed duty adjustments (no PDF bytes) for review or dashboards."""
    try:
        db = get_db()
        docs = list(db['duty_adjustments'].find({}).sort('submitted_at', -1).limit(200))
        out = []
        for d in docs:
            # convert ObjectId to str and remove heavy/unnecessary fields
            d = dict(d)
            d['_id'] = str(d.get('_id'))
            d.pop('raw', None)
            d.pop('pdf_path', None)
            d.pop('pdf_saved', None)
            out.append(d)
        return jsonify({'adjustments': out})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route("/generate", methods=["POST"])
def generate():
    """
    Main pipeline: Filter -> DSATUR -> Rooms -> Teachers
    """
    # -----------------------------------------------
    # 0. Get Config from request
    # -----------------------------------------------
    year_filter    = request.form.get("year", "ALL")
    branch_filter  = request.form.get("branch", "ALL")
    exam_type      = request.form.get("exam_type", "Endsem")
    selected_slots_raw = request.form.get("selected_slots", "").strip()
    frontend_slots     = [s.strip() for s in selected_slots_raw.split(",") if s.strip()] if selected_slots_raw else []
    # NOTE: Date selection removed from UI. Exam dates are read from the Courses sheet

    # Default slot labels per exam type (used when user picks no slots in UI)
    EXAM_TYPE_DEFAULT_SLOTS = {
        "Insem"     : ["10:00 – 11:00", "14:00 – 15:00", "16:00 – 17:00"],
        "Endsem"    : ["10:00 – 12:30", "14:00 – 16:30"],
        "Practical" : ["12:00 – 14:00", "14:30 – 16:30", "17:00 – 19:00"],
    }
    if not frontend_slots:
        frontend_slots = EXAM_TYPE_DEFAULT_SLOTS.get(exam_type, ["10:00 – 12:30", "14:00 – 16:30"])

    # -----------------------------------------------
    # 1. Receive and validate uploaded file
    # -----------------------------------------------
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    file_content = file.read()
    temp_path = "temp_data.xlsx"
    with open(temp_path, "wb") as f:
        f.write(file_content)

    # -----------------------------------------------
    # 1b. Validate required Excel sheets are present
    # -----------------------------------------------
    REQUIRED_SHEETS = ["Student_Courses", "Courses", "Rooms", "Teachers"]
    try:
        import openpyxl as _openpyxl
        _wb = _openpyxl.load_workbook(temp_path, read_only=True)
        missing = [s for s in REQUIRED_SHEETS if s not in _wb.sheetnames]
        _wb.close()
        if missing:
            return jsonify({
                "error": (
                    f"Missing required sheet(s) in uploaded Excel: {', '.join(missing)}. "
                    f"Expected sheets: {', '.join(REQUIRED_SHEETS)}."
                )
            }), 400
    except Exception as val_err:
        return jsonify({"error": f"Could not open Excel file: {str(val_err)}"}), 400

    # -----------------------------------------------
    # 2. Load and Filter Data
    # -----------------------------------------------
    try:
        student_courses, loaded_slots, course_meta, enrolled_counts = load_data(temp_path)
    except Exception as e:
        return jsonify({"error": f"Error loading data: {str(e)}"}), 500

    # Helper: normalize date strings to dd-Mon-YYYY (e.g., 26-Feb-2026)
    import pandas as _pd
    def _normalize_date(s):
        if not s: return ''
        try:
            dt = _pd.to_datetime(str(s), dayfirst=True, errors='coerce')
            if pd.isna(dt): return str(s)
            return dt.strftime('%d-%b-%Y')
        except Exception:
            return str(s)

    # 2a. Filter data based on year and branch
    if year_filter != "ALL" or branch_filter != "ALL":
        year_map = {
            "FY"     : ["1", "FY", "FIRST"],
            "SY"     : ["2", "SY", "SECOND"],
            "TY"     : ["3", "TY", "THIRD"],
            "B.TECH" : ["4", "BT", "BTECH", "FINAL", "4TH"]
        }
        
        # DEBUG PRINTS
        if course_meta:
            first_c = next(iter(course_meta.values()))
            print(f"DEBUG: Excel Year is {first_c['year']} (type: {type(first_c['year'])}) and Frontend Year is {year_filter} (type: {type(year_filter)})")


        def is_match(val, target, category):
            v = str(val).strip().upper()
            t = str(target).strip().upper()
            if t == "ALL": return True

            if category == "year":
                allowed_vals = year_map.get(t, [t])
                return v in allowed_vals

            if category == "branch":
                branch_keywords = {
                    "CE"   : ["COMPUTER", "CE", "COMP"],
                    "ENTC" : ["ENTC", "E&TC", "ELECTRONICS"],
                    "AIDS" : ["AIDS", "AI&DS", "AI", "DS"],
                    "IT"   : ["IT", "INFORMATION TECHNOLOGY", "INFO TECH"]
                }
                allowed_branches = branch_keywords.get(t, [t])
                return any(m in v for m in allowed_branches) or t in v

            return t in v

        filtered_courses = {
            cid: info for cid, info in course_meta.items()
            if is_match(info["year"], year_filter, "year") and
               is_match(info["department"], branch_filter, "branch")
        }

        if not filtered_courses:
            return jsonify({
                "error": (
                    f"Filter Mismatch: No courses found for Branch '{branch_filter}' "
                    f"and Year '{year_filter}'. "
                    "Check if your Excel 'year' column uses 1, 2, 3, 4 instead of FY, SY, TY."
                )
            }), 400

        filtered_student_courses = {}
        for sid, courses in student_courses.items():
            valid_courses = [c for c in courses if c in filtered_courses]
            if valid_courses:
                filtered_student_courses[sid] = valid_courses

        course_meta     = filtered_courses
        student_courses = filtered_student_courses
        
        enrolled_counts = {c: count for c, count in enrolled_counts.items() if c in course_meta}
        print(f"Filter Applied: Year={year_filter}, Branch={branch_filter}. Remaining Courses: {len(course_meta)}")

    # 2b. Build conflict graphs and assign slots
    use_dynamic_dates = any(not info.get("exam_date") for info in course_meta.values())

    slot_meta = {}
    final_data = []

    if use_dynamic_dates:
        print("Using dynamic slot scheduling from Slots sheet (or generated).")
        # Build global conflict graph for all courses
        subgraph = build_conflict_graph(student_courses, all_courses=list(course_meta.keys()))
        coloring = dsatur_coloring(subgraph)

        orig_colors = sorted(set(coloring.values())) if coloring else []
        remap = {orig: idx for idx, orig in enumerate(orig_colors)}

        has_loaded_slots = bool(loaded_slots)
        start_date = date.today()
        L = len(frontend_slots)
        if L == 0:
            frontend_slots = ["Morning", "Afternoon"]
            L = 2

        color_to_global = {}
        for orig_color in orig_colors:
            idx = remap[orig_color]
            color_to_global[orig_color] = idx
            
            if has_loaded_slots and idx in loaded_slots:
                slot_info = loaded_slots[idx]
                slot_meta[idx] = {
                    "display_id": slot_info.get("display_id", idx + 1),
                    "date": slot_info.get("date"),
                    "session": slot_info.get("session")
                }
            else:
                day_offset = idx // L
                slot_in_day = idx % L
                slot_date = start_date + timedelta(days=day_offset)
                date_str = slot_date.strftime("%d-%b-%Y")
                slot_meta[idx] = {
                    "display_id": slot_in_day + 1,
                    "date": date_str,
                    "session": frontend_slots[slot_in_day]
                }

        for course_id in course_meta.keys():
            c_info = course_meta[course_id]
            enrolled = enrolled_counts.get(course_id, 0)
            orig_color = coloring.get(course_id, 0)
            new_color = remap.get(orig_color, 0)
            global_slot = color_to_global.get(orig_color, None)
            
            slot_info = slot_meta.get(new_color, {})
            display_slot = slot_info.get("display_id", 1)
            exam_date = slot_info.get("date", "TBD")
            session = slot_info.get("session", "TBD")

            final_data.append({
                "course_id": course_id,
                "course_name": c_info.get("course_name", "Unknown"),
                "year": c_info.get("year", "N/A"),
                "department": c_info.get("department", "General"),
                "slot": display_slot,
                "date": exam_date,
                "session": session,
                "declared_students": c_info.get("students_count", 0),
                "enrolled_students": enrolled,
                "_global_slot": global_slot
            })
        print(f"DEBUG: Dynamic final_data length: {len(final_data)}")
    else:
        # Pre-assigned dates mode: group by exam date from Courses sheet
        print("Using pre-assigned exam dates from Courses sheet.")
        from collections import defaultdict
        date_groups = defaultdict(list)
        for cid, info in course_meta.items():
            date_groups[info["exam_date"]].append(cid)

        global_slot_counter = 0
        for exam_date, courses_on_date in sorted(date_groups.items()):
            subgraph = build_conflict_graph(student_courses, all_courses=list(courses_on_date))
            coloring = dsatur_coloring(subgraph)

            orig_colors = sorted(set(coloring.values())) if coloring else []
            remap = {orig: idx for idx, orig in enumerate(orig_colors)}

            color_to_global = {}
            L = len(frontend_slots)
            for orig_color in orig_colors:
                new_color = remap[orig_color]
                color_to_global[orig_color] = global_slot_counter
                session_val = frontend_slots[new_color % L] if L > 0 else None
                slot_meta[global_slot_counter] = {
                    "display_id": new_color + 1,
                    "date": exam_date,
                    "session": session_val
                }
                global_slot_counter += 1

            for course_id in courses_on_date:
                c_info = course_meta.get(course_id, {})
                enrolled = enrolled_counts.get(course_id, 0)
                orig_color = coloring.get(course_id, 0)
                new_color = remap.get(orig_color, 0)
                global_slot = color_to_global.get(orig_color, None)
                display_slot = new_color + 1
                session = None
                if global_slot is not None and global_slot in slot_meta:
                    session = slot_meta[global_slot]["session"]
                if not session:
                    session = c_info.get("session") or (frontend_slots[0] if L>0 else "TBD")

                final_data.append({
                    "course_id": course_id,
                    "course_name": c_info.get("course_name", "Unknown"),
                    "year": c_info.get("year", "N/A"),
                    "department": c_info.get("department", "General"),
                    "slot": display_slot,
                    "date": exam_date,
                    "session": session,
                    "declared_students": c_info.get("students_count", 0),
                    "enrolled_students": enrolled,
                    "_global_slot": global_slot
                })
        print(f"DEBUG: Pre-assigned final_data length: {len(final_data)}")

    # -----------------------------------------------
    # 2b. Adjust exam dates (move same-day exams to different days where possible)
    # -----------------------------------------------
    try:
        from main import adjust_exam_dates
        adjust_exam_dates(final_data, student_courses, slot_meta)
    except Exception as adj_err:
        print(f"adjust_exam_dates warning (non-fatal): {adj_err}")

    # -----------------------------------------------
    # 3. Run Stage 2 — Room Allocation
    # -----------------------------------------------
    # Build global conflict graph (used in summary)
    try:
        graph = build_conflict_graph(student_courses, all_courses=list(course_meta.keys()))
    except Exception:
        graph = {}
    try:
        rooms = load_room_data(temp_path)
        if branch_filter != "ALL":
            # Include rooms matching the branch OR rooms marked as 'General' (shared across branches)
            rooms = [
                r for r in rooms
                if is_match(r.get("department", "General"), branch_filter, "branch")
                or str(r.get("department", "")).strip().upper() in ("GENERAL", "COMMON", "SHARED", "")
            ]
        room_assignments, unallocated = allocate_rooms(final_data, rooms)
    except Exception as e:
        return jsonify({"error": f"Error in room allocation: {str(e)}"}), 500

    # Build a mapping of assigned rooms for quick lookup: (course_id, slot, date) -> rooms_assigned
    room_map = {}
    for r in room_assignments:
        key = (str(r.get("course_id")), str(r.get("slot")), str(r.get("date")))
        room_map[key] = r.get("rooms_assigned", "TBD")

    # Attach room assignment to timetable rows
    for row in final_data:
        key = (str(row.get("course_id")), str(row.get("slot")), str(row.get("date")))
        row["room_assigned"] = room_map.get(key, "TBD")

    # -----------------------------------------------
    # 4. Run Stage 3 — Teacher Assignment
    # -----------------------------------------------
    faculty_status = {}
    try:
        teachers = load_teacher_data(temp_path)
        if branch_filter != "ALL":
            # Include teachers from the selected branch OR those marked as 'General' (cross-department)
            teachers = {
                tid: t for tid, t in teachers.items()
                if is_match(t.get("department", "General"), branch_filter, "branch")
                or str(t.get("department", "")).strip().upper() in ("GENERAL", "COMMON", "SHARED", "")
            }
            
        try:
            init_faculty(teachers)
            faculty_status = get_all_faculty_status()
            
            # Filter out admin teachers from the teachers pool
            teachers = {
                tid: t for tid, t in teachers.items()
                if not faculty_status.get(tid, {}).get("is_admin", False)
            }
            
            # Extract db_duty_counts
            db_duty_counts = {}
            for tid, f in faculty_status.items():
                dc = f.get("duty_counts")
                if not dc:
                    flat = f.get("duty_count", 0)
                    role = f.get("role", "Junior")
                    role_key = role.lower() if role.lower() in ["junior", "senior", "squad"] else "junior"
                    dc = {
                        "squad": flat if role_key == "squad" else 0,
                        "junior": flat if role_key == "junior" else 0,
                        "senior": flat if role_key == "senior" else 0
                    }
                db_duty_counts[tid] = dc

            fairness_map   = {tid: f["has_served_high_role"] for tid, f in faculty_status.items()}
        except Exception as db_err:
            print(f"DB Warning: {db_err}")
            fairness_map   = None
            db_duty_counts = None

        duties = build_duties(final_data, room_assignments=room_assignments)
        assignments, unassigned, teacher_duty_count = assign_teachers(
            teachers, duties, fairness_map=fairness_map, db_duty_counts=db_duty_counts
        )
        # -----------------------
        # Apply parsed duty adjustment swaps recorded via PDF uploads
        # -----------------------
        try:
            db = get_db()
            parsed_adjs = list(db['adjustments'].find({"status": "parsed"}))
            if parsed_adjs:
                # Helper to normalize dates
                import pandas as _pd
                def _norm(d):
                    try:
                        dt = _pd.to_datetime(str(d), dayfirst=True, errors='coerce')
                        if pd.isna(dt): return str(d)
                        return dt.strftime('%d-%b-%Y')
                    except Exception:
                        return str(d)

                # Work on assignments in-place
                for adj in parsed_adjs:
                    applicant = adj.get('applicant_id') or adj.get('identifier')
                    hod = adj.get('through_hod')
                    sig = adj.get('signature_present', False)
                    swaps = adj.get('swaps', [])
                    applied_any = False
                    # Require applicant identity and signature; HOD is optional
                    if not applicant or not sig:
                        # mark rejected
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$set': {'status': 'rejected', 'rejected_reason': 'missing applicant or signature'}})
                        continue

                    for s in swaps:
                        # support both legacy keys (old_date/new_date) and enriched keys
                        old_d = _norm(s.get('old_date') or s.get('swapped_duty_date'))
                        new_d = _norm(s.get('new_date') or s.get('next_duty_date'))
                        old_sess = (s.get('old_session') or s.get('swapped_duty_session') or '').strip().lower()
                        new_sess = (s.get('new_session') or s.get('next_duty_session') or '').strip().lower()
                        adjusted_id = s.get('adjusted_id')

                        # find assignment entries
                        a_idx = next((i for i, a in enumerate(assignments) if (str(a.get('teacher_id')) == str(applicant) or str(a.get('teacher_name')).lower() == str(applicant).lower()) and _norm(a.get('date')) == old_d and (old_sess in str(a.get('session','')).lower() or old_sess=='')), None)
                        b_idx = next((i for i, a in enumerate(assignments) if (str(a.get('teacher_id')) == str(adjusted_id) or str(a.get('teacher_name')).lower() == str(adjusted_id).lower()) and _norm(a.get('date')) == new_d and (new_sess in str(a.get('session','')).lower() or new_sess=='')), None)

                        if a_idx is None or b_idx is None:
                            # cannot apply this swap; log skip
                            db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'skipped': {'swap': s, 'reason': 'matching assignment not found'}}})
                            continue

                        if assignments[a_idx]['teacher_id'] == assignments[b_idx]['teacher_id']:
                            db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'skipped': {'swap': s, 'reason': 'same teacher'}}})
                            continue

                        # Check duplicate/conflict: ensure swapping does not create duplicate assignment for either teacher at the other's slot
                        # For applicant, ensure they don't already have another assignment at the new slot (excluding a_idx)
                        def has_conflict(tid, date, slot):
                            return any(True for i,a in enumerate(assignments) if i!=a_idx and str(a.get('teacher_id'))==str(tid) and _norm(a.get('date'))==date and a.get('slot')==slot)

                        a_slot = assignments[a_idx].get('slot')
                        b_slot = assignments[b_idx].get('slot')

                        if has_conflict(assignments[a_idx]['teacher_id'], new_d, b_slot) or has_conflict(assignments[b_idx]['teacher_id'], old_d, a_slot):
                            db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'skipped': {'swap': s, 'reason': 'would create conflict'}}})
                            continue

                        # Perform swap of teacher_id and teacher_name
                        t1_id = assignments[a_idx]['teacher_id']
                        t1_name = assignments[a_idx]['teacher_name']
                        t2_id = assignments[b_idx]['teacher_id']
                        t2_name = assignments[b_idx]['teacher_name']

                        assignments[a_idx]['teacher_id'] = t2_id
                        assignments[a_idx]['teacher_name'] = t2_name
                        assignments[b_idx]['teacher_id'] = t1_id
                        assignments[b_idx]['teacher_name'] = t1_name

                        # record applied swap with normalized fields
                        applied_record = {
                            'applicant_id': applicant,
                            'adjusted_id': adjusted_id,
                            'swapped_duty_date': old_d,
                            'swapped_duty_session': old_sess,
                            'next_duty_date': new_d,
                            'next_duty_session': new_sess,
                            'raw': s.get('raw','')
                        }
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$push': {'applied': {'swap': applied_record, 'applied_at': datetime.utcnow().isoformat()}}})
                        applied_any = True

                    # finalize status
                    if applied_any:
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$set': {'status': 'applied', 'applied_at': datetime.utcnow().isoformat()}})
                    else:
                        # if none applied, mark rejected if had skipped entries
                        db['adjustments'].update_one({'_id': adj['_id']}, {'$set': {'status': 'skipped'}})
        except Exception as adj_err:
            print(f"Adjustment apply warning: {adj_err}")
    except Exception as e:
        return jsonify({"error": f"Error in teacher assignment: {str(e)}"}), 500
    # -----------------------
    # Sanitize outputs: convert numpy types to native Python types and ensure strings
    # -----------------------
    def _clean_row(r):
        out = {}
        for k, v in r.items():
            # drop internal-only keys
            if k.startswith("_"): continue
            try:
                if isinstance(v, (int, float, str, bool)) or v is None:
                    out[k] = v
                else:
                    out[k] = int(v) if (hasattr(v, 'item') and isinstance(v.item(), (int,))) else str(v)
            except Exception:
                out[k] = str(v)
        return out

    final_data = [_clean_row(r) for r in final_data]
    room_assignments = [_clean_row(r) for r in room_assignments]
    # -----------------------------------------------
    # 5. Build and return JSON response
    # -----------------------------------------------
    # Build mapping course_id -> (date, slot_display)
    course_slot_map = {r["course_id"]: (r["date"], r["slot"]) for r in final_data}

    violations = []
    for student, courses in student_courses.items():
        slots_taken = {}
        for course in courses:
            cs = course_slot_map.get(course)
            if not cs: continue
            date_taken, slot_taken = cs
            key = (date_taken, slot_taken)
            if key in slots_taken:
                violations.append(f"{student}: {course} and {slots_taken[key]} on {date_taken} Slot {slot_taken}")
            else:
                slots_taken[key] = course

    preferred_count = sum(1 for a in assignments if a.get("cost", 9999) <= 1)
    pref_pct = round(preferred_count / len(assignments) * 100, 1) if assignments else 0

    response = {
        "summary": {
            "total_courses"    : len(final_data),
            "total_students"   : len(student_courses),
            # Number of actual exam sessions (unique date + display slot pairs)
            "slots_used"       : len({(r["date"], r["slot"]) for r in final_data}),
            "conflict_edges"   : sum(len(v) for v in graph.values()) // 2,
            "conflicts_found"  : len(violations),
            "rooms_allocated"  : len(room_assignments),
            "rooms_failed"     : len(unallocated),
            "duties_assigned"  : len(assignments),
            "duties_failed"    : len(unassigned),
            "pref_satisfaction": pref_pct,
        },
        "timetable": final_data,
        "room_allocation": room_assignments,
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
                "last_role"    : faculty_status.get(a["teacher_id"], {}).get("last_role", "N/A"),
                "is_priority"  : "Yes" if not faculty_status.get(a["teacher_id"], {}).get("has_served_high_role", True) else "No",
                "room_assigned" : (
                    room_map.get((str(a.get("course_id")), str(a.get("slot")), str(a.get("date"))))
                    if a.get("course_id") != "ALL" else ("Control" if a.get("role_required") == "Senior" else "Roaming")
                )
            }
            for a in sorted(assignments, key=lambda x: (x["date"], x["slot"], x["role_required"]))
        ],
        "warnings": {
            "unallocated_rooms" : [u["course_id"] for u in unallocated],
            "unassigned_duties" : [u["duty_id"] for u in unassigned],
            "student_conflicts" : violations
        }
    }

    return jsonify(response)


@app.route("/export_excel", methods=["POST"])
def export_excel():
    """Accepts JSON with keys 'timetable', 'room_allocation', 'teacher_duties' and returns a styled xlsx."""
    data = request.get_json()
    if not data:
        return jsonify({"error": "JSON body required"}), 400

    tables = {
        'Timetable': data.get('timetable', []),
        'Room Allocation': data.get('room_allocation', []),
        'Teacher Duties': data.get('teacher_duties', [])
    }

    # Build room_map from Room Allocation for merging into other sheets
    room_map = {}
    for r in tables.get('Room Allocation', []):
        key = (str(r.get('course_id')), str(r.get('slot')), str(r.get('date')))
        room_map[key] = r.get('rooms_assigned', '')

    # Merge Room Assigned into Timetable
    tt = []
    for row in tables.get('Timetable', []):
        key = (str(row.get('course_id')), str(row.get('slot')), str(row.get('date')))
        new = dict(row)
        new['Room Assigned'] = room_map.get(key, '')
        tt.append(new)
    tables['Timetable'] = tt

    # Prepare Teacher Duties: drop 'preferred' and add 'Room Assigned'
    td = []
    for row in tables.get('Teacher Duties', []):
        new = {k: v for k, v in row.items() if k != 'preferred'}
        key = (str(row.get('course_id')), str(row.get('slot')), str(row.get('date')))
        if row.get('course_id') == 'ALL':
            # infer control/roaming for senior/squad duties
            role = row.get('role') or row.get('role_required')
            new['Room Assigned'] = 'Control' if role == 'Senior' else 'Roaming'
        else:
            new['Room Assigned'] = room_map.get(key, '')
        td.append(new)
    tables['Teacher Duties'] = td

    # Build workbook in memory
    out = BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for name, rows in tables.items():
            df = pd.DataFrame(rows)
            if df.empty:
                df = pd.DataFrame(columns=["Empty"])
            df.to_excel(writer, sheet_name=name[:31], index=False)
        # Note: writer.save() is deprecated in pandas >= 2.0.
        # The context manager (with block) handles saving automatically on __exit__.
    out.seek(0)

    # Post-process with openpyxl for styling
    wb = load_workbook(out)
    thin = Side(border_style="thin", color="000000")
    for ws in wb.worksheets:
        # Header style
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Column widths: approximate by max length
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is None: continue
                v = str(cell.value)
                if len(v) > max_len: max_len = len(v)
            adjusted_width = (max_len + 2)
            ws.column_dimensions[col_letter].width = adjusted_width

        # Borders
        for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
            for cell in row:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Freeze header
        ws.freeze_panes = 'A2'

    # Save workbook back to bytes
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    return send_file(bio, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='timetable_export.xlsx')


@app.route('/admin/duty_adjustments', methods=['GET'])
def admin_list_adjustments():
    status = request.args.get('status')
    db = get_db()
    q = {}
    if status:
        q['status'] = status
    docs = list(db['duty_adjustments'].find(q))
    for d in docs:
        d['_id'] = str(d.get('_id'))
    return jsonify({'results': docs})


@app.route('/duty_adjustments', methods=['GET'])
def list_my_adjustments():
    identifier = (request.args.get('identifier') or '').strip()
    if not identifier:
        return jsonify({'error': 'identifier required'}), 400
    db = get_db()
    # Resolve identifier: allow either email or teacher_id
    teacher = db['teachers'].find_one({'$or': [{'email': identifier}, {'teacher_id': identifier}]})
    canonical_tid = teacher.get('teacher_id') if teacher else None

    q = {'$or': []}
    q['$or'].append({'teacher1_id': identifier})
    q['$or'].append({'teacher2_id': identifier})
    if canonical_tid:
        q['$or'].append({'teacher1_id': canonical_tid})
        q['$or'].append({'teacher2_id': canonical_tid})

    docs = list(db['duty_adjustments'].find(q))
    for d in docs:
        d['_id'] = str(d.get('_id'))
    return jsonify({'results': docs})


@app.route('/duty_adjustments/<adj_id>/pdf', methods=['GET'])
def get_adj_pdf(adj_id):
    db = get_db()
    from bson.objectid import ObjectId
    try:
        doc = db['duty_adjustments'].find_one({'_id': ObjectId(adj_id)})
    except Exception:
        return jsonify({'error': 'Invalid id'}), 400
    if not doc:
        return jsonify({'error': 'Not found'}), 404
    path = doc.get('pdf_path')
    if not path or not os.path.exists(path):
        return jsonify({'error': 'File not available'}), 404
    return send_file(path)


@app.route('/admin/duty_adjustments/<adj_id>/approve', methods=['POST'])
def approve_adjustment(adj_id):
    approver = (request.json.get('approver') or '').strip()
    db = get_db()
    from bson.objectid import ObjectId
    try:
        adj = db['duty_adjustments'].find_one({'_id': ObjectId(adj_id)})
    except Exception:
        return jsonify({'error': 'Invalid id'}), 400
    if not adj:
        return jsonify({'error': 'Request not found'}), 404
    if adj.get('status') != 'Pending':
        return jsonify({'error': 'Request not pending'}), 400

    # Find latest timetable
    tdoc = db['timetables'].find_one(sort=[('created_at', -1)])
    if not tdoc:
        return jsonify({'error': 'No timetable saved in DB'}), 400

    duties = tdoc.get('teacher_duties', [])

    # helper to match a duty for a teacher on date and session
    def match_duty(tid, date, session):
        for i, d in enumerate(duties):
            if str(d.get('teacher_id')) == str(tid) and str(d.get('date')) == str(date):
                # match session loosely
                if not session or session.lower() in str(d.get('session','')).lower():
                    return i, d
        return None, None

    # Support both legacy and normalized field names
    t1 = adj.get('teacher1_id') or adj.get('applicant_id') or adj.get('identifier')
    t2 = adj.get('teacher2_id') or adj.get('swap_applicant_id') or adj.get('adjusted_id')
    t1_date = adj.get('teacher1_date') or adj.get('duty_date')
    t1_sess = adj.get('teacher1_session') or adj.get('session')
    t2_date = adj.get('teacher2_date') or adj.get('swap_duty_date')
    t2_sess = adj.get('teacher2_session') or adj.get('swap_session')

    i1, d1 = match_duty(t1, t1_date, t1_sess)
    i2, d2 = match_duty(t2, t2_date, t2_sess)

    if d1 is None or d2 is None:
        return jsonify({'error': 'Matching duties not found for one or both teachers'}), 400

    # Ensure they are not the same duty
    if i1 == i2:
        return jsonify({'error': 'Both duties refer to the same record'}), 400

    # Swap teacher assignments (teacher_id & teacher_name)
    d1_tid, d1_tname = d1.get('teacher_id'), d1.get('teacher_name')
    d2_tid, d2_tname = d2.get('teacher_id'), d2.get('teacher_name')

    duties[i1]['teacher_id'] = d2_tid
    duties[i1]['teacher_name'] = d2_tname
    duties[i2]['teacher_id'] = d1_tid
    duties[i2]['teacher_name'] = d1_tname

    # Persist updated timetable document (insert new snapshot to preserve history)
    new_tdoc = dict(tdoc)
    new_tdoc['_id'] = None
    new_tdoc['created_at'] = datetime.utcnow().isoformat()
    new_tdoc['teacher_duties'] = duties
    db['timetables'].insert_one(new_tdoc)

    # mark adjustment approved
    db['duty_adjustments'].update_one({'_id': ObjectId(adj_id)}, {'$set': {'status': 'Approved', 'approved_by': approver, 'approved_at': datetime.utcnow().isoformat()}})

    return jsonify({'success': True})


@app.route('/admin/duty_adjustments/<adj_id>/reject', methods=['POST'])
def reject_adjustment(adj_id):
    reason = (request.json.get('reason') or '').strip()
    approver = (request.json.get('approver') or '').strip()
    db = get_db()
    from bson.objectid import ObjectId
    try:
        adj = db['duty_adjustments'].find_one({'_id': ObjectId(adj_id)})
    except Exception:
        return jsonify({'error': 'Invalid id'}), 400
    if not adj:
        return jsonify({'error': 'Request not found'}), 404
    if adj.get('status') != 'Pending':
        return jsonify({'error': 'Request not pending'}), 400

    db['duty_adjustments'].update_one({'_id': ObjectId(adj_id)}, {'$set': {'status': 'Rejected', 'rejected_reason': reason, 'approved_by': approver, 'approved_at': datetime.utcnow().isoformat()}})
    return jsonify({'success': True})


@app.route('/teacher_duties', methods=['GET'])
def teacher_duties():
    identifier = (request.args.get('identifier') or '').strip()
    if not identifier:
        return jsonify({'error': 'identifier required'}), 400
    db = get_db()
    # Resolve identifier: allow either email or teacher_id
    teacher = db['teachers'].find_one({'$or': [{'email': identifier}, {'teacher_id': identifier}]})
    canonical_tid = None
    if teacher:
        canonical_tid = teacher.get('teacher_id')

    tdoc = db['timetables'].find_one(sort=[('created_at', -1)])
    if not tdoc:
        return jsonify({'error': 'No timetable saved'}), 400
    duties = tdoc.get('teacher_duties', [])
    res = []
    for d in duties:
        tid = str(d.get('teacher_id') or '')
        if canonical_tid and tid == str(canonical_tid):
            res.append(d)
        elif tid == str(identifier):
            res.append(d)

    # Also fetch teacher record to return persisted history and counts
    teacher_rec = None
    try:
        teacher_rec = db['teachers'].find_one({'$or': [{'email': identifier}, {'teacher_id': identifier}]})
    except Exception:
        teacher_rec = None

    history = teacher_rec.get('history', []) if teacher_rec else []
    duty_counts = teacher_rec.get('duty_counts', {}) if teacher_rec else {}

    return jsonify({'duties': res, 'history': history, 'duty_counts': duty_counts, 'timetable_created_at': tdoc.get('created_at')})


@app.route("/confirm", methods=["POST"])
def confirm():
    data = request.json
    assignments = data.get("assignments", [])
    dept = data.get("department", "IT")
    timetable_data = data.get('timetable')

    try:
        # Normalize dates consistently and write new assignments
        import pandas as _pd
        def _norm_date(s):
            try:
                dt = _pd.to_datetime(str(s), dayfirst=True, errors='coerce')
                if _pd.isna(dt):
                    return str(s)
                return dt.strftime('%d-%b-%Y')
            except Exception:
                return str(s)

        db = get_db()

        for a in assignments:
            date_norm = _norm_date(a.get('date'))
            is_high = a.get("role") in ["Senior", "Squad"]
            update_faculty_duty(
                a.get("teacher_id"),
                a.get("role"),
                date_norm,
                a.get("slot"),
                is_high_role=is_high
            )

        # Save timetable snapshot to DB if provided
        try:
            if timetable_data:
                db = get_db()
                tdoc = {
                    'created_at': datetime.utcnow().isoformat(),
                    'department': dept,
                    'timetable': timetable_data.get('timetable', []),
                    'room_allocation': timetable_data.get('room_allocation', []),
                    'teacher_duties': timetable_data.get('teacher_duties', []),
                    'summary': timetable_data.get('summary', {})
                }
                db['timetables'].insert_one(tdoc)
        except Exception as save_err:
            print(f"Warning: failed to save timetable to DB: {save_err}")

        # After adding new assignments, remove any previous/original duty history entries
        # for teachers affected by applied swaps so their 'My Duties' won't show obsolete entries.
        try:
            applied_adjs = list(db['adjustments'].find({"status": "applied"}))
            for adj in applied_adjs:
                for rec in adj.get('applied', []):
                    s = rec.get('swap', {})
                    applicant = s.get('applicant_id') or adj.get('applicant_id') or adj.get('identifier')
                    adjusted = s.get('adjusted_id')
                    old_date = _norm_date(s.get('old_date') or s.get('swapped_duty_date'))
                    new_date = _norm_date(s.get('new_date') or s.get('next_duty_date'))

                    # remove applicant's old duty (old_date) from their history
                    if applicant:
                        tdoc = db['teachers'].find_one({'$or': [{'teacher_id': applicant}, {'email': applicant}]})
                        if tdoc:
                            # find matching history entry
                            hist = next((h for h in tdoc.get('history', []) if str(h.get('exam_date')) == str(old_date)), None)
                            if hist:
                                role_assigned = hist.get('role_assigned') or hist.get('role') or 'Junior'
                                role_key = role_assigned.lower() if role_assigned.lower() in ['junior', 'senior', 'squad'] else 'junior'
                                db['teachers'].update_one({'_id': tdoc['_id']}, {'$pull': {'history': {'exam_date': old_date}}})
                                db['teachers'].update_one({'_id': tdoc['_id']}, {'$inc': {f'duty_counts.{role_key}': -1}})

                    # remove adjusted teacher's old duty (new_date) from their history
                    if adjusted:
                        tdoc = db['teachers'].find_one({'$or': [{'teacher_id': adjusted}, {'email': adjusted}]})
                        if tdoc:
                            hist = next((h for h in tdoc.get('history', []) if str(h.get('exam_date')) == str(new_date)), None)
                            if hist:
                                role_assigned = hist.get('role_assigned') or hist.get('role') or 'Junior'
                                role_key = role_assigned.lower() if role_assigned.lower() in ['junior', 'senior', 'squad'] else 'junior'
                                db['teachers'].update_one({'_id': tdoc['_id']}, {'$pull': {'history': {'exam_date': new_date}}})
                                db['teachers'].update_one({'_id': tdoc['_id']}, {'$inc': {f'duty_counts.{role_key}': -1}})
        except Exception as adj_err:
            print(f"Adjustment finalize warning: {adj_err}")

        reset_triggered = check_reset_fairness(department=dept)
        return jsonify({"success": True, "reset_triggered": reset_triggered})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("Running Flask server...")
    print("   -> Open http://127.0.0.1:5000 in your browser")
    app.run(debug=False, port=5000)