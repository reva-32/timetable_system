"""PDF rendering for the multi-department examination circular.

Two distinct documents are produced:

``build_public_timetable_pdf``
    The document that is published for candidates.  It mirrors the college
    circular: a cover/instructions page followed by one timetable page per
    department (Day & Date, Course Code, Name of the Course, Time).

``build_detailed_timetable_pdf``
    An internal document for the examination cell containing blocks, rooms,
    every duty seat (Junior / Squad / Senior) and the assigned faculty.  It is
    deliberately kept separate from the public timetable.
"""

from __future__ import annotations

import re
from datetime import datetime
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from department_config import (
    CATEGORY_LABELS,
    DEPARTMENT_ORDER,
    DEPARTMENT_BLOCKS,
    department_title,
)

NAVY = colors.HexColor("#17365d")
BLUE = colors.HexColor("#1f5fa9")
LINK_BLUE = colors.HexColor("#1769aa")
ORANGE = colors.HexColor("#c0501f")
PINK = colors.HexColor("#c2185b")
LIGHT_BLUE = colors.HexColor("#f2f6fb")
GRID = colors.HexColor("#111111")
MUTED = colors.HexColor("#5b6573")


def _text(value, default=""):
    if value is None:
        return default
    text = str(value).strip()
    return text if text else default


def _safe(value):
    return escape(_text(value)).replace("\n", "<br/>")


def _display_time(session):
    """Format a stored session as the circular does: ``3.30 PM TO 4.30 PM``."""
    text = _text(session, "Time TBA")
    text = text.replace("–", "-").replace("—", "-").replace("to", "TO")
    text = re.sub(r"\s+", " ", text).strip()
    match = re.search(r"(\d{1,2}):(\d{2})\s*(AM|PM)?\s*-\s*(\d{1,2}):(\d{2})\s*(AM|PM)?", text, re.I)
    if match:
        start_hour, start_min, start_mer, end_hour, end_min, end_mer = match.groups()
        # If only the end has a meridiem, apply it to the start too.
        start_mer = (start_mer or end_mer or "").upper()
        end_mer = (end_mer or start_mer or "").upper()
        return (
            f"{int(start_hour)}.{start_min} {start_mer} TO "
            f"{int(end_hour)}.{end_min} {end_mer}"
        ).strip()
    return text.upper()


def _display_date(value):
    text = _text(value, "TBA")
    for fmt in ("%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.strftime("%A"), parsed.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text, ""


def _sortable_date(value):
    text = _text(value)
    for fmt in ("%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return datetime.max


def _exam_heading(run):
    year = _text(run.get("year"), "TY").upper()
    year_map = {
        "FY": "FIRST YEAR (B. TECH)",
        "SY": "SECOND YEAR (B. TECH)",
        "TY": "THIRD YEAR (B. TECH)",
        "BTECH": "FOURTH YEAR (B. TECH)",
        "B.TECH": "FOURTH YEAR (B. TECH)",
    }
    year_text = year_map.get(year, f"{year} (B. TECH)")
    exam_type = _text(run.get("exam_type"), "INSEM").upper()
    if exam_type == "ENDSEM":
        exam_text = "END SEMESTER"
    elif exam_type == "PRACTICAL":
        exam_text = "PRACTICAL"
    else:
        exam_text = "INSEM"
    start = _text(run.get("start_date"), "")
    month_year = ""
    for fmt in ("%d-%b-%Y", "%Y-%m-%d"):
        try:
            month_year = datetime.strptime(start, fmt).strftime("%B-%Y").upper()
            break
        except ValueError:
            continue
    return year_text, f"{exam_text} Examination of {month_year or 'the Examination Period'}"


def _draw_logo(canvas, x, y, size=26 * mm):
    canvas.saveState()
    canvas.setStrokeColor(NAVY)
    canvas.setLineWidth(1.2)
    canvas.circle(x + size / 2, y - size / 2, size / 2, stroke=1, fill=0)
    canvas.setLineWidth(0.5)
    canvas.circle(x + size / 2, y - size / 2, size / 2 - 3, stroke=1, fill=0)
    canvas.setFillColor(NAVY)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawCentredString(x + size / 2, y - size / 2 + 3, "PICT")
    canvas.setFont("Helvetica", 4.5)
    canvas.drawCentredString(x + size / 2, y - size / 2 - 2, "PUNE")
    canvas.restoreState()


def _draw_header(canvas, doc):
    width, height = doc.pagesize
    canvas.saveState()
    _draw_logo(canvas, 13 * mm, height - 5.5 * mm, 13 * mm)
    canvas.setFillColor(colors.black)
    canvas.setFont("Times-Bold", 8.5)
    canvas.drawCentredString(width / 2, height - 10.5 * mm, "Society for Computer Technology & Research's")
    canvas.setFont("Times-Bold", 17)
    canvas.drawCentredString(width / 2, height - 15.5 * mm, "PUNE INSTITUTE OF COMPUTER TECHNOLOGY")
    canvas.setFont("Times-Roman", 8.5)
    canvas.drawCentredString(
        width / 2,
        height - 19.5 * mm,
        "An Autonomous Institute affiliated to the Savitribai Phule Pune University",
    )
    canvas.setStrokeColor(colors.black)
    canvas.setLineWidth(1.2)
    canvas.line(15 * mm, height - 28 * mm, width - 15 * mm, height - 28 * mm)
    canvas.restoreState()


def _draw_cover_header(canvas, doc):
    _draw_header(canvas, doc)


def _draw_page_number(canvas, doc, total_pages):
    canvas.saveState()
    canvas.setFont("Times-Roman", 9)
    canvas.drawRightString(doc.pagesize[0] - 10 * mm, 18 * mm, f"Page {doc.page} of {total_pages}")
    canvas.restoreState()


def _draw_signature_seal(canvas, doc):
    width, _ = doc.pagesize
    centre_x = width / 2
    centre_y = 44 * mm
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#243f7d"))
    canvas.setFillColor(colors.HexColor("#243f7d"))
    canvas.setLineWidth(1.1)
    canvas.circle(centre_x, centre_y, 14 * mm, stroke=1, fill=0)
    canvas.setLineWidth(0.5)
    canvas.circle(centre_x, centre_y, 12.2 * mm, stroke=1, fill=0)
    canvas.setFont("Times-Bold", 7.5)
    canvas.drawCentredString(centre_x, centre_y + 1.5 * mm, "PUNE 43")
    canvas.setFont("Times-Roman", 5)
    canvas.drawCentredString(centre_x, centre_y - 3 * mm, "INSTITUTE OF COMPUTER TECHNOLOGY")
    canvas.restoreState()


def _signature_block(canvas, doc):
    width, _ = doc.pagesize
    canvas.saveState()
    canvas.setStrokeColor(colors.black)
    canvas.setLineWidth(0.4)
    canvas.line(30 * mm, 58 * mm, 66 * mm, 58 * mm)
    canvas.line(width - 66 * mm, 58 * mm, width - 30 * mm, 58 * mm)
    canvas.setFont("Times-Bold", 9.5)
    canvas.drawCentredString(48 * mm, 51 * mm, "Controller of Examinations")
    canvas.drawCentredString(width - 48 * mm, 51 * mm, "Principal")
    canvas.setFont("Times-Roman", 8)
    canvas.drawCentredString(48 * mm, 45 * mm, "Controller of Examinations (COE)")
    canvas.drawCentredString(48 * mm, 40.5 * mm, "SCTR's Pune Institute of Computer Technology")
    canvas.drawCentredString(width - 48 * mm, 45 * mm, "Principal")
    canvas.drawCentredString(width - 48 * mm, 40.5 * mm, "SCTR's Pune Institute of Computer Technology")
    canvas.restoreState()
    _draw_signature_seal(canvas, doc)


def _draw_department_footer(canvas, doc, total_pages):
    _signature_block(canvas, doc)
    _draw_page_number(canvas, doc, total_pages)


def _circular_date(run):
    raw = _text(run.get("created_at") or run.get("published_at"))
    if raw:
        try:
            return datetime.strptime(raw[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass
    return datetime.now().strftime("%d/%m/%Y")


def _cover_story(run):
    styles = getSampleStyleSheet()
    story = []
    story.append(Spacer(1, 1 * mm))
    circular_row = Table(
        [[
            Paragraph(
                "Examination Circular no. 8 of 2026",
                ParagraphStyle("circ", parent=styles["Normal"], fontName="Times-Roman", fontSize=10.5),
            ),
            Paragraph(
                f"Date: {_circular_date(run)}",
                ParagraphStyle("circDate", parent=styles["Normal"], fontName="Times-Roman",
                               fontSize=10.5, alignment=TA_RIGHT),
            ),
        ]],
        colWidths=[90 * mm, 90 * mm],
    )
    circular_row.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(circular_row)
    story.append(Spacer(1, 13 * mm))
    year_text, exam_text = _exam_heading(run)
    story.append(Paragraph(
        f'<font color="#b23a1a"><b>{_safe(year_text)}</b></font>',
        ParagraphStyle("coverYear", parent=styles["Normal"], alignment=TA_CENTER,
                       fontName="Times-Bold", fontSize=16, leading=20),
    ))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        f'<font color="#b23a1a"><b>{_safe(exam_text)}</b></font>',
        ParagraphStyle("coverExam", parent=styles["Normal"], alignment=TA_CENTER,
                       fontName="Times-Bold", fontSize=14, leading=18),
    ))
    story.append(Spacer(1, 9 * mm))
    story.append(Paragraph(
        '<u><font color="#b23a1a"><b>INSTRUCTIONS FOR CANDIDATES</b></font></u>',
        ParagraphStyle("coverInstr", parent=styles["Normal"], alignment=TA_CENTER,
                       fontName="Times-Bold", fontSize=15, leading=19),
    ))
    story.append(Spacer(1, 8 * mm))
    instructions = [
        "Candidates are required to be present at the examination centre, <b>THIRTY MINUTES</b> before the stipulated time.",
        "Candidates are forbidden from taking any material into the examination hall that can be treated as a malpractice.",
        "Candidates are requested to see the Notice Board at their centre of examination regularly for changes if any that may be notified later in the program.",
        '<font color="#b23a1a"><b>Candidates are requested to note the Day, Date and Time of Paper.</b></font>',
        "Candidate must wear simple; light colour clothes without excessive pockets (Jackets &amp; Cargo Pants) also simple sandals or slipper are preferred; avoid bulky footwear.",
        "Before start of the examination do not carry bag/ sack, any study material inside the exam hall.",
        "The exchanges of side-rules, drawing instruments of other materials used in the exam hall is not permitted at the time of examination. Candidates must bring their own instruments and will not be allowed to borrow from each other under any circumstances.",
        "Use of non-programmable battery-operated electronic pocket size Calculator is allowed. The exchange of Calculators is not allowed. Electronics Devices including mobile are not allowed at the time of examination.",
        "The written examination will be conducted in the following order.",
    ]
    body = ParagraphStyle(
        "instructions", parent=styles["Normal"], fontName="Times-Roman", fontSize=11,
        leading=15.5, leftIndent=15 * mm, bulletIndent=5 * mm, spaceAfter=2.5 * mm,
    )
    for item in instructions:
        story.append(Paragraph(f"<bullet>&bull;</bullet>{item}", body))
    return story


def _draw_cover_seal(canvas, doc):
    width, _ = doc.pagesize
    centre_x = width / 2
    centre_y = 44 * mm
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#243f7d"))
    canvas.setFillColor(colors.HexColor("#243f7d"))
    canvas.setLineWidth(1.2)
    canvas.circle(centre_x, centre_y, 15 * mm, stroke=1, fill=0)
    canvas.setLineWidth(0.5)
    canvas.circle(centre_x, centre_y, 13.2 * mm, stroke=1, fill=0)
    canvas.setFont("Times-Bold", 8)
    canvas.drawCentredString(centre_x, centre_y + 1.5 * mm, "PUNE 43")
    canvas.setFont("Times-Roman", 5.2)
    canvas.drawCentredString(centre_x, centre_y - 3 * mm, "INSTITUTE OF COMPUTER TECHNOLOGY")
    canvas.restoreState()


def _department_story(department):
    styles = getSampleStyleSheet()
    dept = _text(department.get("department"), "Department")
    title = _text(department.get("title"), department_title(dept))
    story = [Spacer(1, 3.5 * mm)]
    story.append(Paragraph(
        f'<font color="#b23a1a"><b>{_safe(title)}</b></font>',
        ParagraphStyle("deptTitle", parent=styles["Normal"], alignment=TA_CENTER,
                       fontName="Times-Bold", fontSize=15, leading=18),
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        f'<font color="#{ORANGE.hexval()[2:]}"><b>SEMESTER-V</b></font>',
        ParagraphStyle("deptSem", parent=styles["Normal"], alignment=TA_CENTER,
                       fontName="Times-Bold", fontSize=12, leading=14.5),
    ))
    story.append(Spacer(1, 5 * mm))

    sessions = department.get("sessions") or []
    if not sessions:
        story.append(Paragraph("No papers are scheduled for this department.", styles["Normal"]))
        return story

    cell_head = ParagraphStyle("th", parent=styles["Normal"], fontName="Times-Bold",
                               fontSize=12.5, leading=14.5, alignment=TA_CENTER, textColor=NAVY)
    cell_day = ParagraphStyle("day", parent=styles["Normal"], fontName="Times-Bold",
                              fontSize=11, leading=16.5, alignment=TA_CENTER, textColor=NAVY)
    cell_code = ParagraphStyle("code", parent=styles["Normal"], fontName="Times-Bold",
                               fontSize=10.8, leading=24, alignment=TA_CENTER, textColor=LINK_BLUE)
    cell_name = ParagraphStyle("name", parent=styles["Normal"], fontName="Times-Bold",
                               fontSize=11, leading=16.2, alignment=TA_LEFT)
    cell_time = ParagraphStyle("time", parent=styles["Normal"], fontName="Times-Bold",
                               fontSize=10.2, leading=13, alignment=TA_CENTER, textColor=PINK)

    data = [[
        Paragraph("Day &amp; Date", cell_head),
        Paragraph("Course Code", cell_head),
        Paragraph("Name of the Course", cell_head),
        Paragraph("Time", cell_head),
    ]]
    for session in sessions:
        courses = session.get("courses") or []
        if not courses:
            continue
        day, date_text = _display_date(session.get("date"))
        codes = []
        names = []
        for course in courses:
            codes.append(_safe(_text(course.get("course_code") or course.get("course_id"), "")))
            names.append(_safe(_text(course.get("course_name"), "Course")))
        if len(courses) > 1:
            name_html = "<br/>".join(f"&bull; {name}" for name in names)
        else:
            name_html = names[0] if names else "Course"
        data.append([
            Paragraph(f"{_safe(day)}<br/>{_safe(date_text)}", cell_day),
            Paragraph("<br/>".join(codes), cell_code),
            Paragraph(name_html, cell_name),
            Paragraph(_safe(_display_time(session.get("session"))), cell_time),
        ])

    table = Table(data, colWidths=[28 * mm, 35 * mm, 68 * mm, 49 * mm], repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1.0, GRID),
        ("BOX", (0, 0), (-1, -1), 2.2, GRID),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, 0), 9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 9),
        ("TOPPADDING", (0, 1), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 1.5),
    ]))
    story.append(table)
    return story


def build_public_timetable_pdf(run):
    """Return the public timetable PDF as bytes."""
    from io import BytesIO

    buffer = BytesIO()
    departments = run.get("departments") or []
    total_pages = 1 + len(departments)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=30 * mm,
        bottomMargin=64 * mm,
        title="PICT Examination Timetable",
        author="Pune Institute of Computer Technology",
    )

    def cover_page(canvas, doc_obj):
        _draw_cover_header(canvas, doc_obj)
        _draw_cover_seal(canvas, doc_obj)
        _draw_page_number(canvas, doc_obj, total_pages)

    def later_page(canvas, doc_obj):
        _draw_header(canvas, doc_obj)
        _draw_department_footer(canvas, doc_obj, total_pages)

    story = _cover_story(run)
    for index, department in enumerate(departments):
        story.append(PageBreak())
        story.extend(_department_story(department))

    doc.build(story, onFirstPage=cover_page, onLaterPages=later_page)
    return buffer.getvalue()


def _detail_header(canvas, doc):
    width, height = doc.pagesize
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawString(15 * mm, height - 15 * mm, "EXAMINATION CELL — DETAILED TIMETABLE")
    canvas.setFillColor(MUTED)
    canvas.setFont("Helvetica", 8.5)
    canvas.drawString(15 * mm, height - 20 * mm, "Rooms, blocks and invigilation duties (internal document)")
    canvas.setStrokeColor(colors.black)
    canvas.line(15 * mm, height - 23 * mm, width - 15 * mm, height - 23 * mm)
    canvas.restoreState()


def _detail_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(doc.pagesize[0] - 15 * mm, 9 * mm, f"Page {doc.page}")
    canvas.restoreState()


def _detail_table(data, widths):
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        "detailTH", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=7.5, leading=9.5, textColor=NAVY,
    )
    body_style = ParagraphStyle(
        "detailTD", parent=styles["Normal"], fontName="Helvetica",
        fontSize=7.5, leading=9.5,
    )
    wrapped = []
    for row_index, row in enumerate(data):
        style = header_style if row_index == 0 else body_style
        wrapped.append([
            cell if isinstance(cell, Paragraph) else Paragraph(_safe(cell), style)
            for cell in row
        ])
    table = Table(wrapped, colWidths=widths, repeatRows=1)
    table.hAlign = "LEFT"
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#9aa4b2")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eef7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return table


def build_detailed_timetable_pdf(run):
    """Return the detailed rooms + duties PDF as bytes."""
    from io import BytesIO

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=27 * mm,
        bottomMargin=15 * mm,
        title="PICT Detailed Examination Timetable",
        author="Pune Institute of Computer Technology",
    )
    styles = getSampleStyleSheet()
    heading = ParagraphStyle("detailHeading", parent=styles["Heading2"], fontName="Helvetica-Bold",
                             fontSize=12, leading=15, textColor=NAVY, spaceBefore=4, spaceAfter=4)
    subheading = ParagraphStyle("detailSub", parent=styles["Normal"], fontName="Helvetica-Bold",
                                fontSize=8.5, leading=11, textColor=ORANGE, spaceBefore=5, spaceAfter=2)
    body = ParagraphStyle("detailBody", parent=styles["Normal"], fontName="Helvetica",
                          fontSize=8, leading=10.5)

    story = []
    year_text, exam_text = _exam_heading(run)
    story.append(Paragraph(
        f"{_safe(year_text)} — {_safe(exam_text)}",
        ParagraphStyle("title", parent=styles["Title"], fontName="Helvetica-Bold",
                       fontSize=15, leading=18, textColor=NAVY),
    ))
    story.append(Paragraph(
        f"Examination period: <b>{_safe(run.get('start_date'))}</b> to "
        f"<b>{_safe(run.get('end_date'))}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Departments: <b>{_safe(', '.join(run.get('departments_requested') or DEPARTMENT_ORDER))}</b>",
        body,
    ))
    story.append(Spacer(1, 3 * mm))

    # Department/block overview.
    overview = [["Department", "Blocks", "Junior", "Squad", "Senior", "Papers"]]
    for department in run.get("departments") or []:
        overview.append([
            department.get("department"),
            len(department.get("blocks") or []),
            len(department.get("blocks") or []),
            department.get("squad_count", 0),
            1,
            len(department.get("sessions") or []),
        ])
    story.append(Paragraph("Department block staffing", heading))
    story.append(_detail_table(overview, [28 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm]))
    story.append(Spacer(1, 3 * mm))

    # Calendar overview: one row per examination date.
    calendar = {}
    for department in run.get("departments") or []:
        for session in department.get("sessions") or []:
            entry = calendar.setdefault(session.get("date"), {
                "sessions": set(),
                "papers": set(),
                "departments": set(),
            })
            entry["sessions"].add(_text(session.get("session")))
            for course in session.get("courses") or []:
                entry["papers"].add(_text(course.get("course_code") or course.get("course_id")))
            entry["departments"].add(department.get("department"))
    calendar_rows = [["Date", "Day", "Session(s)", "Papers", "Departments"]]
    for date_text in sorted(calendar, key=lambda value: _sortable_date(value)):
        entry = calendar[date_text]
        day, formatted = _display_date(date_text)
        calendar_rows.append([
            formatted or date_text,
            day,
            " / ".join(sorted(entry["sessions"])),
            len(entry["papers"]),
            ", ".join(sorted(entry["departments"])),
        ])
    story.append(Paragraph("Examination calendar", heading))
    story.append(_detail_table(calendar_rows, [28 * mm, 28 * mm, 60 * mm, 18 * mm, 60 * mm]))

    story.append(PageBreak())
    for index, department in enumerate(run.get("departments") or []):
        if index:
            story.append(PageBreak())
        dept = department.get("department")
        title = department.get("title") or department_title(dept)
        story.append(Paragraph(f"{_safe(dept)} — {_safe(title)}", heading))
        blocks = department.get("blocks") or []
        if blocks:
            block_rows = [["Block", "Room", "Capacity", "Courses using the block"]]
            for block in blocks:
                block_rows.append([
                    block.get("block"),
                    block.get("room_id"),
                    block.get("capacity"),
                    ", ".join(block.get("courses") or []),
                ])
            story.append(Paragraph("Room / block allocation", subheading))
            story.append(_detail_table(block_rows, [20 * mm, 25 * mm, 20 * mm, 130 * mm]))
        story.append(Paragraph("Papers and rooms", subheading))
        paper_rows = [["Date", "Time", "Paper(s)", "Rooms / blocks", "Students", "Status"]]
        for session in department.get("sessions") or []:
            courses = session.get("courses") or []
            code_names = ", ".join(
                f"{_text(c.get('course_code') or c.get('course_id'))} {_text(c.get('course_name'))}"
                for c in courses
            )
            paper_rows.append([
                session.get("date"),
                session.get("session"),
                code_names,
                ", ".join(session.get("room_ids") or []),
                session.get("student_count"),
                session.get("status") or "Allocated",
            ])
        story.append(_detail_table(paper_rows, [23 * mm, 34 * mm, 92 * mm, 70 * mm, 16 * mm, 19 * mm]))
        story.append(Spacer(1, 2 * mm))

    duty_rows = [["Date", "Time", "Dept", "Role", "Block", "Room", "Faculty", "Faculty dept", "Status"]]
    for duty in sorted(
        run.get("duties") or [],
        key=lambda item: (
            str(item.get("date")),
            str(item.get("exam_department")),
            str(item.get("role_required")),
        ),
    ):
        duty_rows.append([
            duty.get("date"),
            duty.get("session"),
            duty.get("exam_department") or duty.get("department"),
            duty.get("role_required"),
            duty.get("block"),
            duty.get("room"),
            duty.get("teacher_name") or "VACANT",
            duty.get("teacher_department") or "",
            "Assigned" if duty.get("teacher_id") else "VACANT",
        ])
    story.append(PageBreak())
    story.append(Paragraph("Invigilation duties", heading))
    story.append(_detail_table(
        duty_rows,
        [21 * mm, 32 * mm, 14 * mm, 16 * mm, 18 * mm, 25 * mm, 52 * mm, 24 * mm, 18 * mm],
    ))
    story.append(Spacer(1, 4 * mm))

    summary = {}
    for duty in run.get("duties") or []:
        teacher_id = duty.get("teacher_id")
        if not teacher_id:
            continue
        entry = summary.setdefault(teacher_id, {
            "name": duty.get("teacher_name") or teacher_id,
            "department": duty.get("teacher_department") or "",
            "role": duty.get("teacher_role") or duty.get("role_required") or "",
            "dates": set(),
            "roles": set(),
        })
        entry["dates"].add(_text(duty.get("date")))
        entry["roles"].add(_text(duty.get("role_required")))
    faculty_rows = [["Faculty", "Department", "Role", "Duties", "Dates"]]
    for entry in sorted(summary.values(), key=lambda item: item["name"]):
        faculty_rows.append([
            entry["name"],
            entry["department"],
            entry["role"],
            len(entry["dates"]),
            ", ".join(sorted(entry["dates"])),
        ])
    story.append(Paragraph("Faculty duty summary", heading))
    story.append(_detail_table(faculty_rows, [50 * mm, 26 * mm, 24 * mm, 18 * mm, 130 * mm]))

    def detail_page(canvas, doc_obj):
        _detail_header(canvas, doc_obj)
        _detail_footer(canvas, doc_obj)

    doc.build(story, onFirstPage=detail_page, onLaterPages=detail_page)
    return buffer.getvalue()
