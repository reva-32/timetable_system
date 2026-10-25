"""Academic-year and department eligibility rules for invigilation duties.

The college rule is:
  * Teachers whose department is First Year (FY) invigilate only FY papers.
  * Teachers from CE / IT / AIDS / ECE / ENTC (and any other non-FY
    department) invigilate only 2nd, 3rd and 4th year (B.Tech) papers.

`teachers.teaching_years` in MongoDB is an *additional* allow-list.  It can
never widen the department policy: an IT teacher configured with ["FY"] is
still not allowed to invigilate FY papers.  An empty list means "use the
department policy", never "no invigilation".
"""

from __future__ import annotations

import ast
import re

import pandas as pd

FY_DEPARTMENTS = {
    "FY", "FE", "F.Y.", "FIRST", "FIRSTYEAR", "FIRST YEAR", "1", "1ST",
    "FIRSTYEARENGINEERING", "FYENGINEERING",
}
SECOND_TO_FOURTH = {"SY", "TY", "BTECH"}

_ROLE_ALIASES = {
    "JUNIOR": "Junior",
    "JS": "Junior",
    "JUNIORSUPERVISOR": "Junior",
    "JUNIOR SUPERVISOR": "Junior",
    "SENIOR": "Senior",
    "SS": "Senior",
    "SENIORSUPERVISOR": "Senior",
    "SENIOR SUPERVISOR": "Senior",
    "SQUAD": "Squad",
    "INTERNALSQUAD": "Squad",
    "INTERNAL SQUAD": "Squad",
    "INTERNAL SQUAD SUPERVISOR": "Squad",
    "ISS": "Squad",
    "ROAMING": "Squad",
}


def safe_str(val, default="N/A"):
    if val is None:
        return default
    if not isinstance(val, str) and pd.isna(val):
        return default
    text = str(val).strip()
    return text if text and text.lower() not in ("nan", "none", "") else default


def normalize_year(value):
    """Normalize FY/SY/TY/B.Tech labels to canonical year names."""
    v = safe_str(value, default="ALL").upper().replace(" ", "").replace("-", "").replace(".", "")
    aliases = {
        "1": "FY", "FY": "FY", "FE": "FY", "FIRST": "FY", "FIRSTYEAR": "FY",
        "2": "SY", "SY": "SY", "SE": "SY", "SECOND": "SY", "SECONDYEAR": "SY",
        "3": "TY", "TY": "TY", "TE": "TY", "THIRD": "TY", "THIRDYEAR": "TY",
        "4": "BTECH", "BTECH": "BTECH", "BT": "BTECH", "FINAL": "BTECH",
        "FINALYEAR": "BTECH", "4TH": "BTECH",
    }
    return aliases.get(v, v)


def normalize_role(value):
    text = safe_str(value, default="Junior")
    key = text.upper().replace("_", "").replace("-", " ").strip()
    return _ROLE_ALIASES.get(key, text)


def normalize_department(value):
    text = safe_str(value, default="General").upper()
    return " ".join(text.replace(".", " ").replace("_", " ").split())


def is_fy_department(department):
    dept = normalize_department(department)
    compact = dept.replace(" ", "")
    return dept in FY_DEPARTMENTS or compact in {item.replace(" ", "") for item in FY_DEPARTMENTS}


def department_allowed_years(department):
    """Return the years a department is allowed to invigilate (never empty)."""
    if is_fy_department(department):
        return {"FY"}
    return set(SECOND_TO_FOURTH)


def _year_list(raw):
    """Accept a list, a set, or a legacy stringified list."""
    if raw is None:
        return []
    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            return []
        if text.startswith("[") and text.endswith("]"):
            try:
                parsed = ast.literal_eval(text)
                if isinstance(parsed, (list, tuple, set)):
                    return list(parsed)
            except (ValueError, SyntaxError):
                pass
        return [part for part in text.replace(";", ",").split(",") if part.strip()]
    if isinstance(raw, (list, tuple, set)):
        return list(raw)
    return [raw]


def _explicit_years(teacher):
    raw = teacher.get("teaching_years")
    if raw is None:
        raw = teacher.get("teaching_year")
    return {normalize_year(value) for value in _year_list(raw) if safe_str(value, default="")}


def effective_teaching_years(teacher):
    """Years this teacher may actually invigilate.

    Department policy always applies.  An explicit allow-list narrows it.
    Empty/ALL means the full department policy.
    """
    policy = department_allowed_years(teacher.get("department"))
    explicit = _explicit_years(teacher)
    if not explicit or "ALL" in explicit:
        return set(policy)
    return {year for year in explicit if year in policy}


_DEPARTMENT_ALIASES = {
    "CE": {"CE", "COMPUTER", "COMPUTER ENGINEERING", "COMP"},
    "IT": {"IT", "INFORMATION TECHNOLOGY", "INFO TECH"},
    "ENTC": {
        "ENTC", "E&TC", "ETC", "ELECTRONICS",
        "ELECTRONICS AND TELECOMMUNICATION",
        "ELECTRONICS AND TELECOMMUNICATIONS",
    },
    "ECE": {
        "ECE", "E&CE", "EC",
        "ELECTRONICS AND COMPUTER ENGINEERING",
        "ELECTRONICS AND COMPUTER",
    },
    "AIDS": {"AIDS", "AI&DS", "AI DS", "AI", "DATA SCIENCE", "DS"},
}


def canonical_department(value):
    """Return the canonical branch key used by the scheduling rules."""
    dept = normalize_department(value)
    compact = dept.replace(" ", "")
    for key, values in _DEPARTMENT_ALIASES.items():
        if dept in values or compact in {item.replace(" ", "") for item in values}:
            return key
    return dept


# Kept as a private alias for backwards compatibility with older imports.
_canonical_department = canonical_department


def department_matches(teacher_department, duty_department):
    teacher = canonical_department(teacher_department)
    duty = canonical_department(duty_department)
    if teacher in {"GENERAL", "COMMON", "SHARED", ""} or duty in {"GENERAL", "COMMON", "SHARED", ""}:
        return True
    return teacher == duty


def cross_department_allowed(teacher, exam_department):
    """Rules for the multi-department examination cycle.

    * First-year faculty never invigilate department (2nd-4th year) papers.
    * A faculty member never invigilates their own department's paper.
    * General/shared exam seats may be staffed by any non-FY department.
    """
    teacher_department = teacher.get("department")
    if is_fy_department(teacher_department):
        return False
    exam = canonical_department(exam_department)
    teacher_branch = canonical_department(teacher_department)
    if exam in {"GENERAL", "COMMON", "SHARED", ""}:
        return True
    if teacher_branch in {"GENERAL", "COMMON", "SHARED", ""}:
        return True
    return teacher_branch != exam


def duty_allowed_years(duty):
    raw = duty.get("allowed_years")
    if raw:
        return {normalize_year(value) for value in _year_list(raw) if safe_str(value, default="")}
    year = normalize_year(duty.get("year", "ALL"))
    return {"FY", "SY", "TY", "BTECH"} if year == "ALL" else {year}


def teacher_can_invigilate(teacher, duty):
    """Role-independent year + department eligibility check.

    Legacy duties (no ``exam_department``/``cross_department`` marker) keep the
    historic same-department rule.  Multi-department duties explicitly require
    a *different* department and always exclude FY faculty.
    """
    allowed = effective_teaching_years(teacher)
    if not allowed:
        return False
    duty_years = duty_allowed_years(duty)
    if not (allowed & duty_years):
        return False

    exam_department = duty.get("exam_department")
    cross_department = duty.get("cross_department")
    if cross_department is None:
        cross_department = exam_department is not None
    if cross_department:
        return cross_department_allowed(
            teacher, exam_department or duty.get("department")
        )
    return department_matches(teacher.get("department"), duty.get("department"))


def session_start_hour(session):
    """Return the 24-hour start hour from a session label/range, if present."""
    text = safe_str(session, default="").upper()
    match = re.search(r"(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?", text)
    if not match:
        return None
    hour = int(match.group(1))
    meridiem = (match.group(3) or "").lower()
    if meridiem == "pm" and hour < 12:
        hour += 12
    elif meridiem == "am" and hour == 12:
        hour = 0
    return hour


def session_code(session):
    """Short code used in the college roster matrix.

    Named sessions map to M/A/E.  The UI stores time ranges such as
    ``10:00 - 12:30`` or ``2:00 pm to 3:00 pm``, so derive the code from the
    start hour as well (respecting AM/PM).
    """
    text = safe_str(session, default="").upper()
    if not text:
        return "?"
    if text.startswith("MORNING") or text.startswith("FN") or text.startswith("M"):
        return "M"
    if text.startswith("AFTERNOON") or text.startswith("AN") or text.startswith("A"):
        return "A"
    if text.startswith("EVENING") or text.startswith("E"):
        return "E"
    hour = session_start_hour(text)
    if hour is not None:
        if hour < 12:
            return "M"
        if hour < 17:
            return "A"
        return "E"
    return text[0]
