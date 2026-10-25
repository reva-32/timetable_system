"""Department-level examination configuration.

The examination cell defines a fixed number of classroom *blocks* for every
department.  Staffing is derived from those blocks instead of from the number
of rooms a particular paper happens to need:

* one Junior Supervisor per block,
* one Internal Squad member for a single block, two for 2-5 blocks and three
  for 6-10 blocks,
* exactly one Senior Supervisor for the department paper.

Keeping the rules in one module makes the scheduler, the PDF writer and the
API validation use exactly the same numbers.
"""

from __future__ import annotations

from eligibility import canonical_department

DEPARTMENT_ORDER = ["CE", "ENTC", "IT", "AIDS", "ECE"]

DEPARTMENT_BLOCKS = {
    "CE": 10,
    "ENTC": 10,
    "IT": 6,
    "AIDS": 2,
    "ECE": 2,
}

DEPARTMENT_TITLES = {
    "CE": "T. Y. COMPUTER ENGINEERING",
    "ENTC": "T. Y. ELECTRONICS & TELECOMMUNICATION ENGINEERING",
    "IT": "T. Y. INFORMATION TECHNOLOGY ENGINEERING",
    "AIDS": "T. Y. ARTIFICIAL INTELLIGENCE & DATA SCIENCE ENGINEERING",
    "ECE": "T. Y. ELECTRONICS & COMPUTER ENGINEERING",
}

DEPARTMENT_SHORT_TITLES = {
    "CE": "Computer Engineering",
    "ENTC": "Electronics & Telecommunication Engineering",
    "IT": "Information Technology Engineering",
    "AIDS": "Artificial Intelligence & Data Science Engineering",
    "ECE": "Electronics & Computer Engineering",
}

# Course groups used by the college circular.  PC = programme core,
# PE = programme elective, MD = multidisciplinary, HC = honours/minor.
CATEGORY_ORDER = ["PC", "PE", "MD", "HC"]

CATEGORY_LABELS = {
    "PC": "Programme Core",
    "PE": "Programme Elective",
    "MD": "Multidisciplinary",
    "HC": "Honours / Minor",
}

# The uploaded circular is an INSEM timetable: core/elective/honours papers
# run in the afternoon slot while multidisciplinary papers run in the morning.
INSEM_CATEGORY_SESSIONS = {
    "PC": "3:30 PM - 4:30 PM",
    "PE": "3:30 PM - 4:30 PM",
    "MD": "11:00 AM - 12:00 PM",
    "HC": "3:30 PM - 4:30 PM",
}

CATEGORY_SESSION_TEMPLATES = {
    "Insem": INSEM_CATEGORY_SESSIONS,
    "Endsem": {category: "10:00 AM - 12:30 PM" for category in CATEGORY_ORDER},
    "Practical": {category: "10:00 AM - 12:00 PM" for category in CATEGORY_ORDER},
}

DEFAULT_SESSION = "3:30 PM - 4:30 PM"


def squads_for_blocks(block_count):
    """Internal Squad demand for a department paper."""
    count = max(0, int(block_count or 0))
    if count <= 1:
        return 1 if count else 0
    if count <= 5:
        return 2
    return 3


def department_block_count(department):
    return DEPARTMENT_BLOCKS.get(canonical_department(department), 0)


def department_title(department):
    return DEPARTMENT_TITLES.get(canonical_department(department), str(department or "Department"))


def department_short_title(department):
    return DEPARTMENT_SHORT_TITLES.get(canonical_department(department), str(department or "Department"))


def category_session(category, exam_type="Insem"):
    template = CATEGORY_SESSION_TEMPLATES.get(exam_type) or INSEM_CATEGORY_SESSIONS
    return template.get(str(category or "").upper(), DEFAULT_SESSION)


def block_names(department):
    """Stable labels such as ``CE-B01`` even when a room sheet is incomplete."""
    dept = canonical_department(department)
    count = department_block_count(dept) or 1
    return [f"{dept}-B{index:02d}" for index in range(1, count + 1)]
