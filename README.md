# ExamSched

ExamSched builds and publishes examination timetables for a college examination
cell. It combines an uploaded workbook (courses, students, rooms and faculty)
with MongoDB-backed eligibility rules, then produces a conflict-free timetable,
room allocation, invigilation duty roster and college-format rosters.

The application is a Flask API with a single-page frontend (`frontend/index.html`).
MongoDB is the source of truth after publication.

## What the system does

1. **Scheduling** — builds a student-conflict graph and colours it with DSATUR,
   then spreads exams across dates and sessions while preserving the
   one-exam-per-student-per-slot invariant. Same-day exams can be moved to a
   later day when that resolves a conflict.
2. **Room allocation** — assigns rooms by capacity for each exam, reusing rooms
   only when they are free in that date/session.
3. **Duty generation** — creates one Junior Supervisor seat per allocated room,
   a Senior Supervisor seat per session, and configurable Internal Squad seats.
4. **Duty assignment** — assigns roles with academic-year, department,
   one-duty-per-day and workload constraints. Pinned duties are placed first.
5. **Review and publication** — the coordinator reviews the draft, then
   publishes it. Publication is versioned; the old version is archived, not
   deleted.
6. **Swaps** — faculty request a change, the coordinator chooses a compatible
   mutual swap, and the server revalidates both teachers before applying the
   change atomically.
7. **Rosters** - downloadable Junior Supervisor and Internal Squad rosters in
   the college matrix format (PDF and Excel), with reporting instructions,
   session codes, stable serial numbers, signatures and revision numbers.
8. **Signed reports and email** - publication emails every assigned faculty
   member a Word and Excel duty report; approved changes generate daily Word and
   Excel change reports. Every report carries COE and Principal digital
   signature blocks.

## Requirements

- Python 3.12+
- MongoDB (MongoDB Atlas or a replica set for swap transactions)
- The Python packages in `backend/requirements.txt` (includes `reportlab` for
  PDF rosters, `python-docx` for Word reports and `openpyxl` for Excel reports)

## Run locally

```bash
cd backend
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
copy .env.example .env        # Windows
# cp .env.example .env        # macOS/Linux
python app.py
```

Open <http://127.0.0.1:5000>.

### Faculty login

Faculty sign in with their institutional email only (for example
`t001@pict.edu`). Bare Teacher IDs are not accepted as login identifiers. The
default password is the lowercase Teacher ID (for example `t001`). To apply
that default to every existing non-admin faculty account:

```bash
cd backend
python reset_teacher_passwords.py
```

Coordinator/admin passwords are not changed by that script.

`backend/.env` is ignored by Git. Fill in at least `MONGODB_URI`, `DB_NAME` and
`SECRET_KEY` before starting. The application refuses to start without a
`SECRET_KEY`.

### Environment variables

| Variable | Purpose |
| --- | --- |
| `MONGODB_URI` | MongoDB connection string. Use a replica set/Atlas for transactions. |
| `DB_NAME` | Database name (default `TimeTable`). |
| `SECRET_KEY` | Flask session signing key. Generate with `python -c "import secrets; print(secrets.token_urlsafe(64))"`. |
| `FRONTEND_ORIGIN` | Optional CORS origin. Leave blank when Flask serves the frontend. |
| `SESSION_COOKIE_SAMESITE`, `SESSION_COOKIE_SECURE` | Session cookie policy. Set `SESSION_COOKIE_SECURE=true` behind HTTPS. |
| `MAX_UPLOAD_SIZE_MB` | Maximum workbook upload size (default 10). |
| `MIGRATION_DEFAULT_PASSWORD` | Temporary password only for legacy accounts handled by `migrate_users.py`. |
| `COORDINATOR_EMAIL` / `COORDINATOR_PASSWORD` | Bootstrap coordinator. Created on first start if missing; an existing account is not overwritten. |
| `COORDINATOR_NAME`, `COORDINATOR_DEPARTMENT` | Display name and department for the bootstrap coordinator. |
| `MAX_DUTIES` | Default per-teacher cap for one run (default 5). Can be overridden per teacher in MongoDB. |
| `JS_PER_SESSION` | Minimum Junior Supervisor seats per session. One per allocated room is always kept; this value adds roaming seats only when the room count is lower. |
| `SQUAD_PER_SESSION` | Exact Internal Squad seats per session (default 9). |
| `LAST_DAY_SQUAD` | Squad seats per session on the final examination day (default 2). |
| `JS_REPORTING_MINUTES` | Reporting time before a session for Junior Supervisors (default 30). |
| `SQUAD_REPORTING_MINUTES` | Reporting time before a session for Internal Squad (default 40). |
| `CONTROL_ROOM` | Examination control room printed on rosters (default `Room A1-110`). |
| `ALLOW_DIRECT_MOVE` | Keep `false` for mutual swaps only. Direct moves can change a day's headcount. |
| `MAIL_ENABLED` | `auto` enables mail once `SMTP_HOST` and a sender are set; `false` disables it. |
| `MAIL_TEST_MODE`, `MAIL_TEST_RECIPIENT` | Redirect the delivery of every personalized message to one demo/test mailbox; the content still names its real faculty recipient. Turn test mode off only when the real faculty mailboxes are ready. |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME`, `SMTP_PASSWORD` | Outgoing SMTP relay/login. Gmail requires an App Password, not the normal account password. |
| `SMTP_FROM_EMAIL`, `SMTP_FROM_NAME`, `SMTP_USE_TLS`, `SMTP_USE_SSL` | Sender identity and transport security for SMTP. |
| `MAIL_ASYNC`, `MAIL_DRY_RUN` | Send in a background worker, or write `.eml` files without contacting SMTP. |
| `MAIL_SEND_DAILY_CHANGES`, `REPORT_RECIPIENTS` | Enable daily-change emails and set the COE/Principal/coordinator recipients. |
| `COE_NAME`, `COE_TITLE`, `PRINCIPAL_NAME`, `PRINCIPAL_TITLE` | Names and titles printed in every signature block. |
| `COE_SIGNATURE_IMAGE`, `PRINCIPAL_SIGNATURE_IMAGE` | Optional PNG/JPG signature images (relative to `backend/` or absolute). A clearly marked **DIGITALLY SIGNED** text block is used when empty. |
| `REPORT_OUTPUT_DIR` | Folder for generated Word/Excel reports (default `backend/generated_reports`). |

The coordinator account is created the first time the application handles a
request. The password in `.env` is used only when the account does not exist;
changing `.env` later does not reset an existing password.

## Workbook format

The uploaded workbook must contain these sheets:

- `Student_Courses` — `student_id`, `course_id`
- `Courses` — `course_id`, `course_name`, `year`, `students_count`, `department`;
  optional `course_code` (the code printed in the circular), `exam_category`
  (`PC`, `PE`, `MD` or `HC`), `exam_order` (the order inside its category),
  `exam_date` and `session`
- `Slots` — optional `slot_id`, `date` and `session`; when the sheet contains
  enough dates for the generated paper groups, those dates drive the whole
  multi-department calendar
- `Rooms` — `room_id`, `capacity`, `department`
- `Teachers` — `teacher_id`, `name`, `role`, `department`, optional `preferred_slots`

A paper may keep a unique internal `course_id` while displaying a shared
circular code in `course_code`. This is required for the multidisciplinary (MD)
and honours (HC) papers, whose codes repeat across departments.

Recognised roles are `Junior`, `Senior` and `Squad` (aliases such as `JS`, `SS`
and `Internal Squad` are normalized). If the workbook has no `Teachers` sheet,
faculty are loaded from MongoDB.

`teaching_years` is deliberately **not** read from the workbook. It is stored in
MongoDB so the examination cell can change eligibility without regenerating the
workbook. A sample workbook can be produced with:

```bash
cd backend
python generate_sample_excel.py
```

## Eligibility rules

The department rule is authoritative:

- A teacher whose department is First Year (`FY`, `FE`, `First Year`, etc.) may
  invigilate **FY papers only**.
- Every other department (`IT`, `CE`, `ENTC`, `AIDS`, etc.) may invigilate
  **SY, TY and B.Tech papers only**. They can never invigilate FY papers.

`teachers.teaching_years` is an optional allow-list that can only **narrow** the
department rule. An empty list or `[]` means “use the department policy”. An
allow-list with no overlap (for example `["FY"]` for an IT teacher) makes that
teacher ineligible.

The same check is applied in three places: draft assignment, the swap dropdown,
and again on approval. An approved swap cannot bypass it.

## Multi-department examination cycle

The coordinator dashboard defaults to **All Departments**. A single run builds
the timetable, room blocks and duty roster for CE, ENTC, IT, AIDS and ECE
together, using the dates between the selected **Exam Start Date** and **Exam
End Date** (Sundays are skipped).

Each department has a fixed number of classroom blocks and a fixed staffing
pattern per paper:

| Department | Blocks | Junior Supervisors | Internal Squad | Senior Supervisor |
| --- | ---: | ---: | ---: | ---: |
| CE | 10 | 10 | 3 | 1 |
| ENTC | 10 | 10 | 3 | 1 |
| IT | 6 | 6 | 3 | 1 |
| AIDS | 2 | 2 | 2 | 1 |
| ECE | 2 | 2 | 2 | 1 |

The rules are implemented in `backend/department_config.py`:

- one Junior Supervisor seat per department block;
- one Squad member for a single block, two for 2–5 blocks and three for 6–10
  blocks;
- exactly one Senior Supervisor per department paper.

Multi-department duties are **cross-department**: a faculty member never
invigilates their own department's paper and First Year (FY/FE) faculty are
never allocated to department papers. The scheduler also enforces one duty per
faculty member per date and the per-teacher duty cap. The generated preview
exposes:

- `POST /generate_departments` — create a draft run;
- `GET /generation_runs/<run_id>/timetable.pdf` — public timetable PDF (cover
  plus one page per department, with Day & Date, Course Code, Name of Course and
  Time only);
- `GET /generation_runs/<run_id>/detailed.pdf` — coordinator-only rooms and
  invigilation duties PDF;
- `POST /confirm_departments` — publish every department timetable and the
  combined view in one operation.

The public PDF follows the college circular: PC, PE and HC papers use the
afternoon session and MD papers use the morning session for an INSEM
examination. The detailed PDF is intentionally separate so rooms and duties are
never mixed into the candidate-facing timetable.

The coordinator dashboard separates the two outputs:

- **Main Timetable** — the candidate-facing circular, with one tab and one page
  per department. It shows Day & Date, Course Code, Name of Course and Time,
  and **Download Main Timetable PDF** produces the same layout with one
  department per page.
- **Detailed Timetable** — coordinator-only rooms and invigilation duties. The
  **Teacher Duties** tab names the allocated faculty (with Teacher ID and
  priority marker) for every duty, and the detailed PDF is downloadable here.
  The legacy result tables and Excel exports remain available for compatibility.

The coordinator's **Exam Start Date** and **Exam End Date** are always the
outer boundary of the generated calendar. If the workbook's `Slots` sheet
contains enough usable dates inside that range, those dates are used; otherwise
the system spreads exams evenly across the selected range (Sundays excluded).
Changing the range therefore changes the dates in the Main Timetable and in the
downloaded PDF — the August 2026 dates are sample-workbook data, not hardcoded.

The legacy **Single Department** mode remains available in the generation-mode
selector for compatibility, but it is no longer the default. A multi-department
run opens on **Main Timetable**; a legacy single-department run opens on
**Detailed Timetable**.

## Duty demand and fairness

Duty demand is generated per date/session:

- **Junior Supervisor**: one seat for every allocated room. `JS_PER_SESSION` is
  a minimum, so extra “Roaming” seats are added only when there are fewer rooms
  than the configured minimum.
- **Senior Supervisor**: one Control-room seat per session.
- **Internal Squad**: exactly `SQUAD_PER_SESSION`, or `LAST_DAY_SQUAD` on the
  final day.

`JS_PER_SESSION`, `SQUAD_PER_SESSION`, `LAST_DAY_SQUAD` and the per-date/session
JSON overrides can be set in the generation form. Overrides accept:

```json
{
  "26-Nov-2026|Morning": {"js": 4, "squad": 2},
  "27-Nov-2026": {"squad": 1}
}
```

The assignment engine:

- assigns one date at a time, so a teacher never gets two duties on one date;
- targets `ceil(total duties / eligible teachers)` and penalises overload;
- applies a per-teacher `max_duties` cap for the current run;
- assigns pinned duties first; a pin may exceed the normal cap but may not
  violate role, year/department or one-duty-per-day rules;
- excludes coordinator/admin accounts from the invigilation pool.

## Per-teacher caps and pins

Coordinators can update eligibility, caps and pins without touching the
workbook. Use the coordinator UI or `PUT /admin/teachers` with the CSRF token:

```json
{
  "teacher_id": "T001",
  "department": "IT",
  "teaching_years": ["SY", "TY"],
  "max_duties": 6,
  "pinned_duties": [
    {"date": "26-Nov-2026", "slot": "1", "role": "Junior", "course_id": "ALL"}
  ]
}
```

Pins use the same date format as the generated timetable (`dd-Mmm-yyyy`) and
the displayed slot value. If a pin conflicts with another pin or with the
teacher's existing duty on that date, it is skipped rather than breaking the
one-duty-per-day rule.

## Publishing, versions and swaps

- Every published timetable is a new version. Previous versions are retained
  and can be inspected or rolled back by a coordinator.
- `assignments` is the single source of truth for published seats, including
  vacant seats. Teacher history, duty counts and exports are derived from it.
- A mutual swap updates two assignment documents inside one MongoDB
  transaction, checks the assignment version, updates the published timetable,
  and writes an entry to `swap_log`.
- Before applying a swap the server re-checks same role, both teachers' year and
  department eligibility, and that neither teacher already has a duty on the
  other's date.
- Faculty can track every change request from their dashboard. Pending requests
  appear as **Under Scrutiny**; approval and rejection decisions, coordinator
  notes, and the new duty details appear in the same section after resolution.
- When no swap partner is available, the coordinator sees the reason (same-role
  shortage, date clashes, or eligibility rules) instead of only a generic error.
- Direct moves are disabled by default. Set `ALLOW_DIRECT_MOVE=true` only if
  the examination cell accepts headcount changes on both days.

## Rosters

After publication, coordinators can download:

- `/roster/junior.pdf` and `/roster/junior.xlsx`
- `/roster/squad.pdf` and `/roster/squad.xlsx`

The rosters contain the college header, reporting instructions, configurable
session line, the date/session-code matrix, stable Sr. No. values, signature
blocks, issue date, and `Page x of y`. A swap increments the roster revision
(for example, “Revision 2”) so the signed copy can be replaced.

Reporting minutes and the control room are configured through the environment
variables above.

## Email delivery and signed Word/Excel reports

Publishing a timetable (`Confirm & Publish`) queues one personalised email per
assigned faculty member. Each email contains:

- the faculty member's duty table and reporting instructions;
- a Word duty report and an Excel duty report;
- COE and Principal **DIGITALLY SIGNED** approval blocks with the signing time.

When `MAIL_TEST_MODE=true`, delivery of every personalized message is redirected
to `MAIL_TEST_RECIPIENT` while the subject and body still name the real intended
faculty recipient. The number of messages is always one per assigned faculty
member, not one generic message.
The coordinator toolbar has **Email Faculty** (resend/queue the active
timetable) and **Test Mail** (send one sample faculty report).

When an adjustment is approved, the system also generates a **Daily Faculty
Duty Change Report** for each affected examination date, in both `.docx` and
`.xlsx`. Both formats show the previous and new faculty member, slot, session,
course, room, reason and approval time, and carry the same COE/Principal
signature blocks. Files are stored under `backend/generated_reports/`; they can
be downloaded from:

- `/admin/reports/daily-change/<dd-Mmm-yyyy>.docx` (coordinator)
- `/admin/reports/daily-change/<dd-Mmm-yyyy>.xlsx` (coordinator)
- `/reports/faculty/<teacher_id>.docx` (that faculty member or a coordinator)
- `/reports/faculty/<teacher_id>.xlsx`

Daily-change emails are sent to `REPORT_RECIPIENTS` (falling back to
`COE_EMAIL`, `PRINCIPAL_EMAIL`, then `COORDINATOR_EMAIL`). They are generated
even when no mail recipient is configured.

Use `MAIL_DRY_RUN=true` to test the complete workflow without SMTP: complete
`.eml` messages are written under `generated_reports/mail_outbox/`. The
signature blocks are visible digital approvals; if the college requires a
cryptographic PDF signature, provide the certificate/PAdES integration details
separately.

### Demo redirect vs. real delivery

There is no hardcoded faculty address in the delivery code. For every assigned
faculty member the application reads that teacher's `email` from MongoDB,
builds a separate report, and addresses a separate message. `MAIL_TEST_MODE`
only changes the **delivery address**:

- `MAIL_TEST_MODE=true` → every personalized message is delivered to
  `MAIL_TEST_RECIPIENT`; the subject and body still show the real intended
  faculty member.
- `MAIL_TEST_MODE=false` → each message is delivered to that faculty member's
  own `@pict.edu` address.

The coordinator's **Email Faculty** button sends all personalized messages.
**Test Mail** sends one sample report. Both use the same environment switch.

### Verifying SMTP credentials

Run the safe credential checker from `backend/`:

```bash
python verify_mail.py
```

It prints the active host/port, transport security, masked username, whether a
password is present, sender, demo recipient, daily recipients, and signature
configuration. It then connects to the SMTP server and verifies STARTTLS/SSL and
login. It **does not send an email or change any timetable/faculty data**.

To also send one minimal message to `MAIL_TEST_RECIPIENT`:

```bash
python verify_mail.py --send
```

For Gmail:

1. Turn on 2-Step Verification.
2. Create a Google **App Password** (Google Account → Security → App passwords).
3. Put the 16-character App Password in `SMTP_PASSWORD`; do not use the normal
   Gmail password.
4. Use `smtp.gmail.com`, port `587`, `SMTP_USE_TLS=true`,
   `SMTP_USE_SSL=false`, and your Gmail address for `SMTP_USERNAME` and
   `SMTP_FROM_EMAIL`.

For a PICT/institutional relay, obtain the host, port, TLS/SSL mode, whether
authentication is required, and the permitted sender address from the mail
administrator, then use the same checks. A successful `verify_mail.py` followed
by a successful **Test Mail** confirms the complete signed-report path.

## Tests

```bash
cd backend
python -m unittest discover -s tests -p "test_*.py" -v
```

The suite covers DSATUR, date adjustment, room allocation, year/department
eligibility, duty demand, one-duty-per-day assignment, pinned-duty validation,
fairness reset behaviour, JSON sanitisation, CSRF, roster rendering, signed
Word/Excel reports, email redirects/attachments, dry-run delivery, workbook-date
preference, custom-range date selection, and the sample workbook's alignment
with the August 2026 circular (dates, shared MD/HC codes, unique internal course
IDs and exam order).

## Docker

```bash
docker build -t examsched .
docker run --env-file backend/.env -p 5000:5000 examsched
```

The container runs `gunicorn` from `/app/backend` and serves the frontend from
the same Flask application.

## Operational notes

- **MongoDB indexes**: the application creates the required indexes on first
  connection. The `teachers.teacher_id` and `teachers.email` indexes are unique
  and sparse; do not create duplicate teacher identities.
- **Transactions**: Atlas/replica sets support the atomic swap path. A
  standalone `mongod` does not support transactions; the code reports a clear
  conflict rather than silently leaving a half-applied swap.
- **Sample data**: `backend/sample_exam_data.xlsx` is the only sample workbook
  and the file to upload for the All Departments flow. It is aligned with the
  `TY BTECH Insem Time table August 2026.pdf` circular: dates 17, 18, 19, 20,
  22 and 24 August 2026, PC/PE/HC afternoon sessions, MD on 22 August at
  11:00–12:00, and the circular's course codes and category order (including
  the shared MD/HC display codes with unique internal course IDs). It contains
  all five departments, exactly the configured blocks, 8 Junior + 4 Squad +
  2 Senior faculty per department, and 5 FY faculty. The FY faculty are
  deliberately present so the eligibility filter can be verified; they are
  never assigned. The All Departments run assigns all 288 duty seats from this
  workbook. Regenerate it with `python generate_sample_excel.py` (close the file
  in Excel first).
- **Legacy eligibility values**: old stringified lists such as
  `"['SY']"` are normalised to real arrays on startup and are still honoured.
- **Rotate secrets** before a production handover. `.env` is ignored by Git.
