ExamSched – Issues to Fix

Scope: only the points agreed after review. Priority: P1 = blocks real use, P2 = should fix
before handover, P3 = tuning.

1. Year-based invigilation rule (P1)
Rule (from teacher): FY teachers invigilate only FY papers. Teachers from IT, AIDS, ECE,
ENTC and CE (2nd–4th year faculty) invigilate only 2nd–4th year papers, never FY.
Problem: Eligibility is stored per teacher in 
teaching_years , and an empty list means "no
invigilation". On a fresh database nobody gets any duty (tested: 183 duties, 0 assigned).
Fix: Derive eligibility from the teacher's department: FY dept → FY only; all other depts →
2nd–4th year. Never leave it empty by default. Apply the same rule in the swap dropdown
and again on approval.

2. Squad demand is hard-coded to 1 per session (P1)
build_duties creates 1 Squad duty per date/session. The college roster needs about 9
Squad per day (2 on the last day), and about 35 JS per day.
Fix: Make JS count and Squad count configurable per date and session.

3. Load spread is wider than the college's (P3)
Simulation at real scale gave JS loads of 1 duty (4 people), 2 (76), 3 (7). The college roster
has 2 duties for 84 people and 3 for 3 people, and nobody with only 1.
76 is close to 84, but the engine is not yet better than the manual roster: 4 people get half
the load and 7 get extra. Likely cause: the Hungarian step ignores same-day clashes, then
drops clashing duties and refills greedily (not proven).
Fix: Handle the one-duty-per-day rule inside the assignment step, targeting ceil(total
duties / people) per person.



4. Single global duty cap (P2)
What it means: 
MAX_DUTIES (default 5) applies equally to every teacher. The college
roster has Squad members with 5 and 6 duties (all six days), who look like standing
coordinators. The engine cannot reproduce that.
Fix: Allow a per-teacher maximum and pinned duties (fixed dates) that are assigned first.


5. Direct move and approval re-validation (P2)
Mutual consent is already covered by the signed swap PDF, so that is not an issue.
Remaining points:
Direct move (no partner) changes the headcount of two days without checking
vacancy: a day needing 35 can become 34 or 36. If the college only does mutual swaps,
remove this path.
Approval does not re-check state. Eligibility is computed when the dropdown loads,
not when the coordinator approves. If another swap is approved in between, a clash
can slip in.
Fix: On approval, re-check on the server that both teachers are free on each other's dates
and that roles and year rules still hold.


6. Role and year rules on swaps (P2)
Same-role-only swaps (JS with JS, Squad with Squad) are correct, and the dropdown
already filters by role. The server swap path also rejects a role mismatch, which is good.
Remaining: the dropdown and approval do not apply the year rule from Issue 1.


7. Swaps are not atomic and data is stored in three places (P1)
Where data lives: 
teachers.history , 
timetables.teacher_duties , and 
Problems:
adjustments .
1. A swap runs separate updates (requester, partner, timetable sync). If the second fails,
one teacher has changed and the other has not, leaving duplicate or lost duties.
2. If the timetable sync fails, the error is only printed and the API still returns success, so
the roster the college sees can differ from teacher histories.
Page 2 of 3
ExamSched – Issues to Fix
3. In the timetable sync, if the partner's duty is not found, the requester's duty is still
handed to the partner and nothing comes back.
4. History updates match 
history.exam_date and 
history.slot_id as two separate
conditions with the positional 
$ operator. MongoDB does not guarantee both match
the same array element (use 
$elemMatch). This is riskier here because every duty has
the same slot id (one Evening session). Not run against a live MongoDB.
5. modified_count is not checked after updates.
Fix: One 
assignments collection, one document per duty (date, session, role, seat,
teacher, version). A swap changes 
teacher_id on two documents in a single transaction,
with the version checked. History and exports are derived from it. Log every swap.


8. New requirement: roster PDF in the college format (P1)
Generate separate PDFs for Junior Supervisors and Internal Squad, matching the college
sheets:
Header: "SPPU BE In Semester Examination <month/year>" and "Junior Supervisor
Duty Schedule" (or "Internal Squad Supervisor Duty Schedule").
Reporting line: JS report 30 minutes before, Internal Squad 40 minutes before, in the
examination control room (Room A1-110). Make the minutes and room configurable.
Session line: e.g. "Evening: 2:00 pm to 3:00 pm".
Table: Sr. No., Name, one column per date with a second header row showing the
session code (E); "E" in the cell where the person has duty. Keep the original Sr. No.
even after removals.
Footer: signature blocks for Senior Supervisor, COE and Principal, plus the date of
issue and "Page x of y".
Revisions: after approved swaps, regenerate as "Revision n" for re-signing.
Also offer the same matrix as Excel.


9. Test with real data (P2)
Use the college roster as the fixture. Check that each day has the required JS and Squad
count, nobody has two duties on one day, and load stays within 1 of the target (except
pinned coordinators)