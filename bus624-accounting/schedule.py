#!/usr/bin/env python3
"""
BUS 624 group project - the deadline spine.
Emits (1) an .ics all five members can import and (2) a tracker spreadsheet.

The graded dates come straight from the syllabus. Everything tagged INTERNAL is a
backward-planned milestone - our own deadline, not hers - built on the rule that
nothing gets submitted the day it is finished.
"""
import csv
from datetime import date, timedelta

# ---------------------------------------------------------------- graded dates
# (date, time_or_None, title, kind)  kind: CLASS | GRADED | INTERNAL
CLASSES = [
    ("2026-10-13", "18:00-21:00", "Class 1 (Zoom) - Mod 5. *** WE PRESENT CFS PART 1 *** (lead: Halia)", "CLASS"),
    ("2026-10-20", "18:00-21:00", "Class 2 (Zoom) - Mod 6. CFS Part 1 - other groups", "CLASS"),
    ("2026-10-24", "08:00-12:00", "Class 3 (Zoom) - Mod 7, 8. Types of Equity case study", "CLASS"),
    ("2026-10-27", "18:00-21:00", "Class 4 (Zoom) - Mod 9. *** WE PRESENT CFS PART 2 *** (lead: Richie)", "CLASS"),
    ("2026-11-03", "18:00-21:00", "Class 5 (Zoom) - Mod 10. *** WE PRESENT CFS PART 3 *** (lead: Sean)", "CLASS"),
    ("2026-11-07", "08:00-12:00", "Class 6 (Zoom) - Mod 15, 16", "CLASS"),
    ("2026-11-10", "18:00-21:00", "Class 7 (Zoom) - Mod 17. *** WE PRESENT CFS PART 4 *** (lead: Wade)", "CLASS"),
    ("2026-11-17", "18:00-21:00", "Class 8 (Zoom) - Mod 18. CFS Part 4 - other groups", "CLASS"),
    ("2026-11-24", "18:00-21:00", "Class 9 (Zoom) - Mod 19, 20", "CLASS"),
    ("2026-12-01", "18:00-21:00", "Class 10 (Zoom) - Mod 21. FINAL MEMO DUE", "CLASS"),
    ("2026-12-08", "18:00-21:00", "Class 11 (Zoom) - Mod 22. Master budget problem", "CLASS"),
    ("2026-12-12", "08:00-12:00", "Class 12 (IN PERSON) - Mod 23 + FINAL EXAM", "CLASS"),
]

GRADED_OWNER = {"2026-10-25": "Richie", "2026-11-01": "Sean",
                "2026-11-08": "Wade", "2026-12-01": "Micah"}

GRADED = [
    ("2026-10-25", "GRADED: Submit Excel common-size B/S + I/S and PPT w/ notes for review"),
    ("2026-11-01", "GRADED: Submit Excel financial ratios and PPT w/ notes for review"),
    ("2026-11-08", "GRADED: Submit draft CFS Part 4 - strategies, recommendations, negative impacts"),
    ("2026-12-01", "GRADED: FINAL MEMO due - email to rantal@hawaii.edu w/ company site + 10-K links"),
]

# ------------------------------------------------- internal backward-planned plan
# (date, title, owner_role, phase)
INTERNAL = [
    # Every milestone lands ON a Friday 12pm call so blockers get resolved live.
    # The three Sunday submissions each sit exactly 2 days after a call - so the
    # pattern is: freeze on the call, submit Friday night, present the following week.

    # PHASE 0 - foundation. Done in the 41-day gap before the first class, so the
    # October deadlines become interpretation deadlines rather than research ones.
    ("2026-09-04", "CALL: kickoff. Lock the Step #3 ask. Confirm domains + leads. Walk the plan.", "All", "P0 Foundation"),
    ("2026-09-11", "CALL: all five have read WMT FY2026 10-K (Items 1, 1A, 7, 8 + both audit reports)", "All", "P0 Foundation"),
    ("2026-09-11", "Run pull_edgar.py --years 4 and build_workbook.py; circulate workbook v0", "Micah", "P0 Foundation"),
    ("2026-09-18", "CALL: Walmart FY24-FY26 figures hand-verified against the printed 10-K", "Richie/Sean/Wade", "P0 Foundation"),
    ("2026-09-18", "Costco, Target, Kroger figures verified; comparability caveats logged", "Micah", "P0 Foundation"),

    # PHASE 1 - industry data and Part 1
    ("2026-09-25", "Mergent Online industry averages via UH Hamilton Library -> Ratios tab", "Micah", "P1 Part 1"),
    ("2026-09-25", "CALL: lock the straw-man strategy hypothesis (Part 1 takeaways must support it)", "All", "P1 Part 1"),
    ("2026-10-02", "CALL: Part 1 content complete - overview + 2-3 takeaways each from MC, MD&A, AR", "Halia", "P1 Part 1"),
    ("2026-10-09", "CALL: FREEZE Part 1 + full dress rehearsal. No edits after this call.", "All", "P1 Part 1"),

    # PHASE 2 - common-size. Analysis already exists from Phase 0.
    ("2026-10-12", "START drafting Part 4 strategic options in background (runs parallel to P2/P3)", "Wade", "P4 Strategy"),
    ("2026-10-16", "CALL: common-size interpreted; the 2-3 takeaways that carry the strategy chosen", "Richie", "P2 Common-size"),
    ("2026-10-16", "Pre-select which 15 of the 20 candidate ratios carry the argument", "Sean", "P3 Ratios"),
    ("2026-10-23", "CALL: FREEZE Part 2. Review deck + Excel live on the call.", "All", "P2 Common-size"),
    ("2026-10-23", "SUBMIT Part 2 Friday night, not Sunday - buys 2 days of Antal feedback before presenting", "Richie", "P2 Common-size"),

    # PHASE 3 - ratios. Only 5 days from the Part 2 presentation to this submission,
    # which is why ratio selection already happened on 16 Oct.
    ("2026-10-30", "CALL: FREEZE Part 3. Ratios interpreted vs competitors + industry; deck reviewed.", "All", "P3 Ratios"),
    ("2026-10-30", "SUBMIT Part 3 Friday night rather than Sunday", "Sean", "P3 Ratios"),

    # PHASE 4 - the heaviest deliverable, drafted in parallel since mid-October.
    ("2026-11-06", "CALL: FREEZE Part 4. Two options costed, SMART framing, timeline, investment $.", "All", "P4 Strategy"),
    ("2026-11-06", "Negative impacts (financial/operational/ethical) + mitigation for each", "Wade", "P4 Strategy"),
    ("2026-11-06", "SUBMIT Part 4 draft Friday night rather than Sunday", "Wade", "P4 Strategy"),

    # PHASE 5 - final memo. Pulled a week earlier than ideal because the 27 Nov
    # call falls the day after Thanksgiving.
    ("2026-11-13", "CALL: incorporate Antal's feedback from the Part 4 presentation", "Micah", "P5 Final memo"),
    ("2026-11-13", "Update the Part 1 section to match FINAL recommendations (syllabus requires this)", "Micah", "P5 Final memo"),
    ("2026-11-20", "INDIVIDUAL reflections due - one per person, written alone, no collaboration", "Each of the 5", "P5 Final memo"),
    ("2026-11-20", "CALL: FREEZE final memo. Assembled, reviewed. No new analysis after this call.", "All", "P5 Final memo"),
    ("2026-11-30", "Final QC: every figure traced to source; company site + 10-K links collected", "Micah", "P5 Final memo"),
]

TZ = """BEGIN:VTIMEZONE
TZID:Pacific/Honolulu
BEGIN:STANDARD
DTSTART:19700101T000000
TZOFFSETFROM:-1000
TZOFFSETTO:-1000
TZNAME:HST
END:STANDARD
END:VTIMEZONE"""


def esc(s):
    return s.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\;")


def build_ics(path):
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//BUS624//Walmart FSA//EN",
         "CALSCALE:GREGORIAN", "X-WR-CALNAME:BUS 624 - Walmart Group Project", TZ]
    n = 0

    def ev(uid, lines):
        L.extend(["BEGIN:VEVENT", f"UID:{uid}@bus624", "DTSTAMP:20260902T000000Z"] + lines + ["END:VEVENT"])

    for d, t, title, _ in CLASSES:
        n += 1
        ds = d.replace("-", "")
        s, e = t.split("-")
        ev(f"class{n}", [
            f"DTSTART;TZID=Pacific/Honolulu:{ds}T{s.replace(':','')}00",
            f"DTEND;TZID=Pacific/Honolulu:{ds}T{e.replace(':','')}00",
            f"SUMMARY:{esc(title)}",
            "LOCATION:Zoom https://us02web.zoom.us/j/83896185964 (ID 838 9618 5964, passcode dlemba8)",
            "BEGIN:VALARM\nTRIGGER:-PT60M\nACTION:DISPLAY\nDESCRIPTION:Class starts in 1 hour\nEND:VALARM",
        ])

    for d, title in GRADED:
        n += 1
        ds = d.replace("-", "")
        nd = (date.fromisoformat(d) + timedelta(days=1)).isoformat().replace("-", "")
        ev(f"graded{n}", [
            f"DTSTART;VALUE=DATE:{ds}", f"DTEND;VALUE=DATE:{nd}",
            f"SUMMARY:{esc(title)}", "PRIORITY:1",
            "BEGIN:VALARM\nTRIGGER:-P2D\nACTION:DISPLAY\nDESCRIPTION:Graded deliverable due in 2 days\nEND:VALARM",
            "BEGIN:VALARM\nTRIGGER:-PT12H\nACTION:DISPLAY\nDESCRIPTION:Due today\nEND:VALARM",
        ])

    for d, title, owner, phase in INTERNAL:
        n += 1
        ds = d.replace("-", "")
        nd = (date.fromisoformat(d) + timedelta(days=1)).isoformat().replace("-", "")
        ev(f"int{n}", [
            f"DTSTART;VALUE=DATE:{ds}", f"DTEND;VALUE=DATE:{nd}",
            f"SUMMARY:{esc('[' + phase + '] ' + title)}",
            f"DESCRIPTION:{esc('Owner: ' + owner + '  |  Internal milestone, not a graded date.')}",
            "BEGIN:VALARM\nTRIGGER:-P1D\nACTION:DISPLAY\nDESCRIPTION:Internal milestone tomorrow\nEND:VALARM",
        ])

    add_weekly_calls(L)
    L.append("END:VCALENDAR")
    open(path, "w").write("\r\n".join("\r\n".join(x.split("\n")) for x in L) + "\r\n")
    return n


def build_tracker(path):
    rows = []
    for d, t, title, _ in CLASSES:
        rows.append({"date": d, "type": "CLASS", "phase": "", "task": title,
                     "owner": "All", "status": "", "notes": f"{t} HST"})
    for d, title in GRADED:
        rows.append({"date": d, "type": "GRADED", "phase": "", "task": title,
                     "owner": GRADED_OWNER[d], "status": "Not started", "notes": ""})
    for d, title, owner, phase in INTERNAL:
        rows.append({"date": d, "type": "INTERNAL", "phase": phase, "task": title,
                     "owner": owner, "status": "Not started", "notes": ""})
    rows.sort(key=lambda r: (r["date"], {"GRADED": 0, "INTERNAL": 1, "CLASS": 2}[r["type"]]))
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "type", "phase", "task", "owner", "status", "notes"])
        w.writeheader()
        w.writerows(rows)
    return rows


def add_weekly_calls(L, start="2026-09-04", until="2026-12-04"):
    """Standing Friday 12pm HST call, weekly, through the final memo."""
    L.extend([
        "BEGIN:VEVENT", "UID:weeklycall@bus624", "DTSTAMP:20260902T000000Z",
        f"DTSTART;TZID=Pacific/Honolulu:{start.replace('-','')}T120000",
        f"DTEND;TZID=Pacific/Honolulu:{start.replace('-','')}T130000",
        f"RRULE:FREQ=WEEKLY;BYDAY=FR;UNTIL={until.replace('-','')}T235959Z",
        "SUMMARY:BUS 624 group call - Walmart project",
        "DESCRIPTION:Richie\\, Halia\\, Sean\\, Wade\\, Micah. Standing weekly. "
        "Freeze dates land on this call deliberately.",
        "BEGIN:VALARM", "TRIGGER:-PT30M", "ACTION:DISPLAY",
        "DESCRIPTION:Group call in 30 minutes", "END:VALARM", "END:VEVENT"])


if __name__ == "__main__":
    n = build_ics("BUS624_Walmart_Project.ics")
    rows = build_tracker("BUS624_tracker.csv")
    print(f"BUS624_Walmart_Project.ics  ({n} events)")
    print(f"BUS624_tracker.csv          ({len(rows)} rows)")
    today = date(2026, 9, 2)
    print("\nNext 6 actionable items:")
    for r in rows:
        if r["type"] != "CLASS" and date.fromisoformat(r["date"]) >= today:
            print(f"  {r['date']}  T-{(date.fromisoformat(r['date'])-today).days:>3}d  [{r['owner']:12}] {r['task'][:66]}")
