#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import dataclasses
import datetime as dt
import re
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from icalendar import Calendar, Event


@dataclasses.dataclass
class CourseMeeting:
    name: str
    weekday: str  # MO..SU
    start_time: dt.time
    end_time: dt.time
    location: str


WEEKDAY_ALIASES = {
    "mon": "MO", "monday": "MO", "1": "MO",
    "tue": "TU", "tues": "TU", "tuesday": "TU", "2": "TU",
    "wed": "WE", "weds": "WE", "wednesday": "WE", "3": "WE",
    "thu": "TH", "thur": "TH", "thurs": "TH", "thursday": "TH", "4": "TH",
    "fri": "FR", "friday": "FR", "5": "FR",
    "sat": "SA", "saturday": "SA", "6": "SA",
    "sun": "SU", "sunday": "SU", "7": "SU",
    "周一": "MO", "周二": "TU", "周三": "WE", "周四": "TH", "周五": "FR", "周六": "SA", "周日": "SU",
}


def normalize_weekday(text: str) -> Optional[str]:
    t = re.sub(r"\s+", " ", (text or "")).strip().lower()
    return WEEKDAY_ALIASES.get(t)


def parse_hhmm(text: str) -> dt.time:
    m = re.fullmatch(r"(\d{1,2}):(\d{2})", text.strip())
    if not m:
        raise ValueError(f"Invalid time: {text}")
    h, mnt = map(int, m.groups())
    return dt.time(h, mnt)


def first_date_for_weekday(term_start_date: dt.date, weekday_code: str) -> dt.date:
    target = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"].index(weekday_code)
    offset = (target - term_start_date.weekday()) % 7
    return term_start_date + dt.timedelta(days=offset)


def build_calendar(courses: Sequence[CourseMeeting], term_start_date: dt.date, weeks: int, tz_name: str) -> Calendar:
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(tz_name)
    except Exception:
        import pytz
        tz = pytz.timezone(tz_name)

    cal = Calendar()
    cal.add("prodid", "-//CSV Course Importer//iOS//")
    cal.add("version", "2.0")

    for c in courses:
        first_date = first_date_for_weekday(term_start_date, c.weekday)
        dtstart = dt.datetime.combine(first_date, c.start_time).replace(tzinfo=tz)
        dtend = dt.datetime.combine(first_date, c.end_time).replace(tzinfo=tz)

        event = Event()
        event.add("summary", c.name)
        if c.location:
            event.add("location", c.location)
        event.add("dtstart", dtstart)
        event.add("dtend", dtend)
        event.add("rrule", {"freq": "weekly", "count": weeks})
        cal.add_component(event)

    return cal


def read_csv(path: Path) -> List[CourseMeeting]:
    meetings: List[CourseMeeting] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"name", "weekday", "start", "end", "location"}
        missing = [c for c in required if c not in reader.fieldnames]
        if missing:
            raise SystemExit(f"CSV missing columns: {', '.join(missing)}. Required columns: name, weekday, start, end, location")
        for row in reader:
            weekday = normalize_weekday(row.get("weekday", ""))
            if not weekday:
                raise SystemExit(f"Invalid weekday: {row.get('weekday')} (row: {row})")
            name = (row.get("name") or "").strip() or "Course"
            start_time = parse_hhmm(row.get("start", ""))
            end_time = parse_hhmm(row.get("end", ""))
            location = (row.get("location") or "").strip()
            meetings.append(CourseMeeting(name=name, weekday=weekday, start_time=start_time, end_time=end_time, location=location))
    return meetings


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert a CSV of courses to an iOS .ics file.")
    parser.add_argument("--csv", required=True, help="Path to CSV with columns: name, weekday, start, end, location")
    parser.add_argument("--term-start", required=True, help="First Monday of week 1 (YYYY-MM-DD)")
    parser.add_argument("--weeks", type=int, default=16, help="Number of teaching weeks (default: 16)")
    parser.add_argument("--tz", default="Asia/Hong_Kong", help="IANA timezone (default: Asia/Hong_Kong)")
    parser.add_argument("--outfile", default="courses_from_csv.ics", help="Output ICS path")

    args = parser.parse_args()

    meetings = read_csv(Path(args.csv))
    if not meetings:
        raise SystemExit("No meetings in CSV.")

    term_start = dt.date.fromisoformat(args.term_start)
    cal = build_calendar(meetings, term_start, args.weeks, args.tz)
    Path(args.outfile).write_bytes(cal.to_ical())
    print(f"Wrote {args.outfile} with {len(meetings)} classes (weekly recurrence for {args.weeks} weeks)")


if __name__ == "__main__":
    main()


