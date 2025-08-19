#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import dataclasses
import datetime as dt
import re
from pathlib import Path
from typing import Dict, List, Optional

from icalendar import Calendar, Event


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


def parse_date(text: str) -> dt.date:
    s = (text or "").strip()
    # Normalize separators
    s = s.replace(".", "-").replace("/", "-")
    # Try YYYY-M-D
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        return dt.date(y, mo, d)
    # Try D-M-YYYY (common in HK)
    m = re.fullmatch(r"(\d{1,2})-(\d{1,2})-(\d{4})", s)
    if m:
        d, mo, y = map(int, m.groups())
        return dt.date(y, mo, d)
    # Fallback ISO exact
    return dt.date.fromisoformat(s)


def get_tz(tz_name: str):
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo(tz_name)
    except Exception:
        import pytz
        return pytz.timezone(tz_name)


def add_exdates(event: Event, tz, exdates_csv: str) -> None:
    for part in (exdates_csv or "").split(','):
        part = part.strip()
        if not part:
            continue
        try:
            d = parse_date(part)
            event.add('exdate', dt.datetime.combine(d, dt.time.min).replace(tzinfo=tz))
        except Exception:
            continue


def add_rdates(event: Event, tz, rdates_csv: str, start_time: dt.time, end_time: dt.time) -> None:
    for part in (rdates_csv or "").split(','):
        part = part.strip()
        if not part:
            continue
        try:
            d = parse_date(part)
            # add two RDATEs? We only add start, since end is covered by duration
            event.add('rdate', dt.datetime.combine(d, start_time).replace(tzinfo=tz))
        except Exception:
            continue


def create_event_recurring(row: Dict[str, str], tz_name: str) -> Event:
    tz = get_tz(tz_name)
    name = (row.get('name') or '').strip() or 'Course'
    location = (row.get('location') or '').strip()
    weekday = normalize_weekday(row.get('weekday', ''))
    if not weekday:
        raise SystemExit(f"Invalid weekday: {row.get('weekday')} for {name}")
    start = parse_hhmm(row['start'])
    end = parse_hhmm(row['end'])
    start_date = parse_date(row['start_date'])
    dtstart_date = _first_date_for_weekday(start_date, weekday)
    dtstart = dt.datetime.combine(dtstart_date, start).replace(tzinfo=tz)
    dtend = dt.datetime.combine(dtstart_date, end).replace(tzinfo=tz)

    ev = Event()
    ev.add('summary', name)
    if location:
        ev.add('location', location)
    ev.add('dtstart', dtstart)
    ev.add('dtend', dtend)

    # RRULE
    interval = int((row.get('interval') or '1').strip() or '1')
    rrule: Dict[str, object] = {'freq': 'weekly', 'interval': interval}

    count_text = (row.get('count') or '').strip()
    end_date_text = (row.get('end_date') or '').strip()
    if count_text:
        try:
            rrule['count'] = int(count_text)
        except Exception:
            pass
    elif end_date_text:
        try:
            until_date = parse_date(end_date_text)
            # UNTIL in local time: set at end of that day
            until_dt = dt.datetime.combine(until_date, dt.time(23, 59)).replace(tzinfo=tz)
            rrule['until'] = until_dt
        except Exception:
            pass
    else:
        # Safety net: limit to 30 occurrences
        rrule['count'] = 30

    ev.add('rrule', rrule)

    # Exceptions and additional dates
    add_exdates(ev, tz, row.get('exceptions', ''))
    add_rdates(ev, tz, row.get('rdates', ''), start, end)

    return ev


def create_event_single(row: Dict[str, str], tz_name: str) -> Event:
    tz = get_tz(tz_name)
    name = (row.get('name') or '').strip() or 'Course'
    location = (row.get('location') or '').strip()
    start = parse_hhmm(row['start'])
    end = parse_hhmm(row['end'])
    date = parse_date(row['date'])
    dtstart = dt.datetime.combine(date, start).replace(tzinfo=tz)
    dtend = dt.datetime.combine(date, end).replace(tzinfo=tz)

    ev = Event()
    ev.add('summary', name)
    if location:
        ev.add('location', location)
    ev.add('dtstart', dtstart)
    ev.add('dtend', dtend)
    return ev


def _first_date_for_weekday(start_date: dt.date, weekday_code: str) -> dt.date:
    target = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"].index(weekday_code)
    offset = (target - start_date.weekday()) % 7
    return start_date + dt.timedelta(days=offset)


def build_calendar_from_csv(csv_path: Path, tz_name: str) -> Calendar:
    with csv_path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = {h.strip().lower() for h in (reader.fieldnames or [])}
        required_base = {'name', 'start', 'end'}
        missing = [h for h in required_base if h not in headers]
        if missing:
            raise SystemExit(f"CSV missing columns: {', '.join(missing)}. Always required: name,start,end")

        cal = Calendar()
        cal.add('prodid', '-//Flexible CSV Course Importer//iOS//')
        cal.add('version', '2.0')

        is_recurring_row: Optional[bool] = None
        for row in reader:
            row_lc = {k.lower(): v for k, v in row.items()}
            has_date = bool(row_lc.get('date'))
            has_rec = bool(row_lc.get('weekday') or row_lc.get('start_date') or row_lc.get('end_date') or row_lc.get('count'))

            if is_recurring_row is None:
                is_recurring_row = not has_date
            # Mix of types allowed; we handle per-row

            if has_date:
                ev = create_event_single(row_lc, tz_name)
                cal.add_component(ev)
                continue

            # recurring requires weekday and start_date at minimum
            for req in ('weekday', 'start_date'):
                if not row_lc.get(req):
                    raise SystemExit(f"Recurring row requires '{req}': {row}")

            ev = create_event_recurring(row_lc, tz_name)
            cal.add_component(ev)

    return cal


def main() -> None:
    parser = argparse.ArgumentParser(description='Flexible CSV to ICS: supports date ranges, intervals, exceptions, and single-date rows.')
    parser.add_argument('--csv', required=True, help='Path to CSV')
    parser.add_argument('--tz', default='Asia/Hong_Kong', help='IANA timezone (default: Asia/Hong_Kong)')
    parser.add_argument('--outfile', default='courses_flexible.ics', help='Output ICS path')

    args = parser.parse_args()

    cal = build_calendar_from_csv(Path(args.csv), args.tz)
    Path(args.outfile).write_bytes(cal.to_ical())
    print(f"Wrote {args.outfile}")


if __name__ == '__main__':
    main()


