"""Which business days does a payment file cover?

A file generated on a working day D contains the payments of every day from the
previous working day up to (and including) D-1.  Examples (Mon=0 ... Sun=6):

* normal Monday      -> Fri, Sat, Sun
* Friday non-working -> Monday file covers Thu, Fri, Sat, Sun
* Monday non-working -> Tuesday file covers Fri, Sat, Sun, Mon
"""
from __future__ import annotations

from datetime import date, timedelta

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DEFAULT_NON_WORKING = frozenset({5, 6})  # Saturday, Sunday


class CalendarError(ValueError):
    pass


def previous_working_day(d: date, non_working: set[int] | frozenset[int]) -> date:
    if len(set(non_working)) >= 7:
        raise CalendarError("At least one day of the week must be a working day.")
    prev = d - timedelta(days=1)
    while prev.weekday() in non_working:
        prev -= timedelta(days=1)
    return prev


def covered_dates(file_date: date, non_working: set[int] | frozenset[int]) -> list[date]:
    """Payment dates that belong in the file generated on ``file_date``."""
    if file_date.weekday() in non_working:
        raise CalendarError(
            f"{file_date:%A %d %b %Y} is marked as a non-working day, so no file is generated on it."
        )
    start = previous_working_day(file_date, non_working)
    return [start + timedelta(days=i) for i in range((file_date - start).days)]
