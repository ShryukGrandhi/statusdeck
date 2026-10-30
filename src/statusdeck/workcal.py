"""Working-day calendar: maps working-day indices to dates (Mon-Fri minus holidays)."""

from __future__ import annotations

import bisect
from datetime import date, timedelta


class WorkCalendar:
    def __init__(self, start: date, holidays: list[date] | None = None, days: int = 0):
        self.start = start
        self.holidays = frozenset(holidays or [])
        self._dates: list[date] = []
        self._cursor = start
        if days:
            self._extend(days)

    def is_workday(self, d: date) -> bool:
        return d.weekday() < 5 and d not in self.holidays

    def _extend(self, n: int) -> None:
        while len(self._dates) < n:
            if self.is_workday(self._cursor):
                self._dates.append(self._cursor)
            self._cursor += timedelta(days=1)

    def date_of(self, idx: int) -> date:
        """Date of working day `idx` (0 = first working day on/after start)."""
        if idx < 0:
            raise ValueError("negative working-day index")
        self._extend(idx + 1)
        return self._dates[idx]

    def elapsed(self, d: date) -> int:
        """Working days strictly before `d` (= COUNTIF(CalDates, "<"&d) in the workbook)."""
        while not self._dates or self._dates[-1] < d:
            self._extend(len(self._dates) + 64)
        return bisect.bisect_left(self._dates, d)

    def dates(self, n: int) -> list[date]:
        self._extend(n)
        return self._dates[:n]
