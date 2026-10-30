"""Earned value management as of the status date.

Planned % complete for a task with ES, duration D, and E working days elapsed
before the status date (straight-line baseline):
  D > 0:  clamp((E - ES) / D, 0, 1)
  D = 0:  1 if E >= ES else 0          (milestone)
The workbook's EVM sheet uses the same formula per row.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from .cpm import CPMResult
from .model import Program
from .workcal import WorkCalendar


def planned_fraction(es: float, dur: float, elapsed: float) -> float:
    if dur <= 0:
        return 1.0 if elapsed >= es else 0.0
    return min(1.0, max(0.0, (elapsed - es) / dur))


def _ratio(a: float, b: float) -> float | None:
    return a / b if b else None


@dataclass
class EVMResult:
    elapsed: int
    planned_pct: list[float]
    pv_by_task: list[float]
    ev_by_task: list[float]
    bac: float
    pv: float
    ev: float
    ac: float
    curve: list[tuple[int, date, float]]  # (working day, date, cumulative PV)

    @property
    def sv(self) -> float:
        return self.ev - self.pv

    @property
    def cv(self) -> float:
        return self.ev - self.ac

    @property
    def spi(self) -> float | None:
        return _ratio(self.ev, self.pv)

    @property
    def cpi(self) -> float | None:
        return _ratio(self.ev, self.ac)

    @property
    def eac(self) -> float | None:
        return _ratio(self.bac, self.cpi) if self.cpi else None

    @property
    def vac(self) -> float | None:
        return None if self.eac is None else self.bac - self.eac

    @property
    def tcpi(self) -> float | None:
        return _ratio(self.bac - self.ev, self.bac - self.ac)


def compute(program: Program, cpm: CPMResult, cal: WorkCalendar, step: int = 5) -> EVMResult:
    elapsed = cal.elapsed(program.status_date)
    tasks = program.tasks
    planned = [planned_fraction(cpm.es[i], cpm.duration[i], elapsed) for i in range(len(tasks))]
    pv_t = [t.budget * f for t, f in zip(tasks, planned)]
    ev_t = [t.budget * t.pct_complete for t in tasks]
    curve = []
    horizon = int(cpm.finish + 0.999999)
    points = list(range(0, horizon + 1, step))
    if points[-1] != horizon:
        points.append(horizon)
    for k in points:
        cum = sum(t.budget * planned_fraction(cpm.es[i], cpm.duration[i], k)
                  for i, t in enumerate(tasks))
        curve.append((k, cal.date_of(k), cum))
    return EVMResult(
        elapsed=elapsed, planned_pct=planned, pv_by_task=pv_t, ev_by_task=ev_t,
        bac=sum(t.budget for t in tasks), pv=sum(pv_t), ev=sum(ev_t),
        ac=sum(t.actual_cost for t in tasks), curve=curve,
    )
