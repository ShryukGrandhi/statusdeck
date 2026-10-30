"""Plain data model for a program: tasks, risks, dates."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

_PRED_RE = re.compile(r"^\s*([A-Za-z0-9_.\-]+?)\s*(?:([+-])\s*(\d+))?\s*$")


@dataclass
class Task:
    id: str
    name: str
    workstream: str = ""
    owner: str = ""
    # Finish-to-start predecessors as (task id, lag in working days).
    preds: list[tuple[str, int]] = field(default_factory=list)
    optimistic: float = 0.0
    likely: float = 0.0
    pessimistic: float = 0.0
    budget: float = 0.0
    pct_complete: float = 0.0
    actual_cost: float = 0.0

    @property
    def duration(self) -> float:
        """Baseline (deterministic) duration: the most-likely estimate."""
        return self.likely

    def preds_text(self) -> str:
        return format_preds(self.preds)


@dataclass
class Risk:
    id: str
    title: str
    owner: str = ""
    probability: int = 1  # 1..5
    impact: int = 1  # 1..5
    mitigation: str = ""

    @property
    def score(self) -> int:
        return self.probability * self.impact

    @property
    def rating(self) -> str:
        return risk_rating(self.score)


def risk_rating(score: float) -> str:
    """Same thresholds as the IF() formula written into the workbook."""
    if score >= 15:
        return "High"
    if score >= 8:
        return "Medium"
    return "Low"


@dataclass
class Program:
    name: str
    start: date
    status_date: date
    tasks: list[Task]
    risks: list[Risk] = field(default_factory=list)
    holidays: list[date] = field(default_factory=list)

    def task(self, task_id: str) -> Task:
        for t in self.tasks:
            if t.id == task_id:
                return t
        raise KeyError(task_id)


def parse_preds(text: str | None) -> list[tuple[str, int]]:
    """Parse 'T1, T4+2, T7-1' into [('T1', 0), ('T4', 2), ('T7', -1)]."""
    if text is None:
        return []
    text = str(text).strip()
    if not text:
        return []
    out = []
    for part in re.split(r"[,;]", text):
        if not part.strip():
            continue
        m = _PRED_RE.match(part)
        if not m:
            raise ValueError(f"cannot parse predecessor {part!r}")
        pid, sign, lag = m.groups()
        lag_v = int(lag) if lag else 0
        out.append((pid, -lag_v if sign == "-" else lag_v))
    return out


def format_preds(preds: list[tuple[str, int]]) -> str:
    parts = []
    for pid, lag in preds:
        if lag > 0:
            parts.append(f"{pid}+{lag}")
        elif lag < 0:
            parts.append(f"{pid}{lag}")
        else:
            parts.append(pid)
    return ", ".join(parts)
