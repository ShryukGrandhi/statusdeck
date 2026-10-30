"""Input checks run before any scheduling. Returns human-readable problems."""

from __future__ import annotations

from .model import Program
from .network import NetworkError, build_network


class ValidationError(ValueError):
    def __init__(self, problems: list[str]):
        self.problems = problems
        super().__init__("; ".join(problems))


def check(program: Program) -> list[str]:
    problems: list[str] = []
    if not program.tasks:
        problems.append("tracker has no tasks")
    if program.status_date < program.start:
        problems.append("status date is before project start")
    for t in program.tasks:
        where = f"task {t.id}"
        if not (0 <= t.optimistic <= t.likely <= t.pessimistic):
            problems.append(f"{where}: need 0 <= optimistic <= likely <= pessimistic "
                            f"(got {t.optimistic}, {t.likely}, {t.pessimistic})")
        if not 0 <= t.pct_complete <= 1:
            problems.append(f"{where}: % complete must be within 0..100%")
        if t.budget < 0 or t.actual_cost < 0:
            problems.append(f"{where}: budget and actual cost must be non-negative")
    for r in program.risks:
        for label, v in (("probability", r.probability), ("impact", r.impact)):
            if v not in (1, 2, 3, 4, 5):
                problems.append(f"risk {r.id}: {label} must be an integer 1..5 (got {v})")
    try:
        build_network(program.tasks)
    except NetworkError as e:
        problems.append(str(e))
    return problems


def require_valid(program: Program) -> None:
    problems = check(program)
    if problems:
        raise ValidationError(problems)
