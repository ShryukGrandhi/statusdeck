"""One-call pipeline: validate -> CPM -> EVM -> Monte Carlo. Shared by CLI, reports, bench."""

from __future__ import annotations

from dataclasses import dataclass

from .cpm import CPMResult, schedule
from .evm import EVMResult, compute
from .model import Program
from .montecarlo import MCResult, simulate
from .network import build_network
from .validate import require_valid
from .workcal import WorkCalendar


def rag(index: float | None) -> str:
    """RAG status for a performance index (SPI/CPI). Same thresholds as the workbook."""
    if index is None:
        return "N/A"
    if index >= 0.95:
        return "Green"
    if index >= 0.85:
        return "Amber"
    return "Red"


@dataclass
class Analysis:
    program: Program
    calendar: WorkCalendar
    cpm: CPMResult
    evm: EVMResult
    mc: MCResult

    def date_at(self, workdays: float):
        """Calendar date on which a task ending at `workdays` finishes."""
        return self.calendar.date_of(max(int(workdays + 0.999999) - 1, 0))

    @property
    def baseline_finish_date(self):
        return self.date_at(self.cpm.finish)

    def task_start(self, i: int):
        return self.calendar.date_of(int(self.cpm.es[i]))

    def task_finish(self, i: int):
        return self.calendar.date_of(int(max(self.cpm.ef[i] - 1, self.cpm.es[i])))


def analyze(program: Program, iterations: int = 10_000, seed: int = 0) -> Analysis:
    require_valid(program)
    net = build_network(program.tasks)
    plan = schedule(program.tasks, net)
    cal = WorkCalendar(program.start, program.holidays)
    evm = compute(program, plan, cal)
    mc = simulate(program.tasks, iterations, seed, network=net, deterministic_finish=plan.finish)
    return Analysis(program, cal, plan, evm, mc)
