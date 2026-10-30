"""Critical Path Method: forward/backward pass, total and free float.

Conventions (mirrored exactly by the formulas in the Excel workbook):
  ES = MAX(0, EF(pred) + lag ...)          EF = ES + duration
  LF = MIN(project finish, LS(succ) - lag)  LS = LF - duration
  total float = LS - ES; free float = MIN(ES(succ) - lag, finish) - EF
Times are working-day offsets from the project start.
"""

from __future__ import annotations

from dataclasses import dataclass

from .model import Task
from .network import Network, build_network

EPS = 1e-9


@dataclass
class CPMResult:
    network: Network
    duration: list[float]
    es: list[float]
    ef: list[float]
    ls: list[float]
    lf: list[float]
    total_float: list[float]
    free_float: list[float]
    finish: float

    def is_critical(self, i: int) -> bool:
        return self.total_float[i] <= EPS

    @property
    def critical_ids(self) -> list[str]:
        crit = [i for i in range(self.network.size) if self.is_critical(i)]
        crit.sort(key=lambda i: (self.es[i], self.ef[i]))
        return [self.network.ids[i] for i in crit]

    def driving_path(self) -> list[str]:
        """One zero-float chain from project start to the latest-finishing task."""
        net = self.network
        pos = {i: k for k, i in enumerate(net.order)}
        end = max((i for i in range(net.size) if self.is_critical(i)),
                  key=lambda i: (self.ef[i], pos[i]))
        path = [end]
        cur = end
        while True:
            nxt = None
            for p, lag in net.preds[cur]:
                if self.is_critical(p) and abs(self.ef[p] + lag - self.es[cur]) <= EPS:
                    nxt = p
                    break
            if nxt is None:
                break
            path.append(nxt)
            cur = nxt
        path.reverse()
        return [net.ids[i] for i in path]


def schedule(tasks: list[Task], network: Network | None = None,
             durations: list[float] | None = None) -> CPMResult:
    net = network or build_network(tasks)
    dur = list(durations) if durations is not None else [float(t.duration) for t in tasks]
    n = net.size
    es = [0.0] * n
    ef = [0.0] * n
    for i in net.order:
        start = 0.0
        for p, lag in net.preds[i]:
            v = ef[p] + lag
            if v > start:
                start = v
        es[i] = start
        ef[i] = start + dur[i]
    finish = max(ef) if n else 0.0
    ls = [0.0] * n
    lf = [0.0] * n
    ff = [0.0] * n
    for i in reversed(net.order):
        late = finish
        early_succ = finish
        for s, lag in net.succs[i]:
            late = min(late, ls[s] - lag)
            early_succ = min(early_succ, es[s] - lag)
        lf[i] = late
        ls[i] = late - dur[i]
        ff[i] = early_succ - ef[i]
    tf = [ls[i] - es[i] for i in range(n)]
    return CPMResult(net, dur, es, ef, ls, lf, tf, ff, finish)
