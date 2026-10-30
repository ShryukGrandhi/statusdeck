"""Monte Carlo schedule risk analysis over the precedence network.

Each task duration is drawn from a triangular(optimistic, likely, pessimistic)
distribution by inverse-CDF sampling (handles zero-width ranges, which
numpy.random.triangular rejects). All iterations are pushed through the
forward and backward CPM passes together as numpy vectors, one task at a time
in topological order, so cost is O(edges * iterations) with no per-iteration
Python loop. Tasks already 100% complete are held at their likely duration.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from .model import Task
from .network import Network, build_network

PERCENTILES = (10, 50, 80, 90, 95)


def triangular_inv(u: np.ndarray, a: np.ndarray, c: np.ndarray, b: np.ndarray) -> np.ndarray:
    """Inverse CDF of triangular(min=a, mode=c, max=b), broadcasting over columns."""
    width = b - a
    safe = np.where(width > 0, width, 1.0)
    fc = np.where(width > 0, (c - a) / safe, 0.0)
    left = a + np.sqrt(u * safe * (c - a))
    right = b - np.sqrt((1.0 - u) * safe * (b - c))
    x = np.where(u < fc, left, right)
    return np.where(width > 0, x, a)


@dataclass
class MCResult:
    iterations: int
    finish: np.ndarray  # (iterations,) project finish in working days
    criticality: np.ndarray  # (tasks,) share of iterations with zero total float
    deterministic_finish: float

    def percentile(self, q: float) -> float:
        return float(np.percentile(self.finish, q))

    @property
    def percentiles(self) -> dict[int, float]:
        return {q: self.percentile(q) for q in PERCENTILES}

    @property
    def p_on_time(self) -> float:
        """Probability of finishing no later than the deterministic (most-likely) plan."""
        return float(np.mean(self.finish <= self.deterministic_finish + 1e-9))

    @property
    def mean(self) -> float:
        return float(self.finish.mean())

    def histogram(self, bins: int = 20) -> tuple[np.ndarray, np.ndarray]:
        lo, hi = float(self.finish.min()), float(self.finish.max())
        if hi - lo < 1e-9:
            hi = lo + 1.0
        counts, edges = np.histogram(self.finish, bins=bins, range=(lo, hi))
        return counts, edges


def sample_durations(tasks: list[Task], n: int, rng: np.random.Generator) -> np.ndarray:
    a = np.array([t.optimistic for t in tasks], dtype=float)
    c = np.array([t.likely for t in tasks], dtype=float)
    b = np.array([t.pessimistic for t in tasks], dtype=float)
    done = np.array([t.pct_complete >= 1.0 for t in tasks])
    a = np.where(done, c, a)
    b = np.where(done, c, b)
    u = rng.random((n, len(tasks)))
    return triangular_inv(u, a, c, b).T.copy()  # (tasks, n): row per task is contiguous


def simulate(tasks: list[Task], iterations: int = 10_000, seed: int = 0,
             network: Network | None = None, durations: np.ndarray | None = None,
             deterministic_finish: float | None = None) -> MCResult:
    net = network or build_network(tasks)
    rng = np.random.default_rng(seed)
    d = durations if durations is not None else sample_durations(tasks, iterations, rng)
    k, n = d.shape
    es = np.zeros((k, n))
    ef = np.empty((k, n))
    for i in net.order:
        row = es[i]
        for p, lag in net.preds[i]:
            np.maximum(row, ef[p] + lag, out=row)
        np.add(row, d[i], out=ef[i])
    finish = ef.max(axis=0)
    ls = np.empty((k, n))
    crit = np.zeros(k)
    for i in reversed(net.order):
        lf = finish.copy()
        for s, lag in net.succs[i]:
            np.minimum(lf, ls[s] - lag, out=lf)
        np.subtract(lf, d[i], out=ls[i])
        crit[i] = np.count_nonzero(ls[i] - es[i] <= 1e-9) / n
    if deterministic_finish is None:
        from .cpm import schedule
        deterministic_finish = schedule(tasks, net).finish
    return MCResult(n, finish, crit, deterministic_finish)
