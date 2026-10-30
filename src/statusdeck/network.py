"""Compiled precedence network shared by the CPM and Monte Carlo engines."""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass

from .model import Task


class NetworkError(ValueError):
    pass


class CycleError(NetworkError):
    def __init__(self, cycle: list[str]):
        self.cycle = cycle
        super().__init__("dependency cycle: " + " -> ".join(cycle))


@dataclass
class Network:
    ids: list[str]
    index: dict[str, int]
    preds: list[list[tuple[int, int]]]  # per task: (pred index, lag)
    succs: list[list[tuple[int, int]]]  # per task: (succ index, lag)
    order: list[int]  # topological order

    @property
    def size(self) -> int:
        return len(self.ids)


def build_network(tasks: list[Task]) -> Network:
    ids = [t.id for t in tasks]
    index: dict[str, int] = {}
    for i, tid in enumerate(ids):
        if tid in index:
            raise NetworkError(f"duplicate task id {tid!r}")
        index[tid] = i
    preds: list[list[tuple[int, int]]] = [[] for _ in tasks]
    succs: list[list[tuple[int, int]]] = [[] for _ in tasks]
    for i, t in enumerate(tasks):
        for pid, lag in t.preds:
            if pid not in index:
                raise NetworkError(f"task {t.id!r} depends on unknown task {pid!r}")
            p = index[pid]
            if p == i:
                raise CycleError([t.id, t.id])
            preds[i].append((p, lag))
            succs[p].append((i, lag))
    order = _topo_order(ids, preds, succs)
    return Network(ids, index, preds, succs, order)


def _topo_order(ids, preds, succs) -> list[int]:
    """Kahn's algorithm; ties broken by input order so output is stable."""
    indeg = [len(p) for p in preds]
    q = deque(i for i, d in enumerate(indeg) if d == 0)
    order = []
    while q:
        i = q.popleft()
        order.append(i)
        for s, _ in succs[i]:
            indeg[s] -= 1
            if indeg[s] == 0:
                q.append(s)
    if len(order) != len(ids):
        raise CycleError(_find_cycle(ids, preds, {i for i, d in enumerate(indeg) if d > 0}))
    return order


def _find_cycle(ids, preds, remaining: set[int]) -> list[str]:
    """Walk predecessor links inside the unsorted remainder until a node repeats."""
    start = min(remaining)
    seen: dict[int, int] = {}
    path: list[int] = []
    node = start
    while node not in seen:
        seen[node] = len(path)
        path.append(node)
        node = next(p for p, _ in preds[node] if p in remaining)
    cyc = path[seen[node]:] + [node]
    cyc.reverse()  # predecessor walk -> report in dependency direction
    return [ids[i] for i in cyc]
