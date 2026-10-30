import pytest

from statusdeck.cpm import schedule
from statusdeck.model import Task, format_preds, parse_preds
from statusdeck.network import CycleError, NetworkError, build_network
from statusdeck.sample import generate


def by_id(res, field):
    return dict(zip(res.network.ids, getattr(res, field)))


def test_textbook_values(textbook):
    r = schedule(textbook)
    es, ls = by_id(r, "es"), by_id(r, "ls")
    tf, ff = by_id(r, "total_float"), by_id(r, "free_float")
    # A3 B4 D5 F2 = 14 ; A3 C2 E6 F2 = 13
    assert r.finish == 14
    assert es == {"A": 0, "B": 3, "C": 3, "D": 7, "E": 5, "F": 12, "G": 14}
    assert ls == {"A": 0, "B": 3, "C": 4, "D": 7, "E": 6, "F": 12, "G": 14}
    assert tf == {"A": 0, "B": 0, "C": 1, "D": 0, "E": 1, "F": 0, "G": 0}
    assert ff == {"A": 0, "B": 0, "C": 0, "D": 0, "E": 1, "F": 0, "G": 0}
    assert r.critical_ids == ["A", "B", "D", "F", "G"]
    assert r.driving_path() == ["A", "B", "D", "F", "G"]


def test_lags_positive_and_negative():
    tasks = [Task("A", "a", likely=5), Task("B", "b", likely=3, preds=[("A", 2)]),
             Task("C", "c", likely=4, preds=[("A", -3)])]
    r = schedule(tasks)
    es = dict(zip(r.network.ids, r.es))
    assert es == {"A": 0, "B": 7, "C": 2}
    assert r.finish == 10
    # C: LF = finish 10 (no successors), LS = 6 -> float 4
    assert dict(zip(r.network.ids, r.total_float))["C"] == 4


def test_negative_lag_clamped_at_project_start():
    tasks = [Task("A", "a", likely=1), Task("B", "b", likely=2, preds=[("A", -5)])]
    r = schedule(tasks)
    assert r.es[1] == 0


def test_cycle_is_reported_with_path():
    tasks = [Task("A", "a", likely=1, preds=[("C", 0)]), Task("B", "b", likely=1, preds=[("A", 0)]),
             Task("C", "c", likely=1, preds=[("B", 0)]), Task("D", "d", likely=1)]
    with pytest.raises(CycleError) as e:
        build_network(tasks)
    cyc = e.value.cycle
    assert cyc[0] == cyc[-1] and set(cyc) == {"A", "B", "C"}
    # consecutive pairs must be real dependency edges (pred -> succ)
    edges = {("C", "A"), ("A", "B"), ("B", "C")}
    assert all((a, b) in edges for a, b in zip(cyc, cyc[1:]))


def test_self_dependency_and_unknown_and_duplicate():
    with pytest.raises(CycleError):
        build_network([Task("A", "a", preds=[("A", 0)])])
    with pytest.raises(NetworkError, match="unknown"):
        build_network([Task("A", "a", preds=[("Z", 0)])])
    with pytest.raises(NetworkError, match="duplicate"):
        build_network([Task("A", "a"), Task("A", "b")])


def test_topological_order_respects_edges_on_random_network():
    tasks = generate(400, seed=3).tasks
    net = build_network(tasks)
    pos = {i: k for k, i in enumerate(net.order)}
    for i, preds in enumerate(net.preds):
        for p, _ in preds:
            assert pos[p] < pos[i]


def test_cpm_invariants_on_random_network():
    tasks = generate(300, seed=11).tasks
    r = schedule(tasks)
    net = r.network
    for i in range(net.size):
        assert r.total_float[i] >= -1e-9
        assert r.free_float[i] >= -1e-9
        assert r.free_float[i] <= r.total_float[i] + 1e-9
        for p, lag in net.preds[i]:
            assert r.es[i] >= r.ef[p] + lag - 1e-9
    assert r.critical_ids, "a network always has at least one critical task"
    path = r.driving_path()
    assert r.ef[net.index[path[-1]]] == r.finish
    assert len(net.preds[net.index[path[0]]]) == 0 or r.es[net.index[path[0]]] == 0


def test_parse_and_format_preds():
    assert parse_preds("T1, T4+2; T7-1") == [("T1", 0), ("T4", 2), ("T7", -1)]
    assert parse_preds(None) == [] and parse_preds("  ") == []
    assert format_preds([("T1", 0), ("T4", 2), ("T7", -1)]) == "T1, T4+2, T7-1"
    with pytest.raises(ValueError):
        parse_preds("T1 + x")
