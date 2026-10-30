import numpy as np
import pytest

from statusdeck.cpm import schedule
from statusdeck.model import Task
from statusdeck.montecarlo import simulate, triangular_inv
from statusdeck.sample import generate


def test_triangular_inverse_cdf_moments():
    u = np.random.default_rng(0).random(400_000)
    a, c, b = 2.0, 5.0, 14.0
    x = triangular_inv(u, np.float64(a), np.float64(c), np.float64(b))
    assert x.min() >= a and x.max() <= b
    assert x.mean() == pytest.approx((a + b + c) / 3, rel=5e-3)
    var = (a * a + b * b + c * c - a * b - a * c - b * c) / 18
    assert x.var() == pytest.approx(var, rel=1e-2)
    # mode split: P(X < c) = (c - a) / (b - a)
    assert np.mean(x < c) == pytest.approx((c - a) / (b - a), abs=3e-3)


def test_degenerate_estimates_reproduce_cpm(textbook):
    mc = simulate(textbook, iterations=500, seed=1)
    assert np.all(mc.finish == schedule(textbook).finish)
    assert mc.p_on_time == 1.0
    crit = dict(zip([t.id for t in textbook], mc.criticality))
    assert crit == {"A": 1, "B": 1, "C": 0, "D": 1, "E": 0, "F": 1, "G": 1}


def test_serial_chain_matches_analytic_sum():
    tasks = []
    for i in range(30):
        tasks.append(Task(f"T{i}", "t", optimistic=4, likely=6, pessimistic=12,
                          preds=[(f"T{i - 1}", 0)] if i else []))
    mc = simulate(tasks, iterations=40_000, seed=2)
    mean = 30 * (4 + 6 + 12) / 3
    sd = np.sqrt(30 * (16 + 36 + 144 - 24 - 48 - 72) / 18)
    assert mc.mean == pytest.approx(mean, rel=2e-3)
    assert mc.finish.std() == pytest.approx(sd, rel=2e-2)
    assert np.all(mc.criticality == 1.0)


def test_parallel_paths_merge_bias():
    # Two identical independent paths: E[max] > E[single path] (merge bias).
    tasks = [Task("A", "a", optimistic=5, likely=10, pessimistic=20),
             Task("B", "b", optimistic=5, likely=10, pessimistic=20),
             Task("M", "m", preds=[("A", 0), ("B", 0)])]
    mc = simulate(tasks, iterations=50_000, seed=3)
    single = (5 + 10 + 20) / 3
    assert mc.mean > single + 1.0
    assert mc.criticality[0] + mc.criticality[1] == pytest.approx(1.0, abs=1e-3)
    assert mc.criticality[2] == 1.0


def test_completed_tasks_are_fixed():
    tasks = [Task("A", "a", optimistic=1, likely=5, pessimistic=30, pct_complete=1.0)]
    assert np.all(simulate(tasks, iterations=200).finish == 5)


def test_seed_reproducible_and_percentiles_ordered():
    tasks = generate(80, seed=5).tasks
    a = simulate(tasks, iterations=2000, seed=9)
    b = simulate(tasks, iterations=2000, seed=9)
    assert np.array_equal(a.finish, b.finish)
    p = list(a.percentiles.values())
    assert p == sorted(p)
    counts, edges = a.histogram(20)
    assert counts.sum() == 2000 and len(edges) == 21


def test_vectorised_matches_scalar_cpm_per_iteration():
    tasks = generate(50, seed=4).tasks
    rng = np.random.default_rng(0)
    from statusdeck.montecarlo import sample_durations
    d = sample_durations(tasks, 25, rng)
    mc = simulate(tasks, durations=d)
    crit_counts = np.zeros(len(tasks))
    for k in range(25):
        r = schedule(tasks, durations=list(d[:, k]))
        assert mc.finish[k] == pytest.approx(r.finish)
        crit_counts += [r.is_critical(i) for i in range(len(tasks))]
    assert np.allclose(mc.criticality, crit_counts / 25)
