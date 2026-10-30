"""Benchmark: CPM and Monte Carlo throughput, Monte Carlo accuracy against
closed-form answers, Excel formula fidelity (pycel recalculation), and
end-to-end report build time.

    python bench/run_bench.py            # full run, writes bench/results.md
    python bench/run_bench.py --quick    # smaller sizes (used in CI)
"""

from __future__ import annotations

import argparse
import platform
import statistics
import sys
import tempfile
import time
from pathlib import Path

import numpy as np

from statusdeck.analysis import analyze
from statusdeck.cpm import schedule
from statusdeck.deck import build_deck
from statusdeck.excel_report import build_workbook
from statusdeck.model import Task
from statusdeck.montecarlo import simulate
from statusdeck.network import build_network
from statusdeck.sample import generate
from statusdeck.tracker import read_tracker, write_tracker
from statusdeck.verify import verify_workbook


def timed(fn, repeat=5):
    times = []
    out = None
    for _ in range(repeat):
        t0 = time.perf_counter()
        out = fn()
        times.append(time.perf_counter() - t0)
    return statistics.median(times), out


def bench_cpm(sizes):
    rows = ["| Tasks | Edges | CPM time (median of 5) | Tasks / s |", "|---:|---:|---:|---:|"]
    for n in sizes:
        tasks = generate(n, seed=1).tasks
        edges = sum(len(t.preds) for t in tasks)
        t, _ = timed(lambda: schedule(tasks, build_network(tasks)))
        rows.append(f"| {n:,} | {edges:,} | {t * 1e3:.2f} ms | {n / t:,.0f} |")
    return rows


def bench_mc(cases):
    rows = ["| Tasks | Iterations | Time (median of 3) | Iterations / s |", "|---:|---:|---:|---:|"]
    for n, it in cases:
        tasks = generate(n, seed=1).tasks
        net = build_network(tasks)
        t, _ = timed(lambda: simulate(tasks, it, seed=0, network=net, deterministic_finish=0),
                     repeat=3)
        rows.append(f"| {n:,} | {it:,} | {t:.3f} s | {it / t:,.0f} |")
    return rows


def bench_accuracy(iterations):
    rows = ["| Check | Expected | Simulated | Error |", "|---|---:|---:|---:|"]
    a, c, b, k = 4.0, 6.0, 12.0, 30
    chain = [Task(f"T{i}", "t", optimistic=a, likely=c, pessimistic=b,
                  preds=[(f"T{i - 1}", 0)] if i else []) for i in range(k)]
    mc = simulate(chain, iterations, seed=123)
    mean = k * (a + b + c) / 3
    sd = (k * (a * a + b * b + c * c - a * b - a * c - b * c) / 18) ** 0.5
    rows.append(f"| 30-task serial chain, mean finish | {mean:.3f} | {mc.mean:.3f} | "
                f"{abs(mc.mean - mean) / mean:.3%} |")
    rows.append(f"| 30-task serial chain, std dev | {sd:.3f} | {mc.finish.std():.3f} | "
                f"{abs(mc.finish.std() - sd) / sd:.3%} |")
    # Two parallel triangular(5,10,20) paths: E[max] by numeric integration of the CDF.
    par = [Task("A", "a", optimistic=5, likely=10, pessimistic=20),
           Task("B", "b", optimistic=5, likely=10, pessimistic=20),
           Task("M", "m", preds=[("A", 0), ("B", 0)])]
    mc2 = simulate(par, iterations, seed=321)
    x = np.linspace(5, 20, 200_001)
    cdf = np.where(x < 10, (x - 5) ** 2 / (15 * 5), 1 - (20 - x) ** 2 / (15 * 10))
    exp_max = 5 + np.trapezoid(1 - cdf ** 2, x) if hasattr(np, "trapezoid") else \
        5 + np.trapz(1 - cdf ** 2, x)
    rows.append(f"| 2 parallel paths, E[max] (merge bias) | {exp_max:.3f} | {mc2.mean:.3f} | "
                f"{abs(mc2.mean - exp_max) / exp_max:.3%} |")
    # Degenerate estimates: every iteration must equal the CPM answer exactly.
    det = generate(200, seed=2).tasks
    for t in det:
        t.optimistic = t.pessimistic = t.likely
    mc3 = simulate(det, 1000, seed=0)
    exact = int(np.sum(mc3.finish == schedule(det).finish))
    rows.append(f"| 200-task network, zero-width estimates: iterations equal to CPM | 1,000 | "
                f"{exact:,} | {1000 - exact} |")
    return rows


def bench_excel(sizes, tmp):
    rows = ["| Tasks | Formula cells | Matching after pycel recalculation | Write time | "
            "Recalc time |", "|---:|---:|---:|---:|---:|"]
    for n in sizes:
        a = analyze(generate(n, seed=1), iterations=1000, seed=0)
        path = tmp / f"w{n}.xlsx"
        t0 = time.perf_counter()
        _, expected = build_workbook(a, path)
        tw = time.perf_counter() - t0
        rep = verify_workbook(path, expected)
        rows.append(f"| {n:,} | {rep.checked:,} | {rep.checked - len(rep.mismatches):,} "
                    f"({(rep.checked - len(rep.mismatches)) / rep.checked:.1%}) | "
                    f"{tw:.2f} s | {rep.seconds:.2f} s |")
    return rows


def bench_end_to_end(n, iterations, tmp):
    tracker = tmp / "tracker.xlsx"
    write_tracker(generate(n, seed=7), tracker)
    t = {}
    t0 = time.perf_counter()
    prog = read_tracker(tracker)
    t["read tracker"] = time.perf_counter() - t0
    t0 = time.perf_counter()
    a = analyze(prog, iterations=iterations, seed=0)
    t[f"analyze (CPM + EVM + {iterations:,} MC iterations)"] = time.perf_counter() - t0
    t0 = time.perf_counter()
    build_workbook(a, tmp / "w.xlsx")
    t["write Excel workbook"] = time.perf_counter() - t0
    t0 = time.perf_counter()
    build_deck(a, tmp / "d.pptx")
    t["write PowerPoint deck"] = time.perf_counter() - t0
    rows = ["| Stage | Time |", "|---|---:|"]
    rows += [f"| {k} | {v:.3f} s |" for k, v in t.items()]
    rows.append(f"| **total** | **{sum(t.values()):.3f} s** |")
    sizes = (f"Output sizes: workbook {(tmp / 'w.xlsx').stat().st_size / 1024:.0f} KiB, "
             f"deck {(tmp / 'd.pptx').stat().st_size / 1024:.0f} KiB.")
    return rows, sizes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--out", default=None, help="write markdown here (default bench/results.md "
                                                "for a full run)")
    args = ap.parse_args()
    q = args.quick
    out: list[str] = [
        "# statusdeck benchmark results", "",
        f"Command: `python bench/run_bench.py{' --quick' if q else ''}`  ",
        f"Machine: {platform.machine()} / {platform.system()} {platform.release()}, "
        f"Python {platform.python_version()}, numpy {np.__version__}", "",
    ]
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        sections = [
            ("CPM forward/backward pass", bench_cpm([100, 1000] if q else [100, 1000, 10000])),
            ("Monte Carlo throughput (vectorised over iterations)",
             bench_mc([(60, 2000)] if q else [(60, 10000), (500, 10000), (2000, 10000)])),
            ("Monte Carlo accuracy against closed-form answers",
             bench_accuracy(20000 if q else 200000)),
            ("Excel formula fidelity (workbook recalculated by pycel vs. Python engine)",
             bench_excel([60] if q else [60, 300], tmp)),
        ]
        e2e, sizes = bench_end_to_end(60, 2000 if q else 10000, tmp)
        sections.append(("End-to-end: 60-task tracker to workbook + deck", e2e))
        for title, rows in sections:
            out += [f"## {title}", "", *rows, ""]
        out.append(sizes)
    text = "\n".join(out) + "\n"
    print(text)
    path = args.out or (None if q else str(Path(__file__).with_name("results.md")))
    if path:
        Path(path).write_text(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
