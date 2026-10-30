"""Deterministic synthetic tracker: a vehicle service-campaign (recall) rollout program.

Used for the example, the tests and the benchmark. The generator builds a
layered DAG (every task depends only on earlier tasks), right-skewed
three-point estimates, and progress/cost data consistent with a status date.
"""

from __future__ import annotations

from datetime import date

import numpy as np

from .cpm import schedule
from .model import Program, Risk, Task
from .workcal import WorkCalendar

WORKSTREAMS = {
    "Engineering": ["Root-cause analysis", "Fix design review", "Software update build",
                    "Validation testing", "Field test", "Design release"],
    "Parts & Supply": ["Supplier quote", "Tooling change", "Parts PPAP", "Initial parts build",
                       "Distribution to service centers", "Safety stock build"],
    "Service Ops": ["Repair procedure draft", "Technician training", "Service bulletin",
                    "Labor time study", "Mobile service scheduling", "Center readiness check"],
    "Compliance": ["Regulatory filing", "Defect report submission", "Owner notification letter",
                   "Quarterly status report", "Remedy adequacy review", "Records audit"],
    "Customer Comms": ["FAQ and scripts", "In-app notification", "Call center briefing",
                       "Dealer communication", "Customer survey", "Website update"],
}
RISKS = [
    ("Supplier capacity shortfall delays parts", "Qualify second source; weekly supply call"),
    ("Regulator requests additional remedy data", "Pre-brief regulator; hold test data package"),
    ("Software update fails field validation", "Staged rollout with rollback plan"),
    ("Technician training completion below target", "Add remote sessions; track by center"),
    ("Owner contact data incomplete", "Cross-match registration data; mail fallback"),
    ("Parts quality escape at first build", "100% inspection on first lot"),
    ("Service center capacity constrained", "Extended hours; mobile service"),
    ("Scope grows to additional model years", "Pre-scope VIN ranges with engineering"),
    ("Labor time underestimated", "Time study at pilot centers before release"),
    ("Translation of notices delayed", "Start localisation from draft text"),
    ("Key engineer unavailable", "Named backup on each critical task"),
    ("Logistics disruption at port", "Pre-position stock regionally"),
]
OWNERS = ["A. Rivera", "B. Chen", "C. Okafor", "D. Patel", "E. Novak", "F. Haddad", "G. Kim"]


def generate(n_tasks: int = 60, seed: int = 7, start: date = date(2027, 1, 4),
             progress: float = 0.4, n_risks: int = 12) -> Program:
    rng = np.random.default_rng(seed)
    names = list(WORKSTREAMS)
    tasks: list[Task] = []
    for i in range(n_tasks):
        ws = names[i % len(names)]
        base = WORKSTREAMS[ws][(i // len(names)) % len(WORKSTREAMS[ws])]
        wave = i // (len(names) * len(WORKSTREAMS[ws]))
        name = base if wave == 0 else f"{base} (wave {wave + 1})"
        preds = []
        if i:
            k = int(rng.integers(1, min(3, i) + 1))
            window = range(max(0, i - 12), i)
            for p in sorted(rng.choice(list(window), size=min(k, len(window)), replace=False)):
                lag = int(rng.choice([0, 0, 0, 0, 1, 2]))
                preds.append((f"T{p + 1}", lag))
        milestone = i > 0 and rng.random() < 0.06
        likely = 0.0 if milestone else float(rng.integers(2, 16))
        opt = round(likely * rng.uniform(0.6, 0.9))
        pess = round(likely * rng.uniform(1.2, 2.2))
        tasks.append(Task(
            id=f"T{i + 1}", name=name, workstream=ws, owner=OWNERS[int(rng.integers(len(OWNERS)))],
            preds=preds, optimistic=float(min(opt, likely)), likely=likely,
            pessimistic=float(max(pess, likely)),
            budget=float(likely * int(rng.choice([800, 1200, 1500, 2000]))) if likely else 0.0,
        ))
    holidays = [date(2027, 1, 18), date(2027, 2, 15), date(2027, 5, 31), date(2027, 7, 5),
                date(2027, 9, 6), date(2027, 11, 25), date(2027, 12, 24)]
    plan = schedule(tasks)
    cal = WorkCalendar(start, holidays)
    status_idx = int(plan.finish * progress)
    status = cal.date_of(status_idx)
    for i, t in enumerate(tasks):
        if t.likely == 0:
            planned = 1.0 if status_idx >= plan.es[i] else 0.0
        else:
            planned = min(1.0, max(0.0, (status_idx - plan.es[i]) / t.likely))
        # Actual progress drifts from plan: most tasks a little behind.
        actual = planned if planned in (0.0, 1.0) else planned * rng.uniform(0.45, 1.05)
        t.pct_complete = round(min(1.0, actual), 2)
        t.actual_cost = round(t.budget * t.pct_complete * rng.uniform(0.9, 1.25), 0)
    risks = []
    for k in range(n_risks):
        title, mitigation = RISKS[k % len(RISKS)]
        risks.append(Risk(id=f"R{k + 1}", title=title, owner=OWNERS[k % len(OWNERS)],
                          probability=int(rng.integers(1, 6)), impact=int(rng.integers(1, 6)),
                          mitigation=mitigation))
    return Program(name="Service Campaign Rollout", start=start, status_date=status,
                   tasks=tasks, risks=risks, holidays=holidays)


