from datetime import date

import pytest

from statusdeck.analysis import analyze
from statusdeck.model import Program, Risk, Task
from statusdeck.sample import generate


def textbook_tasks():
    """Classic 7-activity AON network with a known answer.

    A(3) -> B(4), C(2); B -> D(5); C -> E(6); D, E -> F(2); F -> G(0, milestone)
    Paths: A-B-D-F = 14, A-C-E-F = 13  => critical A,B,D,F,G; C/E have 1 day float.
    """
    spec = [("A", 3, ""), ("B", 4, "A"), ("C", 2, "A"), ("D", 5, "B"), ("E", 6, "C"),
            ("F", 2, "D, E"), ("G", 0, "F")]
    from statusdeck.model import parse_preds
    return [Task(id=i, name=f"Task {i}", likely=d, optimistic=d, pessimistic=d,
                 preds=parse_preds(p), budget=1000 * d) for i, d, p in spec]


@pytest.fixture
def textbook():
    return textbook_tasks()


@pytest.fixture
def small_program():
    tasks = textbook_tasks()
    for t, pct, ac in zip(tasks, [1.0, 0.5, 1.0, 0.0, 0.25, 0, 0], [3200, 2100, 1900, 0, 1700, 0, 0]):
        t.pct_complete, t.actual_cost = pct, ac
    risks = [Risk("R1", "Vendor late", "Ann", 4, 4), Risk("R2", "Spec churn", "Bo", 2, 4),
             Risk("R3", "Minor", "Cy", 1, 2)]
    # 2027-01-04 is a Monday; status Monday 2027-01-11 -> 5 working days elapsed.
    return Program("Small", date(2027, 1, 4), date(2027, 1, 11), tasks, risks, [])


@pytest.fixture(scope="session")
def sample_analysis():
    return analyze(generate(60, seed=7), iterations=4000, seed=1)
