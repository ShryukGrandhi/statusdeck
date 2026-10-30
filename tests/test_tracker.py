from datetime import date

from openpyxl import Workbook

from statusdeck.sample import generate
from statusdeck.tracker import read_tracker, write_tracker
from statusdeck.validate import check


def test_round_trip(tmp_path):
    prog = generate(40, seed=2)
    path = tmp_path / "t.xlsx"
    write_tracker(prog, path)
    back = read_tracker(path)
    assert back.name == prog.name
    assert back.start == prog.start and back.status_date == prog.status_date
    assert back.holidays == prog.holidays
    assert [(t.id, t.preds, t.likely, t.pessimistic, t.pct_complete, t.budget)
            for t in back.tasks] == \
        [(t.id, t.preds, t.likely, t.pessimistic, t.pct_complete, t.budget) for t in prog.tasks]
    assert [(r.id, r.score) for r in back.risks] == [(r.id, r.score) for r in prog.risks]


def test_reads_reordered_columns_aliases_and_percent_strings(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Program"
    ws.append(["Program", "Hand made"])
    ws.append(["Start Date", "2027-03-01"])
    ws.append(["Status Date", date(2027, 3, 10)])
    t = wb.create_sheet("Tasks")
    t.append(["Duration", "Name", "ID", "Depends On", "% Complete", "Budget", "Notes"])
    t.append([5, "Design", "D1", None, "100%", "$1,000", "x"])
    t.append([3, "Build", "B1", "D1+1", 40, 600, None])
    t.append([None, None, None, None, None, None, None])  # blank row ignored
    path = tmp_path / "h.xlsx"
    wb.save(path)
    prog = read_tracker(path)
    assert prog.start == date(2027, 3, 1)
    d, b = prog.tasks
    assert (d.id, d.likely, d.optimistic, d.pessimistic, d.pct_complete, d.budget) == \
        ("D1", 5, 5, 5, 1.0, 1000)
    assert b.preds == [("D1", 1)] and b.pct_complete == 0.4
    assert prog.risks == [] and check(prog) == []


def test_validation_messages():
    prog = generate(10, seed=1)
    prog.tasks[3].optimistic = prog.tasks[3].likely + 1
    prog.tasks[4].pct_complete = 1.5
    prog.tasks[0].preds = [("T9", 0)]  # creates a cycle through the layered DAG
    prog.risks[0].impact = 7
    prog.status_date = date(2020, 1, 1)
    msgs = "\n".join(check(prog))
    for needle in ("optimistic <= likely", "% complete", "impact must be", "before project start",
                   "cycle"):
        assert needle in msgs
