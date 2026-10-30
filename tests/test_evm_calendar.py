from datetime import date

import pytest

from statusdeck.analysis import analyze, rag
from statusdeck.cpm import schedule
from statusdeck.evm import compute, planned_fraction
from statusdeck.workcal import WorkCalendar


def test_calendar_skips_weekends_and_holidays():
    cal = WorkCalendar(date(2027, 1, 1), holidays=[date(2027, 1, 18)])  # Fri start
    assert cal.date_of(0) == date(2027, 1, 1)
    assert cal.date_of(1) == date(2027, 1, 4)  # skips Sat/Sun
    assert cal.date_of(11) == date(2027, 1, 19)  # skips the holiday Monday
    assert cal.elapsed(date(2027, 1, 1)) == 0
    assert cal.elapsed(date(2027, 1, 2)) == 1  # Saturday: Friday already done
    assert cal.elapsed(date(2027, 1, 4)) == 1
    assert cal.elapsed(date(2027, 1, 19)) == 11


def test_calendar_start_on_weekend():
    cal = WorkCalendar(date(2027, 1, 2))  # Saturday
    assert cal.date_of(0) == date(2027, 1, 4)


def test_planned_fraction():
    assert planned_fraction(es=3, dur=4, elapsed=5) == 0.5
    assert planned_fraction(es=3, dur=4, elapsed=1) == 0
    assert planned_fraction(es=3, dur=4, elapsed=99) == 1
    assert planned_fraction(es=3, dur=0, elapsed=3) == 1
    assert planned_fraction(es=3, dur=0, elapsed=2) == 0


def test_evm_hand_computed(small_program):
    plan = schedule(small_program.tasks)
    e = compute(small_program, plan, WorkCalendar(small_program.start))
    assert e.elapsed == 5
    assert e.bac == 22000
    assert e.pv == 7000  # A 3000 + B 4000*0.5 + C 2000
    assert e.ev == 8500  # A 3000 + B 2000 + C 2000 + E 1500
    assert e.ac == 8900
    assert e.sv == 1500 and e.cv == -400
    assert e.spi == pytest.approx(8500 / 7000)
    assert e.cpi == pytest.approx(8500 / 8900)
    assert e.eac == pytest.approx(22000 * 8900 / 8500)
    assert e.tcpi == pytest.approx(13500 / 13100)
    assert e.curve[0][2] == 0 and e.curve[-1][2] == pytest.approx(e.bac)
    cum = [v for _, _, v in e.curve]
    assert cum == sorted(cum), "planned value curve must be non-decreasing"


def test_evm_ratios_none_without_progress(small_program):
    for t in small_program.tasks:
        t.pct_complete = t.actual_cost = 0
    small_program.status_date = small_program.start
    e = compute(small_program, schedule(small_program.tasks), WorkCalendar(small_program.start))
    assert e.pv == 0 and e.spi is None and e.cpi is None and e.eac is None


def test_rag_thresholds():
    assert [rag(x) for x in (None, 1.0, 0.95, 0.94, 0.85, 0.84)] == \
        ["N/A", "Green", "Green", "Amber", "Amber", "Red"]


def test_analyze_rejects_invalid(small_program):
    from statusdeck.validate import ValidationError
    small_program.tasks[0].optimistic = 99
    with pytest.raises(ValidationError, match="optimistic"):
        analyze(small_program, iterations=10)
