"""Recalculate a generated workbook with pycel and compare to the Python engine.

pycel is an independent Python implementation of Excel's formula language. It
is an optional (dev) dependency: `pip install statusdeck[dev]`.
"""

from __future__ import annotations

import math
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

EXCEL_EPOCH = date(1899, 12, 30)


def _compat() -> None:
    # pycel 1.0b30 reads `defined_names.definedName`, which openpyxl 3.1 removed.
    from openpyxl.workbook.defined_name import DefinedNameDict

    if not hasattr(DefinedNameDict, "definedName"):
        DefinedNameDict.definedName = property(lambda self: list(self.values()))


def excel_serial(d: date) -> int:
    if isinstance(d, datetime):
        d = d.date()
    return (d - EXCEL_EPOCH).days


def same(got, want, tol: float = 1e-6) -> bool:
    if isinstance(want, str):
        return got == want or (want == "" and got is None)
    if isinstance(want, date):
        if isinstance(got, (date, datetime)):
            return excel_serial(got) == excel_serial(want)
        want = excel_serial(want)
    try:
        g = float(got)
    except (TypeError, ValueError):
        return False
    return math.isclose(g, float(want), rel_tol=tol, abs_tol=tol)


@dataclass
class VerifyReport:
    checked: int
    seconds: float
    mismatches: list[tuple[str, object, object]] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.mismatches


def verify_workbook(path: str | Path, expected: dict[str, object]) -> VerifyReport:
    _compat()
    from pycel import ExcelCompiler

    t0 = time.perf_counter()
    xl = ExcelCompiler(filename=str(path))
    bad = []
    for ref, want in expected.items():
        got = xl.evaluate(ref)
        if not same(got, want):
            bad.append((ref, got, want))
    return VerifyReport(len(expected), time.perf_counter() - t0, bad)
