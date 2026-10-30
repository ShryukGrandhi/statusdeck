"""Command line: statusdeck {build,validate,sample}."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from .analysis import analyze, rag
from .deck import build_deck
from .excel_report import build_workbook
from .tracker import read_tracker, write_tracker
from .validate import ValidationError, check


def _build(args) -> int:
    t0 = time.perf_counter()
    program = read_tracker(args.tracker)
    try:
        a = analyze(program, iterations=args.iterations, seed=args.seed)
    except ValidationError as e:
        for p in e.problems:
            print(f"error: {p}", file=sys.stderr)
        return 2
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx, pptx = out / "status_workbook.xlsx", out / "status_deck.pptx"
    _, expected = build_workbook(a, xlsx)
    build_deck(a, pptx)
    e, mc = a.evm, a.mc
    print(f"{program.name}: {len(program.tasks)} tasks, status {program.status_date}")
    print(f"  baseline finish  {a.baseline_finish_date} ({a.cpm.finish:g} working days)")
    print(f"  critical path    {' -> '.join(a.cpm.driving_path())}")
    print(f"  SPI {e.spi:.3f} ({rag(e.spi)})  CPI {e.cpi:.3f} ({rag(e.cpi)})"
          if e.spi is not None and e.cpi is not None else "  SPI/CPI n/a (no PV or AC yet)")
    print(f"  Monte Carlo P50 {a.date_at(mc.percentile(50))}  P80 {a.date_at(mc.percentile(80))}"
          f"  P(on time) {mc.p_on_time:.1%}")
    print(f"  wrote {xlsx} ({len(expected)} formula cells) and {pptx}")
    if args.verify:
        from .verify import verify_workbook
        rep = verify_workbook(xlsx, expected)
        print(f"  pycel recalculation: {rep.checked - len(rep.mismatches)}/{rep.checked} "
              f"formula cells match")
        for ref, got, want in rep.mismatches[:10]:
            print(f"    {ref}: workbook {got!r} != engine {want!r}")
        if not rep.ok:
            return 1
    print(f"  done in {time.perf_counter() - t0:.2f}s")
    return 0


def _validate(args) -> int:
    problems = check(read_tracker(args.tracker))
    for p in problems:
        print(f"error: {p}")
    print("ok" if not problems else f"{len(problems)} problem(s)")
    return 1 if problems else 0


def _sample(args) -> int:
    from .sample import generate
    write_tracker(generate(args.tasks, seed=args.seed), args.path)
    print(f"wrote {args.path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="statusdeck", description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="tracker.xlsx -> status workbook + PowerPoint deck")
    b.add_argument("tracker")
    b.add_argument("--out", default="out")
    b.add_argument("--iterations", type=int, default=10_000)
    b.add_argument("--seed", type=int, default=0)
    b.add_argument("--verify", action="store_true",
                   help="recalculate the workbook with pycel and compare (needs [dev] extra)")
    b.set_defaults(fn=_build)
    v = sub.add_parser("validate", help="check a tracker for errors")
    v.add_argument("tracker")
    v.set_defaults(fn=_validate)
    s = sub.add_parser("sample", help="write a synthetic example tracker")
    s.add_argument("path")
    s.add_argument("--tasks", type=int, default=60)
    s.add_argument("--seed", type=int, default=7)
    s.set_defaults(fn=_sample)
    args = ap.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
