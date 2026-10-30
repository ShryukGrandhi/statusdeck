"""Excel status workbook with live formulas.

The schedule (CPM passes), earned value, and risk sheets are written as Excel
formulas, not pasted values: change a duration, a % complete, or the status
date in Excel and every dependent number recalculates. Monte Carlo output is
written as static values because Excel has no vectorised simulation engine.

`build_workbook` also returns the value Python computed for every formula cell
it wrote ("expected"), so tests and the benchmark can recalculate the saved
file with an independent formula engine (pycel) and compare cell by cell.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from .analysis import Analysis, rag
from .model import risk_rating

NAVY = "1F3864"
RED, AMBER, GREEN = "C00000", "FFC000", "70AD47"
LIGHT = {"High": "F8CBAD", "Medium": "FFE699", "Low": "C6E0B4"}
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '"$"#,##0;[Red]-"$"#,##0'
DATE = "yyyy-mm-dd"
MAX_GANTT_WEEKS = 104

# Schedule sheet layout
S_FIRST = 3
SC = dict(id="A", name="B", ws="C", owner="D", preds="E", opt="F", likely="G", pess="H",
          dur="I", es="J", ef="K", ls="L", lf="M", tf="N", ff="O", crit="P", start="Q",
          finish="R", mc="S")
GANTT_COL = 20  # column T


def _header(ws, row: int, labels: list[str], col: int = 1) -> None:
    for j, label in enumerate(labels):
        c = ws.cell(row, col + j, label)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _widths(ws, widths: dict[str, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _name(wb, name: str, ref: str) -> None:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


class _Book:
    def __init__(self, a: Analysis):
        self.a = a
        self.wb = Workbook()
        self.expected: dict[str, object] = {}
        self.n = len(a.program.tasks)
        self.s_last = S_FIRST + self.n - 1

    def expect(self, ws, ref: str, value) -> None:
        self.expected[f"{ws.title}!{ref}"] = value

    # ------------------------------------------------------------------ build
    def build(self):
        summary = self.wb.active
        summary.title = "Summary"
        self.schedule = self.wb.create_sheet("Schedule")
        self.evm = self.wb.create_sheet("EVM")
        self.risks = self.wb.create_sheet("Risks")
        self.mc = self.wb.create_sheet("MonteCarlo")
        self.cal = self.wb.create_sheet("Calendar")
        self._calendar()
        self._summary_inputs(summary)
        self._schedule()
        self._evm()
        self._risks()
        self._montecarlo()
        self._summary_kpis(summary)
        self.wb.calculation.fullCalcOnLoad = True
        return self.wb

    # --------------------------------------------------------------- calendar
    def _calendar(self):
        a, ws = self.a, self.cal
        days = max(int(a.mc.finish.max()), int(a.cpm.finish), a.evm.elapsed) * 2 + 260
        ws["A1"], ws["B1"], ws["D1"] = "Working day", "Date", "Holidays"
        _header(ws, 1, ["Working day", "Date"])
        _header(ws, 1, ["Holidays"], col=4)
        for k, d in enumerate(a.calendar.dates(days)):
            ws.cell(k + 2, 1, k)
            ws.cell(k + 2, 2, d).number_format = DATE
        for j, h in enumerate(a.program.holidays):
            ws.cell(j + 2, 4, h).number_format = DATE
        _widths(ws, {"A": 12, "B": 12, "D": 12})
        ws.freeze_panes = "A2"
        _name(self.wb, "CalDates", f"Calendar!$B$2:$B${days + 1}")
        self.cal_days = days

    # ---------------------------------------------------------------- summary
    def _summary_inputs(self, ws):
        a = self.a
        ws["A1"] = f"{a.program.name} - Status Report"
        ws["A1"].font = Font(bold=True, size=16, color=NAVY)
        rows = [
            ("Project start", a.program.start, None),
            ("Status date (edit me)", a.program.status_date, None),
            ("Working days elapsed", '=COUNTIF(CalDates,"<"&StatusDate)', a.evm.elapsed),
            ("Baseline finish (working days)",
             f"=MAX(Schedule!$K${S_FIRST}:$K${self.s_last})", a.cpm.finish),
            ("Baseline finish date", "=INDEX(CalDates,MAX(ROUNDUP(ProjFinish,0)-1,0)+1)",
             a.baseline_finish_date),
        ]
        for k, (label, value, exp) in enumerate(rows):
            r = 3 + k
            ws.cell(r, 1, label).font = Font(bold=True)
            c = ws.cell(r, 2, value)
            c.border = BOX
            if k in (0, 1, 4):
                c.number_format = DATE
            if exp is not None:
                self.expect(ws, f"B{r}", exp)
        ws["B4"].fill = PatternFill("solid", fgColor="FFF2CC")
        _name(self.wb, "ProjStart", "Summary!$B$3")
        _name(self.wb, "StatusDate", "Summary!$B$4")
        _name(self.wb, "Elapsed", "Summary!$B$5")
        _name(self.wb, "ProjFinish", "Summary!$B$6")
        _widths(ws, {"A": 34, "B": 16, "C": 12, "D": 4, "E": 30, "F": 16})

    def _summary_kpis(self, ws):
        a, e = self.a, self.a.evm
        _header(ws, 9, ["Earned value", "Value", "RAG"])
        rows = [("BAC", "BAC", MONEY), ("Planned value (PV)", "PV", MONEY),
                ("Earned value (EV)", "EV", MONEY), ("Actual cost (AC)", "AC", MONEY),
                ("Schedule variance (SV)", "SV", MONEY), ("Cost variance (CV)", "CV", MONEY),
                ("SPI", "SPI", "0.00"), ("CPI", "CPI", "0.00"), ("EAC", "EAC", MONEY),
                ("VAC", "VAC", MONEY), ("TCPI", "TCPI", "0.00")]
        for k, (label, key, fmt) in enumerate(rows):
            r = 10 + k
            ws.cell(r, 1, label)
            c = ws.cell(r, 2, f"=EVM!{self.kpi_cell[key]}")
            c.number_format, c.border = fmt, BOX
            self.expect(ws, f"B{r}", self.kpi_expected[key])
            if key in ("SPI", "CPI"):
                ws.cell(r, 3, f'=IF(B{r}="","N/A",IF(B{r}>=0.95,"Green",IF(B{r}>=0.85,"Amber","Red")))')
                self.expect(ws, f"C{r}", rag(e.spi if key == "SPI" else e.cpi))
        rag_range = "C16:C17"
        for text, color in (("Green", GREEN), ("Amber", AMBER), ("Red", RED)):
            ws.conditional_formatting.add(rag_range, CellIsRule(
                operator="equal", formula=[f'"{text}"'], fill=PatternFill("solid", fgColor=color)))

        _header(ws, 3, ["Schedule risk (Monte Carlo, static)", "Working days", "Date"], col=5)
        pct = a.mc.percentiles
        rows = [(f"P{q} finish", pct[q]) for q in (50, 80, 90)]
        for k, (label, v) in enumerate(rows):
            r = 4 + k
            ws.cell(r, 5, label)
            ws.cell(r, 6, round(v, 2)).number_format = "0.0"
            ws.cell(r, 7, a.date_at(v)).number_format = DATE
        ws.cell(7, 5, "P(finish <= baseline)")
        ws.cell(7, 6, round(a.mc.p_on_time, 4)).number_format = "0.0%"

        _header(ws, 9, ["Counts", "Value"], col=5)
        counts = [
            ("Tasks", f"=COUNT(Schedule!$I${S_FIRST}:$I${self.s_last})", self.n),
            ("Critical tasks", f'=COUNTIF(Schedule!$P${S_FIRST}:$P${self.s_last},"Yes")',
             len(a.cpm.critical_ids)),
        ]
        rlast = 1 + max(len(a.program.risks), 1)
        for level in ("High", "Medium", "Low"):
            counts.append((f"{level} risks", f'=COUNTIF(Risks!$G$2:$G${rlast},"{level}")',
                           sum(1 for r in a.program.risks if r.rating == level)))
        for k, (label, f, exp) in enumerate(counts):
            r = 10 + k
            ws.cell(r, 5, label)
            ws.cell(r, 6, f).border = BOX
            self.expect(ws, f"F{r}", exp)
        _widths(ws, {"E": 34, "F": 14, "G": 12})
        ws.sheet_view.showGridLines = False

    # --------------------------------------------------------------- schedule
    def _schedule(self):
        a, ws = self.a, self.schedule
        ws["A1"] = "Schedule - critical path method (live formulas: edit Most Likely to re-plan)"
        ws["A1"].font = TITLE_FONT
        _header(ws, 2, ["ID", "Task", "Workstream", "Owner", "Predecessors", "Optimistic (d)",
                        "Most Likely (d)", "Pessimistic (d)", "Duration", "ES", "EF", "LS",
                        "LF", "Total float", "Free float", "Critical", "Start", "Finish",
                        "MC criticality"])
        net, cpm = a.cpm.network, a.cpm
        row_of = {i: S_FIRST + i for i in range(self.n)}
        for i, t in enumerate(a.program.tasks):
            r = row_of[i]
            ref = lambda col: f"{SC[col]}{r}"  # noqa: E731
            vals = {"id": t.id, "name": t.name, "ws": t.workstream, "owner": t.owner,
                    "preds": t.preds_text(), "opt": t.optimistic, "likely": t.likely,
                    "pess": t.pessimistic, "mc": round(float(a.mc.criticality[i]), 4)}
            for col, v in vals.items():
                ws[ref(col)] = v
            terms = [f"{SC['ef']}{row_of[p]}{_lag(lag, '+')}" for p, lag in net.preds[i]]
            late = [f"{SC['ls']}{row_of[s]}{_lag(lag, '-')}" for s, lag in net.succs[i]]
            early = [f"{SC['es']}{row_of[s]}{_lag(lag, '-')}" for s, lag in net.succs[i]]
            formulas = {
                "dur": (f"={ref('likely')}", cpm.duration[i]),
                "es": (f"=MAX({','.join(['0'] + terms)})", cpm.es[i]),
                "ef": (f"={ref('es')}+{ref('dur')}", cpm.ef[i]),
                "lf": (f"=MIN({','.join(['ProjFinish'] + late)})", cpm.lf[i]),
                "ls": (f"={ref('lf')}-{ref('dur')}", cpm.ls[i]),
                "tf": (f"={ref('ls')}-{ref('es')}", cpm.total_float[i]),
                "ff": (f"=MIN({','.join(['ProjFinish'] + early)})-{ref('ef')}", cpm.free_float[i]),
                "crit": (f'=IF({ref("tf")}<=0.000001,"Yes","No")',
                         "Yes" if cpm.is_critical(i) else "No"),
                "start": (f"=INDEX(CalDates,INT({ref('es')})+1)", a.task_start(i)),
                "finish": (f"=INDEX(CalDates,INT(MAX({ref('ef')}-1,{ref('es')}))+1)",
                           a.task_finish(i)),
            }
            for col, (f, exp) in formulas.items():
                ws[ref(col)] = f
                self.expect(ws, ref(col), exp)
            for col in ("start", "finish"):
                ws[ref(col)].number_format = DATE
            for col in ("tf", "ff"):
                ws[ref(col)].number_format = "0.0"
            ws[ref("mc")].number_format = "0%"
        last = self.s_last
        ws.freeze_panes = "C3"
        ws.auto_filter.ref = f"A2:S{last}"
        _widths(ws, {"A": 7, "B": 36, "C": 14, "D": 12, "E": 16, **{c: 9 for c in "FGHIJKLMNOP"},
                     "Q": 11, "R": 11, "S": 10})
        ws.conditional_formatting.add(f"P{S_FIRST}:P{last}", CellIsRule(
            operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor="F8CBAD"),
            font=Font(bold=True, color=RED)))
        ws.conditional_formatting.add(f"S{S_FIRST}:S{last}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=1, color="ED7D31"))
        self._gantt(ws, last)

    def _gantt(self, ws, last):
        a = self.a
        first_week = a.program.start - timedelta(days=a.program.start.weekday())
        end = max(a.baseline_finish_date, a.program.status_date)
        weeks = min((end - first_week).days // 7 + 1, MAX_GANTT_WEEKS)
        c0 = get_column_letter(GANTT_COL)
        c1 = get_column_letter(GANTT_COL + weeks - 1)
        for w in range(weeks):
            col = GANTT_COL + w
            c = ws.cell(2, col, first_week + timedelta(weeks=w))
            c.number_format = "dd-mmm"
            c.fill, c.font = HDR_FILL, Font(color="FFFFFF", size=8)
            c.alignment = Alignment(text_rotation=90, horizontal="center")
            ws.column_dimensions[get_column_letter(col)].width = 3.2
        area = f"{c0}{S_FIRST}:{c1}{last}"
        bar = f"AND({c0}$2<=$R{S_FIRST},{c0}$2+6>=$Q{S_FIRST})"
        ws.conditional_formatting.add(area, FormulaRule(
            formula=[f'AND($P{S_FIRST}="Yes",{bar})'], fill=PatternFill("solid", fgColor=RED),
            stopIfTrue=True))
        ws.conditional_formatting.add(area, FormulaRule(
            formula=[bar], fill=PatternFill("solid", fgColor="4472C4"), stopIfTrue=True))
        ws.conditional_formatting.add(area, FormulaRule(
            formula=[f"AND({c0}$2<=StatusDate,{c0}$2+6>=StatusDate)"],
            fill=PatternFill("solid", fgColor="D9D9D9")))
        self.gantt_weeks = weeks

    # -------------------------------------------------------------------- evm
    def _evm(self):
        a, ws, e = self.a, self.evm, self.a.evm
        ws["A1"] = "Earned value (live formulas: edit % Complete, Actual Cost, or the status date)"
        ws["A1"].font = TITLE_FONT
        _header(ws, 2, ["ID", "Task", "Budget", "ES", "Duration", "Planned %", "% Complete",
                        "PV", "EV", "AC", "SV", "CV"])
        for i, t in enumerate(a.program.tasks):
            r = 3 + i
            sr = S_FIRST + i
            ws[f"A{r}"] = f"=Schedule!A{sr}"
            ws[f"B{r}"] = t.name
            ws[f"C{r}"] = t.budget
            ws[f"G{r}"] = t.pct_complete
            ws[f"J{r}"] = t.actual_cost
            cells = {
                "D": (f"=Schedule!J{sr}", a.cpm.es[i]),
                "E": (f"=Schedule!I{sr}", a.cpm.duration[i]),
                "F": (f"=IF(E{r}=0,IF(Elapsed>=D{r},1,0),MAX(0,MIN(1,(Elapsed-D{r})/E{r})))",
                      e.planned_pct[i]),
                "H": (f"=C{r}*F{r}", e.pv_by_task[i]),
                "I": (f"=C{r}*G{r}", e.ev_by_task[i]),
                "K": (f"=I{r}-H{r}", e.ev_by_task[i] - e.pv_by_task[i]),
                "L": (f"=I{r}-J{r}", e.ev_by_task[i] - t.actual_cost),
            }
            for col, (f, exp) in cells.items():
                ws[f"{col}{r}"] = f
                if col != "A":
                    self.expect(ws, f"{col}{r}", exp)
            for col in "CHIJKL":
                ws[f"{col}{r}"].number_format = MONEY
            for col in "FG":
                ws[f"{col}{r}"].number_format = "0%"
        last = 2 + self.n
        tot = last + 1
        ws[f"B{tot}"] = "Total"
        ws[f"B{tot}"].font = Font(bold=True)
        for col, exp in (("C", e.bac), ("H", e.pv), ("I", e.ev), ("J", e.ac),
                         ("K", e.sv), ("L", e.cv)):
            c = ws[f"{col}{tot}"]
            c.value = f"=SUM({col}3:{col}{last})"
            c.number_format, c.font = MONEY, Font(bold=True)
            self.expect(ws, f"{col}{tot}", exp)
        dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                            showErrorMessage=True, error="0% to 100%")
        dv.add(f"G3:G{last}")
        ws.add_data_validation(dv)
        ws.conditional_formatting.add(f"K3:L{last}", CellIsRule(
            operator="lessThan", formula=["0"], font=Font(color=RED)))

        # KPI block
        _header(ws, 2, ["Metric", "Value"], col=14)
        kpis = [
            ("BAC", f"=C{tot}", e.bac), ("PV", f"=H{tot}", e.pv), ("EV", f"=I{tot}", e.ev),
            ("AC", f"=J{tot}", e.ac), ("SV", "=O5-O4", e.sv), ("CV", "=O5-O6", e.cv),
            ("SPI", '=IF(O4=0,"",O5/O4)', e.spi), ("CPI", '=IF(O6=0,"",O5/O6)', e.cpi),
            ("EAC", '=IF(OR(O10="",O10=0),"",O3/O10)', e.eac),
            ("VAC", '=IF(O11="","",O3-O11)', e.vac),
            ("TCPI", '=IF(O3=O6,"",(O3-O5)/(O3-O6))', e.tcpi),
        ]
        self.kpi_cell, self.kpi_expected = {}, {}
        for k, (label, f, exp) in enumerate(kpis):
            r = 3 + k
            ws.cell(r, 14, label).font = Font(bold=True)
            c = ws.cell(r, 15, f)
            c.number_format = "0.00" if label in ("SPI", "CPI", "TCPI") else MONEY
            c.border = BOX
            self.kpi_cell[label] = f"$O${r}"
            self.kpi_expected[label] = "" if exp is None else exp
            self.expect(ws, f"O{r}", self.kpi_expected[label])

        # S-curve (cumulative PV is a static series from Python; EV/AC plotted at status date)
        r0 = 17
        _header(ws, r0, ["Date", "Planned value", "Earned value", "Actual cost"], col=14)
        status_k = min(range(len(e.curve)), key=lambda k: abs(e.curve[k][0] - e.elapsed))
        for k, (_, d, pv) in enumerate(e.curve):
            r = r0 + 1 + k
            ws.cell(r, 14, d).number_format = DATE
            ws.cell(r, 15, round(pv, 2)).number_format = MONEY
            if k == status_k:
                ws.cell(r, 16, f"=I{tot}").number_format = MONEY
                ws.cell(r, 17, f"=J{tot}").number_format = MONEY
        r1 = r0 + len(e.curve)
        ch = LineChart()
        ch.title, ch.height, ch.width = "Cumulative planned value vs. EV / AC", 9, 18
        ch.y_axis.title, ch.y_axis.numFmt = "USD", '"$"#,##0'
        ch.add_data(Reference(ws, min_col=15, max_col=17, min_row=r0, max_row=r1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=14, min_row=r0 + 1, max_row=r1))
        for s in ch.series[1:]:
            s.marker.symbol, s.marker.size = "circle", 9
            s.graphicalProperties.line.noFill = True
        ws.add_chart(ch, "S2")
        _widths(ws, {"A": 7, "B": 36, **{c: 12 for c in "CDEFGHIJKL"}, "N": 12, "O": 14,
                     "P": 14, "Q": 14})
        ws.freeze_panes = "C3"

    # ------------------------------------------------------------------ risks
    def _risks(self):
        a, ws = self.a, self.risks
        _header(ws, 1, ["ID", "Risk", "Owner", "Probability (1-5)", "Impact (1-5)", "Score",
                        "Rating", "Mitigation"])
        risks = a.program.risks
        last = 1 + max(len(risks), 1)
        for k, rk in enumerate(risks):
            r = 2 + k
            for col, v in zip("ABCDEH", (rk.id, rk.title, rk.owner, rk.probability, rk.impact,
                                         rk.mitigation)):
                ws[f"{col}{r}"] = v
            ws[f"F{r}"] = f"=D{r}*E{r}"
            ws[f"G{r}"] = f'=IF(F{r}>=15,"High",IF(F{r}>=8,"Medium","Low"))'
            self.expect(ws, f"F{r}", rk.score)
            self.expect(ws, f"G{r}", rk.rating)
        dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5",
                            showErrorMessage=True, error="Score 1 (low) to 5 (high)")
        dv.add(f"D2:E{last + 50}")
        ws.add_data_validation(dv)
        ws.conditional_formatting.add(f"F2:F{last}", ColorScaleRule(
            start_type="num", start_value=1, start_color="63BE7B", mid_type="num", mid_value=9,
            mid_color="FFEB84", end_type="num", end_value=25, end_color="F8696B"))
        for level, color in LIGHT.items():
            ws.conditional_formatting.add(f"G2:G{last}", CellIsRule(
                operator="equal", formula=[f'"{level}"'], fill=PatternFill("solid", fgColor=color)))
        ws.auto_filter.ref = f"A1:H{last}"
        ws.freeze_panes = "A2"

        # 5x5 heat map of risk counts (COUNTIFS) - impact across, probability down
        ws["J1"] = "Risk heat map (count of risks)"
        ws["J1"].font = Font(bold=True, color=NAVY)
        ws["J2"] = "Probability \\ Impact"
        for i in range(1, 6):
            c = ws.cell(2, 10 + i, i)
            c.fill, c.font, c.alignment = HDR_FILL, HDR_FONT, Alignment(horizontal="center")
        for row, p in enumerate(range(5, 0, -1)):
            r = 3 + row
            c = ws.cell(r, 10, p)
            c.fill, c.font, c.alignment = HDR_FILL, HDR_FONT, Alignment(horizontal="center")
            for i in range(1, 6):
                cell = ws.cell(r, 10 + i, f"=COUNTIFS($D$2:$D${last},{p},$E$2:$E${last},{i})")
                cell.fill = PatternFill("solid", fgColor=LIGHT[risk_rating(p * i)])
                cell.alignment, cell.border = Alignment(horizontal="center"), BOX
                self.expect(ws, cell.coordinate,
                            sum(1 for x in risks if x.probability == p and x.impact == i))
        _widths(ws, {"A": 7, "B": 46, "C": 14, "D": 11, "E": 11, "F": 8, "G": 9, "H": 46,
                     "J": 20, **{get_column_letter(c): 6 for c in range(11, 16)}})

    # ------------------------------------------------------------- montecarlo
    def _montecarlo(self):
        a, ws, mc = self.a, self.mc, self.a.mc
        ws["A1"] = f"Monte Carlo schedule risk - {mc.iterations:,} iterations (static values)"
        ws["A1"].font = TITLE_FONT
        _header(ws, 3, ["Percentile", "Finish (working days)", "Finish date"])
        for k, (q, v) in enumerate(mc.percentiles.items()):
            ws.cell(4 + k, 1, f"P{q}")
            ws.cell(4 + k, 2, round(v, 2)).number_format = "0.0"
            ws.cell(4 + k, 3, a.date_at(v)).number_format = DATE
        r = 4 + len(mc.percentiles)
        ws.cell(r, 1, "Mean")
        ws.cell(r, 2, round(mc.mean, 2)).number_format = "0.0"
        ws.cell(r + 1, 1, "Deterministic plan")
        ws.cell(r + 1, 2, a.cpm.finish).number_format = "0.0"
        ws.cell(r + 2, 1, "P(finish <= plan)")
        ws.cell(r + 2, 2, round(mc.p_on_time, 4)).number_format = "0.0%"

        counts, edges = mc.histogram(20)
        h0 = 3
        _header(ws, h0, ["Bin from", "Bin to", "Iterations", "Cumulative %"], col=5)
        for k, c in enumerate(counts):
            rr = h0 + 1 + k
            ws.cell(rr, 5, round(float(edges[k]), 2))
            ws.cell(rr, 6, round(float(edges[k + 1]), 2))
            ws.cell(rr, 7, int(c))
            cum = ws.cell(rr, 8, f"=SUM($G${h0 + 1}:G{rr})/{mc.iterations}")
            cum.number_format = "0.0%"
            self.expect(ws, f"H{rr}", float(counts[:k + 1].sum()) / mc.iterations)
        hl = h0 + len(counts)
        ch = BarChart()
        ch.title, ch.height, ch.width, ch.gapWidth = "Simulated finish (working days)", 8, 16, 10
        ch.legend = None
        ch.add_data(Reference(ws, min_col=7, min_row=h0, max_row=hl), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=6, min_row=h0 + 1, max_row=hl))
        ws.add_chart(ch, "J3")

        t0 = hl + 3
        _header(ws, t0, ["ID", "Task", "Criticality index"])
        top = sorted(range(self.n), key=lambda i: -mc.criticality[i])[:15]
        for k, i in enumerate(top):
            ws.cell(t0 + 1 + k, 1, a.program.tasks[i].id)
            ws.cell(t0 + 1 + k, 2, a.program.tasks[i].name)
            ws.cell(t0 + 1 + k, 3, round(float(mc.criticality[i]), 4)).number_format = "0%"
        _widths(ws, {"A": 22, "B": 36, "C": 16, "E": 10, "F": 10, "G": 11, "H": 13})


def _lag(lag: int, sign: str) -> str:
    """Render a lag term for a formula: '+2' / '-2' adjusted for the pass direction."""
    if lag == 0:
        return ""
    v = lag if sign == "+" else -lag
    return f"+{v}" if v > 0 else f"{v}"


def build_workbook(analysis: Analysis, path: str | Path | None = None):
    """Build the status workbook. Returns (workbook, expected values per formula cell)."""
    book = _Book(analysis)
    wb = book.build()
    if path is not None:
        wb.save(path)
    return wb, book.expected
