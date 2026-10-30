# statusdeck benchmark results

Command: `python bench/run_bench.py`  
Machine: arm64 / Darwin 25.5.0, Python 3.12.14, numpy 2.5.4

## CPM forward/backward pass

| Tasks | Edges | CPM time (median of 5) | Tasks / s |
|---:|---:|---:|---:|
| 100 | 199 | 0.10 ms | 1,016,955 |
| 1,000 | 1,992 | 1.19 ms | 838,370 |
| 10,000 | 20,010 | 26.35 ms | 379,441 |

## Monte Carlo throughput (vectorised over iterations)

| Tasks | Iterations | Time (median of 3) | Iterations / s |
|---:|---:|---:|---:|
| 60 | 10,000 | 0.009 s | 1,163,450 |
| 500 | 10,000 | 0.073 s | 136,482 |
| 2,000 | 10,000 | 0.340 s | 29,389 |

## Monte Carlo accuracy against closed-form answers

| Check | Expected | Simulated | Error |
|---|---:|---:|---:|
| 30-task serial chain, mean finish | 220.000 | 219.971 | 0.013% |
| 30-task serial chain, std dev | 9.309 | 9.307 | 0.032% |
| 2 parallel paths, E[max] (merge bias) | 13.444 | 13.436 | 0.065% |
| 200-task network, zero-width estimates: iterations equal to CPM | 1,000 | 1,000 | 0 |

## Excel formula fidelity (workbook recalculated by pycel vs. Python engine)

| Tasks | Formula cells | Matching after pycel recalculation | Write time | Recalc time |
|---:|---:|---:|---:|---:|
| 60 | 1,127 | 1,127 (100.0%) | 0.04 s | 0.35 s |
| 300 | 5,207 | 5,207 (100.0%) | 0.11 s | 2.12 s |

## End-to-end: 60-task tracker to workbook + deck

| Stage | Time |
|---|---:|
| read tracker | 0.006 s |
| analyze (CPM + EVM + 10,000 MC iterations) | 0.009 s |
| write Excel workbook | 0.041 s |
| write PowerPoint deck | 0.066 s |
| **total** | **0.121 s** |

Output sizes: workbook 41 KiB, deck 75 KiB.
