# VBA Demand Forecasting Suite — v2.3.0

**Released:** 2026-08-13
**Branch:** `claude/vba-demand-forecasting-QIF7f`

Drop-in module set for the consensus forecast workbook. Import every `.bas`
in this folder, then run `Run_All_Passes`.

---

## Import

1. Open your `.xlsm` workbook and press `Alt+F11`.
2. **Remove every existing `mod*` module first** (right-click → Remove → No to
   export). A stale module from an older version will break the compile.
3. `File → Import File…` and import all 14 `.bas` files in this folder.
4. `Debug → Compile VBAProject` — this must come back clean.
5. Save as macro-enabled (`.xlsm`).

## Run

| Macro | What it does |
|---|---|
| `Run_All_Passes` | The monthly forecast. Preflight, then all seven passes. |
| `Run_All_Tests` | Self-check on synthetic data. **Use a scratch workbook — it clears the Consensus and SOF sheets.** |
| `Build_Deficit_Report` | Deficit / excess inventory report for last completed month. |
| `Run_Country_Split` | Writes `USA = NA Opportunity − Canada` on the Country Split sheet. |
| `Build_Variance_Summary` | Marketing vs Field variance sheet. Run after `Run_All_Passes`. |

Nothing needs editing month to month. Every pass keys off today's calendar
month, so the forecast starts at the current month automatically and past
months are never overwritten.

---

## Contents

| File | Role |
|---|---|
| `modConfig.bas` | Constants, thresholds, colors |
| `modUtilities.bas` | Block detection, month parsing, accuracy math, grading cache |
| `modMain.bas` | Orchestrator + preflight check |
| `modFieldGrading.bas` | Pass 1 — grade field forecast accuracy |
| `modFFvsBP.bas` | Pass 2 — seed Opportunity from FF or BP |
| `modSOFOverride.bas` | Pass 3 — override SOF months |
| `modSoFarRunRate.bas` | Pass 4 — current month = SO FAR + Back Order |
| `modForecastSelect.bas` | Pass 5 — constrained path + Consensus N-1 hold |
| `modInventoryPass.bas` | Pass 6 — red font where projected inventory < demand |
| `modInfoBlock.bas` | Pass 7 — notes column + dissonance fill |
| `modDeficitReport.bas` | Standalone — deficit / excess report |
| `modCountrySplit.bas` | Standalone — US / Canada split |
| `modVarianceSummary.bas` | Standalone — Marketing vs Field variance summary |
| `modTests.bas` | Synthetic workbook + assertions |

---

## What changed in v2.3.0

New module `modVarianceSummary.bas`. Nothing else changed.

### Marketing vs Field variance summary

`Build_Variance_Summary` creates a **`Mktg vs FF Variance`** sheet comparing
Marketing Demand Forecast against Field Forecast across the forward window —
current month through the last month column, which is Aug–Mar on a sheet
ending at 27-Mar. Three tables:

1. **Summary by model** — window totals for both rows, variance in units and
   percent, direction, FF accuracy grade, and a plain-language "why the
   Opportunity landed where it did" column
2. **Monthly variance (units)** — model × month
3. **Monthly variance (%)** — model × month

Percent variances at or beyond 15% are filled gold. A TOTAL row closes
table 1, written as live formulas so it recalculates if you edit the sheet.

**Run it after `Run_All_Passes`.** The "why" column replays the same decision
walk the selection pass uses, so it only reflects reality once that pass has
run.

| Constant | Default | Meaning |
|---|---|---|
| `KF_MKTG_DEMAND` | `"Marketing Forecast Qty"` | Which Marketing row to compare |
| `VAR_PCT_OF_FF` | `True` | Percent is of Field Forecast; `False` divides by Marketing |
| `VAR_MATERIAL` | `0.15` | Threshold for calling a variance material |

The window is derived from today's date, so the sheet re-scopes itself every
month with no edits.

---

## What changed in v2.2.0

Only `modCountrySplit.bas` changed. Everything else is identical to v2.0.0.

**CANADA + USA now always reconciles to the NA total.** When the current-month
SO FAR uplift fires, the Canada row is restated to the uprated figure rather
than left at its original forecast, so the two halves add back up.

```
' current month, uplift fires
Canada row  := Canada SO FAR × 1.20      ' restated in place
USA row     := NA Opportunity − that same figure
```

`WRITE_UPLIFT_TO_CANADA` now defaults to `True`. Set it to `False` to go back
to leaving the Canada row untouched.

This is the only case in which the pass writes to a Canada cell. Outside an
uplifted current month, Canada is still read-only.

---

## What changed in v2.1.0

Only `modCountrySplit.bas` changed. Everything else is identical to v2.0.0.

### Canada SO FAR uplift — current month only

`Run_Country_Split` still writes `USA = NA Opportunity − Canada` for the
current month forward. One rule is new, and it applies to the **current month
only**: if Canada's SO FAR is running at or above the Canada forecast, Canada
is already tracking to beat its number, so the subtraction uses an uprated
figure instead.

```
if Canada SO FAR >= Canada forecast × 0.9      ' "at or above, or close"
    Canada effective = Canada SO FAR × 1.20
else
    Canada effective = Canada forecast

USA = NA Opportunity − Canada effective
```

Every other month is untouched by this rule. A model with no Canada SO FAR
row behaves exactly as it did in v2.0.0.

| Constant | Default | Meaning |
|---|---|---|
| `SOFAR_CLOSE_RATIO` | `0.9` | SO FAR at ≥ 90% of forecast counts as "close" |
| `SOFAR_UPLIFT` | `0.2` | Add 20% over SO FAR |
| `WRITE_UPLIFT_TO_CANADA` | `True` (as of v2.2.0) | Restate the Canada row so the halves reconcile |

### Two Canada rows per model

With SO FAR imported, each model now has more than one `CANADA TOTAL` row.
Row lookup keys on the **Key Figure** column (C) to tell them apart: a row
whose key figure reads `SO FAR` is the run-rate row, any other row for that
segment is the forecast row. The old lookup matched on model + segment alone
and took the first hit, which could have grabbed the wrong row.

The run summary reports how many models had an uplift applied, and each one
is logged to the Immediate window with the before and after figures.

---

## What changed in v2.0.0

### Forecast selection replaced (Pass 5, new)

Constraint is **inferred** per model *and* per month — there is no flag column:

```
availability(month) < max(Consensus N-1, Field N-1)
```

`availability` is the running inventory pool: it starts at the current month's
Available Inventory (on hand today), adds each later month's incoming, and is
reduced by whatever consensus consumes as the walk moves forward.

**Constrained path**

| Month | Consensus value |
|---|---|
| N to N+3 | Availability |
| N+4 | Availability + backorder, or Field when Field exceeds that |
| N+5 onward | Falls through to the normal path |

**Normal path**

| Month | Consensus value |
|---|---|
| N to N+2 | Unchanged — Passes 2–4 stand |
| N+3 onward | Consensus N-1, held even when Field varies past 15% |

Variance never rewrites a number. It only raises a flag and a fill.

### Color coding cut back

`modLegend` is gone, along with all Opportunity-row fills from Passes 2–4.
Three things still write formatting:

| What | Where | Written by |
|---|---|---|
| Accuracy grade (green / yellow / red) | FF Qty Final label cell, column B | Pass 1 |
| Inventory constrained (red font) | Opportunity month cells | Pass 6 |
| Dissonance flag (gold fill) | Opportunity month cells | Pass 7 |

### Notes column added (Pass 7, new)

A plain-language notes column in the first free column after the month data
(`last month column + INFO_BLOCK_COL_OFFSET`, default 2). One section per
model, starting on that model's first row:

```
ES-400 II - notes
26-Aug to 26-Nov: Constrained - availability
26-Dec: Constrained + backorder
27-Jan to 27-Jul: Consensus N-1 hold
Field movement vs Field N-1: 26-Sep +22.4%, 26-Oct -18.0%
SOF reference (N..N+2): 26-Aug 4,512 | 26-Sep 3,254 | 26-Oct 1,434
FLAG 26-Sep: Field +22.4% vs Consensus N-1; SOF -19.0% vs Consensus N-1
```

Decision labels collapse into month ranges. The notes column and all
Opportunity fills are cleared at the start of every run, so nothing stale
survives.

**Dissonance tests** — three tests, all at 15% (`VARIANCE_DISSONANCE`), all
measured **as a percentage of Consensus N-1**; it is the denominator every
time:

- Field N-1 vs Field
- Field vs Consensus N-1
- Consensus N-1 vs SOF

A month with no Consensus N-1 value cannot be tested and is skipped.

---

## Settings worth knowing

All in `modConfig.bas`:

| Constant | Default | Meaning |
|---|---|---|
| `CONSTRAINT_USE_FIELD_N1` | `True` | Demand signal is `max(Consensus N-1, Field N-1)`. Set `False` to use the current Field submission instead. |
| `VARIANCE_DISSONANCE` | `0.15` | Dissonance flag threshold |
| `SO_FAR_VARIANCE` | `0.5` | Current-month run-rate flag threshold |
| `INFO_BLOCK_COL_OFFSET` | `2` | Notes column position, right of the last month column |
| `ACCURACY_GREEN` / `ACCURACY_YELLOW` | `0.8` / `0.6` | Field grading bands |

---

## Required sheet structure

**`Consensus`** — model name in column A (repeated on every row of a block),
key figure in column B, month headers from column C. The preflight check
requires these key figure rows per model:

- `Business Plan Quantity`
- `Sell-In Quantity (Gross)`
- `SO FAR`
- `Field Forecast Qty Final`
- `Opportunity Qty Final`
- `Available Inventory (Manual)`

Also read when present (a model missing them is skipped, not blocked):
`Back Order Qty`, `Field Forecast Qty N-1`, `Consensus Fcst Qty Final N-1`.

**`SOF`** — model in column A, three month columns from column B.

**`Country Split`** (only for `Run_Country_Split`) — model in column A,
segment in column B (`CANADA TOTAL` / `USA TOTAL`), key figure in column C,
months from column D.

If the preflight fails it lists every problem in one dialog and changes
nothing on the sheet.
