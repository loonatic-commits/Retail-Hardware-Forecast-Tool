# VBA Demand Forecasting Suite

Modular VBA scripts for the consensus forecast workbook. Each pass is its
own module so failures stay isolated.

## Files

| File | Module name | Role |
|---|---|---|
| `modConfig.bas` | `modConfig` | Constants, thresholds, colors |
| `modUtilities.bas` | `modUtilities` | Block detection, month parsing, accuracy math, grading cache |
| `modMain.bas` | `modMain` | Orchestrator + preflight check |
| `modFieldGrading.bas` | `modFieldGrading` | Pass 1 - score FF accuracy, grade the FF label cell |
| `modFFvsBP.bas` | `modFFvsBP` | Pass 2 - seed Opportunity from FF or BP (values only) |
| `modSOFOverride.bas` | `modSOFOverride` | Pass 3 - override SOF months (values only) |
| `modSoFarRunRate.bas` | `modSoFarRunRate` | Pass 4 - current month = SO FAR + Back Order (values only) |
| `modForecastSelect.bas` | `modForecastSelect` | Pass 5 - constrained path + Consensus N-1 hold |
| `modInventoryPass.bas` | `modInventoryPass` | Pass 6 - red font where projected inventory < demand |
| `modInfoBlock.bas` | `modInfoBlock` | Pass 7 - notes column + dissonance fill |
| `modDeficitReport.bas` | `modDeficitReport` | Standalone - month-dynamic deficit / excess report |
| `modCountrySplit.bas` | `modCountrySplit` | Standalone - splits Opportunity into US / Canada |
| `modTests.bas` | `modTests` | Synthetic workbook + assertions |

## Import

1. Open the `.xlsm` workbook.
2. `Alt+F11` to open the VBA editor.
3. Remove any previously imported `mod*` modules first, then
   `File > Import File...` for each `.bas` under `vba/`.
4. `Debug > Compile VBAProject` to confirm everything resolves.
5. Run `Run_All_Passes` from `modMain`.

Standalone runs (not part of `Run_All_Passes`):

- `Build_Deficit_Report` - deficit / excess inventory report
- `Run_Country_Split` - writes `USA = NA Opportunity - Canada`

## Forecast selection logic

`Run_Forecast_Selection` decides each month from the current month (N) forward.

**Constraint is inferred**, per model *and* per month - there is no flag column:

```
availability(month) < max(Consensus N-1, Field N-1)
```

`availability` is the running inventory pool: it starts at the current month's
Available Inventory (on hand today), adds each later month's incoming, and is
reduced by whatever consensus consumes as the walk moves forward.

To compare against the current Field submission instead of Field N-1, flip
`CONSTRAINT_USE_FIELD_N1` to `False` in `modConfig`.

### Constrained path

| Month | Consensus value |
|---|---|
| N to N+3 | Availability |
| N+4 | Availability + backorder, or Field when Field exceeds that |
| N+5 onward | Falls through to the normal path |

### Normal path

| Month | Consensus value |
|---|---|
| N to N+2 | Unchanged - Passes 2-4 stand |
| N+3 onward | Consensus N-1, held even when Field varies past 15% |

Variance never rewrites a number. It only raises a flag and a fill.

## Color coding

The legend module is gone. Three things still write formatting:

| What | Where | Written by |
|---|---|---|
| Accuracy grade (green / yellow / red) | FF Qty Final label cell, column B | Pass 1 |
| Inventory constrained (red font) | Opportunity month cells | Pass 6 |
| Dissonance flag (gold fill) | Opportunity month cells | Pass 7 |

## Notes column

Pass 7 writes a notes column in the first free column after the month data
(`last month column + INFO_BLOCK_COL_OFFSET`, default 2, leaving one spacer).
One section per model, starting on that model's first row. Read it top to
bottom:

```
ES-400 II - notes
26-Jun to 26-Sep: Constrained - availability
26-Oct: Constrained + backorder
26-Nov to 27-Mar: Consensus N-1 hold
Field movement vs Field N-1: 26-Jul +22.4%, 26-Aug -18.0%
SOF reference (N..N+2): 26-Jun 4,512 | 26-Jul 3,254 | 26-Aug 1,434
FLAG 26-Jul: Field +22.4% vs Consensus N-1; SOF -19.0% vs Consensus N-1
```

The notes column and all Opportunity fills are cleared at the start of every
run, so nothing stale survives.

### Dissonance tests

Three tests, all at 15% (`VARIANCE_DISSONANCE`), all measured **as a
percentage of Consensus N-1** - it is the denominator every time:

- Field N-1 vs Field
- Field vs Consensus N-1
- Consensus N-1 vs SOF

A month with no Consensus N-1 value cannot be tested and is skipped.

## Running it every month

Nothing needs editing month to month. Every pass keys off today's calendar
month, so the forecast starts at the current month automatically and past
months are never overwritten.
