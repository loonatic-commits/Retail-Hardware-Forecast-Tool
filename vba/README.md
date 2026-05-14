# VBA Demand Forecasting Suite

Modular VBA scripts for the consensus forecast workbook. Each pass is its
own module so failures stay isolated.

## Files

| File | Module name | Role |
|---|---|---|
| `modConfig.bas` | `modConfig` | Constants, color initialization |
| `modUtilities.bas` | `modUtilities` | Block detection, month parsing, accuracy math, grading cache |
| `modMain.bas` | `modMain` | Orchestrator + preflight check |
| `modFieldGrading.bas` | `modFieldGrading` | Pass 1 - score FF accuracy |
| `modFFvsBP.bas` | `modFFvsBP` | Pass 2 - seed Opportunity from FF or BP |
| `modSOFOverride.bas` | `modSOFOverride` | Pass 3 - override next 3 months with SOF |
| `modSoFarRunRate.bas` | `modSoFarRunRate` | Pass 4 - override current month with SO FAR |
| `modInventoryPass.bas` | `modInventoryPass` | Pass 5 - red font on inventory-constrained months |
| `modLegend.bas` | `modLegend` | Pass 6 - write color legend to corner of Consensus |
| `modTests.bas` | `modTests` | Synthetic workbook + assertions |

## Import

1. Open the `.xlsm` workbook.
2. `Alt+F11` to open the VBA editor.
3. `File > Import File...` and import each `.bas` file under `vba/`.
4. Run `Run_All_Passes` from `modMain`.
5. To test on synthetic data, run `Run_All_Tests` from `modTests` (this overwrites the Consensus and SOF sheets - use a scratch workbook).

## Preflight check

`Run_All_Passes` calls `PreflightCheck` first and aborts (without touching
any cells) if any of these are missing or malformed:

- `Consensus` sheet
- `SOF` sheet
- At least one parseable month header on `Consensus` row 1
- At least one model block in Consensus column A
- Per-block required key figures: Business Plan Quantity, Sell-In Quantity,
  SO FAR, Field Forecast Qty Final, Opportunity Qty Final,
  Available Inventory (Manual)
- SOF month headers (columns B-D row 1) and at least one SOF model row

All issues are reported in a single dialog so you can fix them in one pass.

## Build status

- [x] `modConfig`
- [x] `modUtilities`
- [x] `modMain` (orchestrator + preflight check)
- [x] `modFieldGrading` (Pass 1)
- [x] `modFFvsBP` (Pass 2)
- [x] `modSOFOverride` (Pass 3)
- [x] `modSoFarRunRate` (Pass 4)
- [x] `modInventoryPass` (Pass 5)
- [x] `modLegend` (Pass 6)
- [x] `modTests`
