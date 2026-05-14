# VBA Demand Forecasting Suite

Modular VBA scripts for the consensus forecast workbook. Each pass is its
own module so failures stay isolated.

## Files

| File | Module name | Role |
|---|---|---|
| `modConfig.bas` | `modConfig` | Constants, color initialization |
| `modUtilities.bas` | `modUtilities` | Block detection, month parsing, accuracy math, grading cache |
| `modMain.bas` | `modMain` | Orchestrator |
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
4. Ensure the Microsoft Scripting Runtime reference is enabled if you
   prefer early binding (current code uses `CreateObject("Scripting.Dictionary")`
   so no reference change is required).
5. Run `Run_All_Passes` from `modMain`.

## Build status

- [x] `modConfig`
- [x] `modUtilities`
- [ ] `modMain` (orchestrator)
- [ ] `modFieldGrading` (Pass 1)
- [ ] `modFFvsBP` (Pass 2)
- [ ] `modSOFOverride` (Pass 3)
- [ ] `modSoFarRunRate` (Pass 4)
- [ ] `modInventoryPass` (Pass 5)
- [ ] `modLegend` (Pass 6)
- [ ] `modTests`
