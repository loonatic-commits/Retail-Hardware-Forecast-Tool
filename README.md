# Retail-Hardware-Forecast-Tool

## PromoUpdateTemplate.bas

An Excel VBA standard module that builds the **SAP IBP Promo Update Template (US)** from
four input sheets in the active workbook.

### Install

1. Open the workbook that holds the input sheets.
2. `Alt+F11` → **File → Import File…** → `PromoUpdateTemplate.bas`.
3. Run `GeneratePromoUpdateTemplate` (`Alt+F8`).

No references are required — `Scripting.Dictionary` is created late bound.

### Inputs (read only, never modified)

| Sheet | Purpose |
|---|---|
| `Promo Grid` | `Model Desc`, `MSRP`, `Product ID`, then one column per week; a cell holds that week's promo net price, blank = no promo |
| `CY Template` | Last year's submitted template — which channels carry which product, `Policy Price`, and the account contribution % |
| `NY Template` | Header row only; defines the output column order and the `WB` week calendar |
| `Full Funding Events` | `Event Name`, `Start Week`, `End Week`, `Funding Mode` (`First N Days` / `Full Period`), `Model` |
| `Config` | `Setting` / `Value`: Financial Year, Key Figure, Country, Full Funding Days, Full Funding Accounts, Pricing Policy ID, Blank Placeholder, Date Format |

Every column is located by header text at run time through an alias table, so the CY and NY
spellings need not match and inserted columns are harmless. `Pricing Policy FY` is
deliberately **not** aliased to `Pricing Policy ID`.

### Outputs (deleted and rebuilt on every run)

- **`Promo Update Output`** — one row per contiguous block of weeks at the same promo depth,
  per product × channel, split further into one row per funding segment where a Full Funding
  event applies. Values only, no formulas.
- **`Run Log`** — `Issue | Product ID | Channel | Week | Detail` plus a run summary.

### Contribution math

```
CY, per row:   row pct     = Accn Cont. / (Policy Price - Net)      averaged per product x channel
NY, per row:   Total Disc  = Policy Price - Net
               Accn Cont.  = pct x Total Disc
               Epson Cont. = Total Disc - Accn Cont.
```

Rows with zero discount and last year's fully funded rows (`Accn Cont.` = 0 with a populated
`Special Promo Start Date`) are excluded from the CY average. Missing product × channel data
falls back to the channel average across all products, then to 0 — both are logged.

Assumptions are listed in the header comment of `PromoUpdateTemplate.bas`.
