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

- **`Promo Update Output`** — one row per **promo depth**, per product × channel. All weeks at
  the same price share one row whether or not they are consecutive: a gap in the promo does
  **not** start a new row. A row is split only where the account funding changes partway
  through, so a Full Funding event yields one row per funding segment. Values only, no formulas.
- **`Run Log`** — `Issue | Product ID | Channel | Week | Detail` plus a run summary.

Because a depth group can straddle a gap, funding windows are intersected with the weeks
actually promoted at that depth — an event landing in a gap cannot split the row.

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

## PromoUpdateTemplateTests.bas

A standalone test module. Import it the same way and run
`TestAccountContributionConsistency`.

It checks that the contribution percentage carried forward from the CY Template is the same
across models within a channel — a channel negotiates one funding split and applies it to
every model it carries, so a model that disagrees points at a CY data error that would
otherwise be carried straight into next year's template.

**Amazon is exempt** (its funding structure changed). The exempt list is Config
`Pct Test Exempt Accounts`, falling back to `Full Funding Accounts`, then to `AMAZON`.
Tolerance is Config `Pct Test Tolerance` as a fraction (default `0.005` = 0.5 points).

Results go to a rebuilt `Pct Consistency Test` sheet:

- **Channel summary** — models with data, min / max %, spread, and `PASS` / `FAIL` /
  `EXEMPT` / `NO DATA`
- **Product detail** — one row per product × channel with its %, CY rows used, deviation
  from the channel median, and `OK` / `OUTLIER` / `EXEMPT` / `NO DATA`

The module re-reads the CY Template itself rather than calling into
`PromoUpdateTemplate`, so a bug in the generator cannot mask a bug in the data. It reads
only, and never writes to an input sheet.
