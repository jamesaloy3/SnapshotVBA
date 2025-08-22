# SP.FINANCIALS — SinglePane Excel Function
_Last updated: 2025-08-22_

## What it does
Returns a value from the SinglePane financial database for a specific property and USALI account, for a selected time aggregation, year, and version. Typical sources include hotel P&Ls, budgets, monthly forecasts, and owner/operator pro‑formas.

## Syntax
```excel
=SP.FINANCIALS(property_code, usali, month_or_aggregation, year, version)
```

- **`property_code`** — 3‑letter SinglePane property code (e.g., `"XYZ"`). These appear in the **My Properties** sheet added by the add‑in.
- **`usali`** — USALI GL account to pull (e.g., `"Total Revenue - 100"`). All codes are listed in the **Usali Reference** sheet added by the add‑in.
- **`month_or_aggregation`** — A 3‑letter month (e.g., `"Jul"`) _or_ one of the aggregations below.
- **`year`** — 4‑digit year (e.g., `2025`).
- **`version`** — Data version to return (see list).

## Time aggregations for the `month_or_aggregation` argument
- `Total Year`
- `Q1`, `Q2`, `Q3`, `Q4`
- `MMMYTD` — e.g., `JulYTD` (no space)
- `MMMBOY` — Balance of Year, excludes the named month (e.g., `JulBOY` returns Aug–Dec)
- `MMMTTM` — Trailing‑12 (e.g., `JulTTM` is current Jul through prior Aug)

## Versions
- `Actual` (from hotel P&Ls)
- `Budget` (final/current)
- `Budget1` … `Budget12` (budget drafts)
- `Forecast1` … `Forecast12` — monthly forecasts as of that month’s close (e.g., `Forecast6` is the forecast at end of June, with June Actuals)
- `Proforma`
- `LY_Actual` — same month last year
- `Var_LY_Actual` — variance vs. last year
- `Var_Budget` — variance vs. budget

## Examples
Literal values:
```excel
=SP.FINANCIALS("XYZ","Total Revenue - 100","JulYTD",2023,"Actual")
```

Cell references (recommended for model transparency):
```excel
=SP.FINANCIALS(B2, B3, B4, B5, B6)
```
Where:
- **B2** = `XYZ` (property code)
- **B3** = `Total Revenue - 100` (USALI)
- **B4** = `JulYTD` (aggregation)
- **B5** = `2023` (year)
- **B6** = `Actual` (version)

## Notes & Best Practices
- Keep a small **code map** table to validate property codes and USALI labels from the **My Properties** and **Usali Reference** sheets; use data validation to avoid typos.
- Prefer named ranges for frequently used inputs (`Property`, `USALI`, `MonthAgg`, `Year`, `Version`).
- Use `TEXT(DATE(Year, MonthNum, 1), "MMM")` to generate 3‑letter month abbreviations for dynamic reports.
- For YTD/TTM/BOY aggregations, ensure your model clearly displays the implied date span in the report header (e.g., “As of: July 2025 (YTD)”).

## Troubleshooting
- **#VALUE! / blank result** — Confirm the property is authorized to your account and the USALI label matches the **Usali Reference** sheet exactly.
- **Unexpected totals** — Check that the `version` aligns with the period (e.g., `Forecast6` with June close).
- **Slow recalc** — Avoid repeating expensive helper logic; centralize inputs and reference them.

## See also
- **SP.FINANCIALS_AGG** for aggregating multiple properties.
- **SP.FILTER** for generating dynamic property lists (to feed into AGG).

