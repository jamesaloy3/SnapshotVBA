# SP.FINANCIALS_AGG — SinglePane Excel Function
_Last updated: 2025-08-22_

## What it does
Returns an **aggregate** value across multiple properties from the SinglePane financial database for a given USALI account, time aggregation, year, and version. Works identically to `SP.FINANCIALS` but accepts a _set of property codes_.

## Syntax
```excel
=SP.FINANCIALS_AGG(property_codes, usali, month_or_aggregation, year, version)
```

### `property_codes` can be provided three ways
1. **Range reference** — e.g., `A2:A20` where each cell contains a property code.
2. **Quoted list** — e.g., `"ABC, DEF, GHI"`
3. **Nested `SP.FILTER()`** — returns a list of property codes matching filters (brand, manager, fund, etc.).

Other arguments match `SP.FINANCIALS` (USALI, aggregation, year, version).

## Time aggregations (same as `SP.FINANCIALS`)
- `Total Year`
- `Q1`, `Q2`, `Q3`, `Q4`
- `MMMYTD` (e.g., `JulYTD`)
- `MMMBOY` (e.g., `JulBOY`, returns Aug–Dec)
- `MMMTTM` (e.g., `JulTTM`, trailing‑12)

## Versions (same as `SP.FINANCIALS`)
- `Actual`, `Budget`, `Budget1` … `Budget12`,
  `Forecast1` … `Forecast12`, `Proforma`,
  `LY_Actual`, `Var_LY_Actual`, `Var_Budget`

## Examples
Quoted list:
```excel
=SP.FINANCIALS_AGG("XYZ,ABC,DEF","Total Revenue - 100","JulYTD",2023,"Actual")
```

With a range (best for dynamic models):
```excel
=SP.FINANCIALS_AGG($A$2:INDEX($A:$A, MATCH("zzzzz",$A:$A)), $B$1, $B$2, $B$3, $B$4)
```

With `SP.FILTER` (recommended pattern):
```excel
=LET(codes, SP.FILTER("Manager=Sage","Brand=Hilton"),
     SP.FINANCIALS_AGG(codes,"Total Revenue - 100","Q2",2025,"Budget"))
```

## Performance Tips
- **Calculate `SP.FILTER` once** on the sheet (e.g., in an empty helper column) and **pass the resulting range** into all `SP.FINANCIALS_AGG` formulas rather than nesting `SP.FILTER` inside each cell. This reduces redundant recalculation and speeds up large reports.
- Keep your code list contiguous (no blanks) to avoid partial ranges and mismatched totals.

## Troubleshooting
- **Unexpected blank** — Check that every code in the list is authorized to your account and spelled exactly as in **My Properties**.
- **Double‑count risk** — Ensure each property code appears only once in the range.
- **Performance** — Replace repeated quoted lists with a single maintained range or filtered spill.

## See also
- **SP.FINANCIALS** — single‑property retrieval.
- **SP.FILTER** — build filtered property lists for AGG inputs.
