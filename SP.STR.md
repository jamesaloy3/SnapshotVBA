# SP.STR — SinglePane Excel Function (STR data)
_Last updated: 2025-08-22_

## What it does
Returns STR metrics for a property for a given date, aggregation, subject/comp/market view, and market segment.

## Syntax
```excel
=SP.STR(property_code, date, aggregate_type, metric, subject_comp_market, market_segment)
```

## Arguments
- **`property_code`** — 3‑letter SinglePane property code (from **My Properties**).
- **`date`** — Either a text date in `YYYY-MM-DD` format or a normal Excel date (e.g., `6/30/2025`). The exact date to use depends on `aggregate_type` (see below).
- **`aggregate_type`** — **Case‑sensitive**. Use exactly one of:
  - `month` — Use the **last day of the month** in `date`. If monthly STR is not yet available, returns MTD where available.
  - `monthToDate` — Use the **Saturday** for the week ending (STR week end).
  - `day` — Any calendar date.
  - `currentWeek` — Use the **Saturday** for the week ending.
  - `running28Days` — Use the **Saturday** for the week ending.
  - `yearToDate` — Use the **last day of the month**.
  - `running3Month` — Use the **last day of the month**.
  - `running12Month` — Use the **last day of the month**.
- **`metric`** — One of: `Occ`, `ADR`, `RevPAR`, `Occ % Chg`, `ADR % Chg`, `RevPAR % Chg`, `MPI`, `ARI`, `RGI`, `MPI % Chg`, `ARI % Chg`, `RGI % Chg`, `Occ Rank`, `ADR Rank`, `RevPAR Rank`, `Occ % Chg Rank`, `ADR % Chg Rank`, `RevPAR % Chg Rank`.
- **`subject_comp_market`** — One of: `Subject`, `CS1`, `CS2`, …, or `Market Scale`.  
  _Notes_: For Index/Rank metrics you can use `Subject` to reference the **primary** compset, or specify a particular `CSx`/`Market Scale` to choose the comparison basis.
- **`market_segment`** — One of: `Total`, `Group`, `Contract`, `Transient`.

## Examples
Monthly RGI vs Comp Set 2:
```excel
=SP.STR("ACD","2023-06-30","month","RGI","CS2","Total")
```

Current week RevPAR index (primary compset), Transient:
```excel
=SP.STR(A2, B2, "currentWeek", "RGI", "Subject", "Transient")
```

## Tips
- Because `aggregate_type` is **case‑sensitive**, consider a validated dropdown to avoid typos.
- For week‑based aggregations (`currentWeek`, `running28Days`, `monthToDate`), pass the **Saturday of week‑end** in the `date` argument for consistent results.
- Keep a small metric dictionary table to standardize labels in your reports (e.g., “RGI — Revenue Generation Index”).

## Troubleshooting
- **Blank/NA** — Verify STR is available for the requested period and that the `date` matches the expected “anchor date” for that aggregation.
- **Wrong comp set** — For Index/Rank metrics, ensure `subject_comp_market` matches the intended compset (`Subject` for primary, or `CSx`/`Market Scale`).

## See also
- Your model may combine `SP.STR` with `SP.FINANCIALS`/`AGG` to align top‑line performance with P&L.
