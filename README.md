# SnapshotVBA
Snapshot SP

A VBA macro suite that builds a portfolio “Snapshot” report for hotel assets using the SinglePane Excel add-in (`SP.FINANCIALS` / `SP.FINANCIALS_AGG`). It organizes results by **Fund**, adds **Fund Subtotals**, **Total Managed Portfolio** (excludes a specific fund), and **Total Portfolio**, with three stacked modes: **MTD**, **YTD**, and **FY**. The module also auto-creates a **USALI Map** for human-friendly metric names and manages the named ranges and inputs the report depends on.

---

## Table of Contents

* [What this does](#what-this-does)
* [Prerequisites](#prerequisites)
* [Sheets this macro creates/uses](#sheets-this-macro-createsuses)
* [How data is read](#how-data-is-read)
* [How the report is built](#how-the-report-is-built)
* [How to run it (Quick Start)](#how-to-run-it-quick-start)
* [Core macros you’ll use](#core-macros-youll-use)
* [Metric bands & calculations](#metric-bands--calculations)
* [USALI Map: display names → USALI codes](#usali-map-display-names--usali-codes)
* [Formatting choices](#formatting-choices)
* [Configuration knobs (customize behavior)](#configuration-knobs-customize-behavior)
* [Name management & stability](#name-management--stability)
* [Troubleshooting](#troubleshooting)
* [FAQ](#faq)
* [License / Notes](#license--notes)

---

## What this does

* Builds a fully formatted **Snapshot** worksheet with three stacked tables: **MTD**, **YTD**, **FY**.
* Appends an **STR performance** section with MTD, YTD and Running 12‑Month metrics (Occ, MPI/ARI/RGI and Change).
* Groups hotels **by Fund**; within each fund, properties are **sorted by Hotel name**.
* Adds **Fund Subtotal** rows and two portfolio lines:

  * **Total Managed Portfolio** (excludes a designated fund),
  * **Total Portfolio** (all properties).
* Automatically generates/maintains:

* An **Input** sheet (month/year selectors plus **Generate Report** and **Export to PDF** buttons).
  * A **USALI Map** sheet to translate human-friendly metric labels into tenant-specific USALI strings.
  * A hidden **Helper** sheet that stores named ranges of property codes per fund for aggregate queries.
  * All required workbook **Named Ranges**, while removing sheet-scoped duplicates that cause “shadowing.”

---

## Prerequisites

* **Excel** with dynamic arrays (for `A1#` spills). Office 365 / Excel 2019+ recommended.
* **SinglePane Excel add-in** installed and logged in.
* Two data sheets in the workbook:

  1. **`My Properties`** with a spilled or contiguous table starting at **A1**, containing at minimum the columns:

     * `Code`, `HotelName`, `ManagementCompany`, `Rooms`, `Fund`
  2. **`Usali Reference`** with a spilled or contiguous table starting at **A1**, containing a column header **`usali`** (case-insensitive).

> The code automatically tolerates spilled ranges or standard tables by detecting `A1#` or falling back to `CurrentRegion`.

---

## Sheets this macro creates/uses

* **`Snapshot`** (report output): created if missing, cleared and rebuilt on each run.
* **`Input`** (user controls): created if missing; holds Month/Year selectors and two buttons: **Generate Report** and **Export to PDF**.
* **`USALI Map`** (editable mapping): created/updated to map display metrics → tenant USALI strings. Flags missing codes.
* **`Helper`** (hidden): holds named ranges of codes by fund and totals:

  * `Codes_<FundSanitized>`
  * `Codes_TotalManaged`
  * `Codes_TotalPortfolio`

---

## How data is read

* **`My Properties`**:

  * The code locates key columns by header text (case-insensitive): `Code`, `HotelName`, `ManagementCompany`, `Rooms`, `Fund`.
  * Builds an in-memory dictionary: `Fund → Collection of (Hotel, Code, Manager, Rooms)`.
  * Sorts hotels alphabetically within each fund.
* **Ordering of funds**:

  * Funds sorted alphabetically; the constant `FUND_EXCLUDE` (default **"Stonebridge Legacy"**) is moved to the end.
* **`Usali Reference`**:

  * Must contain a `usali` column. Used to validate mapping rows in **USALI Map**.

---

## How the report is built

**Three stacked tables:** **MTD**, **YTD**, **FY**.
Each table includes:

* Two-row header with metric “bands” (Actual, Budget, Var vs Bud, LY, Var vs LY).
* Property rows grouped by Fund.
* A **Fund Subtotal** row per fund.
* A spacer.
* **Total Managed Portfolio (ex. FUND\_EXCLUDE)**.
* **Total Portfolio**.

Data cells are formulas to `SP.FINANCIALS` (property-level) or `SP.FINANCIALS_AGG` (aggregate rows) using:

* A three-letter month token (e.g., `Jun`) or `MMMYTD` for YTD, or `"Total Year"` for FY.
* Year number from the Input sheet.
* Version tokens: `Actual`, `Budget`, `LY_Actual`, and `ForecastN` for FY (N = month number).

---

## How to run it (Quick Start)

1. **Open** the workbook that contains **`My Properties`** and **`Usali Reference`**.
2. **Enable macros** and ensure you’re **logged into SinglePane**.
3. If this is your first run, execute **`BuildFormatRun`** from the Macro dialog *or* go to the **Input** sheet and press **Generate Report**.

   * First run auto-creates **Input**, **USALI Map**, **Helper**, and required **Named Ranges**.
4. On the **Input** sheet, pick the **Month** (full name dropdown) and **Year**.
5. Press **Generate Report** again any time you change Month/Year.
6. Use **Export to PDF** to save the current snapshot as a PDF file.

> You can also run **`BuildSnapshot`** directly (it builds all three tables).
> **`BuildFormatRun`** = Auto-setup + Build + Final formatting.

---

## Core macros you’ll use

* **`BuildFormatRun`** — One-click workflow: ensures inputs & names, builds, and formats the Snapshot.
* **`BuildSnapshot`** — Builds the three stacked tables (MTD/YTD/FY).
* **`AutoSetupOnOpen`** — Ensures the Input sheet and named ranges exist (call from `Workbook_Open` if you want).
* **`HardResetSnapshotConfig`** — Resets old inputs/names; use this if you’ve had naming conflicts.
* **`FixNamesNow`** — Purges sheet-scoped duplicates and rebinds critical names + USALI map.

---

## Metric bands & calculations

### Bands per table

* **MTD**: `Actual`, `Budget`, `Var vs Bud`, `LY`, `Var vs LY`
* **YTD**: same as MTD, with token `MMMYTD`
* **FY**: `Forecast` (as **ForecastN**, where N = selected month number), `Budget`, `Var vs Bud`, `LY`, `Var vs LY`

### Metrics in each band (default)

`Occ`, `ADR`, `RevPAR`, `Total Rev (000's)`, `NOI (000's)`, `NOI Margin`
(See `MetricsList()`; you can customize these.)

### Variances

* **As % of base** for: `ADR`, `RevPAR`, `Total Rev (000's)`, `NOI (000's)`.
* **Point difference (p.p.)** for: `Occ`, `NOI Margin`.
* All variances formatted with `+0.0%;-0.0%;0.0%`.

### Aggregates

* Aggregate rows (Fund subtotals & Portfolio totals) use `SP.FINANCIALS_AGG` with named code lists:

  * `Codes_<Fund>`, `Codes_TotalManaged`, `Codes_TotalPortfolio`.
* **NOI Margin on aggregate rows** is calculated manually as `NOI / Total Rev` (per band) to avoid weighted-average errors.

---

## USALI Map: display names → USALI codes

* The **`USALI Map`** sheet contains two mapping blocks:

  * Columns **A:C** mirror the traditional USALI mapping with headers:
    1. `DisplayMetric`
    2. `USALI`
    3. `Notes` (flags “NOT FOUND in Usali Reference” if a mapping isn’t present in your tenant’s reference)
  * Columns **E:F** provide an editable map for **STR metrics** (`STR_Display` → `STR_Code`).
* The report headers use `DisplayMetric` (friendly names). Each data cell looks up the **USALI** code via:

  ```text
  XLOOKUP(<header cell>, UsaliMap_Display, UsaliMap_Code)
  ```
* **Add/edit rows freely** (keep the named ranges bound to the columns). The builder refreshes names every run:

  * `UsaliMap_Display` → column A
  * `UsaliMap_Code` → column B
  * `StrMap_Display` → column E
  * `StrMap_Code` → column F

> Tip: If you rename or add metrics in `MetricsList()`, be sure the `DisplayMetric` text in the header appears in `USALI Map` with the correct tenant-specific USALI string.

---

## Formatting choices

* **Header bar** (row 2) styled in brand red (`#E03C31`), white bold text.
* **Two-row band headers** with merged left columns: **Hotel**, **Rooms**, **Manager**.
* **Alternating row shading** for property rows.
* **Medium borders** around fund subtotals and portfolio totals.
* **Vertical separators** between bands and to the **left of the first value column**.
* **Font sizes**:

  * Data rows: **12 pt**
  * Fund/Portfolio total rows: **13 pt**
  * Title (“Hospitality Portfolio Snapshot”): **16 pt**
  * “As of” and Month/Year: **14 pt**
* **Number formats**:

  * Currency: `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`
  * Thousands: `_($* #,##0,_) ;_($* (#,##0,)_);_($* "-"_) ;_(@_)`
  * Percent: `0.0%`
  * Variance: `+0.0%;-0.0%;0.0%`

---

## Configuration knobs (customize behavior)

Edit these **constants/functions** at the top of the module to tailor the report:

* **`FUND_EXCLUDE`**: default `"Stonebridge Legacy"`.

  * This fund is listed last and excluded from **Total Managed Portfolio**.
* **Brand color**: change `RED_HEX` in `WriteTwoRowHeader` / `FormatSnapshotShell` (default `"E03C31"`).
* **Metrics**: edit `MetricsList()` (order matters).
* **Manager name shortening**: `ShortManagerName()` keeps **“Great Wolf”** as is; otherwise first word of name.
* **Widths**: column widths set in `WriteTwoRowHeader` (metric-dependent).

---

## Name management & stability

To prevent broken links and shadowed names, the module:

* Removes **sheet-scoped duplicates** before adding workbook-scoped names (`KillAllSheetScoped`).
* Maintains these key names:

  * **Inputs**: `MonthText` (full month name), `YearNum` (year), `MonthNum` (1–12).
  * **Snapshot header**: `Snap_MonthFull`, `Snap_MonthNum`, `Snap_MonthMMM`, `Snap_YearNum`.
  * **USALI mapping**: `UsaliMap_Display`, `UsaliMap_Code`.
  * **Aggregate code lists**: `Codes_<Fund>`, `Codes_TotalManaged`, `Codes_TotalPortfolio`.

If you run into name conflicts:

* Use **`FixNamesNow`** (quick clean and rebind), or
* Use **`HardResetSnapshotConfig`** (full reset of legacy inputs + names).

---

## Troubleshooting

**“'My Properties' sheet not found.”**
→ Add the sheet, ensure the data starts at A1 (spill or region), and headers include `Code`, `HotelName`, `ManagementCompany`, `Rooms`, `Fund`.

**“'Usali Reference' missing 'usali' header.”**
→ Ensure a column header named `usali` (case-insensitive). The sheet must spill or have a `CurrentRegion` starting at A1.

**Blank values for metrics**

* Verify you’re **logged into SinglePane** and have access to the asset codes.
* Confirm the **USALI Map** has correct tenant strings for your metrics.
* Check that **Month/Year** inputs are set correctly on the **Input** sheet.

**Weird month/year on the Snapshot header**

* Don’t overwrite `G2`/`H2` on `Snapshot`. They’re formulas bound to `MonthText`/`YearNum`.
* Change values only on the **Input** sheet.

**Wrong NOI Margin on subtotals**

* By design, **aggregate NOI Margin** is computed as `NOI / Total Rev` per band (not a weighted average of property margins).

**Name shadowing errors**

* Run **`FixNamesNow`** or **`HardResetSnapshotConfig`** and rebuild.

---

## FAQ

**Q: Can I run just one table (e.g., YTD only)?**
A: The provided entry points build all three. You can adapt `BuildSnapshot` to call `BuildOneTable` selectively.

**Q: How does FY Forecast work?**
A: FY uses `"Total Year"` for most versions; the **Forecast** band uses `"ForecastN"`, where `N = MonthNum` from the Input sheet (e.g., `Forecast6` for June).

**Q: Can I add more metrics?**
A: Yes. Add to `MetricsList()`, then add corresponding `DisplayMetric → USALI` rows on **USALI Map**. If a metric is a percent or a “thousands” metric, update `MetricIsPercent()` / `MetricIsThousands()`.

**Q: Can I exclude multiple funds from Managed Portfolio?**
A: Out of the box, only one exclusion is supported (`FUND_EXCLUDE`). To exclude more, adjust `PrepareHelperCodeLists` logic where `managedSet` is built.

**Q: Where do the aggregate code lists come from?**
A: The **Helper** sheet. The macro writes hotel **Codes** per Fund into blocks named `Codes_<Fund>`, plus compiled lists for `Codes_TotalManaged` / `Codes_TotalPortfolio`.

---

## License / Notes

* Internal reporting utility for environments using the **SinglePane** Excel add-in.
* This module uses **late binding** for `Scripting.Dictionary` (no explicit library reference required).
* Tested on Windows Excel 365. Mac behavior with Form Controls may vary.

---

### Handy macro list (copy for your convenience)

* `BuildFormatRun` — recommended “one-button” run
* `BuildSnapshot` — build only
* `AutoSetupOnOpen` — ensure inputs/names (optional to call from `Workbook_Open`)
* `HardResetSnapshotConfig` — full reset of legacy inputs/names
* `FixNamesNow` — purge sheet-scoped duplicates & rebind names

---

**That’s it!** Drop this module into your workbook, confirm the two source sheets, pick your Month/Year on **Input**, and click **Generate Report**. When you're ready to share, use **Export to PDF** to create a snapshot file.
