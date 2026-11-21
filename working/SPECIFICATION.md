# APAC Treatment Dashboard — Specification

Derived from the VBA source in this folder (v25.9). Where the code is ambiguous, items
are marked **(verify)**. Update this file whenever behaviour changes.

## 1. Purpose

A senior-management dashboard of daily **Budget vs Actual** treatment volumes per
Country / Region / Center, with achievement % and variance, refreshed by importing
country workbooks with one click, plus PDF export of charts.

## 2. Users and workbook layout

- **Users**: country data submitters/analysts (import, view), IT admin (maintenance).
- **Sheets**
  - `Dashboard`: KPIs, slicers, 3 PivotCharts, 2 breakdown tables (Country x Month,
    Country x Center), Import buttons.
  - `Data`: table `tblData`, the single data store (append/upsert only via importer).
  - `Pivots`: all macro-built PivotTables.
  - `Upload Summary`: import audit.
  - Per-country dashboard and pivot sheets (one per country key, e.g. MY/SG/CN/JP).
  - Hidden/very hidden: calc sheet for KPI outputs; `Pivots_Scratch` and `CS__*`
    chart sheets exist only transiently during PDF export.
- Sheet-structure changes by users are blocked by workbook protection.

## 3. Data model — `tblData`

| Column | Meaning |
|---|---|
| Country, Region, Center | From source sheet name `Country - Center` (+ region lookup) |
| Date | Day (yyyy-mm-dd) |
| Month | Month label of the date |
| Budget, Actual | Numbers, 2 decimals |
| Variance | Budget - Actual |
| AchievementPct | Actual / Budget (pivot calculated field; hence Budget > 0) |
| UploadId, UploadedAt | Import batch identity/time |
| LatestDay, LatestUploadedAt | Latest-record helpers for KPIs |

- **Key**: `Country|Region|Center|Date` (`BuildKey`). Re-importing the same key updates
  the row (`UpsertRow`); it does not duplicate.
- Blank rows are removed (`RemoveBlankRowsInTblData`).
- Adjustment log: `AdjLog` table (see 6) records previous-month adjustments with `ADJ_VERSION` "v1".

## 4. Source workbook format (country files)

- One worksheet per center. Sheet name: exactly 3 parts separated by `-`:
  part 1 in {MY, JP, SG, CN}, parts 2 and 3 alphanumeric; total length <= 30 including
  separators.
- `C3` = full month name (January..December); required.
- Row 4 headers fixed: `C4`=Day, `D4`=Budget, `E4`=Actual, `F4`=Variance. Additional
  month blocks repeat to the right, each starting with a "Day" header in row 4
  (block = Day, Budget, Actual, Variance, %).
- Data from row 5 down to last row of the first Day column.

## 5. Import flow (`ImportAndRefresh`)

1. User picks one or more country workbooks (FileDialog); files opened read-only; locked
   files detected (`IsFileLocked`).
2. **Validate all** sheets of all files (`ValidateSelectedWorkbooks`). Any failure ->
   one consolidated message, **no data written** (all-or-nothing).
3. Backup workbook copy (`BackupWorkbookCopy`, stage suffix + token) to the OneDrive
   library folder.
4. Parse each sheet (`ImportOneSheet`) and upsert into `tblData`.
5. Write upload summary; refresh pivot caches (`RefreshAllPivotCaches`); build or
   refresh pivots/slicers/charts/tables (`EnsurePivotsSlicersChartsAndTables`; smart gate:
   `DashboardIsBuilt` -> refresh-only vs full rebuild); build country dashboards.
6. Always restore Excel state and re-apply protection, even on failure.

## 6. Validation rules (`modImportValidation`)

| # | Rule |
|---|---|
| 1 | `C3` is a full month name |
| 2 | Day column is yyyy-mm-dd and cell format may not be changed/other |
| 3 | Budget and Actual numeric (or blank) with 2-decimal format `####.00` |
| 4 | Variance must equal Budget - Actual (formula-consistent, not hand-tallied) |
| 5 | Month at `C3` present with data |
| 6 | Headers fixed in row 4 (Day/Budget/Actual/Variance at C-F) |
| 7 | **(verify)** not present as a labelled rule in code comments; confirm numbering |
| 8 | Sheet naming convention (section 4) |
| 9 | Budget > 0 (prevents #DIV/0! in Ach%) |

Errors accumulate and are shown together (`MAX_MSG_CHARS` 30,000 cap).

## 7. Previous-month adjustment import (`ImportPrevMonthAdj`)

Separate button. Select one adjustment output file; validate
(`ValidatePrevMonthAdjWorkbook`, month name, country name -> code); check coverage in
existing data (`ValidateAdjCoverageInData`); apply row adjustments to the matching
previous-month data (`ApplyOneAdjustmentRow`, using `SumActualForMonth` and
`LookupRegionByCenter`); record each in the adjustment log table with an upload id
(`BuildAdjUploadId`); skipped rows are logged (`AppendSkip`).

## 8. Dashboard and charts

- KPIs on `Dashboard` from a hidden calc sheet.
- 3 PivotCharts (Budget/Actual columns with Achievement % line on secondary axis;
  `ApplyPivotChartComboFormat`, `ApplyAchOnlySeriesFormat`) and two formula-driven
  breakdown tables (`RenderCountryMonthTable`, `RenderCountryCenterTable`).
- Slicers filter charts; cross-filter hides items without data (Excel 2019-safe
  numeric constants). Slicer changes trigger a guarded, deferred chart refresh
  (`gChartRefreshBusy`, `OnTime`), with `RunPendingChartRefresh` / `RunPendingAxisRebuild`.
- Per-country dashboards built from `tblData` (`BuildOneCountryDashboard`) with health
  checks (`IsCountryPivotHealthy`, `...SourceCurrent`) and `RepairCountryDashboard`.

## 9. PDF export (`modExportPdfChartDirect`)

`ExportAllCharts_Fast_ChartSheets_Direct`: temporarily creates chart sheets (`CS__` prefix)
from dashboard charts and per-country/region scratch pivots (`Pivots_Scratch`), adds a
footer with generation time, exports to PDF, then deletes all temp objects. Output
folder is environment-aware (OneDrive/local); older PDFs are archived and cleaned up
(`ArchiveExistingPdfs`, `CleanupOldPdfExports`).

## 10. Security and administration

- Workbook structure and sheets protected; `ApplyHardening` is called on each entry
  point; pivot-friendly protection allows slicer use; data sheet strictly locked.
- IT admin actions (delete non-core sheets, enable sheet delete) require a password
  prompt (`PromptPasswordAndValidate`) and confirmation (`ConfirmAction`).
- `DEV_MODE` must be False in released builds.
- User identity (`GetBestUserName/Key`) and OneDrive path discovery are used for
  backups/audit.

## 11. Non-functional requirements

- Excel 2019 compatible; Windows desktop Excel only (OneDrive, FileDialog, OnTime).
- Excel state (`ScreenUpdating`, `EnableEvents`, `Calculation`, alerts) is always restored.
- Import is atomic from the user's view: validation before write, backup before write.
- Progress UI for operations longer than a few seconds.

## 12. Known gaps / recommendations

- Hardcoded passwords in source (`modSecurity`); move out of repo or accept VBA-password
  limits. Treat as non-secret.
- Stray `?` characters in messages are encoding loss of symbols from the export.
- `Main.bas` is ~7.5k lines; consider splitting (import / pivots / charts / country dashboards).
- No automated tests; add a manual regression checklist per release (import valid file,
  import each invalid-rule file, re-import same file = no duplicates, slicer refresh,
  PDF export, protection intact after error).
