# CLAUDE.md — APAC Treatment Dashboard (Excel + VBA)

Excel workbook `APAC_Treatment_Dashboard_v25_9_live_v1.xlsm` (repo root) that imports
country KPI workbooks, builds PivotTables / PivotCharts / slicers, and exports PDFs.
All logic is VBA. The `.bas` files in this `working/` folder are the **text export** of
the workbook's VBA modules. Full behaviour is in `SPECIFICATION.md` (same folder).

## Ground rules

- **Edit and create code only inside `working/`.** Never modify the `.xlsm`, `.docx`
  or anything outside `working/` (binary files, and the macro-enabled workbook is the
  user's to save).
- The `.xlsm` is the runtime; `.bas` here is the source of truth for review/diff. After
  changing a `.bas`, tell the user to **re-import it into the workbook** (VBE: File →
  Import File, remove the old module first) and re-export before committing.
- Do not run or test VBA here: this is a Linux container with no Excel. Verify by
  careful reading, and list manual test steps for the user.
- Commit only when asked. Develop on the branch the session names; no PRs unless asked.

## Environment constraints (do not break)

- **Excel 2019 compatible.** Do not use enum members missing in 2019 (see the
  `XL_SLICER_*` numeric constants in `Main.bas`). Use numeric constants, not new enums.
- `Option Explicit` in every module. Declare all variables.
- VBA has no UTF-8 in code files: keep `.bas` ASCII (the `?` in message strings are
  mojibake of emoji; don't "fix" them to emoji).
- Modules export with `Attribute VB_Name` header and **CRLF** line endings. Preserve
  both when editing (do not convert to LF).

## Code conventions (follow existing style)

- Step tracing: set `gStep = "Module.Proc: what"` before each significant action; error
  handlers show `gStep`. Keep this pattern in new code.
- Entry procedures: `On Error GoTo CleanFail`, then a `CleanExit` that restores
  `ScreenUpdating`, `DisplayAlerts`, `EnableEvents`, `Calculation`, re-applies protection
  (`modSecurity.ApplyHardening`) and closes progress (`ProgressFinish`/`ProgressForceClose`).
  Never leave Excel with events/calculation off.
- Long operations use `ProgressStart / ProgressSet / ProgressStep / ProgressFinish`.
- Import is **validate-all-then-import**: `modImportValidation` must pass for every sheet
  before any write. Add new rules there, not in the importer.
- Upsert by key `BuildKey(country, region, center, date)`; never blindly append to `tblData`.
- Sheet names used by code: `Dashboard`, `Data` (table `tblData`), `Pivots`,
  `Upload Summary`, `Pivots_Scratch` (temp), per-country dashboard/pivot sheets,
  `CS__*` temp chart sheets. Do not rename without updating every reference.
- Protection: sheets/workbook are locked by `modSecurity`. Code that edits protected
  objects must unprotect, act, and re-protect in the same call path (see
  `BeginStructureUnlock`, `gRunUnlocked`).
- Re-entrancy: slicer-driven chart refresh is guarded by `gChartRefreshBusy` and deferred
  with `Application.OnTime` (`gPendingChartCountry`, `gPendingAxisCountry`). Keep it.
- Reserved globals (`gRebuild*`, `gDeferCombo*`) are intentionally unused; leave them.

## Security notes

- Passwords are currently **hardcoded** in `modSecurity.bas` (`SYS_*_PWD`) and
  `DEV_MODE` is a `Public Const`. Never print them into docs/chat; never set
  `DEV_MODE = True` in a committed change. Flag to the user that VBA passwords in a
  public/shared repo are exposed; recommend moving secrets out of source.
- Never add network calls, shell execution, or file deletion beyond what the
  backup/PDF-cleanup modules already do, without explicit user approval.

## Module map

| Module | Role |
|---|---|
| `Main.bas` | Import orchestration, pivot/chart/slicer build, country dashboards, repair/refresh |
| `modImportValidation.bas` | Pre-import rule checks (rules 1-9), consolidated error report |
| `modImportMthAdj.bas` | Previous-month adjustment import + `AdjLog` |
| `modExportPdfChartDirect.bas` | Chart-sheet PDF export, scratch pivots, PDF archive/cleanup |
| `modSecurity.bas` / `modSecurityManager.bas` | Protection, hardening, sheet-delete toggle |
| `modITUtils.bas` | IT admin utilities, password prompt, user identity, OneDrive path |
| `modBackup.bas` | Backup copy to OneDrive library folder, token, folder creation |
| `ProgressManager.bas` | Progress UI |
| `modFileUtil.bas`, `modMISC.bas`, `Module1.bas`, `modDummyPivotCache.bas`, `modSecurityManager.bas` | Small helpers / button creation / debug |

## Workflow for a change

1. Read `SPECIFICATION.md` section for the area; read the affected procedures fully.
2. Make the smallest change; keep style and `gStep` tracing.
3. Run the `vba-reviewer` agent on the changed `.bas` before reporting done.
4. Update `SPECIFICATION.md` if behaviour, rules, or sheet/column names changed.
5. Give the user: modules to re-import, and a manual test checklist.

## Git

- Remote `tanhs/apac_kpi`; local Windows repo `C:\dev\apac_kpi` (cmd, not bash: no `printf`).
- `.gitignore` excludes `~$*` lock files. Binary `.xlsm` diffs are opaque: review the `.bas`.
