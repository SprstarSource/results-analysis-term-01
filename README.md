# Term 01 School Results & Orientation Toolkit (Excel VBA)

A set of VBA modules that automate the first-term workflow of an Algerian secondary school (Arabic, right-to-left workbooks):

1. **Collect and distribute** term marks exported from the national school platform into per-class analysis workbooks.
2. **Orient** first-year common-core pupils (*جذع مشترك*) into second-year streams, using pupils' wishes, subject profiles and per-stream seat ceilings.
3. **Analyse** marks (pass rates, means, medians, modes, standard deviations, rankings, student cards).

The code is delivered as plain `.bas` files to be imported into one or more macro-enabled Excel workbooks (the "tool workbooks"). There is no installer, no external dependency, and no network access.

> **Status:** working internal tooling, published as-is. See [Known limitations](#known-limitations) before reusing it in another school or school year.

---

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Repository layout](#repository-layout)
- [Quick start](#quick-start)
- [Orientation workflow](#orientation-workflow)
- [Configuration](#configuration)
- [Safety behaviour](#safety-behaviour)
- [Known limitations](#known-limitations)
- [Data privacy](#data-privacy)
- [Documentation](#documentation)

---

## Features

| Area | What it does |
|---|---|
| Collection | Converts `.xls` / `.xlsb` / `.xlsm` exports in `Source01\` to `.xlsx` (originals are moved to `Source01\_Original`), then merges every source file into `Term01Results.xlsx`. |
| Distribution | Pastes each class section into the matching target workbook (`Data` sheet, from row 7), with a timestamped backup of every target first. |
| Orientation | Validates pupil data, imports wishes by name + birth date, assigns a stream per passing pupil, exports the final lists per section. |
| Analysis builder | Generates one `.xlsx` per class group with `Data`, `Analysis01`, `AnalyseCompaire` (multi-section groups only) and `Cart_Tec` sheets, driven by live formulas. |
| Exam analysis | Consolidates exam marks from `ExamenSource01\` into a new workbook with a general sheet, six class sheets and a data-transfer summary. Analysis cells are static values. |
| Control panels | Button-based panels (`ControlPanel`, `Term01Panel`) that run the steps in order and report progress. |

---

## Requirements

- **Windows** desktop **Microsoft Excel** (2019 or Microsoft 365). The code uses `MAXIFS`, `MINIFS`, `AGGREGATE`, `MODE.SNGL`, `Formula2` (with a `FormulaArray` fallback), `Scripting.FileSystemObject`, `Scripting.Dictionary`, and the Windows registry (font detection). Hard-coded `\` path separators are used in several modules, so macOS is not supported.
- The tool workbook must be saved in a **local folder** (not a OneDrive/SharePoint URL path). Every entry point refuses to run otherwise.
- Macros enabled; "Trust access to the VBA project object model" is optional (used only to verify that panel macros exist).
- Recommended fonts (automatic fallbacks are applied when missing): Cairo, Dubai, Dubai Medium, Aptos Slab, Arabic UI Display Light, Bahnschrift SemiBold; fallbacks are Arial and Calibri.
- All Arabic strings in the source are stored as hexadecimal code points, so every `.bas` file is **pure ASCII** and safe to import under any system locale.

---

## Repository layout

All files are flat in the repository root. The `VB_Name` inside a file is what other modules reference.

| File | `VB_Name` | Role |
|---|---|---|
| `modBuildPanel01.bas` | `modBuildPanel` | Builds `ControlPanel`, enforces workflow order (`RunTermStep01`) |
| `modCollector01.bas` | `modCollector` | `CollectTerm01Results` -> `Term01Results.xlsx` |
| `modResultFlow01.bas` | `modResultFlow01` | `DistributeDataFromTerm01Results` (self-contained) |
| `modClearDataOnly01.bas` | `modClearDataOnly01` | `ClearAllTargetFilesData01` (backup, then clear) |
| `modTransfer01.bas` | `modTransfer01` | `Run_Transfer01`: source -> `Term01Guidance.xlsx` |
| `modStreamPreferences01.bas` | `modStreamPreferences01` | `Run_StreamPreferences01`: wishes -> `Term01Guidance.xlsx` |
| `modOrientation01.bas` | `modOrientation01` | `Run_Orientation01`: stream assignment + live report |
| `modOrientationResults01.bas` | `modOrientationResults01` | `Run_OrientationResults01` -> `OrientationResults01.xlsx` |
| `modExamenAnalysis01.bas` | `modExamenAnalysis01` | `Run_ExamenAnalysis01` |
| `modBiuldAnalysisfilesBuildAll01.bas` | `modBiuldAnalysisfilesBuildAll01` | `Build_All`, `Build_One` orchestrator |
| `modBiuldAnalysisfilesData01.bas` | `modBiuldAnalysisfilesData01` | Reads `Source01`, builds the `Data` sheet, file map |
| `modBiuldAnalysisfilesAnalysis01.bas` | `modBiuldAnalysisfilesAnalysis01` | Builds `Analysis01` |
| `modBiuldAnalysisfilesCompaire01.bas` | `modBiuldAnalysisfilesCompaire01` | Builds `AnalyseCompaire` |
| `modBiuldAnalysisfilesCartTec01.bas` | `modBiuldAnalysisfilesCartTec01` | Builds `Cart_Tec` |
| `modBiuldAnalysisfilesCommon01.bas` | `modBiuldAnalysisfilesCommon01` | Shared constants, styles, data discovery |
| `modBiuldAnalysisfilesPanel01.bas` | `modBiuldAnalysisfilesPanel01` | Builds `Term01Panel` and progress display |

> File names keep their original spelling (`Biuld`). Renaming them is safe as long as `VB_Name` is left unchanged.

The `01` suffix means *Term 1*. The design assumes per-term copies of the toolkit (Term 2, Term 3) that are deliberately independent of one another.

---

## Quick start

### 1. Prepare a tool workbook

1. Create a new workbook, save it as `.xlsm` in a local working folder.
2. In the VBA editor choose **File > Import File...** and import the modules for the workflow you need:
   - **Guidance and data-flow tool:** `modBuildPanel01`, `modCollector01`, `modResultFlow01`, `modClearDataOnly01`, `modTransfer01`, `modStreamPreferences01`, `modOrientation01`, `modOrientationResults01`, `modExamenAnalysis01`.
   - **Analysis-builder tool:** all `modBiuldAnalysisfiles*01` modules.
3. Run `BuildControlPanel` (guidance tool) or `Panel_Create` (analysis-builder tool) once to create the button panel.

### 2. Folder structure expected next to the tool workbook

```text
<working folder>\
  ToolWorkbook.xlsm
  Source01\                      <- platform exports (.xls/.xlsb/.xlsm/.xlsx)
  ExamenSource01\                <- exam-mark workbooks (.xlsx), for Examen analysis
  Term01Guidance.xlsx            <- orientation workbook (you provide the template)
  StreamPreferences01.xlsx       <- pupils' wishes (sheets "Tecno" and/or "Litr")
  1LitAnalysis01.xlsm ... 3SEGlobalAnalysis01.xlsm   <- distribution targets
```

Folders created by the tools: `Source01\_Original`, `_Backup\<timestamp>\`, and the Arabic-named output folder used by `Build_All` ("first-term results").

### 3. Templates you must supply

The repository contains code only. These workbooks are **not included**:

- `Term01Guidance.xlsx`: a sheet `Data` holding a table named `Data` that starts at **A2** and has at least **21 columns**. The arts sheet `Data01` (15 columns) is created on demand when arts source sheets exist.
- The 12 target workbooks listed in `modClearDataOnly01` / `modResultFlow01`, each with a `Data` sheet (header row 6, data from row 7) and a section column.

---

## Orientation workflow

The control panel enforces this order. A step is enabled only when the previous one completed successfully and the files have not changed since.

```text
1. Data Transfer            Run_Transfer01             Source01 -> Term01Guidance.xlsx
2. Stream Preferences       Run_StreamPreferences01    StreamPreferences01.xlsx -> wishes columns
3. First Term Orientation   Run_Orientation01          assigns a decision per pupil, builds report
4. Export Results           Run_OrientationResults01   -> OrientationResults01.xlsx
```

Each step validates its entire input first and writes nothing if any error is found; the dialog lists the first 20 offending cells.

Decision values written to the guidance sheet:

- Science (`Data`, column U): Mathematics, Technical Mathematics, Experimental Sciences, Management & Economics, *Orientation* (outside the stream system), or *Repeat year*.
- Arts (`Data01`, column O): Arts & Philosophy, Foreign Languages, *Orientation*, or *Repeat year*.

The assignment rules are documented in [ARCHITECTURE.md](ARCHITECTURE.md#5-orientation-algorithm).

---

## Configuration

Tunable parameters are constants at the top of each module.

| Constant | Module | Default | Meaning |
|---|---|---|---|
| `PASS_MARK` | Orientation, Examen, Common | `10` | Pass threshold (defined separately in each module) |
| `TOP_PCT` | Orientation | `0.05` | Share of top pupils whose first wish is granted unconditionally |
| `CAP_MATH` / `CAP_TECH` | Orientation | `0.18` / `0.22` | Seat ceilings as a share of passers (rounded down) |
| `CAP_EXP` / `CAP_MGMT` | Orientation | `0.55` / `1` | Seat ceilings (rounded up) |
| `COMPAT_GAP` | Orientation | `1.0` | Maximum profile advantage of the alternative over wish 1 |
| `COMPAT_MIN` | Orientation | `12.99` | Profile level that exempts wish 1 from subject conditions (experimental and management only) |
| `AGE_OFFSET` | Orientation | `17` | Age rule for *Repeat year* vs *Orientation* |
| `WORK_YEAR_OVERRIDE` | Orientation | `0` | `0` = use `Year(Date)`; set a year to pin the calculation |
| `ALLOW_PARTIAL` | ResultFlow | `False` | If `False`, a target workbook is skipped when any required section is missing |
| `MAX_STUDENTS` | Examen | `200` | Per-cohort capacity |

The file map (target workbooks and the sections each one requires) is hard-coded in `modResultFlow01`, `modClearDataOnly01` and `GetFileMap` in `modBiuldAnalysisfilesData01`. Adapting the toolkit to another school means editing those three places.

---

## Safety behaviour

- Target workbooks are backed up to `_Backup\<timestamp>\` before any change, and a file is saved only if every step for that file succeeded.
- Workbooks that are already open, read-only, or missing the expected sheet are skipped and reported; they are never modified.
- Source `.xls/.xlsb/.xlsm` files are archived rather than deleted. Macros are force-disabled while opening them.
- Application state (screen updating, events, alerts, calculation mode) is restored on exit in most modules (see limitations).
- Sheet protection on the guidance workbook uses an **empty password**. It prevents accidental edits and is not a security control.

---

## Known limitations

- **Date-dependent rule:** unless `WORK_YEAR_OVERRIDE` is set, the age rule uses `Year(Date)`, so re-running the same data in a different calendar year can change *Repeat year* / *Orientation* outcomes.
- **Zero vs. absent mark:** analysis formulas treat `0` as "absent" (`>0` masks); orientation and transfer treat `0` as a real mark; Examen includes `0` in the weighted average but excludes it from per-subject statistics.
- **School-specific configuration** (file names, section counts, coefficient tables) is hard-coded.
- **Duplicated helper code** across modules (about 56 procedure names appear more than once). Most copies are identical; `IsNum`, `LockAndProtect`, `IsWorkbookOpen` and `ClearDataArea` have diverging variants.
- **Windows only**, and only for local folders.
- No automated tests. The orientation algorithm in particular has not been validated against small cohorts (fewer than ~12 passers) or heavy score ties.

A fuller list with locations is in [ARCHITECTURE.md](ARCHITECTURE.md#9-known-limitations-and-technical-debt).

---

## Data privacy

The workbooks processed by these tools contain pupils' names, birth dates and marks. **Never commit real data** to a public repository. Add at least the following to `.gitignore`:

```gitignore
*.xlsx
*.xlsm
*.xls
*.xlsb
Source01/
ExamenSource01/
_Backup/
```

---

## Documentation

- [ARCHITECTURE.md](ARCHITECTURE.md): data flow, module responsibilities, file contracts, orientation algorithm, statistical conventions, design decisions and technical debt.

## License

No license has been specified yet. Add a `LICENSE` file before publishing if you want others to reuse the code.
