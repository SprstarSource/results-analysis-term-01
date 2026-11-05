# Changelog

All notable changes to this project are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses [Semantic Versioning](https://semver.org/) once tagged releases exist.

> **Note on history.** No version history was available when this changelog was created. The `0.1.0` entry is a **baseline reconstructed from the source code** and describes the state of the toolkit at the time it was first documented. It is not a record of earlier releases.

## How to maintain this file

- Add every user-visible or behaviour-affecting change under **[Unreleased]** in the same pull request that makes it.
- Use these headings only: `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, `Security`.
- Prefix the entry with the module in backticks, e.g. `modOrientation01`, so a reader can find the change.
- Changes that alter **decisions or numbers** produced for pupils (orientation rules, pass mark, profile coefficients, statistics) must be flagged with **(behaviour change)**.
- Moving to a new term (Term 02, Term 03) is a copy of the modules with a new suffix, not a version bump; record it as `Added`.

## Versioning policy

| Increment | When |
|---|---|
| MAJOR | A change to a workbook/file contract that makes existing templates unusable (section 4 of `ARCHITECTURE.md`): column positions, sheet names, table anchors, file names |
| MINOR | New capability, new module, new optional configuration, new report |
| PATCH | Bug fix or message/documentation change that does not alter outputs |

---

## [Unreleased]

### Added
- Project documentation: `README.md`, `ARCHITECTURE.md`, `changelog.md`, `contributing.md`, `security.md`, `test_plan.md`, `troubleshooting_ar.md`, `user_guide_ar.md`.

### Known issues (not yet fixed)

Tracked here so that fixes can be recorded against them. Details and locations are in `ARCHITECTURE.md` section 9.

| ID | Module(s) | Issue |
|---|---|---|
| KI-01 | `modOrientation01` | Repeat/Orientation rule depends on `Year(Date)` unless `WORK_YEAR_OVERRIDE` is set |
| KI-02 | analysis modules, `modTransfer01`, `modOrientation01`, `modExamenAnalysis01` | `0` is "absent" in some modules and a real mark in others |
| KI-03 | `modResultFlow01` | Section column position is not verified before paste; its `ClearDataArea` guesses the clear range when the section header is not found |
| KI-04 | `modOrientation01` | One passer without an eligible branch aborts the whole run; the user's table is re-sorted in place |
| KI-05 | `modBuildPanel` | Workflow state is lost if the tool workbook is not saved; manual wish edits reset the stage; `fullHash` is never used |
| KI-06 | `modCollector`, `modResultFlow01`, `modTransfer01` | Converted-file counter increases when moving the original fails |
| KI-07 | several | About 56 duplicated procedure names; diverging variants of `IsNum`, `LockAndProtect`, `IsWorkbookOpen`, `ClearDataArea` |
| KI-08 | several | School-specific values hard-coded: file map, section counts, coefficient tables, and the school/directorate names printed in the orientation report headers |
| KI-09 | `modTransfer01` | No duplicate-section check when two source sheets share a group number |
| KI-10 | all | No automated tests (see `test_plan.md`) |

---

## [0.1.0] - 2026-11-05 (documented baseline)

### Added

**Control and workflow**
- `modBuildPanel`: `ControlPanel` sheet with eight buttons (First Term Results, First Term Analysis, Clear Tables, Examen Analysis, Stream Preferences, Data Transfer, First Term Orientation, Export Orientation Results); verifies that the macros behind the buttons are imported.
- Ordered orientation workflow (Data Transfer, Stream Preferences, First Term Orientation, Export Results) with stage tracking in hidden names and file-stamp invalidation.

**Data flow**
- `modCollector`: converts `.xls`/`.xlsb`/`.xlsm` exports in `Source01` to `.xlsx` (originals archived in `Source01\_Original`) and merges all sources into `Term01Results.xlsx`.
- `modResultFlow01`: distributes class sections into 12 target workbooks with a timestamped backup per run and per-file save-on-success.
- `modClearDataOnly01`: backs up and clears the data area (row 7 down, up to the section column) of the 12 target workbooks after confirmation.

**Orientation**
- `modTransfer01`: validates source headers and data, writes science (`Data`, 21 columns) and arts (`Data01`, 15 columns) sheets in `Term01Guidance.xlsx`, computes the four science profiles and two arts profiles.
- `modStreamPreferences01`: imports wishes from `StreamPreferences01.xlsx` (sheets `Tecno`, `Litr`) by name plus birth date, with duplicate and ambiguity handling.
- `modOrientation01`: science stream assignment (top-5% grant, seat ceilings, two wish rounds, profile compatibility, leftover assignment, first-wish restoration) and arts assignment (arts vs foreign languages), classification into pass / repeat / orientation, live general report.
- `modOrientationResults01`: exports final lists to `OrientationResults01.xlsx` (`Litr` and one sheet per science section).

**Analysis**
- `modExamenAnalysis01`: consolidates exam marks from `ExamenSource01` into six cohorts with weighted semester averages and static statistics.
- `modBiuldAnalysisfiles*`: builds one `.xlsx` per class group with `Data`, `Analysis01`, `AnalyseCompaire`, `Cart_Tec`; `Term01Panel` with `Term 01 Analysis` and open-output buttons; text build log with one retry per failed file.

**Cross-cutting**
- ASCII-only sources with Arabic text built from hexadecimal code points.
- Fail-closed validation, backup-before-modify, restore of Application state in most modules.

[Unreleased]: ../../compare/v0.1.0...HEAD
[0.1.0]: ../../releases/tag/v0.1.0
