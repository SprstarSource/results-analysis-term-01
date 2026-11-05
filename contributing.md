# Contributing

Thank you for helping improve this toolkit. It processes real pupils' marks and produces orientation decisions, so the bar for changes is **correctness and reversibility first, features second**.

Read `README.md` and `ARCHITECTURE.md` before your first change. Run the relevant cases from `test_plan.md` before opening a pull request.

## Contents

1. [Ground rules](#1-ground-rules)
2. [Development environment](#2-development-environment)
3. [Repository and file conventions](#3-repository-and-file-conventions)
4. [Coding standards](#4-coding-standards)
5. [Working with duplicated code](#5-working-with-duplicated-code)
6. [Changing behaviour that affects pupils](#6-changing-behaviour-that-affects-pupils)
7. [Branches, commits and pull requests](#7-branches-commits-and-pull-requests)
8. [Documentation requirements](#8-documentation-requirements)
9. [Adding a new term or a new school](#9-adding-a-new-term-or-a-new-school)
10. [Reporting issues](#10-reporting-issues)

---

## 1. Ground rules

1. **Never commit real data.** No `.xlsx`/`.xlsm`/`.xls`/`.xlsb` files containing names, birth dates or marks, in code, tests, screenshots, issues or pull requests. Use synthetic data only (`test_plan.md` section 4).
2. **Fail closed.** If input is invalid or ambiguous, stop and report; do not guess. A run that cannot finish must leave every workbook as it found it.
3. **Reversible by default.** Back up before modifying, archive instead of deleting, save only after every step succeeded.
4. **One change, one reason.** Keep pull requests small and focused.
5. **Do not silently change outputs.** Any change that can alter a decision or a number must follow section 6.

## 2. Development environment

- Windows, desktop Microsoft Excel 2019 or Microsoft 365, Arabic-capable system fonts.
- A **development workbook** (`.xlsm`) saved in a **local folder** (not OneDrive/SharePoint), separate from any workbook used with real data.
- Optional: *Trust access to the VBA project object model* (used by `CheckPanelMacros`).
- Git client. Python 3 is optional and only needed for the helper scripts below.

### Importing and exporting modules

1. In the VBA editor use **File > Import File...** to import the `.bas` files you need (see the grouping in `README.md`).
2. Edit in the editor, test, then **File > Export File...** back over the repository file.
3. Keep the `Attribute VB_Name` line unchanged. Other modules call procedures by qualified module name (for example `modTransfer01.Run_Transfer01`), and `modBuildPanel` / `modCollector` are referenced without a suffix.
4. Do not commit the development workbook. Commit only exported `.bas` files.

Recommended `.gitattributes` so the VBA editor can import files checked out on Windows:

```gitattributes
*.bas text eol=crlf
*.md  text eol=lf
```

## 3. Repository and file conventions

| Rule | Reason |
|---|---|
| Module files are named `mod<Name><NN>.bas`, where `NN` is the term number (`01`) | Each term has independent copies |
| `VB_Name` equals the file name without extension, except `modBuildPanel` and `modCollector` | Existing references depend on this |
| Source files are **pure ASCII** | Imports must work under any system locale |
| Arabic text is written as space-separated hexadecimal code points and decoded by the module's helper (`AR`, `ArT`, `ArHex`, `Txt`, `ArtsText`, `L`) | Same reason; also avoids code-page corruption |

Check for accidental non-ASCII characters before committing:

```bash
LC_ALL=C grep -nP '[^\x00-\x7F]' *.bas && echo "NON-ASCII FOUND" || echo "OK"
```

To produce a hex string for a new Arabic literal:

```python
print(" ".join(f"{ord(c):X}" for c in "ملاحظة"))   # paste into the module's hex helper
```

## 4. Coding standards

These are the conventions already used throughout the code. New code must follow them.

**Structure**
- `Option Explicit` at the top of every module.
- Constants at the top of the module; no unexplained literals in logic. (`Allowed()` in `modOrientation01` is a known exception, see `ARCHITECTURE.md` section 9, item 9. Do not add more.)
- Everything is `Private` unless it is an entry point or must be called by another module.
- One entry point per task, named `Run_<Task>01` (or an established name such as `CollectTerm01Results`). Orientation-workflow modules also expose `<Task>01Succeeded() As Boolean`, which `modBuildPanel` uses to decide whether the next step may run.

**Entry-point pattern** (see `ARCHITECTURE.md` section 7.1)

```vb
Public Sub Run_Example01()
    ' 1. validate environment: saved workbook, local path (not a URL)
    ' 2. remember Application state: ScreenUpdating, EnableEvents, DisplayAlerts, Calculation
    On Error GoTo EH
    ' 3. open / build in memory
    ' 4. validate EVERYTHING, collect errors, show at most the first 20
    ' 5. write, then save
CleanUp:
    ' 6. close only what this macro opened; restore the remembered Application state
    ' 7. one summary MsgBox
    Exit Sub
EH:
    ' record message, set failure flag, GoTo CleanUp
End Sub
```

**Error handling**
- Prefer `Err.Raise vbObjectError + n, , "<actionable message>"` and handle it in the entry point.
- Error-number ranges already in use: `1-6` environment/file, `44` dates, `55-59` orientation arithmetic and totals, `80-84` sheet/table structure, `701-702` clearing, `801-802` panel, `901-906` build/exam, `1000-1021` orientation results. Reuse the range of the area you are changing.
- Messages must say **what to do** ("Unprotect X before running."), not only what failed.
- `On Error Resume Next` is allowed only around a single statement whose failure is expected and handled on the next line (existence checks, optional cleanup). Never wrap a block, and always restore `On Error GoTo <handler>` afterwards.

**Files and workbooks**
- Never modify a workbook that is open read-only, already open in the session, or missing its expected sheet; skip and report it.
- Back up before modifying; save only after all steps for that file succeeded; otherwise close with `SaveChanges:=False`.
- Open external files with `UpdateLinks:=False`, `ReadOnly:=True`, `AddToMRU:=False`, and force-disable macros (`msoAutomationSecurityForceDisable`) wherever `.xls`/`.xlsm` files may be opened; restore the previous setting afterwards.
- Build paths with `Application.PathSeparator`. (Hard-coded `"\"` exists in places; do not copy it into new code.)

**Numbers and locale**
- Round with `WorksheetFunction.Round`; compare marks with the module's `EPS` tolerance, not `=`.
- Percentages of counts use `Decimal` arithmetic (`ShareCount`, `ShareCountUp`, `ShareCountNearest`) to avoid binary drift.
- Do not embed `CStr(<non-integer>)` in formulas; it is locale-dependent.
- Parse marks stored as text with `NormalizeMark`, which respects `Application.International`.

**Arabic and RTL**
- Normalise before comparing names or headers (`NormKey` / `NormAr` / `Norm`).
- Use `vbMsgBoxRight + vbMsgBoxRtlReading` for Arabic message boxes (as `modExamenAnalysis01` does).

## 5. Working with duplicated code

About 56 procedure names appear in more than one module. Until they are consolidated, **a fix in one copy must be applied to every copy**, or the divergence must be documented.

Find all copies and detect drift with this script (run it from the repository root):

```python
import re, glob, hashlib, collections
procs = collections.defaultdict(list)
for f in sorted(glob.glob("*.bas")):
    lines = open(f, encoding="latin-1").read().replace("\r", "").split("\n")
    i = 0
    while i < len(lines):
        m = re.match(r"^\s*(?:Public |Private |Friend )?(?:Static )?(Sub|Function)\s+(\w+)", lines[i])
        if m:
            kind, name = m.groups(); j = i
            while j < len(lines) and not re.match(r"^\s*End " + kind, lines[j]): j += 1
            body = "\n".join(l.strip() for l in lines[i:j+1]
                             if l.strip() and not l.strip().startswith("'"))
            procs[name].append((f, hashlib.md5(body.encode()).hexdigest()[:6]))
            i = j
        i += 1
for name, v in sorted(procs.items(), key=lambda x: -len(x[1])):
    if len(v) > 1:
        print(f"{name:30} copies={len(v)} variants={len({h for _, h in v})}")
```

Notes: trailing inline comments change the hash, so a "variant" can be a comment-only difference; read the diff before concluding behaviour differs. If you consolidate helpers into a shared module, do it in its own pull request with no behaviour change, and keep per-term independence (section 9).

## 6. Changing behaviour that affects pupils

Applies to: orientation rules, seat ceilings, pass mark, profile coefficients, age rule, subject conditions, statistics, matching rules.

1. Open an issue first describing the **pedagogical or administrative rule** (cite the circular or instruction if one exists), not just the code change.
2. Change the constant or rule in one place; update every duplicate.
3. Add or update test cases in `test_plan.md` with expected values computed independently (spreadsheet or by hand).
4. Update `ARCHITECTURE.md` (sections 4-6) and mark the `changelog.md` entry **(behaviour change)**.
5. In the pull request, include a before/after comparison on the synthetic cohort.

## 7. Branches, commits and pull requests

- Branch from `main`: `fix/<short-name>`, `feat/<short-name>`, `docs/<short-name>`.
- Commit messages follow [Conventional Commits](https://www.conventionalcommits.org/): `fix(modOrientation01): use pinned work year`. Use the module name as the scope.
- A pull request must include:
  - [ ] Description of the problem and the approach
  - [ ] Modules touched, and every duplicate updated (section 5)
  - [ ] ASCII check passed (section 3)
  - [ ] Applicable test cases from `test_plan.md` run, with results noted
  - [ ] `changelog.md` updated under **[Unreleased]**
  - [ ] `ARCHITECTURE.md` / user documentation updated if a contract, workflow or message changed
  - [ ] No real data in the diff, screenshots or attachments
- Review focuses on: fail-closed behaviour, file safety, rounding/locale, duplicates, and documentation.

## 8. Documentation requirements

| If you change... | Update |
|---|---|
| A workbook/file contract (columns, sheet names, table anchors, file names) | `ARCHITECTURE.md` section 4, `README.md` layout/templates, `user_guide_ar.md`, **MAJOR** bump in `changelog.md` |
| A user-visible message or button caption | `troubleshooting_ar.md`, `user_guide_ar.md` |
| A constant listed in the README configuration table | `README.md` |
| Any security-relevant behaviour (macro security, file handling, protection) | `security.md` |
| A known limitation (fixed or new) | `ARCHITECTURE.md` section 9, `changelog.md` known-issues table |

Messages shown to users are in English (the Examen module uses Arabic); Arabic end-user documents describe them by their exact English text so users can search for them.

## 9. Adding a new term or a new school

**New term (Term 02, Term 03).** Copy the modules, replace the `01` suffix and the term-specific names (`Term01Results.xlsx`, `Term01Guidance.xlsx`, `StreamPreferences01.xlsx`, `OrientationResults01.xlsx`, `Source01`, `ExamenSource01`, `Term01Panel`, hidden `_Term01_*` names, Arabic "first term" labels). Keep each term's copy independent. Do not make Term 02 call Term 01 code. Record the addition in `changelog.md`.

**New school.** Edit, in this order: the file map (`modResultFlow01`, `modClearDataOnly01`, `GetFileMap` in `modBiuldAnalysisfilesData01`), `CoreKeysFor`, the school/directorate text printed in the orientation report headers (`modOrientation01`), coefficient tables and cohort lists in `modExamenAnalysis01`, then the documentation.

## 10. Reporting issues

- Bugs and feature requests: open an issue with the module, the exact message text, the steps, and a **synthetic** reproduction.
- Security or privacy concerns: do **not** open a public issue; follow `security.md`.
