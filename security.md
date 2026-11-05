# Security Policy

This toolkit is a set of Excel VBA modules that run **locally** on a school administrator's Windows PC and process **personal data of minors** (names, birth dates, marks, orientation decisions). Its security concerns are therefore mainly **data protection, file integrity and macro safety**, not network exposure.

## Contents

1. [Supported versions](#1-supported-versions)
2. [Reporting a vulnerability](#2-reporting-a-vulnerability)
3. [Scope](#3-scope)
4. [What the code does and does not do](#4-what-the-code-does-and-does-not-do)
5. [Data protection](#5-data-protection)
6. [Macro and file safety](#6-macro-and-file-safety)
7. [Sheet protection is not security](#7-sheet-protection-is-not-security)
8. [Hardening checklist for operators](#8-hardening-checklist-for-operators)
9. [Hardening checklist for contributors](#9-hardening-checklist-for-contributors)
10. [Known security-relevant limitations](#10-known-security-relevant-limitations)

---

## 1. Supported versions

Only the latest commit on the default branch and the most recent tagged release receive fixes. Each term's copy (`...01`, `...02`) is supported equally.

## 2. Reporting a vulnerability

**Do not open a public issue** for security or privacy problems, and never attach real student data to any report.

Use GitHub's *Private vulnerability reporting* (**Security > Report a vulnerability**) on this repository. If it is not enabled, contact the maintainer privately at the address listed on the maintainer's GitHub profile.

> Maintainer: set the preferred private contact here before publishing: `<security contact>`.

Please include: affected module and procedure, the Excel/Windows version, clear steps using **synthetic** data, and the impact you expect. You will receive an acknowledgement; fixes are prioritised by impact on pupils' data and on the integrity of orientation decisions.

## 3. Scope

**In scope**

- Any path by which the code could disclose, corrupt, overwrite or delete workbooks or data beyond its documented folders.
- Execution of untrusted macros while the toolkit opens source files.
- Incorrect handling of personal data (for example leaving it in logs, backups or temp files unexpectedly).
- Weaknesses that allow silent alteration of orientation decisions or marks.

**Out of scope**

- Attacks that require an attacker to already control the user's Windows account or Excel installation.
- The security of Microsoft Excel, Windows, the national school platform or the files it exports.
- The empty-password sheet protection being removable (documented behaviour, section 7).

## 4. What the code does and does not do

Verified by reading all 16 modules:

| Capability | Status |
|---|---|
| Network access (HTTP, downloads, sockets) | **None found** |
| Shell or process execution (`Shell`, `WScript.Shell`) | **None found** |
| API `Declare` statements | **None found** |
| COM objects used | `Scripting.FileSystemObject`, `Scripting.Dictionary`, `VBScript.RegExp`, WMI `StdRegProv` |
| Registry access | **Read-only**, to detect installed fonts (`modCollector`, `...Data01`) |
| Environment variables read | `WINDIR`, `LOCALAPPDATA` (font folders) |
| File writes | Only inside the tool workbook's folder tree: `Source01\_Original`, `_Backup\<timestamp>`, the output folder, and the named result/target workbooks |
| File deletion | Not performed; originals are **moved** to `_Original`. `Collector` overwrites `Term01Results.xlsx` without asking; `Examen` and `Build_All` pick a new unused file name (`NewOutputPath` / `NextFreePath`); `OrientationResults` asks before replacing |

If you find any behaviour outside this table, treat it as a defect and report it.

## 5. Data protection

The data handled is personal data of minors. Operators are responsible for compliance with the regulations applicable to their institution.

**Repository**

- Never commit workbooks with real data. Use a `.gitignore` that excludes `*.xlsx`, `*.xlsm`, `*.xls`, `*.xlsb`, `Source01/`, `ExamenSource01/`, `_Backup/`.
- If real data was committed, **rewriting history is required** (for example `git filter-repo`), followed by rotating any exposed material; deleting the file in a new commit does not remove it from history.

**On the working machine**

- `_Backup\<timestamp>\` folders contain full copies of target workbooks and are **never pruned** by the toolkit. Treat them with the same sensitivity as the originals and delete them when no longer needed.
- `Source01\_Original` keeps the original platform exports.
- `Term01Results.xlsx`, `Term01Guidance.xlsx`, `OrientationResults01.xlsx`, `BuildLog_*.txt` and the build output folder contain pupil data (the log may contain file names and row counts).
- Store the working folder on an encrypted disk, restrict it with NTFS permissions, and avoid synchronising it to a personal cloud account. The tools themselves refuse to run from OneDrive/SharePoint URL paths for reliability reasons; that is not a privacy control.

## 6. Macro and file safety

- While opening `.xls`, `.xlsb` or `.xlsm` source files, the code sets `Application.AutomationSecurity = msoAutomationSecurityForceDisable` (verified in `modCollector`, `modResultFlow01`, `modClearDataOnly01`, `modTransfer01`, `...Data01`) and restores the previous value afterwards. This blocks `Auto_Open`/`Workbook_Open` code in untrusted exports.
- Source files are opened with links not updated and read-only where applicable.
- Converting a macro workbook to `.xlsx` **removes its macros** and archives the original.
- `modExamenAnalysis01` and the guidance modules open `.xlsx` workbooks only.
- Distribution/clearing modify target workbooks only after a successful backup, and save only after all steps for that file succeeded.
- Open files from untrusted senders with Excel's Protected View first; do not place unknown files in `Source01` or `ExamenSource01`.

## 7. Sheet protection is not security

The guidance workbook's sheets are protected with an **empty password** (`Protect Password:=""`). This prevents accidental edits by operators; it does **not** prevent a deliberate user from unprotecting the sheet, and offers no confidentiality. Do not rely on it to restrict access to pupil data or to prevent tampering with orientation decisions. The decision column is intentionally left editable after orientation.

For integrity, keep the dated backups and the exported `OrientationResults01.xlsx`, and treat the final signed/printed lists as the record of decision.

## 8. Hardening checklist for operators

- [ ] Run the toolkit on a patched Windows/Excel installation under a standard (non-administrator) account.
- [ ] Keep macro security at *Disable all macros with notification*; enable macros only for the tool workbook, ideally via a Trusted Location limited to the working folder.
- [ ] Use a dedicated working folder, not shared with other staff or synchronised publicly.
- [ ] Verify the tool workbook's source (code reviewed from this repository) before enabling macros.
- [ ] After each run, inspect the summary dialog and keep the `BuildLog_*.txt` and `_Backup` folder only as long as needed.
- [ ] Delete `_Backup` and `_Original` contents at the end of the school year according to your retention rules.
- [ ] Never email or upload the working folder; share only the exported final lists to authorised recipients.

## 9. Hardening checklist for contributors

- [ ] No network, shell or `Declare` usage without a documented, reviewed justification in the pull request and an update to section 4.
- [ ] Every new code path that opens an external workbook forces macros disabled and restores the previous setting.
- [ ] No new writes outside the tool workbook's folder tree; no deletions (move to an archive instead).
- [ ] Every new output that contains pupil data is listed in section 5.
- [ ] Source remains pure ASCII (`LC_ALL=C grep -nP '[^\x00-\x7F]' *.bas`), so no hidden or bidirectional-control characters can disguise code.
- [ ] No credentials, contact details of pupils or personal paths in code, comments or test data.

## 10. Known security-relevant limitations

| Area | Limitation |
|---|---|
| Archived originals | Macros of `.xlsm` originals are preserved in `_Original`; opening them manually re-enables the risk |
| Backups | Unencrypted and unbounded growth |
| Logs | `BuildLog_*.txt` is plain text |
| Overwrite | `modCollector` replaces `Term01Results.xlsx` without a confirmation prompt (a previous version is not kept) |
| Hard-coded identity | The orientation report headers contain the school and directorate names as literals; review before sharing the code or outputs outside the institution |
| Integrity | No checksums or signatures on outputs; protection uses an empty password |
| Audit trail | No record of who ran which step; workflow stamps record file size and modification time only |
