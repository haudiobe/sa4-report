# TEMPLATE-002B: first release candidate of the SA4 Report Template runtime

2026-10-01. Branch `template-002b/runtime` (worktree `sa4-report-template-002b`), based on
`template-002a/legacy-parity` `7d1a15b`. Local implementation and automated tests only: nothing
was pushed, deployed or run against Google, and CENTRAL and Legacy were not touched.

**Status: ready for a live smoke test (§10). Not deployed.**

---

## 1. What is in this branch

| From | How | Commits |
|---|---|---|
| `template-002a/legacy-parity` | branch point | everything up to `7d1a15b` (Legacy parity, `Code.js` 2.16.0) |
| `design/report-template-architecture` | `git cherry-pick -x`, one by one, no merge | `6dcff3f` → `3e886ce`, `21c1a3f` → `43123c0`, `fb13147` → `64a319c`, `20fbf59` → `7f95a7e` |
| TEMPLATE-002B | new | `40b8589` runtime, `27dbd5e` tests, plus this documentation commit |

The four design commits only add files (architecture document, `template/`, `tools/`, three
test files and a helper), so they applied without conflict. Each cherry-picked commit names
its original in its message.

Files changed by TEMPLATE-002B itself:

| File | Change |
|---|---|
| `Code.js` | Template runtime hooks; version 2.17.0 |
| `template/ReportCreator.js` | Prototype turned into the runtime file |
| `tools/template-release.js` | Records the `Code.js` version in `Release.js` |
| `tests/template002b-runtime.test.js` | New, 125 checks |
| `tests/template-bootstrap.test.js` | Follows the final function names |
| `tests/legacy-parity-upgrade-00{2,3,4}.test.js`, `tests/template002a-intentional-differences.test.js` | Labels and one check updated for the start-date decision (§5) |
| `docs/` | This document; pointers in the other two |

---

## 2. Source of truth and release model

- **One production `Code.js`.** The template bundle contains it byte-identical (tested).
  There is no template copy of it.
- **A bundle adds two files:** `Release.js` (generated; defines `SA4_RELEASE_`) and
  `ReportCreator.js`. `Code.js` asks `templateRuntimeRelease_()` at call time. Without
  `Release.js` (the CENTRAL add-on, a Legacy copy, the tests) every template hook is inert.
  Load order does not matter (tested with `Release.js` loaded last).
- **Template-only behaviour that must never reach CENTRAL lives in `ReportCreator.js`**: the
  15/30-minute trigger installer, the creator, the first run, the template menus. `Code.js`
  itself contains no sub-hourly trigger call (an existing test asserts this).
- **The master template receives a tagged snapshot.** `tools/template-release.js` refuses a
  dirty tree, an untagged HEAD, the Legacy and CENTRAL script IDs, test files, and Node-only
  code. It never runs clasp.
- **A report is pinned** to the snapshot the template held when it was created. Later template
  releases change the master only. There is no automatic update of existing reports.

---

## 3. Master template, created report, ordinary document

| Kind | How it is recognised | What it can do |
|---|---|---|
| Master template | The document ID equals `SA4_RELEASE_.templateDocumentId` | Create reports; show release information. Build, update, trigger installation and configuration are refused. |
| Report created from the template | Any other document that had the creator's setup information; marked `SA4_BOOTSTRAP_STATE = done` after its first use | Everything a report does |
| Ordinary document that contains the code | Any other document without setup information; marked `SA4_BOOTSTRAP_STATE = manual` after the first check | Everything a report does; configured by hand, as before |

The role comes from the document ID in the code, never from inherited properties (TEMPLATE-001:
a copy starts with empty property stores). The creator refuses to run outside the master, both
in the dialog and in its server calls.

---

## 4. Menus

**Master template** (`⚠️Scripts⚠️`):

- 🆕 Create New SA4 Report
- ℹ️ Template Release Info

**Created report** (`⚠️Scripts⚠️`): the existing menu, with these differences only:

- 🚀 Finish Report Setup, first item, until the report has been set up
- 📝 INITIAL SETUP, 🚀 REPORT OPERATIONS (first item ▶️ Run Full Report Build),
  📋 DOCUMENT MANAGEMENT, 🔧 TOOLS & DIAGNOSTICS, 🎨 FORMATTING & FIXES, 📧 EMAIL EXPORT:
  unchanged
- no ☁️ CENTRAL ADD-ON submenu
- 🗄️ LEGACY (old workflow): "Legacy: Build Initial Report" and "Legacy: Update All" moved
  here, out of the normal workflow
- ℹ️ About This Report, last item

No existing report operation was dropped (tested against the add-on menu). The add-on menu is
unchanged without `Release.js` (tested).

---

## 5. Email Collection Start Date (template decision, 2026-10-01)

Restored as a normal Configure Meeting field **in the template runtime**. It is no longer dead
configuration: since TEMPLATE-002A the collector reads `EMAIL_START_DATE`.

| Requirement | Implementation |
|---|---|
| Visible and editable in Configure Meeting | Date input in "3. Options" |
| One property | The existing `EMAIL_START_DATE`; no new key |
| Default: the meeting start date | The creator puts the resolved start date into the setup information; in the dialog an empty field follows the resolved start date on *Resolve* |
| An explicit stored value is preserved | Shown as "Saved value"; neither *Resolve* nor the first run replaces it |
| The user may move it earlier or later | A typed date survives *Resolve* and is saved |
| Validation | One rule shared with the collector (`isValidCollectorStartDate_`: a real `YYYY-MM-DD`); checked before anything is written; a blank value never erases a stored one |
| Existing report without the property | The dialog proposes the meeting start date; the collector falls back to it, else to `2026-08-21` |

A Collector Configuration table row `EMAIL_START_DATE` still wins, as before.

**Outside the template runtime nothing changed.** The CENTRAL dialog still has no such field
and its collector fallback is still `2026-08-21` (ADDON-007B1 and TEMPLATE-002A, both still
tested). Because of that, five TEMPLATE-002A checks that called this "template" behaviour are
now labelled "add-on runtime", and one source-text check ("the code never writes
`EMAIL_START_DATE`") was replaced by "it is written only behind the template-runtime switch".
No Legacy parity assertion was changed.

---

## 6. Create New SA4 Report

1. Open "SA4 Report Template", choose **Create New SA4 Report**.
2. Enter the meeting ID, **Look up**. The existing discovery pipeline runs unchanged
   (resolve, TDoc list, ad-hoc sources, family inference). The dialog shows meeting, family,
   mailing list, document folder, TDoc list, agenda and the e-mail collection start date.
3. Choose the report family only if asked (main meetings; ambiguous ad-hoc meetings).
4. **Create Report**: `DriveApp.makeCopy(title, same folder)`, then the setup information is
   written to the copy's Drive description. If that write fails, the copy goes to the trash.
5. **Open the new report** (link in the dialog).
6. In the report, the first of *Finish Report Setup*, *Configure Meeting* or *Run Full Report
   Build* stores the setup information as ordinary Document Properties, through the same code
   as a manual Save.
7. **Run Full Report Build**. If a source is still missing (agenda not published yet), the
   build guard says so; *Configure Meeting → Discover* later.
8. Optionally **Manage Auto-Update Trigger**.

Only bootstrap fields are transferred: meeting identity, family, sources, mailing-list
override and the start date. No CENTRAL deployment, Apps Script API, Advanced Drive service,
manual script copying or separate creator project is involved.

---

## 7. Continuous Update trigger (template runtime)

| Requirement | Implementation |
|---|---|
| 15 / 30 / 60 minutes | `everyMinutes(15)`, `everyMinutes(30)`, `everyHours(1)`; never `everyMinutes(60)` |
| Default | 30 minutes |
| Validate before changing anything | An interval that is not offered is refused before ScriptApp is touched |
| Never lose a working trigger | The new trigger is created first; only then are the previous `continuousUpdate` triggers removed. If creation fails, the old trigger, the recorded interval and the abstracts switch stay as they were |
| Idempotent | Repeating the request leaves exactly one trigger |
| Show the interval | Recorded in `CONTINUOUS_UPDATE_INTERVAL_MINUTES`; the dialog shows "Running every: 30 minutes" and preselects it |
| Disable | *Stop Trigger* removes only `continuousUpdate` triggers and forgets the interval |
| Unrelated triggers | Never touched |
| Nothing inherited | A copy has no triggers (TEMPLATE-001); the master cannot own one (guard); the probe is not shipped |

The add-on runtime keeps its hourly-only rule, with the existing tests unchanged.

---

## 8. Version

`Code.js` **2.17.0**. The repository's convention is `MAJOR.MINOR.PATCH` in the `Code.js`
header with a changelog entry per change: new functionality raises the minor number (2.14.0
agenda discovery, 2.15.0 exporter, 2.16.0 Legacy parity), fixes raise the patch number.
TEMPLATE-002B adds functionality to `Code.js`, so it is the next minor version.

The template release ID (`T-2026.10.0`) is a separate, deliberate label for one snapshot
installed into the master template. `Release.js` records both, and *About This Report* shows
them together: "T-2026.10.0 (Code.js 2.17.0, commit …)".

---

## 9. Tests and release-candidate review

Complete suite: **76 test files, 3,899 checks, 0 failures.**

| Group | Files | Checks |
|---|---|---|
| Existing `sa4-report` tests | 62 | 3,224 |
| Legacy parity suites (TEMPLATE-002A) | 9 | 408 |
| TEMPLATE-002A intentional differences | 1 | 20 |
| Template: bootstrap, release tool, probe | 3 | 122 |
| Template runtime (TEMPLATE-002B) | 1 | 125 |

| Question | Answer |
|---|---|
| A. All TEMPLATE-002A parity behaviour present? | **Yes.** The branch is built on `7d1a15b`; the 408 parity checks pass unchanged, with the same 16 skipped checks. |
| B. Exactly one production `Code.js`? | **Yes.** One file; the bundle carries it byte-identical. |
| C. Template vs created report told apart safely? | **Yes** in tests: by document ID; guard on build, update, trigger and configuration; creator refuses outside the master. |
| D. Report creation without CENTRAL, add-on, API or manual script copy? | **Yes.** `DriveApp.makeCopy` plus the Drive description only. |
| E. 15/30/60-minute triggers implemented safely? | **Yes** in tests. Live: creation of a 30-minute trigger in a copy was verified by TEMPLATE-001; its first execution was not waited for. |
| F. Email Collection Start Date a functional normal field? | **Yes**, in the template runtime. |
| G. Anything still needing manual Google validation? | **Yes**, see below. |

Still unverified until the smoke test (none of this code has run in Google yet):

- the menus under the real simple-trigger rules (`onOpen` reading the document ID and its
  own properties without authorization);
- the creator dialog and its server calls from a real dialog, including how long the lookup takes;
- `makeCopy` of the real template (a ~680 KB bound script) and the permission prompt in the copy;
- the first use of the copy (setup information read from the real Drive description);
- a real Full Build in a copy, and the first trigger execution in a copy;
- `clasp push` of the bundle into a new bound project (manifest included).

---

## 10. Proposed live smoke test

Preparation, once (about 10 minutes):

1. Create an empty Google Doc named **SA4 Report Template**. Open Extensions → Apps Script
   once and copy the Script ID (Project Settings). Copy the Doc ID from the URL.
2. In this branch: save both as `template/template-target.json` (see the example file),
   commit it, then `git tag template-release/T-2026.10.0`.
3. `node tools/template-release.js --release T-2026.10.0 --write`, then run the printed
   `cd … && clasp push` from the bundle folder. This is the only push; it goes to the new
   template project and nowhere else.

Smoke test (about 10 minutes):

| # | Action | Expected |
|---|---|---|
| 1 | Reload the template | Menu with exactly: Create New SA4 Report, Template Release Info |
| 2 | Template Release Info (allow the permission prompt, click again) | "T-2026.10.0 (Code.js 2.17.0 …)" |
| 3 | Create New SA4 Report → a current meeting ID → Look up | Family, lists, sources and start date shown; Create enabled |
| 4 | Create Report → Open the new report | A new document with the report title |
| 5 | In the report: check the menu | Report menu, Finish Report Setup first, no CENTRAL submenu |
| 6 | Run Full Report Build (allow the prompt, click again, confirm) | The report is built |
| 7 | Configure Meeting | Pre-filled; Email Collection Start Date = meeting start date |
| 8 | Manage Auto-Update Trigger → Start (30 minutes); open it again | "Running every: 30 minutes"; then Stop |
| 9 | Back in the template: Extensions → Apps Script → Triggers | No trigger; the template body is unchanged |

Pass: all nine rows as expected. If the report is a scratch report, trash it afterwards. The
first trigger execution can be observed later on the first real report.
