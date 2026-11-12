# Large-build checkpoints and the status-dropdown repair (after T-2026.10.22)

Branch `fix/audio-large-build-checkpoint`, on master `fd0621b` (T-2026.10.22, `Code.js` 2.26.6). GitHub issue #6;
evidence and safety gates in `docs/AUDIO-LARGE-BUILD-RECOVERY.md`.

**Status: released as T-2026.10.23 (`Code.js` 2.26.7, issue #8).** PR #7 was squash-merged as `cbf12a3`; the
candidate `6267011` had been tested live in the Audio report's bound script (issue #6). The release commit adds only
the version line 2.26.7 and its changelog entry; the ledger takes them out again. Deployment record: issue #8.

## 1. The failure

Audio SWG SA4#138 Calgary, Full Report Build with T-2026.10.22: 10 agenda items, 111 TDocs, 100 abstracts. Every
phase logged completion; the status-dropdown phase logged Google's error `Too many changes applied before saving
document. Please save changes in smaller batches using Document.saveAndClose(), then reopen the document with
Document.openById().`; the execution ended in error.

**Where it threw.** A Full Build of T-2026.10.22 never saves until its last phase:
`finalizeStatusDropdowns_()` calls `saveAndClose()` once, after skeleton, e-mail, revisions, abstracts and formatting
have made all their DocumentApp changes. That one save carried every change of the build and was refused. The step
catches every error, so the phase was reported as *skipped* with Google's message as its reason, and the build
returned "ok". At the end of the execution Apps Script tried to save the same pending changes and failed again:
that is the execution error. Nothing in a `catch` or `finally` saved a second time; the second attempt is Apps
Script's own.

So the logged "completed" phases are not evidence that anything is in the document. What persisted in Audio must be
inspected (section 6, step 1). The 50-per-batch Docs API writes were never reached: the failure is the DocumentApp
save before them, and the existing 130-TDoc test did not model a limit on unsaved DocumentApp changes.

The test reproduces this with the released `Code.js` of T-2026.10.22 (section 4, A): all phases "done", the
dropdown phase "skipped" with Google's message, no save, and the end-of-execution save refused.

## 2. What changes

### Checkpoints in a Full Build (`LARGE BUILD CHECKPOINTS` section of `Code.js`)

- The build estimates its pending DocumentApp changes (`LARGE_BUILD_UNITS_`: 40 per skeleton TDoc table, 6 per
  registration row, 8 per formatted row, 6 per abstract row, 2 per row of column widths). When about
  `LARGE_BUILD_CHECKPOINT_BUDGET_` (2000) are pending it makes a **checkpoint**: `saveAndClose()`, then
  `DocumentApp.openById()` of the same document (three attempts, 0/1/2 s apart), as Google's message asks.
- Checkpoints are made inside the loops that make most changes -- the TDoc tables of the skeleton, the abstracts,
  the column widths, the formatting -- and at the end of a phase. Once a build has made one checkpoint, it saves at the end
  of every following phase as well.
- **A small build makes none** and saves exactly as T-2026.10.22 did (one save, in the status-dropdown step; its
  log lines are unchanged). Nine TDocs: no checkpoint. 111 TDocs (the test): 14, none above about 2000.
- Checkpoints exist only inside `runFullReportBuildCore_()`. Updates, menu steps and background runs are unchanged.

### Fresh references

A closed document can no longer be changed, and neither can any element read from it. Checkpoints were added only
after making every later access fresh:

- Every code use of `DocumentApp.getActiveDocument()` in `Code.js` (44) now calls `currentReportDocument_()`:
  the document the last checkpoint reopened, else -- always, outside a large build -- the active document. The
  context-based accessors (`getReportDocument_()` and everything built on it) fall back to it too.
- Each phase takes its body when it starts (unchanged), so a phase after a checkpoint works on the reopened document.
- Inside a loop that checkpoints, the loop takes its body (`largeBuildStep_()`, skeleton) or its tables
  (`largeBuildTablesIfDue_()`, abstracts, column widths and formatting) again from the reopened document; the
  formatting also takes its tables again after the column-widths pass, which may have made one. Those loops add or remove
  no table, so the tables are matched by position; if the reopened document has another number of tables the build
  stops instead of writing into the wrong one.
- The skeleton's `body` became a `let`; nothing else in the skeleton keeps an element across a checkpoint.

### Error reporting

- A checkpoint whose save is refused, or whose document cannot be opened again, **fails the phase**, so the build
  stops. The message names the checkpoint, Google's message, and the last checkpoint that was saved (or that nothing
  of the build was saved), and says the report is incomplete.
- A refused save in the status-dropdown step now **fails the build** in that step (`saveFailed`), instead of being
  reported as skipped. In an update the step still never throws; the result carries `saveFailed`.

### The repair of an existing report (no rebuild)

`previewStatusDropdownRepair()` (dry run) and `repairStatusDropdownsAfterBuild()`, run from the Apps Script editor,
finish what the status-dropdown step of a build would have done, on the document as it is:

1. one Docs API read of the report; refused, with nothing changed, when it has no TDoc table, when a TDoc number has
   more than one table, when the Docs API is not available or the template has no "Document Status" definition;
2. the definition is created from the template if the report has none;
3. every Status that is plain text and is a Portal status with an option becomes its dropdown, by the same rules and
   the same revision-guarded batches as the build (`statusDropdownHealCandidates_()`,
   `planStatusDropdownRequests_()`, `sendStatusDropdownInsertPairs_()`). *parked*, *Plenary*, *other*, statuses
   without an option and empty cells stay text;
4. the report is read again: every converted cell must be its dropdown, every other Status cell and everything
   outside the Status cells (`statusDocsOtherContent_()`) unchanged. Only then is the report marked
   (`STATUS_DROPDOWNS`).

DocumentApp changes nothing; nothing is downloaded, cleared or rebuilt; no table is added; minutes, abstracts and
existing dropdown values are kept. A second run finds nothing to do and writes nothing. A run cut short leaves every
cell whole (a dropdown or its text) and the next run finishes it. It holds the document lock an update holds. It
writes its result to the log as `[REPAIR] …` and returns the same text.

It differs from *Convert Status Fields to Dropdowns…* (T-2026.10.8) in purpose and rule: the conversion turns a
manually kept report's statuses into dropdowns and treats every option name as a decision; the repair completes a
build, so a status is mapped as the build maps a Portal status, and a report that does not look built is refused.

## 3. Production diff against T-2026.10.22

| File | Change |
|---|---|
| `Code.js` (stays 2.26.6 until the release) | New section `LARGE BUILD CHECKPOINTS` after `TDOC LIST DOWNLOAD` (14 functions, 4 top-level settings/state). Changed: `runFullReportBuildCore_`, `buildSkeletonWithTdocTables`, `addAbstractsForTables_`, `removeRowHeightAndSpacing`, `setTwoColumnTDocTableWidths_`, `finalizeStatusDropdowns_`. Every code use of `DocumentApp.getActiveDocument()` (44) → `currentReportDocument_()`. |

`appsscript.json` (no new service or scope), `HyperLink.js`, `template/ReportCreator.js` (no menu change) and the
release tooling are unchanged.

## 4. Tests

`tests/large-build-checkpoints.test.js`, against a fake report that models Google's limit on unsaved changes, a
closed document, and `openById()` giving a new document whose old elements throw on any use:

- **A.** the live failure with the released `Code.js` of T-2026.10.22: build "ok", dropdown phase skipped with
  Google's message, no save, end-of-execution save refused, no dropdown;
- **B.** the same 111-TDoc build: completed, every save under the limit, checkpoint = save + `openById`, nothing
  unsaved at the end, no checkpoint above about the budget; 111 tables once each, every abstract once, the e-mail
  phase, the column widths and the formatting writing into the reopened document, all statuses right, report marked; a smaller budget checkpoints inside the abstracts loop too; a
  rebuild doubles nothing and keeps a value chosen by hand;
- **C.** small builds: no checkpoint, one save, the released log line; without the Docs API: no save, as before;
- **D.** a refused checkpoint save, a later refused one, a document that cannot be reopened (and one that opens on
  the third attempt), a refused save in the dropdown step (build fails, no longer skipped), a reopened document with
  a missing table;
- **E.** an update of a large report makes no checkpoint; the helpers are inert outside a build; only
  `currentReportDocument_()` calls `DocumentApp.getActiveDocument()`;
- **F.** the repair on the report a failed build leaves (111 text statuses, minutes typed by hand): dry run reads
  only; repair creates the definition, converts 95, leaves 16 "not treated" as text, keeps minutes and abstracts,
  marks; second run writes nothing and keeps a value chosen by hand; a refused batch heals on the next run; refusals
  (doubled table, no TDoc table, no Docs API, no definition in the template, lock held);
- **G.** ledger: without the layer, `Code.js` is T-2026.10.22 byte for byte; only the six functions above differ
  beyond the document access; the new functions are the section's.

A mutation check was run by hand: dropping the reassignment of the skeleton's body, of the column-widths or the
formatting loop's tables, or the formatting's refresh after the widths, each makes the test fail (stale elements
are used).

Existing tests: ledger layer `LARGE BUILD CHECKPOINTS` in `tests/helpers/before-structure-fix.js` (taken out before
the TDoc-list layer, so every earlier ledger still compares with its release); the harnesses' `openById()` reopens
the report; the fake document counts its unsaved changes (addition only); source-pinned checks updated where the
text moved (`let`, two places that save, a third place that marks, a fourth user of the batch sender); the
status-dropdown test now expects a failed save to fail the build.

**At `e014f5d`: 118 test files, 7,982 checks, 0 failures.** A check is an output line starting with `ok` or
`PASS`; failures are lines starting with `FAIL` or `not ok`, and every file must exit 0. An earlier draft of this
line said 7,980: it was counted before the column-widths checkpoint added two checks to
`tests/large-build-checkpoints.test.js` (the checkpoint-list check became two, and "the widths every row of every TDoc
table" is new); the PR said 7,982. Two runs of the committed tree give identical per-file counts.
With the integration-test kit (`tests/integration-test-kit.test.js`, 33 checks): see `docs/INTEGRATION-TEST-ISSUE6.md`.

## 5. Release preparation (Master Template)

Proposed: the next template release after T-2026.10.22, `Code.js` 2.26.7. Issue #4 (meeting time zones,
`feature/meeting-local-time-zones`) is also a candidate after T-2026.10.22 and is kept separate; whichever is
released second is rebased onto the first and its ledger layer chained after it.

Before the release:

1. Review and approve the PR. Do not merge before.
2. **Validate on a disposable copy of the Audio report** (never the original). Kit, runbook, pass criteria and
   current state: `docs/INTEGRATION-TEST-ISSUE6.md` (status 2026-11-12: blocked on making the copy by hand).
   - a Drive copy of Audio, bound script set to this candidate (adoption to the copy only);
   - one Full Report Build on the copy: the log shows `[CHECKPOINT]` lines and every phase done; no
     `Too many changes`; the execution "Completed";
   - on the copy: 111 TDoc tables, none doubled; abstracts present; status dropdowns present; values as the Portal's;
   - a second Full Build on the copy, and one Update Report Now: no doubled table, values chosen by hand kept;
   - if a checkpoint is still refused: lower `LARGE_BUILD_CHECKPOINT_BUDGET_` (a one-line change with its test) and
     repeat.
   This is the first live proof that a bound script can keep editing its own document after `saveAndClose()` and
   `openById()` in one execution (TEMPLATE-002C §10 called this unproven; the status-dropdown step only proved
   save-then-Docs-API).
3. Merge. Release commit: version line 2.26.7 and its changelog entry, the ledger header pair, version-pinned tests.
   Tag `template-release/<next>`.
4. Bundle: `node tools/template-release.js --release <next>` (dry run), then `--write`. Only `Code.js` and
   `Release.js` differ from T-2026.10.22.
5. Master template: guarded push without `--force`; read-back matching the SHA-256 hashes of
   `RELEASE-MANIFEST.json`.

## 6. Audio SWG SA4#138: recovery (after the release, with approval)

1. **Back up** the Audio document (Drive copy, kept) and pull its bound script read-only; record the revision ID,
   size, script files and SHA-256 sums.
2. **Inspect** what persisted: the number of TDoc tables (expected 111), doubled numbers, abstracts, Status cells,
   whether a "Document Status" definition exists, `STATUS_DROPDOWNS`.
3. Adopt the release into the Audio bound script (`--adopt <next> --label audio-138`): dry run, push without
   `--force`, read-back against the manifest.
4. If the build content is there: run `previewStatusDropdownRepair()` from the editor and compare its counts with
   step 2; then `repairStatusDropdownsAfterBuild()`; then `previewStatusDropdownRepair()` again (expected: nothing to
   make a dropdown of). Minutes typed since are kept.
5. If the build content did **not** persist (the pending changes were lost): the repair refuses or finds little;
   a Full Build is then needed. It clears the document and is a separate decision, first on a copy (section 5,
   step 2), and only with approval on the original.

Other reports (MBS, 6G, …) are not touched by any of this.

## 7. Release and deployment of T-2026.10.23 (2026-11-12, issue #8)

| Step | Result |
|---|---|
| Live state before | **Master template** (`17xPMiY9…`, parent "SA4 Report Template"): last update 2026-11-11 20:25:31 UTC; a read-only pull gave five files byte-identical to the T-2026.10.22 bundle. **Audio** (`1TJAxpo…`, parent "Audio SWG Minutes SA4#138"): last update 2026-11-12 09:20:44 UTC (the guarded push of candidate `6267011`); six files byte-identical to that push's read-back. Neither had changed in between. Issue #4 (time zones, PR #5) is not merged and not deployed. |
| Code | PR #7 squash-merged as `cbf12a3`; its tree is identical to the candidate `6267011` tested live in Audio. |
| Release commit | `edbbdd213f8dfffea7d6a1d455d33abdedda273c`: version line 2.26.7 (2026-11-12) and its changelog entry, the ledger pair in `LARGE_BUILD_CHANGES`, version-pinned tests. Without the header, `Code.js` is the `Code.js` of `cbf12a3`. Tests: 119 files, 8,032 checks, 0 failures. |
| Tag and publication | Lightweight tag `template-release/T-2026.10.23`; `origin/master` fast-forwarded `cbf12a3` → `edbbdd2`, without force. |
| Release bundle | `--release T-2026.10.23`: dry run, then `--write`. Only `Code.js` and `Release.js` differ from T-2026.10.22. SHA-256: `appsscript.json` `c5ac395c…`, `Code.js` `3f9e719e…`, `HyperLink.js` `aeb79994…`, `ReportCreator.js` `c74d3ff0…`, `Release.js` `43faa6f4…`. Compared with the `Code.js` tested live in Audio, only the version line and the changelog entry differ. |
| **Master template: backup** | The five pulled files with `SHA256SUMS`, kept locally in `dist/backups/backup-template-before-T-2026.10.23/` (gitignored). Re-pulled immediately before the push: unchanged. |
| **Master template: push** | 2026-11-12 10:47:32 UTC, from its own staging folder, without `--force`; the staged files matched `RELEASE-MANIFEST.json`. |
| **Master template: read-back** | A fresh pull: exactly five files, each matching its SHA-256 in `RELEASE-MANIFEST.json` and byte-identical to the bundle. `Release.js`: T-2026.10.23 / 2.26.7 / `edbbdd2` / `template-release/T-2026.10.23`. |
| **Audio: backup** | Six pulled files (with the test-only `IntegrationCheck.js`) and `SHA256SUMS` in `dist/backups/backup-audio-138-before-T-2026.10.23/`; re-pulled immediately before the push: unchanged since 09:20:44 UTC. The content backup "BACKUP before issue #6 test - Audio SWG Minutes SA4#138 - do not edit" (`1Ik9pp…`) stays. |
| **Audio: adoption** | `--adopt T-2026.10.23 --label audio-138`: dry run, then `--write`. `Code.js`, `HyperLink.js`, `ReportCreator.js` and `Release.js` were staged byte for byte from the adoption bundle. `appsscript.json` stayed Audio's own (`4c2cc5ea…`): the same content as the release's (same keys and values, scopes in the same order), in Google's formatting. Its bytes differ from the release's, so pushing the release's would have been a manifest overwrite that needs `--force`; that was not needed and not done. |
| **Audio: push** | 2026-11-12 10:50:17 UTC, without `--force`. The test-only `IntegrationCheck.js` was removed (it is in the backup). |
| **Audio: read-back** | A fresh pull: exactly five files. Four match `RELEASE-MANIFEST.json` and the adoption bundle byte for byte; `appsscript.json` is byte-identical to Audio's own from before the push. `Release.js`: T-2026.10.23 / 2.26.7 / `edbbdd2`. |
| Not done | No Full Build, no report function, no trigger, no Document Property. The Audio document's content was not changed by this release. MBS SA4#138, MBS 86172, 6G and every other report untouched; the repository-root `.clasp.json` unused. |

**Rollback.** Master template: push `dist/backups/backup-template-before-T-2026.10.23/script` (or the
T-2026.10.22 bundle) from its own staging folder with the template's `.clasp.json`, without `--force`, then pull and
compare. Audio: the same with `dist/backups/backup-audio-138-before-T-2026.10.23/script` and the Audio Script ID.
