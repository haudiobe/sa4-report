# Issue #6 / PR #7: live integration test on a disposable copy of Audio

Gate before any release of the large-build checkpoints (`docs/LARGE_BUILD_CHECKPOINTS.md` §5 step 2). The original
Audio report, the master template and every other report stay untouched.

## Status (2026-11-12): BLOCKED — waiting on manual steps

| Step | State |
|---|---|
| Repository | `fix/audio-large-build-checkpoint`, in sync with origin; PR #7 open, mergeable, not merged. |
| Backup of the original's script (read-only) | **Done.** `clasp pull` of `1TJAxpo…`: Code.js, HyperLink.js, ReportCreator.js and Release.js are T-2026.10.22 byte for byte (`RELEASE-MANIFEST.json`); appsscript.json has the same content in Google's formatting. Apps Script API: parent = Audio document `1zawfI…`, last update 2026-11-11 21:30:26Z (the adoption). Kept locally in `dist/integration-test/audio-original-backup-20261112T070038Z/` (gitignored). |
| Drive metadata of the original (read-only) | **Done.** "Audio SWG Minutes SA4#138", version 41, created 2026-11-11 21:30:22Z, modified 22:29:47Z, description "SA4 report for meeting 60780 (Audio) … T-2026.10.22" (provenance only, no setup payload). |
| Disposable copy | **Blocked.** clasp's OAuth client has no access to document content (`files.copy` → 403 "has not granted the app … read access"). Needs a person (step M1). |
| Push of the candidate into the copy | Ready: `tools/integration-test/push-to-disposable-copy.mjs`. Needs the copy's IDs (M2). |
| Full Build, checks | Need the Docs UI (the build asks for confirmation through `DocumentApp.getUi()`) and the copy's own authorization (M3–M6). |

Nothing was written to any Google file. No credentials were created or printed.

## The kit (`tools/integration-test/`, not part of any release bundle)

- `push-to-disposable-copy.mjs` — refuses the original Audio script/document, the master template, the repository
  project and the protected projects; requires that the Apps Script API names the copy as the script's parent and
  that the Drive name starts with `DISPOSABLE TEST COPY`; backs the script up first and requires it to be
  T-2026.10.22 (a fresh copy of Audio); pushes Code.js, HyperLink.js, ReportCreator.js of the committed candidate,
  the copy's own appsscript.json, a Release.js that names the candidate commit and says NOT RELEASED, and
  `IntegrationCheck.js`; never `--force`; pulls back and compares SHA-256 of every file. Dry run unless `--write`.
- `IntegrationCheck.js` — run from the copy's editor; reads only (one Docs API read), refuses the original and the
  template. `it6Inspect()`, `it6Snapshot()`, `it6CompareLast()`: TDoc tables, doubled numbers, abstracts, dropdowns,
  text and empty statuses, statuses still to convert, the `STATUS_DROPDOWNS` mark, a digest of everything outside
  the Status cells, and every TDoc's status.
- `tests/integration-test-kit.test.js` — the guard, the candidate Release.js (accepted by the runtime's own
  validation), the summary and the diff.

## Manual steps

**M1. Make the copy** (Docs UI, as the owner). Open the original → *File › Make a copy*:
name `DISPOSABLE TEST COPY (issue #6, PR #7) - Audio SWG Minutes SA4#138 - delete after test`, folder *My Drive*,
**untick** "Share it with the same people" and "Copy comments". Do not edit or run anything in the original.

**M2. Send the two IDs** (not secrets): the copy's document ID (from its URL) and its Script ID (in the copy:
*Extensions › Apps Script › Project Settings › Script ID*). Then:

```
node tools/integration-test/push-to-disposable-copy.mjs --script <copy Script ID> --doc <copy document ID>
node tools/integration-test/push-to-disposable-copy.mjs --script <copy Script ID> --doc <copy document ID> --write
```

Expected: "Read-back MATCHES the bundle"; `dist/integration-test/<time>/evidence.json`.

**M3. Configure the copy.** Reload the copy. *SA4 Report › Configure Meeting…* (authorize the copy's script when
asked — it is a new project): the same settings as the original (meeting 60780, Audio, agenda S4-261857, the
shared SA4#138 TDoc list). To read them in the original, open its dialog and **Cancel**. Make sure the Reviewer API
token is available in the copy (your personal settings carry over; otherwise set it in the copy's dialog) — without
it the abstracts phase is skipped and the test is not representative. Do **not** turn on automatic updates in the copy.

**M4. Full Build #1.** In the copy: *SA4 Report › Report › Build Report from Scratch…* → Yes. Then in the editor:
*Executions* → this run → copy the whole log. Run `it6Snapshot()` (snapshot #1) and `previewStatusDropdownRepair()`;
copy both logs.

**M5. Hand-made changes.** In the copy, for two TDocs whose Status is a dropdown, choose another value by hand
(e.g. *parked*, *noted*); note the numbers. Type the line `IT6 minutes probe` into the minutes of one TDoc. Run
`it6Snapshot()` (#2).

**M6. Idempotence.**
1. *SA4 Report › Report › Update Report Now* → `it6Snapshot()` (#3) → `it6CompareLast()` (#2 → #3).
2. `repairStatusDropdownsAfterBuild()` → `it6Snapshot()` (#4) → `it6CompareLast()` (#3 → #4).
3. *Build Report from Scratch…* again (Full Build #2) → execution log → `it6Snapshot()` (#5) → `it6CompareLast()` (#4 → #5).

Send the logs of M4–M6.

## Pass criteria (all required for GO)

| Check | Expected |
|---|---|
| Full Build #1 execution | Status "Completed"; every `[FULLBUILD] <phase>` done (abstracts **not** skipped); no `Too many changes`; `[CHECKPOINT] … saved … and opened again` lines (the test predicts about 14 for 111 TDocs; any number ≥ 1 passes), each one followed by more work of the build. |
| Snapshot #1 | `tdocNumbers` 111, `tdocTableElements` 111, `duplicates` []; `abstracts` ≈ 100 (as the failed live build: 100); `definition` present; `stillToConvert` 0; `marked` true; dropdown + text statuses = 111, the text ones only *withdrawn*-like statuses with no option. |
| `previewStatusDropdownRepair()` after #1 | "to be made a dropdown: 0", no refusal. |
| #2 → #3 (Update) | no table added or removed, counts unchanged; the two hand-chosen values unchanged; changed statuses only where the Portal changed. |
| #3 → #4 (repair) | empty diff, `otherContentUnchanged` true (the repair writes nothing on a complete report). |
| #4 → #5 (Full Build #2) | 111 tables, none doubled, the two hand-chosen values kept; `stillToConvert` 0. The typed minutes are expected to be gone (a Full Build replaces the content; only the Reallocations table is kept). |

**If a checkpoint is still refused** ("Too many changes" at a `[CHECKPOINT]`): lower `LARGE_BUILD_CHECKPOINT_BUDGET_`
(one line and its test), push again with `--allow-other-base`, make a fresh copy, repeat.
**If a phase after a checkpoint fails on a closed document** ("Document is closed", "Service Documents failed"):
NO-GO — reopening a bound script's own document within one execution does not work live, and the design must change.

## Clean-up after the test

Trash the copy (its script goes with it). Delete `dist/integration-test/` only after the evidence is in PR #7.
