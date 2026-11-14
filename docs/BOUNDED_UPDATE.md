# Bounded update: Update Report Now within the Apps Script time limit (T-2026.10.24, Code.js 2.26.8)

Status: **fix branch `fix/bounded-update`, on `release/T-2026.10.24` (`8edb1a7`). Not merged, not released, not deployed.** GitHub issue #9 (integration test U5).

## What happened

On 2026-11-13, during U5, the SA4#138 copy was set to `adhoc` to test the sessions. An ad-hoc update takes the TDocs of all agenda items. Update Report Now started at 13:24:41 (log time):
- it downloaded and parsed 392 TDocs, in 45 agenda items;
- it found 112 TDocs already in the report;
- it began inserting the other 280, each with its abstract.

At 13:30:39 it failed with `Exceeded maximum execution time`. The last TDoc it worked on was S4-262161.

Snapshots #3 and #4 (before and after) are identical: 114 tables, 111 TDoc numbers, 100 abstracts, 111 dropdowns, no duplicates. `it6CompareLast()` found no difference.

## Cause

`continuousUpdateCore_()` had **no time budget and no intermediate save**. It handled every new TDoc of the list in one loop, in one unsaved batch.

Each new TDoc cost:

| Per new TDoc | Code | Measured with the test harness (392 in the list, 112 in the report, 280 new) |
|---|---|---|
| One Reviewer request, made one after another | `completeInsertedUploadedTdoc_()` → `fetchAndAddAbstract_()` (`UrlFetchApp.fetch`) | 280 requests |
| One scan of the whole document for the end of its agenda section | `findAgendaSectionEndIndex_()`: `getChild`, `getType`, `getText`, `getHeading` per element | 127,332 `getChild` calls (the document grows from ~400 to ~800 elements) |
| A TDoc table: 11 rows × 2 cells, link, style | `insertTDocTableAtIndex_()` | ~30 calls per table |
| Property reads (REVISION_MAP, the Reviewer "no summary" cache) | `findParentRevisedToTable_()`, `isReviewerNoSummaryCached_()` | 611 reads |

The harness has no network latency. With the real Reviewer the requests dominate: at about 1 s per request, 280 requests alone take more than 4½ minutes. The document scans come on top, and grow with every table inserted.

The Full Build checkpoints of issue #6 did **not** protect this path. They were switched on only inside `runFullReportBuildCore_()`.

**What happened to the document.** Apps Script stopped the execution from outside. The Document service had not saved the changes, and they were dropped, so the report stayed as it was. That is not a rollback anyone can rely on: the Document service may save on its own before the end, and properties written in the run (the Reviewer "no summary" cache) are kept. Snapshot #4 shows that nothing was saved in this case.

**Main vs. ad-hoc.** A main meeting takes only the TDocs of its report's agenda items, an ad-hoc meeting takes all of them (`{mode:'all'}`). Any update with a few hundred new TDocs runs into the same limit, whichever kind of meeting it is.

## The fix: BOUNDED UPDATE section of Code.js

1. **Time budget for insertions.** A new TDoc is started only while less than `UPDATE_INSERT_BUDGET_MS_` (3 minutes) has passed in the execution. After that:
   - it, and every later new TDoc, is left to the next run;
   - so is the upload check (link and abstract) of the TDocs already in the report. The link marks that transition, so the next run sees it again.

   A TDoc that was started is finished: its table and its one Reviewer request.
2. **Optional steps.** The abstract sweep and the e-mail/revisions collector start only before 4 minutes (`UPDATE_OPTIONAL_STEPS_BUDGET_MS_`) and only when no new TDoc was left over. Otherwise they are left to the next run.
3. **What finishes the run always runs:**
   - the registration rows;
   - the revision placement;
   - the Disposition Summary;
   - the formatting;
   - the status dropdowns.

   The run then ends normally and is saved.
4. **Checkpoints during insertions.** These apply to an update with the report open (menu or timer), not to the background scheduler. Whenever about 2000 changes are pending before a new TDoc, the report is saved (`saveAndClose`) and opened again. The update then goes on with the reopened document and a fresh TDoc index. What was saved stays saved, even if the execution is stopped later.

   Only the insertions make checkpoints. The later steps make the changes an update always made, so the formatting of a large report adds no extra saves.
5. **Progress is the document.** No progress record exists apart from the document. Every run reads the TDoc tables the report has and inserts only what is missing, as every update always did. So:
   - a run whose changes were dropped leaves nothing that claims otherwise;
   - a TDoc is never inserted twice;
   - two runs at once cannot write competing progress. Both entry points hold the document lock, unchanged.
6. **Registration table on every run.** It gets the rows of the TDocs whose tables the report has, and no rows for TDocs left over. A run stopped after a checkpoint saved tables without rows; the next run adds those rows.
7. **Feedback.** The result has `paused = { newTdocsAdded, newTdocsLeft, stepsLeft }`.
   - **Update Report Now** shows it as an error the user sees: *"The report was updated in part: the update stopped in time before the Apps Script limit of 6 minutes. 180 new TDoc(s) were added; 100 new TDoc(s) are still to be added; … Nothing was lost or added twice. Run Update Report Now again to go on."*
   - **The timer** goes on by itself at its next run.
   - **The scheduler** does not advance `lastRunAt` for a paused report, so the report stays due. The budget counts from the start of the scheduler's execution.

No new permission, service or scope. `appsscript.json` and HyperLink.js are unchanged.

## Guarantees and limits

| | |
|---|---|
| Duplicate TDocs | None: the next run reads the document. Tested with 392 TDocs, a pause, a resume, and an execution stopped after a checkpoint. |
| Hand-written text, minutes | Untouched: an update only inserts tables and rows and writes Status/Disposition/links as before. |
| Dropdowns, statuses, abstracts | Kept. New TDocs get their dropdowns at the end of each run. A Status left as text (a run stopped before its dropdown step) is turned into a dropdown by the next update, as before. |
| A run stopped despite the budget | Saved up to its last checkpoint; the next run continues. Steps after the loop (formatting, dropdowns) are then redone by the next run. |
| Background scheduler | Time budget, but **no checkpoints**: it has no active document to reopen. |

## Release metadata of the test copy

The integration-test kit wrote `codeVersion: base + " + issue #6 candidate (NOT RELEASED)"` into the copy's Release.js, appending to whatever was there. The second push therefore read `2.26.7 + issue #6 candidate (NOT RELEASED) + issue #6 candidate (NOT RELEASED)`.

`candidateReleaseJs()` (tools/integration-test/guard.js) now writes these fields anew from the candidate:
- `codeVersion`: `2.26.8 (candidate of T-2026.10.24, NOT RELEASED)`;
- `gitTag`: `candidate <branch> (…)`;
- the release the copy was made from, kept in `baseReleaseId` / `baseCodeVersion`.

The push tool requires `--candidate-of T-YYYY.MM.N`. The diagnostics (`it9TimeZoneCheck`) name a candidate as such. Only the test kit changed; the runtime and release Release.js are not affected.

## Tests

`tests/bounded-update.test.js` covers:
- the pure rules;
- a small update, unchanged;
- 392 TDocs with 280 new and a deterministic clock: 180 added, 100 left over, checkpoints, the message;
- resume, with no duplicates and all 392 TDocs once in the report and in the registration table;
- idempotence;
- hand-written text;
- dropdowns and their values;
- an execution stopped after a checkpoint, then resumed;
- optional steps left over when time is short;
- the background path (`lastRunAt`, the scheduler's clock);
- regressions of the meeting-number check and of ABT / IANA zones;
- the ledger: without the section, Code.js and ReportCreator.js are those of `8edb1a7`.

Mutation check: without the budget rule, 5 checks fail.

Changed expectations:
- `large-build-checkpoints` (checkpoints are switched on in two places);
- `main-meeting-number` H and `meeting-time-zones` J (compared on the layer below);
- `integration-test-kit` (the candidate Release.js).

## Follow-ups (not in this fix)

- The Reviewer requests could be made in parallel (`UrlFetchApp.fetchAll`), and the agenda-section scan could use an index. Both would speed a large update up; neither is needed for correctness.
