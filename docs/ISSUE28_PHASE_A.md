# Issue #28 phase A: titles, disposition headings, update costs (Code.js 2.28.0, candidate)

**Status: candidate, not released, not deployed.** Branch `feature/i28-a-titles-headings-perf`, on T-2026.10.34
(`21e0b8e`). The code is the section "ISSUE 28 A" at the end of `Code.js`, plus eleven marked changes to released
functions and one menu item in `template/ReportCreator.js` (ledger: `codeWithoutIssue28A`, `creatorWithoutIssue28A`
in `tests/helpers/before-structure-fix.js`; without them both files are those of T-2026.10.34 byte for byte).
Assessment of the whole issue: docs/ISSUE28_ASSESSMENT.md.

## 1. Titles from the TDoc list

**Rule.** The TDoc number is the identity. Every update compares the title of the TDoc list (only after the list
passed its identity check, as every update does) with the Title row of each TDoc table:

| Report shows | List shows | Record says | Result |
|---|---|---|---|
| the list title | — | — | in sync (the record is written the first time) |
| the title last seen in the list | a new title | that title | **written** into the TDoc table and the Registration of Documents row (if that row still shows the old title) |
| something else | the title last seen | — | typed by hand: kept, listed as "Typed by hand" |
| something else | a new title | another title, or none | **listed for review**, never written |
| anything | a title the record lists as earlier | — | an older list: never written, listed |
| anything | the title the rapporteur chose to keep against | — | kept, not listed |

- **Record.** `TDOC_TITLE_SEEN_<tdoc>` in the report state (`v1|<current>|<up to 3 earlier>|<kept-against>`, 8-hex
  FNV-1a hashes of the normalised title; about 30 bytes per TDoc). Written only when it changes.
- **What is never touched:** Minutes, Disposition, Status, e-mail text, any other row, any free text; a Title cell of
  more than one line.
- **Look and link.** The new text is inserted after the old and the old deleted, so it takes the look of the cell; a
  link over the whole old title is put on the whole new one; a partial link is not spread.
- **Generated tables.** The Disposition Summary and the four plenary tables copy the TDoc table, so they show the new
  title when written in the same update. The Undisposed Documents section of a report without Goes to Plenary follows
  at the next update (it reads the run's snapshot).
- **Review.** *Report › Review Title Changes…* lists what is pending (both titles, why) and the last 40 log entries.
  "Use the list title" writes only where the report still shows what the dialog showed; "Keep the report title" stops
  that list title from being offered again (a newer one is offered). It takes the document lock.
- **Audit log.** `TITLE_SYNC_LOG` (newest first, ≤ 8.5 KB): TDoc, action, before, after, UTC time, update or dialog.
  Pending list: `TITLE_SYNC_PENDING` (≤ 8.5 KB). The update log says "Titles: N written, N to review, …".
- **Cost.** A report with Status dropdowns or Goes to Plenary reads the titles from the update's existing Docs API
  snapshot (no extra call). A report with text statuses reads at most 60 Title rows per update until every TDoc has a
  record; afterwards a TDoc is read only when its list title differs from the record.

**First update after deployment.** Every TDoc whose report title equals the list is recorded silently. A TDoc whose
title already changed in 3GU before deployment differs *without* a record, so it is listed for review, not written:
the rapporteur applies it once in the dialog. Expected: a handful per report.

## 2. Heading 2 class lines in the Disposition Summary (opt-in)

**Decision 2026-11-23:** Heading 2 is **off by default**. Without a setting, 2.28.0 writes the class lines as bold Normal
text, exactly as T-2026.10.33/.34 (test F: the same summary as T-2026.10.34 writes). `DISPOSITION_CLASS_HEADINGS=heading2`
switches it on per report — only once 2.28.0 (which reads Heading 2 class lines) is the rollback target in production.

- With the setting, the three (or four) class lines are Heading 2 under the Heading 1 "Disposition Summary", not bold. They appear in the
  document outline. The note, "None.", the tables and their header cells are unchanged; nothing else becomes a heading.
- The section readers know a class heading as part of the summary: `findDispositionSummarySection_` and
  `plenarySummarySectionOfContent_` (ownership, "text typed into it"), `nextPlenaryLandscapeStep_` (no portrait
  break above a class line). Text typed below the summary is kept; text typed between two class lines still stops the
  write, as before.
- **Migration, without a rewrite.** A summary written before keeps its text; the next update gives its class lines the
  new look: in the three-table form through DocumentApp (only the class lines), in the four-table form as two requests
  per line in the existing table-format batch (`planPlenarySummaryFormat_`). The update after it changes nothing.

## 3. Update costs

| Change | Same result because |
|---|---|
| Existing-TDoc scan without a snapshot: one pass builds the numbers and the index | the index of before was built from the same reads, in the same order (test G) |
| `findStatusInDocTable_`: rows counted once, cells counted only for the row sought | same first matching row, same cell, same value (test G, also for a one-cell row) |
| "Is there a summary?": `Body.findText('Disposition Summary')` first; null means none | a summary cannot exist without its heading text; found → the walk as before |

### Measured (DocumentApp calls on the fake document, 250 TDocs, SA4#138 MBS agenda; relative, not a time)

| Update | T-2026.10.34 | Phase A |
|---|---|---|
| Text statuses, unchanged | 35,263 | 26,702 (−24%; includes 600 calls of the one-time title records) |
| Text statuses, unchanged, after the records (150 TDocs, test G) | 21,363 | 15,802 (−26%) |
| Status dropdowns, unchanged | 2,190 | 542 (−75%) |
| Status dropdowns, new TDoc + Portal status | 5,972 | 4,269 (−29%) |
| Goes to Plenary, unchanged | 552 | 552 |
| Goes to Plenary, new TDoc + Portal status | 4,654 | 4,601 |
| Title changed (dropdowns) | 2,384 (title not written) | 3,064 (title written in both tables) |

Docs API reads are unchanged (1 for dropdowns, 3 for Goes to Plenary on an unchanged update). The real time is measured
on the disposable copy (§ Validation).

### Not done (left for later, with reasons)
- Structure check and registration links still scan through DocumentApp in text-status reports (4,148 + 2,293): a
  snapshot there would change "a report with text statuses is never read through the Docs API".
- The plenary chain still makes 3–6 full Docs API reads per changing update; reusing a read across writes needs a
  revision-safe design.
- The TDoc list is still converted to a Sheet on every update; caching it needs chunked storage.

## Configuration

| Document property | Values | Default |
|---|---|---|
| `TITLE_SYNC_MODE` | `auto` (write where safe), `review` (list only), `off` (nothing compared, nothing stored) | `auto` |
| `DISPOSITION_CLASS_HEADINGS` | `heading2` (also `heading 2`, `h2`): Heading 2 class lines; anything else: bold Normal | bold Normal (as T-2026.10.34) |

Both can be set without a code change (Apps Script › Project settings › Script/Document properties, or a one-line
script), and take effect at the next update.

## Migration of existing reports

Opt-in by deployment: nothing changes in a report until it runs 2.28.0. Then the first update records titles, lists
pre-existing differences for review, and restyles the summary's class lines. No rebuild, no status, Plenary choice,
minutes or disposition is changed. Old AHG reports and other SWG reports behave the same way when they get the release.

## Rollback

- **Behaviour only:** set `TITLE_SYNC_MODE=off` and/or remove `DISPOSITION_CLASS_HEADINGS` (the next update restores
  bold Normal class lines).
- **Code, default (Heading 2 never switched on):** push the report's T-2026.10.33 (or .34) read-back (`--rollback`); no
  preparation. Test F runs T-2026.10.33 and .34 over the same document after 2.28.0: each rewrites the summary in place
  (one heading, three class lines), and 2.28.0 reads what they wrote when upgraded again.
- **Code, Heading 2 switched on in that report:** first remove `DISPOSITION_CLASS_HEADINGS` and run **one update with
  2.28.0** (it restyles the class lines back to bold Normal; test F). Only then push the older read-back. Rolling back to
  2.28.0 or later needs no preparation.
  **Why:** older code ends the summary at the first heading after its own, so on a summary with Heading 2 class
  lines its next rewrite writes a new summary under the note and keeps the Heading 2 one below it as "text typed
  below" — a duplicated summary (reproduced on the fake document with the T-2026.10.34 code). Titles written stay
  written (they are the list's titles). The `TDOC_TITLE_*`/`TITLE_SYNC_*` properties are ignored by older code and
  can be left.

## Release procedure (when approved)

Like T-2026.10.33/.34 (docs/ROLLOUT_T-2026.10.34.md): build the bundle from the clean worktree at the release commit;
dry run per target against its last read-back; disposable copy first; then Video, MBS, RTC, Audio, 6G one at a time with
"before T-2026.10.3x" named versions and "nobody editing". **New:** `ReportCreator.js` changes (one menu item), so the
plan must allow Code.js, Release.js **and** ReportCreator.js; each report's appsscript.json stays. The master template
is a separate decision.

## Validation

- Tests: `tests/i28a-titles-headings-perf.test.js` (parts A–H) and the existing suite (133 files, all passing).
  Existing tests changed only where the intended behaviour changed: version strings, the menu tree (+1 item), readers
  of the summary that ended it at any heading, the "bold class lines" assertion, and the .33/.34 layer checks now read
  the .34 code.
- Disposable copy: pending. Needs a copy of a SA4#138 report made by the user (clasp cannot copy Docs); checks: About
  shows 2.28.0, one update (time from the execution log), Review Title Changes lists the expected TDocs, outline shows
  the class lines, a second update changes nothing.

## Independent review (2026-11-23) and what changed

A separate reviewer read the diff without this document. Findings and outcome:

| # | Finding | Outcome |
|---|---|---|
| 1 | More differences than one property holds: the shown list churned each update, review entries were re-logged and pushed "applied" entries out of the log | **Fixed.** "Listed for this list title" is part of each TDoc's record (`…|review:<hash>`), so nothing is re-logged; the shown list is kept in TDoc order with a count of the rest (`__more`); the log cuts review entries before "applied"/"kept" ones (test R) |
| 2 | Text statuses: TDocs in review/kept were re-read every update and used the 60-read budget, so a real change could starve | **Fixed.** A TDoc is read only when its list title differs from the one it was last decided for; recorded TDocs with a changed list title are read first (test R: 60, 9, then 0 reads; the real change written at once) |
| 3 | Code rollback duplicates a summary with Heading 2 class lines | **Decided 2026-11-23:** Heading 2 is opt-in, off by default (below) |
| 4 | Snapshot path joined a multi-line Title into one string and listed it | **Fixed.** The Title is read with its paragraphs (`titleTextOfDocsTable_`), so it is skipped as before |
| 5 | A cell ending with an empty paragraph could lose its look/link on replacement | **Fixed** as far as testable: only the first paragraph is edited; a cell with text in two paragraphs is refused. To be confirmed once on the disposable copy (linked, bold title) |
| 6 | Class headings were set to bold=false | **Fixed.** Bold is cleared so the Heading 2 style decides (DocumentApp `BOLD: null`; Docs API empty `textStyle` with field `bold`) |
| 7 | Add-on background store: per-key Script Properties calls; adopted documents read/write different stores | Documented: adoption is not wired into any menu; title records are not in `reportStateKeysToAdopt_`. To be added when adoption is used |
| 8 | Size of document properties | ~45–65 bytes per TDoc (250 TDocs ≈ 15 KB) on top of the existing per-TDoc caches; check the live stores before rollout |
| 9 | The headings setting read document properties directly | **Fixed** (report state, like `TITLE_SYNC_MODE`) |

### Decision (2026-11-23): Heading 2 opt-in, off by default
Option 1 was chosen, in code rather than by the deploy script: no setting means bold Normal. Test F proves the
backward-compatible reading (2.28.0 reads Heading 2 and bold Normal summaries, and what T-2026.10.33/.34 wrote) and the
rollbacks over the same document. The considered options were:

Older code (≤ 2.27.x) duplicates a summary that has Heading 2 class lines on its next rewrite. An emergency rollback
may not allow "set `DISPOSITION_CLASS_HEADINGS=normal` and run one update" first. Options:
1. **First rollout with `DISPOSITION_CLASS_HEADINGS=normal`** set in each report by the deploy script; switch Heading 2
   on per report one release later, when the rollback target (2.28.0) reads Heading 2 class lines. Safest.
2. Ship a tolerant reader as a separate hotfix (2.27.3) first, then 2.28.0 with Heading 2 on.
3. Accept the documented two-step rollback.

## An update stopped from outside (added 2026-11-25)

The first run on a disposable copy of the 6G Media report was killed at Apps Script's 6-minute limit, and nothing recorded it. The copy started without document properties, as every Drive copy does, so the update ran in text-status mode. In that mode an update of 150 TDocs makes about 15,800 DocumentApp calls in 2.28.0 (21,400 in T-2026.10.34). In the Status-dropdown mode it makes about 340.

- **Journal.** The update keeps `UPDATE_RUN` in the report state: when it started and which stage it reached. The stage is updated by `perfTimed_`, and the record is removed when the update ends. The next update finds a leftover record, keeps it as `UPDATE_INTERRUPTED` and logs it.
- **Repair.** Only after an interruption, a TDoc table that is in the report more than once is handled:
  - **Identical content:** one copy is kept. A move killed between its insert and its remove leaves exactly such a copy.
  - **Different content:** the copies are not touched; they are named in the log and in the result.

  The record closes after an update that completes, with no stage left over and no differing copies.
- **Budgets.** Each budget runs on its stage's own clock, and the rest is left to the next update ("updated only in part"):
  - the TDoc loop (new TDoc tables, statuses): 100 s;
  - the revision moves: 60 s.

  The stages after them keep their budgets on the update's clock.
- **Resuming.** A TDoc loop stopped by its budget records where it stopped (`UPDATE_TDOC_CURSOR`), and the next update starts there, wrapping around to the beginning. Without this, every update would spend its budget on the same first TDocs and never reach the last ones (found by the convergence fixture). New TDocs keep their order, and the stages after the loop keep the list order.
- **A repairing update** reports itself as "updated only in part": its read of the report was taken before the repair. The next update completes it, for example turning a Status that is still text back into a dropdown, and closes the record.
- **Tests.** `tests/i28a-update-journal.test.js`. Part E covers:
  - a move killed between its insert and its remove; after recovery the report equals one an uninterrupted update produces;
  - twelve new TDocs under the budget, converging across several updates to the document of a single unbudgeted update, after which a further update changes nothing.

## One clock for the whole update (added 2026-11-26)

On 2026-11-26 an A update of the clean 140-TDoc main copy took 212.6 s. The next one was ended by the 6-minute limit. The stage budgets above each have their own clock and only decide whether a stage starts; a stage that has started (the e-mail archive pages, the revisions listing) runs on, and a network request cannot be stopped once it runs.

- **The update's clock.** It starts with the journal, on the same clock as the stage budgets. The target is 300 s; 45 s of it are kept for the end (the status dropdowns, the journal).
- **Network requests.** `safeFetch_()` (the RSS feed, the archive pages, the revisions listing) starts no request later than 195 s into an update: a request may take up to a minute, and the end keeps its 45 s. A refused request throws an error marked `updateDeferral`. Its collector stops before it writes anything, so no E-mail Discussion or Revisions cell is rewritten from part of the messages or listings, and no link is lost. The collector is reported as left to the next update ("updated only in part"), not as a failure. The TDoc list (at the start of the update) is not affected.
- **The TDoc loop and the moves** keep their own 100 s and 60 s, and also start no TDoc later than 170 s and no move later than 200 s into the update.
- **The stages the next update does again anyway** (the revision relationships, the Goes to Plenary data, the Disposition Summary, Undisposed Documents) start no later than 255 s. The status dropdowns always end the update.
- **Resuming.** The revisions go on where the last update stopped them (`UPDATE_REVISIONS_CURSOR`), as the TDoc loop does (`UPDATE_TDOC_CURSOR`); a whole pass removes the cursor. Every stage that is left to the next update is idempotent: it compares with the report and writes only what differs.
- **Diagnostics.** `UPDATE_RUN` keeps the stage, the time since the start (`elapsedMs`) and the last 12 stages with their times, the network requests among them. An update stopped from outside is named by the next one with its stage and time (`UPDATE_INTERRUPTED`, `result.interrupted`).
- **The reallocations.** `getReallocationMap_()` already reads from the update's one Docs API read (`reallocationMapFromSnapshot_()`); the 29 s under its timer in the real run were that read itself, which happens there first. A test proves that an update walks no table for it.
- **Tests.** `tests/i28a-update-journal.test.js`, part F: no mail request 196 s into an update and no E-mail Discussion cell touched, while the same report makes the request at 100 s; the late stages left at 256 s and run at 200 s; the TDoc loop and the moves on the update's clock; the revisions cursor (resume, wrap, removal only after a whole pass); the journal's stages and times, named by the next update; the reallocations without a table walk.

## The revision pairs within the time (added 2026-11-26, after Run 5)

In Run 5 the third A update of the 140-TDoc main copy spent more than 200 s in `rearrangeRevisionTables_()` (normally 7-14 s) and was ended by the 6-minute limit; the journal kept the stage. The budget stopped only the moves: every pair's "Revised to" disposition and its position check went on, under a Docs service that was slow that afternoon (the one read of the report took 33 s, 58 s and about 85 s in three updates).

- **Every pair within the time.** The pairs' budget (their own 60 s, and no pair later than 200 s into the update) is checked before each pair and again before its move. Past it the loop stops; the pair it stopped at is kept as the next one to do (`UPDATE_PAIRS_CURSOR`) — the pairs before it are done, and that pair is done again next time (each step of a pair compares before it writes). The update says it is in part ("revision placement").
- **Resuming.** The next update starts at the kept pair and goes round once to the pair before it, so no pair is skipped; a whole pass removes the cursor.
- **The in-place check from the update's read.** Whether a revision already sits below its original comes from the update's Docs API read, without asking the document for positions — only while that read is still the report's structure: no move and no section repair in this update, and as many body elements as the read (an insertion or removal changes that). Otherwise, or with any doubt (a table twice, a section break between, a table that is not in the read), the document is asked as before.
- **Clocks.** A stage clock that starts at 0 (possible only with the clock of the tests) is no longer taken for "not started".
- **Tests.** `tests/i28a-update-journal.test.js`, part G: an update past the pairs' budget stops inside the third of three pairs, says so and keeps that pair; the next one starts there and goes round once; the report equals that of an update without budget (order, statuses, dispositions, typed Minutes and Disposition text), no table twice, "Revised to" once; a further update changes nothing; the in-place check reads no position of the document while the read holds, and asks the document after a section repair or when the body has another number of elements.
- **Independent review (2026-11-26).** Three changes after it: (1) the read is not trusted for positions after any table was inserted (new TDocs, revision-linked tables) or removed (the repair of an interrupted update) earlier in the update, even when the number of elements happens to be the same; (2) only an empty paragraph as `isEmptyBodyParagraph_()` takes it counts as "between" — Normal text, no list item, no positioned object; (3) resuming starts at the beginning of the cursor pair's chain (A→B before B→C), so a chain is never split. Tests for each in part G.
- **Known limitation.** The abstracts stage (`addAbstractsForTables_()`, off by default for automatic updates, not active on the test copies) asks the Reviewer API with its own requests, outside `safeFetch_()` and the update's clock.

## Every stage on one timeline (added 2026-11-26, after Run 6; diagnostics only)

Run 6: one update of the 140-TDoc main copy took 347 s, but its timed stages accounted for only about 95 s. The rest was spent in stages that had no timer, and the evidence could not say which ones.

- **What is timed now.** Besides the stages timed before, each significant stage of `continuousUpdateCore_()` runs in `perfTimed_()`:
  - the report body;
  - the meeting-room availability and the check after an interrupted update;
  - the structure check;
  - the TDoc loop (new TDoc tables, statuses, upload completion);
  - the Registration table links, the ad-hoc sessions and the Session column;
  - the title sync;
  - the revision relationships, Goes to Plenary, the Disposition Summary and Undisposed Documents;
  - the status dropdowns (`finalizeStatusDropdowns_()`).
- **The timeline.** `perfTimed_()` records each stage in `PERF_STAGES_`: its name, its start (ms from the start of the update), its length and its depth. Depth 0 is a stage of the update itself; a stage inside another has depth 1 or more, so nested time is never counted twice. At the end of the update, one `[PERF] stage timeline:` line gives:
  - the time in the timed stages;
  - the time outside them (total minus timed);
  - the slowest stages.
- **Evidence after a stop from outside.** The update journal (`UPDATE_RUN`) records the end of each stage, written with the journal's next stage (no write of its own). It keeps the last 40 stages instead of 12: up to about 5 KB with 80-character names and their ends, well below the 9 KB of a property. An update stopped at the limit therefore shows which stages ended, when, and the stage it was in.
- **Nothing else changes.**
  - The update does the same work in the same order, with the same budgets, cut-offs, network guard, cursors and rules.
  - Each newly timed stage costs one more write of `UPDATE_RUN` at its start, about 15 per update, as the journal already does for the timed stages.
  - The ledger (`tests/helpers/before-structure-fix.js`) undoes these changes first and gives the code of `e6d5a9a`. The byte comparison with T-2026.10.34 still holds.
- **Tests.** `tests/i28a-update-journal.test.js`, part H:
  - **Same report.** The same report and list are updated twice with the timing and twice without it (`perfTimed_` only calling its function). Both give the same report (the whole body, the order of the tables), the same report state and the same results.
  - **Timeline.** Each stage of the update is on the timeline once, in order, without overlap. The summary counts nested stages once. After an error the depth is right again.
  - **Stopped update.** An update stopped in its third stage leaves the first two with their ends. The next update records them.
  - **Journal size.** The journal keeps 40 stages, well below 9 KB.

## The Disposition Summary within the time (added 2026-11-27, after Run 7)

Run 7 timed one update of the 140-TDoc main copy at 326 s, with less than 1.1 s outside the timed stages:

- About 166 s were a one-off identity check of the TDoc list. The copy had no saved anchor (`TDOC_LIST_ANCHOR`, removed by the test cleanup) and no agenda TDoc. After the first update the anchor is there, and the check takes about 1 s.
- The Disposition Summary took 109 s, starting at 216 s. The other stages had already been left to the next update at 214 s, but the summary had no limit of its own. The next two updates spent 53 s and 15 s on it.

The change:

- **The limit.** An update that has run more than 170 s on its clock (`updateElapsedMs_()`, the limit of the TDoc loop) does not start the summary.
  - It leaves the summary to the next update (`deferReportUpdateStage_('Disposition Summary')`).
  - The update says it is updated in part (`result.incomplete`), and its result has `dispositionSummary: 'deferred'`.
  - Nothing of the summary is read or written.
- **The next update.** Every update writes the summary again from the report, so the next update that reaches it within 170 s writes it as it should be. In Run 7 the updates reached it at 100 s and 62 s.
- **Unchanged.**
  - A report without a summary is still `'absent'`, never "in part".
  - A report with Goes to Plenary writes its summary at the end of the update, with its own limit, as before.
  - Past 255 s the earlier limit leaves all the late stages, as before.
  - Undisposed Documents reads from the update's Docs API read when the summary has not run, as it does in a report without a summary.
    - Without the Docs API it would read the TDoc tables itself (at most every 15 minutes), a cost the limit does not save.
    - The reports and the test copy use the Docs API.
  - Outside an update nothing is left.
- **Tests.** `tests/disposition-summary-close-report.test.js`, part X:
  - 170.001 s: left, in part, the report identical, no read for it, logged. 170 s: it runs.
  - The next update writes it; the result is what Update Disposition Summary writes.
  - Repeated updates at 62–169 s: complete and identical.
  - Two late updates in a row, then one in time: written.
  - Typed Minutes and Disposition, the Status dropdowns and the links stay.
  - Goes to Plenary and outside an update: nothing left by the new limit.
