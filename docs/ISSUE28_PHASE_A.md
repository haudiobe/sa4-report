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
