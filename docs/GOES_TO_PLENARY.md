# Goes to Plenary, the compact Status row, the four disposition tables and the revision relationship

**Status: CANDIDATE, not released and not deployed.** Branch `feature/plenary-dispositions`
(PR #26) is based on `af3779f` (T-2026.10.29, `Code.js` 2.26.13). The release candidate is
T-2026.10.30 (`release/T-2026.10.30`). Nothing has been pushed to the master template or to the
five production reports. The candidate runs only in the disposable MBS copy.

Code: the section `GOES TO PLENARY AND THE FOUR DISPOSITION TABLES (after T-2026.10.29)` of
`Code.js`, small changes to 24 released functions, and two menu items in
`template/ReportCreator.js`. Tests: `tests/goes-to-plenary.test.js`. The ledger layer is
`codeWithoutGoesToPlenary()` in `tests/helpers/before-structure-fix.js`; taking it out gives the
T-2026.10.29 files byte for byte.

**A report changes only when its rapporteur runs Report › Set Up Goes to Plenary….** A report
that has not been set up behaves as under T-2026.10.29. The exceptions are general to every
report:
- the revision relationship in the TDoc row (§ 4);
- a faster update (§ 0);
- the read-only check *Advanced and Repair › Check Report Integrity*.

## 0. Fifth candidate: the two blockers found in the fourth candidate's real validation

The fourth candidate's update ran in about 95.5 s in the disposable MBS copy, so the performance
fix held. Two blockers remained.

### Blocker 1: 28 Status values stayed plain text

**Cause.** An update turns a text Status back into a dropdown only in a report marked
`STATUS_DROPDOWNS`. Only *Build Report from Scratch* and *Convert Status Fields to Dropdowns…*
set that mark. *Set Up Goes to Plenary* does not require it. A report that got its Status
dropdowns another way, as the copy did, is therefore never healed.

*Check Report Integrity* uses the same rule for which statuses count as missing, but ignores the
mark. It said "the next Update Report Now makes them dropdowns", which was wrong for this
report.

**Fix.**
- **The update.** In a report with Goes to Plenary, the update's single read is used to check
  whether the report already has Status dropdowns: the definition is present and at least one TDoc
  Status is a dropdown of it (`statusDropdownReportByEvidence_()`).
  - If so, every text Status that has an option is restored as a dropdown, in the same batches as
    before.
  - Only the status part, before "Plenary", is replaced. The Plenary dropdown, the A.I., the
    minutes and the dispositions are untouched.
  - Afterwards the report is marked.
- **No wasted writes.** Dropdowns that are already right are not written. A second update makes
  no API write.
- **Reports with text statuses** (no Status dropdown at all) are unchanged by this. Without the
  mark or Goes to Plenary, the update still does not read the report for this.
- **The targeted command.** If every Status is text, the update has nothing to go by.
  *Advanced and Repair › Convert Status Fields to Dropdowns…* (`repairStatusDropdowns_()`):
  - converts them, compact rows included;
  - checks that nothing else changed;
  - marks the report.
- **Check Report Integrity** now shows "Marked as a report with Status dropdowns: yes/no". For text
  statuses it says whether the update or the command will convert them.
- **A refused batch** now makes the update incomplete: "Status dropdowns (failed: …)".

### Blocker 2: "summary: failed"

**Cause.** The summary writer computed where the summary heading is among DocumentApp's children
from the Docs API's list of body elements. It skipped section breaks.

Google Docs adds a paragraph in front of every section break it inserts. The API lists that
paragraph; DocumentApp counts the body differently. After the landscape section breaks went in,
the computed position no longer held the heading. The writer then returned `failed`, and the
reason ("the summary is not where the Docs API shows it") was not logged.

The fake document used by the tests did not model that paragraph, so the tests missed it. It
does now, in both directions (`api.model` in `tests/helpers/fake-docs-api.js`). Both reproduce
the real "failed" on the fourth candidate.

**Fix.**
- **Finding the summary** (`plenaryLocateSummaryInBody_()`).
  - The heading is looked for near the computed position, within the number of section breaks
    before it.
  - Every element after it must be the kind the read shows (paragraph or table).
  - Otherwise nothing is written, and the reason names the element that differs.
- **Writing it.** The new summary is written right after the old one. The old one is removed only
  when the new one is complete. If writing fails, what was written is taken out again and the
  existing summary stays.
- **Logging.**
  - The reason is logged: `summary: failed (<reason>)`.
  - When DocumentApp's count differs from the API's, the offset is logged too.
- **Incomplete reporting.** A summary that is not written (failed, refused, unreadable, no scope),
  a formatting or landscape step that fails, or a Goes to Plenary step that throws now makes the
  update incomplete.
  - *Update Report Now* says "The report was updated only in part. These parts were not completed
    (out of time, or a step failed) …".
  - It names the stage and the reason.
- **Unchanged update.** An unchanged update still does not rewrite the summary: `unchanged`, no
  write.

### 79 against 80

*Check Report Integrity* counts TDoc tables: two columns, "TDoc" and a TDoc number. The update's
"Found N existing TDOCs" counted every table whose header begins with "TDoc". The four summary
tables begin with "TDoc | Title | …", so one extra entry was counted: the empty key in the fourth
candidate, "Title" in the old scan. The update now counts only TDoc tables, so both say 79. The
"not in the TDoc list" log already ignored the extra entry. Nothing is deleted.

### Evidence

The fake-document tests (`tests/goes-to-plenary.test.js` §§ R, S, T) failed on the fourth
candidate and pass now.

| Section | What it covers |
|---|---|
| R | 28 text statuses in an unmarked report, all restored with their values, Plenary and A.I. kept, nothing else changed, then no write; a refused batch reported as incomplete and completed next time; the command for a report with only text statuses |
| S | Both mismatches: summary replaced with four tables (6, 6, 5, 5 columns), links, minutes and landscape kept, then `unchanged`; a failure while writing: old summary kept, reason logged, update incomplete; a mismatch inside the summary: nothing written, the reason precise |
| T | The existing-TDoc count equals the integrity count; a failing Goes to Plenary step makes the update incomplete |

Performance on the fake is unchanged:

| Update | DocumentApp calls |
|---|---:|
| Unchanged | 198 |
| Unchanged, with the collectors | 620 |
| Summary rewrite | 309 (290 before; 19 calls for the position check) |

Whether the real report's summary is found depends on how Google Docs counts that paragraph. The
fix handles either way, and the log of the first real update says which it is.

## 0b. Performance: the six-minute limit (fourth candidate)

### What failed

The third candidate passed its first update in the disposable MBS copy. The **second** update,
with almost nothing changed, still ran past ten minutes:

- 77 s passed between "existing TDoc tables found" and "Added 0 new";
- the revision step was still running minutes later;
- the log counted 75 TDocs in the list and 76 TDoc tables in the report.

### Root cause

The third candidate fixed the summary, the plenary data and the status sync. It did not touch the
code that runs **before and after** them on every update, and that code still read and wrote the
report one cell at a time through DocumentApp. In a report this size each DocumentApp call costs
about 0.1 s (77 s for about 10 calls per TDoc across 75 TDocs).

| Step (pre-existing code) | What it did on every update | DocumentApp calls, 75 TDocs, ~150 revision links (fake document) |
|---|---|---:|
| **Revisions** (`updateRevisions_`) | For every revision file in the folder, `insertRevisedDocTablesAfter_` scanned the whole document for its table; every Revisions cell was rewritten and its store saved, even when equal | about 85,000 |
| **E-mail discussions** (`checkRSSFeed_`) | Every TDoc's discussion cell was rewritten, and the deadline row re-set, on every run | thousands |
| **Existing TDoc detection, structure check, registration links** | Each walked every table and every row | the 77 s before "Added 0 new" |
| **Deadline extensions, reallocations** | Each walked the whole document for its table | hundreds |

The total for an unchanged second update with the real collectors was **95,154 DocumentApp calls**
(third candidate, `b5893c8`), which at about 0.1 s per call is well over the six minutes.

### What the fourth candidate does

1. **One Docs API read per update** (`reportRunSnapshot_()`). It reads the whole report once: all
   tables, cells, links and dropdowns. Every step below decides from it, and uses DocumentApp only
   for a cell that must change. The read is taken only in an update of a report with Status
   dropdowns or Goes to Plenary; otherwise the update runs as before.
2. **Revisions** (`updateRevisionsFast_()`).
   - The revisions folder is fetched **once** per update.
   - For each TDoc, the planned cell (text, every link, the local-revision markers) is compared
     with the cell in the read. An equal cell is never written, and its store is not saved.
   - A cell whose text is right but whose link is lost is written again.
   - A stale or failed local listing clears **no** local revision that was imported before.
3. **E-mail discussions** (`checkRSSFeedFast_()`): the same comparison. The deadline row is left
   alone when it already reads "No deadline set".
4. **Existing TDocs, structure check, registration links, deadline extensions, reallocations**
   are all taken from the same read.
5. **75 against 76.** The update now logs every TDoc table that is not in the TDoc list
   (`logTdocTablesNotInList_()`), and *Check Report Integrity* lists them as "not in the TDoc
   list". **Nothing is deleted.** Typical causes: a TDoc withdrawn from the list, or a table
   added by hand. The rapporteur decides.
6. **Time budgets, reported honestly.**

   | Elapsed time | Effect |
   |---|---|
   | 150 s | Collectors not started are left to the next update |
   | 215 s | The revision step stops between two TDocs |
   | 240 s | Formatting is left to the next update |
   | 285 s | The summary is left to the next update |

   - A stage left over is logged as `[PERF] left to the next update: …`.
   - It is returned in `result.incomplete`.
   - *Update Report Now* then says "The report was updated only in part: …", never plain success.
   - The dropdowns (Status, Plenary) are always completed.
7. **Resumable and idempotent.** Every step compares before it writes. The next update completes
   what an earlier one left, and writes nothing twice (§ Q of the tests: a stopped revision step
   is completed by the next update, with no duplicate line).
8. **Logging.** Each update logs:
   - the time of the read;
   - the folder fetch and the parsing;
   - for revisions and for e-mail discussions, how many cells were compared, written and stored;
   - the stages left over;
   - a `[PERF] timings summary`.

### Evidence, and what it does not show

**Measured on the fake document, in the tests.** These are counts of DocumentApp calls, not times
(`tests/goes-to-plenary.test.js` § Q: real collectors, 75 TDocs, about 150 revision links,
mail on 5 TDocs):

| Unchanged update with the collectors | DocumentApp calls | API writes | Stores written |
|---|---:|---:|---:|
| third candidate (`b5893c8`) | 95,154 | — | — |
| fourth candidate | **620** | **0** | **0** |

The same tests also show:

- one new revision file changes exactly one TDoc table;
- one new message changes exactly one table;
- a lost link is restored;
- a stale agent clears nothing;
- at 160 s the collectors are deferred and reported as incomplete;
- a stopped revision step is completed by the next update.

**Not shown by the tests:** real durations. How long one Docs API read of this report takes, and
how slow each DocumentApp call really is, can only be measured in Google Docs. The timeout is
**not** claimed to be fixed until two real updates in the disposable copy finish well within the
six minutes (rollout step 0).

### Earlier: the third candidate (the first-update timeout)

On 2026-11-18, *Update Report Now* in the disposable MBS copy (75 TDocs) ran out of Apps Script's
six minutes after about 361 s.

| Part of the reported log | Time |
|---|---:|
| Disposition Summary replacement | about 134 s |
| Mail/RSS and revision collection | 96 s |
| Table-width formatting | 28 s |
| TDoc list | 8 s |
| Not in the log; see below | about 95 s |

Every DocumentApp call is a round trip, and in a large document each one is slow. Where the
calls went:

| Cause | DocumentApp calls (75 TDocs, test document) |
|---|---|
| **Summary.** Every cell was formatted one call at a time (bold, background, width, spacing). Before that, every summary cell was read back to compare, and every TDoc table was read through a reader that scans its rows four or five times. | 18,900 per update in T-2026.10.29 (whenever a summary exists); more in the second candidate |
| **Status sync** (pre-existing). For every TDoc, the Status cell was found and read twice, even when nothing changed. Most likely most of the unattributed 95 s. | 10,400 per update |
| **Formatting** (pre-existing). Any new TDoc re-formatted every row of every table. | about 2 per row of the report |
| **Collection** (pre-existing, unchanged). Mail/RSS and the revisions folder fetches. | network |

What the third candidate does:

1. **The summary is built from one Docs API read** of the saved document (`plenarySnapshot_()`),
   which carries all tables, statuses, dropdowns and links. It never reads cell by cell.
   - What the summary shows is compared from the same read; when equal, nothing is written (0
     DocumentApp calls).
   - When something changed, each table is inserted whole (one call).
   - Its look (landscape widths, header, bold, spacing, links) is set in **one** API batch after
     the save (`formatPlenarySummaryTables_()`).
2. **The Goes to Plenary data come from one Docs API read.** The automatic agenda items come from
   the TDoc list alone. DocumentApp is used only for a table that must change (the update's TDoc
   index finds it).
3. **The status sync** decides from statuses the API has already read when a dropdown stays as it
   is (`statusSyncLeavesDropdown_()`, the same rule). DocumentApp is used only where something
   changes. A report with text statuses is never read for this.
4. **Formatting:** the update formats only tables whose row count is not the one it last
   formatted (`FORMATTED_TABLE_ROWS`).
5. **A time budget.**
   - After 240 s the update leaves the formatting to the next run.
   - After 285 s it leaves the summary to the next run (`deferred`).
   - The dropdowns (Status, Plenary) are always completed.
   - Every step is idempotent, so a retried update completes what an interrupted one left (§
     Recovery).

**Measured on the fake document** (`tests/goes-to-plenary.test.js` § M, and an ad-hoc profile).
These are counts of DocumentApp calls for the same 75-TDoc report, not times:

| Update | T-2026.10.29 | second candidate (`4c55c3c`) | third candidate |
|---|---:|---:|---:|
| unchanged | 34,082 | 48,119 | **5,036** |
| one Status and one Plenary changed | — | 56,249 | **5,128** |
| a new TDoc and a Portal status | — | 57,869 | **7,154** |

What remains of an unchanged update is pre-existing upkeep: structure check, upload completion
and registration links. Real times can only be measured in Google Docs; that is the first step
of the disposable-copy test.

## 1. The compact Status row

```
| TDoc   | S4-262013 (revision of S4-262012)                       |
| …      |                                                         |
| Status | [agreed ▾]  Plenary [☑ Yes ▾]  A.I.: 15.4               |
```

- **No extra rows.** The Status cell holds the status, exactly as before (a dropdown, or text
  for a status without an option), then the code's own part:
  - "Plenary": a native dropdown, ☐ No / ☑ Yes, ☐ No by default.
  - "A.I.: X.X": shown only when Plenary is ☑ Yes and the SA4 plenary agenda item is known.
- **Why a dropdown.** A Google Docs checklist box cannot be read by Apps Script or by the Docs
  API (the API's `Bullet` has no checked state), so it could never decide the tables.
- **The A.I. refreshes on the next update.** Google Docs cannot show or hide text when a
  dropdown changes, so A.I. appears or disappears at the next update, not at once.
- **The selection belongs to the rapporteur.** No code clears it, a change of Status never
  touches it, and it is kept by every update, TDoc-list refresh, Update TDoc Revisions, Import
  Local Inbox Listing and Build Report from Scratch.
- **How the code finds its part.** The code's part begins with an invisible mark: a zero-width
  space and an em space before "Plenary". Every reader of the Status reads only what stands
  before it. That covers the status sync, the Disposition Summary, Close Report, Report Status
  Summary, the discussion e-mails and the Docs API index. The changes are in
  `findStatusInDocTable_`, `readTdocStatus_`, `findStatusText_` and `statusDocsCellEntry_`.
- **How the part is written.**
  - The code's part is written only through the Docs API, at the positions of one read, in
    batches tied to its revision (`finalizePlenarySelection_()`).
  - DocumentApp never rewrites such a cell. A status that is text is replaced in front of the
    mark (`setStatusCellText_()`), and it still becomes a dropdown as soon as it has an option.
- **The Status "Plenary" is retired.** A TDoc cannot be both "agreed" and "Plenary", so (the
  rapporteur's decision, 2026-11-18):
  - a Status "Plenary" becomes Plenary ☑ Yes;
  - the Status becomes the TDoc list's status, or "available" until the next update moves it
    on;
  - the option leaves the report's "Document Status" dropdown.

## 2. The SA4 plenary agenda item

The value lives in document properties, so that it is kept while A.I. is hidden:

| Property | Holds |
|---|---|
| `PLENARY_AGENDA_AUTO:<TDoc>` | what the update derived (§ Mapping) |
| `PLENARY_AGENDA_MANUAL:<TDoc>` | what the rapporteur typed over "A.I.: …" |
| `PLENARY_AGENDA_SHOWN:<TDoc>` | what the code last wrote there, to tell an edit from its own text |

What the rapporteur can do with "A.I.: …":

| Action | Result |
|---|---|
| Type over the value | Manual override; never overwritten, kept while Plenary is ☐ No |
| `-` | Blank on purpose; no A.I. shown |
| Clear it to "A.I.:" | Automatic again |
| A work item changes | Only the automatic value moves |

The Disposition Summary shows the item for every TDoc, also for those not going to plenary.

### Mapping

The sources (nothing is invented):

- **The meeting's agenda TDoc.** For SA4#138 this is S4-261857. It is read once and kept in
  the report.
- **The work items:**
  - the TDoc list's **Related WIs**;
  - the title tag (`[AMD_ARCH_Ph2-MED] …`);
  - the acronym of the SWG agenda item and of the items above it.
- **A CR's Release.** This is the TDoc list's **Release** column.

The plenary parts are found by their titles:

| Part | SA4#138 |
|---|---|
| maintenance: top-level "Release N and earlier matters" | **14** |
| the items under "Release N Features" | 15.1–15.7 |
| the items under "Study Items" | 16.1–16.7 |

Rules, first match wins:

| # | Rule | SA4#138 |
|---|---|---|
| 1 | A CR to the maintenance release or earlier | → 14 |
| 2 | Under an SWG "Release N and earlier matters" item (7.4 / **8.4** / 9.4 / 10.4) | → 14 |
| 3 | The work items name exactly one 15.x/16.x item; more than one → blank | 8.5 → 15.4, 11.x → 16.7 |
| 4 | TEI of the maintenance release or earlier (TEI19) | → 14 |
| 5 | The SWG item has the title of exactly one 15.x/16.x item | 8.8 → 15.7 |
| 6 | Otherwise blank, for the rapporteur | 8.3, 8.9 |

> **To confirm:** "Maintenance: use 8.4". In S4-261857, 8.4 is the MBS SWG item; the plenary
> item of that name is 14. The candidate writes 14 (a one-line change for 8.4).

An ad-hoc meeting has no SA4 plenary agenda: no A.I., and no agenda is read.

## 3. The four disposition tables, on landscape pages

**Content.** The Disposition Summary at the end of the report has four tables, in this order:

1. Agreed, Endorsed and Approved Documents – Goes to Plenary
2. Other Dispositions – Goes to Plenary
3. Agreed, Endorsed and Approved Documents – Not Going to Plenary
4. Other Dispositions – Not Going to Plenary

- Columns of tables 1 and 2 (going to plenary), six: TDoc (the number, with the link of its TDoc
  table), Title, Source, Status, SWG Agenda Item, SA4 Plenary Agenda Item.
- Columns of tables 3 and 4 (not going to plenary), five: the same **without** SA4 Plenary
  Agenda Item. The width goes to Title. The value is still kept for those TDocs and comes back
  when they go to plenary.
- Each contribution is in exactly one table, in the order of the report.
- "Positive" means the statuses agreed, partially agreed, endorsed and approved.
  Everything else is "other", including a missing, unknown or conflicting status. The taxonomy
  is the existing one (`DISPOSITION_OF_STATUS_`).

**Regeneration.** Every update writes the summary, creating it where there is none:
- It rewrites the summary only when the content changes; a repeated update makes no API write.
- It never writes over text typed into the summary. Text typed below it is kept.
- It writes nothing when a status or a selection cannot be read.
- *Update Disposition Summary* stays, and Close Report writes the same tables. Close Report's
  counts and its "not yet disposed" list are unchanged.

**Landscape.** The summary stands in a section of its own:

- **The section:** a NEXT_PAGE section break before its heading, with
  `SectionStyle.flipPageOrientation`. The rest of the report (agenda, TDoc tables, minutes) stays
  portrait. Text typed below the summary gets a portrait section again.
- **Page size:** the document's. The Docs API can change only the orientation of a section, so
  an A4 report gets A4 landscape and a Letter report gets Letter landscape.
- **Column widths:** taken from the real page (its height less the side margins), with the Title
  widest. The tables flow over as many landscape pages as they need.
- **Repeated updates** add no break and change no orientation: each step is checked against a
  fresh read.
- **When it is set:** after every update, Build from Scratch, Update Disposition Summary and Close
  Report.

## 4. The revision relationship in the TDoc row

`S4-26XXXX (revision of S4-26YYYY)`, with both numbers linked where the TDoc list has their
links.

**Every report has it,** whether Goes to Plenary is set up or not: a general part of the TDoc row,
written by every update and every Build from Scratch (`syncTdocRevisionRelationships_()`).
- Nothing else of a report changes with it: not its Status, its dropdowns, its summary or its
  workflow.
- A TDoc cell that holds anything but the number (and this code's own " (revision of …)") is
  never touched.

- **Source.** Only the TDoc list's own metadata counts: the predecessor's "Revised to" names the
  revision, or the revision's "Is revision of" names the predecessor. It never comes from
  numbering, titles, or a newer file in a revisions folder.
- **Several generations.** Each revision names its immediate predecessor (001 ← 012 ← 013).
- **Shown without a link:** a predecessor of another meeting.
- **Nothing shown:** two different predecessors, or a cycle.
- **Kept as it was:** the Revisions row and the minutes.
- **Updates.** Every update with the TDoc list refreshes the relationship, and removes it when
  the list no longer says so.
- **Readers.** Every reader of the TDoc number reads the part before " (revision of …)"
  (`safeCellText_`, the Docs API index, the upload completion).

## Migration

**Reports not set up yet:** run *Report › Set Up Goes to Plenary…*. It asks first, makes every
Status row compact (☐ No), retires the Status "Plenary", and is safe to run again. The next
update fills in A.I., the revision relationships and the four landscape tables.

**The disposable MBS copy (the first candidate's two-row layout):** the next update or Set Up
migrates it automatically:

1. The selection of each "Goes to Plenary" row is kept in `PLENARY_CHOICE_PENDING:<TDoc>`, and a
   manual "Plenary Agenda Item" becomes `PLENARY_AGENDA_MANUAL`.
2. Then the two rows are removed and the compact row is written.
3. A selection stays pending until its dropdown stands. A failed end of the run never loses one;
   the next update completes it.

**Rollback for one report:** restore the version named before the set-up.

## Why "reserved" showed as plain text

**The cause, most likely, was where the run stopped.** An update writes every Status as text
first: a new TDoc's status, and a status the TDoc list changed. Its **last** step turns those
texts into dropdowns (`finalizeStatusDropdowns_()`). The update of 2026-11-18 was stopped by the
time limit during the formatting, which comes *before* that step, so:

- the new TDoc S4-262302 kept its Portal status as text ("reserved");
- S4-262301's change from reserved to withdrawn was not written into its dropdown.

Dropdowns that existed before ("revised", "available") were untouched.

Nothing in the compact row converts a Status dropdown to text, and the migration and the
revision reordering do not either. Tested for every option (§ O), in a compact row and in a
moved revision table.

**The fixes:**
- The update now stays well inside the limit (§ 0).
- Past 4–5 minutes it skips the work that can wait, so the dropdown step always runs.
- The existing repair makes every text Status that has an option a dropdown again at the next
  update.

The integrity check lists any that are left. Without access to the document, I couldn't confirm
from its content that only S4-262302 is affected; run the check first.

## Recovery after an interrupted update

An update stopped by the six-minute limit can leave part of its work undone:

- a TDoc it inserted keeps a text Status and no compact row;
- a Portal status change is not applied to a dropdown;
- the summary is not rewritten.

It never leaves half a table or a duplicate. The next update completes all of it, idempotently
(`tests/goes-to-plenary.test.js` § N):

- **No duplicates.** The TDoc is found by its number, so no second table is made, and the
  registration table keeps one row.
- **Dropdowns.** A text Status that has an option becomes a dropdown again (the existing repair
  of the status dropdowns, also in a compact row). The missing compact row is written.
- **Summary.** It is compared and rewritten once. Its tables and section breaks are recognised
  as its own, so they are never duplicated.

**Before testing further, run *Advanced and Repair › Check Report Integrity*.** It reads the
report once through the Docs API and changes nothing. It lists:
- TDocs with two tables;
- statuses that are text although they have an option (a dropdown missing);
- compact rows without their Plenary dropdown;
- rows left from the two-row layout;
- selections still pending;
- the number of summary headings;
- the section breaks (landscape or portrait);
- text typed into the summary.

Then run *Update Report Now* once; it repairs what the check lists, except a second summary
heading or text typed into the summary, which are left for you. Run the check again.

## Tests

`node tests/goes-to-plenary.test.js` runs the real build, update, set-up, Update TDoc Revisions,
summary and Close Report on the fake document and fake Docs API. The fake now models cells that
mix text and several dropdowns, `insertText`, and section breaks.

| | What is tested |
|---|---|
| A | Every mapping rule on S4-261857; the cell parsers; predecessors (chain, conflict, cycle, other meeting) |
| B | A report that is not set up: unchanged; an ad-hoc report |
| C | Set-up: compact rows, exact layout, status readers, retired "Plenary", writes, run again |
| D | A.I. only for ☑ Yes; the four tables; landscape section and widths; revision relationships and links |
| E | Repeated updates: same document, zero API writes, nothing duplicated |
| F | Status changes, the Portal sync of a text status in a compact row, new TDocs, Update TDoc Revisions |
| G | Manual A.I., hidden while ☐ No, "-", cleared, work-item changes, agenda unreadable |
| H | Text below (portrait again) and into the summary; unreadable selection; Update Disposition Summary |
| I | Build from Scratch keeps selections, manual values, revisions; landscape after the next update |
| J | Close Report |
| L | Migration of the two-row layout; a failed end of run keeps the selection pending |
| M | DocumentApp and Docs API calls of an update of 75 TDocs: the second candidate against this one (fake document; not a time) |
| N | An update interrupted after inserting a TDoc: the next one completes it, with no duplicate |
| O | Every Status option (reserved included) keeps its dropdown in the compact row; a lost dropdown is restored; Check Report Integrity |
| P | An update that has run long defers the summary and formatting, and still completes the dropdowns |
| K | Ledger, never throws, nothing fetched but the agenda, no `setText` on a Status cell, the summary read through the API only, local listing agent untouched |

**All of this is mocked:** a fake DocumentApp and a fake Docs API. No test here runs against
Google Docs. The real integration test is the disposable MBS copy (docs/ROLLOUT_T-2026.10.30.md,
step 0).

## What only the disposable copy can show

The fake API models what the code relies on, but neither Google Docs' rendering nor its time:

0. How long two consecutive *Update Report Now* take, read from the execution log's [PERF]
   lines. The second one should be well under the first.

1. The compact row's look: the spacing of the em spaces, and the two dropdowns side by side.
2. The landscape section:
   - Is the Disposition Summary on landscape pages, with the rest portrait?
   - Is there no blank page? (Google Docs puts an empty paragraph before a section break.)
   - Is the section kept when the summary is rewritten?
3. Removing "Plenary" from the Status dropdown keeps every other selection.
4. Does the SA4#138 TDoc list have "Related WIs", "Release" and "Is revision of" under those
   names? Without "Is revision of", only the predecessor's "Revised to" gives the relationship.
