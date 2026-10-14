# The structure of a report

**Status: RELEASED AND LIVE. Template release T-2026.10.13 (`Code.js` 2.23.0), fix commit
`319084a`, release commit `4668c8eeb3299d876d5c29811284ede8a129bc38`, tag
`template-release/T-2026.10.13`. Deployed to the master template and adopted by the 6G and MBS
live reports on 2026-10-14; the live MBS report was repaired by its first update
(§ Verification).**

## The rule

A report is, in this order:

1. the agenda sections, in the order of their numbers (2.9 before 2.10, 2.5 before 2.5.1
   before 2.6);
2. what closes it: "Close of the session" and, in an ad-hoc report with imported attendance,
   the Attendance section.

Nothing of the agenda — no agenda heading, no TDoc table — stands after the close or after
the Attendance section. An update that runs twice leaves what it left the first time.

## What was wrong

Found in the MBS report of the ad-hoc meeting after SA4#137-e, exported on 2026-10-13: of
1,230 body elements 1,157 were empty paragraphs, 1,129 of them in three runs (571, 215, 343);
and after "2.8 Close of the session" and the Attendance section came a heading "2.7" with
four TDocs and a heading "2.6" with two.

Two separate defects of the update, both in `Code.js` since before the template releases:

| | Blank lines | Sections after the attendance |
|---|---|---|
| Where | `moveTableAfter_()`, called by `rearrangeRevisionTables_()` at the end of every update | `findInsertionPointForAgendaItem_()`, called by `insertNewTdoc_()` |
| What it did | Asked whether the revision is the *very next body child* after its parent | Appended the heading of an agenda item the report did not have to the *end of the document* |
| Why that fails | Google Docs keeps a paragraph between two tables, so it never is. Every revision was "out of place" on every update, was copied to below its parent again, and the original was removed — without its separator | The end of the document is after the close and after the attendance |
| Result | One empty paragraph more per revision per update, gathering below the last table of each revision chain | The section, and every later TDoc of that item, after the Attendance section; in the order the TDocs were found (2.7 before 2.6) |

The three runs stand exactly below the three revision chains of item 2.5 that existed while
the update ran on its trigger. The two late sections are the agenda items the build had not
written a heading for.

The earlier tests did not see the first defect because their document let two tables touch,
which Google Docs does not.

## What changed

- **Revision placement.** A revision is in place when it is the next *table* below its
  parent with only empty paragraphs between them. Moving a table takes its separator with it
  (`removeTableWithItsSeparator_()`), and at most one empty paragraph stays between two
  tables.
- **A heading the report does not have** is written in its place: before the first heading
  of a later agenda item; without one, before the heading that closes the report; the end of
  the document only when the report has neither close nor attendance
  (`findOrderedAgendaInsertIndex_()`). It gets the title of its agenda item where the report
  knows it (the agenda it was built from, else the TDoc list).
- **A TDoc of the closing item** goes before the Attendance section, not after it
  (`findAgendaSectionEndIndex_()`).
- **Repair, at the start of every update** (`repairReportStructure_()`), of what the earlier
  releases left:
  1. an agenda section after the close whose number is before it is moved to its place, with
     everything below its heading, in order. Tables and paragraphs are copied and the
     originals removed only afterwards, as the revision placement has always done;
  2. below a TDoc table, a run of two or more empty paragraphs becomes one.

  A report that needs neither is read once and not changed.

What it does **not** do:

- No general removal of empty paragraphs. Only runs directly below a TDoc table are
  shortened; a paragraph that holds anything — text, an image, a page break — is never
  removed; empty paragraphs anywhere else stay.
- No TDoc table is moved because its "Agenda Item" row names another item than the heading
  it stands under. (In the MBS report S4aI260089, an item-2.7 TDoc, stands under 2.5. Why is
  not known; it stays there until it is moved by hand or reallocated.)
- Build Report from Scratch is unchanged. It still writes only the agenda items of its
  agenda; the TDocs of other items are added by the first update — now in their place.
- The "⚠️ Legacy: Update All" path (`copyTableToDoc_()`) is unchanged.

One changed behaviour to know: an agenda item numbered *after* the closing item (a TDoc
registered under 2.9 when 2.8 is "Close of the session") is written directly before the
close, not after it.

## The MBS report after one update (the test)

From `tests/fixtures/mbs-report-structure-damaged.json`, the sanitized structure of the
exported report, with the fake document that keeps a paragraph between tables:

```
2.4  S4aI260094 096 098 099 108
2.5  S4aI260080 091 100 089 082 102 105 083 092 084 085 086 087 088 093 101 106 107
2.6  S4aI260095 097
2.7  S4aI260081 090 103 104
2.8  Close of the session
     Attendance
```

1,230 body elements become 104; 1,157 empty paragraphs become 31 (one between two tables,
and the few of the opening section and the end). Every TDoc table is there once, with its
minutes. A second update changes nothing.

## Verification

**Real-document smoke**, 2026-10-14, on a Drive copy of the live MBS report (the candidate
`67badd9`, which also carries the configuration guard, `docs/CONFIGURATION_GUARD.md`).
After the meeting was configured:

- first update: agenda sections 2.7 and 2.6 put back before "2.8 Close of the session",
  1,216 empty paragraphs removed, revision chains in place, the MBS sources only (29 TDocs);
  the result checked by eye and in the export, status dropdowns working;
- second update: 0 new TDocs, 0 statuses, revision placement 6 in place / 0 moved, no
  "Structure:" line, formatting skipped -- nothing changed.

An earlier copy, updated before its meeting was saved, had imported unrelated SA4#136 TDocs
(the reason for the configuration guard); its structural results were the same.

**Live, after the deployment of T-2026.10.13:**

| Report | First update | Next update |
|---|---|---|
| MBS | 2.7 and 2.6 put back before the close; 1,222 empty paragraphs removed; 29 TDocs, 0 new; revisions 6 in place, 0 moved | completed (49.7 s) |
| 6G | one empty paragraph removed; 35 TDocs, 0 new | completed (88.6 s); no "Structure:" line, so the removal of the run before had been kept |

What remains as it was:

- That Google Docs keeps a paragraph between two tables is concluded from the documents, not
  probed separately; the code does not depend on where exactly Docs puts it.
- Why the build of the MBS report had no heading for 2.6 and 2.7 is not established.
- S4aI260089, a TDoc of item 2.7, stands under 2.5 in the MBS report; an update does not move
  a table whose section exists. Not part of this fix.

## Tests

`tests/report-structure-fix.test.js`: missing items (middle, before the close, after the
close), reallocation, discovery order, nothing after close or attendance, idempotence over
forty updates, text typed by hand, a report in order is not touched, main-meeting numbering,
the MBS structure, and that `Code.js` without this fix is the `Code.js` of T-2026.10.12.
`tests/adhoc-integration.test.js` § 3 for the ad-hoc report with sessions.
