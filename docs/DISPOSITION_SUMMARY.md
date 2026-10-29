# Disposition Summary and Close Report

**Status: CANDIDATE — not released, not deployed.** Branch
`feature/disposition-summary-close-report`, based on `origin/master` `e2ceddc` (release
T-2026.10.15, `Code.js` 2.25.0). `Code.js` keeps its version line; a release adds it.

Template reports only have the two menu items (SA4 Report › 📄 Report). The functions exist in
CENTRAL and Legacy too, but nothing there calls them.

## What it does

**Update Disposition Summary** writes one section at the end of the report:

```
Disposition Summary                                   (Heading 1)
Written from the Status of the TDoc tables by SA4 Report > Report > Update Disposition
Summary, and kept up to date by every update. Text typed into this section stops it from
being updated.
Agreed and Endorsed Documents (17)                    (bold)
  TDoc | Title | Source | Status                      (table; "None." without TDocs)
Other Disposed Documents (12)
  TDoc | Title | Source | Status
Documents Not Yet Disposed (0)
  None.
```

- Every TDoc table of the report is one row in exactly one table, in the order of the report.
  A TDoc with two tables is one row.
- Title, Source and Status are those of its TDoc table; the TDoc number has the link of its
  TDoc table. Agreed and Endorsed keep their own status.
- Nothing else is written: no status, no revision, nothing in 3GU, no minutes.

**Close Report** writes the same summary (creating it where there is none), checks the report
and says whether it is ready to close. See § Close Report.

## The classification

The disposition of a TDoc is the Status of its TDoc table, read as every reader reads it
(`detectTdocTablesInDocument_()` → `readTdocStatus_()`: the text of the cell, else the value
selected in its dropdown). The vocabulary is the 14 options of the "Document Status" dropdown
of the master template (`docs/STATUS_DROPDOWN_RELEASE_CANDIDATE.md`), plus the Portal wording
"replied to" (`STATUS_DROPDOWN_ALIASES_`). One constant, `DISPOSITION_OF_STATUS_`, classifies
them, and the summary, Close Report and its dialog all use it:

| Class | Statuses | Why |
|---|---|---|
| Agreed and Endorsed | agreed, endorsed | as asked |
| Other disposed | noted, revised, merged, approved, withdrawn, replied, postponed | a decision of the meeting on the document; a revised document is disposed and its revision has a row of its own |
| Not yet disposed | available, reserved, parked, Plenary, other | no decision yet (available, reserved), deferred within the meeting (parked), left to the plenary (Plenary), or no meaning (other) |
| Not yet disposed, "not a known status" | anything else: no status, a typed text such as "agreed, see minutes", a Portal status without an option ("not pursued"), two tables with different statuses | fail safe: nothing counts as disposed unless it is listed |

Case and whitespace are ignored ("  AGREED " is agreed).

Decisions to confirm before release:
- **approved** is "other disposed", not in the Agreed and Endorsed table (only agreed and
  endorsed were asked for). One line moves it.
- **Plenary** is "not yet disposed" (the decision is taken elsewhere). If a report counts it
  as disposed, one line moves it.
- **postponed** is "other disposed" (a decision of this meeting).

## Where it stands

Created at the end of the document. For MBS that is:

```
2.8 Close of the session
    (TDocs of 2.8, text typed under the close)
    Attendance                    (Heading 3; the generated section)
Disposition Summary               (Heading 1)
```

The invariant: **nothing of the agenda and no Attendance section stands after the summary.**
It holds for what an update, a build or an import writes afterwards:

| What is written | Where | Function |
|---|---|---|
| a TDoc of the closing item | before the Attendance section, and before the summary | `findAgendaSectionEndIndex_()` |
| the heading of an agenda item the report does not have | before the close; without close and Attendance, before the summary | `findOrderedAgendaInsertIndex_()` via `findReportTerminalBoundary_()` |
| a new Attendance section | at the end of the closing section, which ends at the summary heading; without a close, before the summary | `findAdhocAttendanceInsertIndex_()` |
| Refresh Attendance | in place; its section ends at the summary heading | unchanged |

**Its section** runs from its heading to the next heading. It is rewritten only when what it
would say differs from what it says (otherwise not touched: no revision, no cursor jump for
the people editing), and only when it holds nothing but what it writes and empty paragraphs.
Text typed into it is never removed: the summary is then not written and the dialog (or the
update log) says why. Text under a heading of its own after the summary is outside the
section. The empty paragraphs at its end are left as they are, so a rewrite adds none.

## During the meeting

- **Only the menu creates a summary.** Update Disposition Summary and Close Report write one
  where there is none.
- **Every update keeps one that is there up to date** — Update Report Now and the automatic
  update — and never writes one where there is none. A report without a summary (every
  existing report, MBS and 6G today) is updated exactly as before; its update result is
  unchanged.
- Not in a background run of the CENTRAL scheduler (`continuousUpdateCore_(context)`): no
  status dropdowns can be read there. Template reports do not run there.
- A status dropdown that cannot be read (Docs API failure) stops the write: the summary is
  not written with unknown statuses.
- **Build Report from Scratch** clears it with the report and says in its result: "The
  Disposition Summary was cleared with the report. Use … Update Disposition Summary to write
  it again." It is not rebuilt automatically: the statuses the rebuild restores into the
  dropdowns are only there after its last step.
- **Report Status Summary** does not count the summary's tables as TDocs.

## Close Report

Close is a check, not a state. It does **not** make the document read-only, remove a trigger,
change sharing, remove data, store a "closed" flag, change any status, or touch 3GU. It can be
run any number of times; the report can be corrected afterwards. (There is no report-state
concept in the project to record a close in, and none was added.)

It:
1. reads the TDocs and writes the summary;
2. downloads the TDoc list of the meeting (the same `downloadAndGroupTdocs_()` as an update;
   after the configuration guard, so never the built-in SA4#136 defaults);
3. says **Report ready to close** or **Report NOT READY TO CLOSE**, with the counts.

Not ready while any of these holds — each is listed with what to do:

| Check | Listed as |
|---|---|
| a TDoc not yet disposed (incl. an unknown status) | TDoc \| Title \| Source \| Current Status (the dialog shows 40; the summary table has all) |
| revised in the report, not in 3GU (below) | Original TDoc \| Revision \| Source \| Report state \| 3GU state, and the action |
| something could not be checked: a status, the TDoc list, the summary (text typed into it) | the reason |

### Revised in the report, not in 3GU

From the TDoc list of real meetings (meeting 85916: 15 TDocs with a "Revised to", all with the
status `revised`): when a revision is reserved in 3GU "as revision of" a TDoc, 3GU gives the
TDoc the status `revised` and the revision number in "Revised to". So, for a TDoc whose
**Status in the report is revised** (classified as above; the visible text "revised" anywhere
else is not used):

| TDoc list (3GU) | Result |
|---|---|
| status revised, with or without "Revised to" | consistent |
| no "Revised to" and any other status (available, reserved, agreed, …) | **revision-not-reserved** — "Author needs to reserve/upload the revision in 3GU." |
| "Revised to" B, and B not uploaded: no hyperlink on B in the list (the upload sign of the upload completion, `tdocListLinkUrl_()`), or without a readable link B's status reserved | **revision-not-uploaded** — "Author needs to upload the revision B in 3GU." |
| the TDoc is not in the TDoc list | **not-in-tdoc-list** — "Check the TDoc in 3GU …" |

The Revision column shows "Revised to" from 3GU where there is one, else the TDoc numbers the
report's Disposition and Revisions cells name (for information only; they decide nothing).

## Tests

`tests/disposition-summary-close-report.test.js` (114 checks): the classification of every
option and of unknown statuses; the summary (every TDoc once, counts, order, links, title,
source, status, "None.", text statuses, duplicates); idempotence (5 runs and 5 updates
element-for-element identical; seven status changes back and forth, each followed by an
update: no element and no empty paragraph more); typed text never removed; an unreadable status; Close Report (clean, available,
reserved, parked, several, unknown, 45 open, TDoc list unreadable, no saved meeting, master
template, nothing changed but the summary, nothing stored); every revision/3GU combination,
on the real update and as pure rules; placement (ad-hoc with and without Attendance, Attendance
imported later, no close, text typed at the end, an annex heading after it, main meeting); the
MBS fixture end to end; a report without a summary is never given one; Build Report from
Scratch; and that without this change `Code.js` and `ReportCreator.js` are those of T-2026.10.15,
byte for byte. The ledger view `codeWithoutDispositionSummary()` in
`tests/helpers/before-structure-fix.js` takes it out for the older ledgers.

## MBS, after a release

Not done; to be done after a release has been deployed to the master template and adopted by
MBS. No Build Report from Scratch.

1. Update Report Now once (the released update; no summary is written by it).
2. Update Disposition Summary. Check: the summary is after "2.8 Close of the session" and the
   Attendance section; 29 (or more) TDocs; Title, Source, Status and links as in the tables.
3. Run it again: "already up to date"; nothing changed (version history).
4. Close Report. Read both lists.
5. Correct the statuses in the minutes; have the authors reserve/upload the revisions in 3GU.
6. Close Report again until "Report ready to close".
7. Update Report Now once more: the summary says "unchanged" in the log; nothing else changed.
