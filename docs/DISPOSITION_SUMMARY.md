# Disposition Summary and Close Report

**Status: CANDIDATE — not released, not deployed.** Branch
`feature/disposition-summary-close-report`, based on `origin/master` `e2ceddc` (release
T-2026.10.15, `Code.js` 2.25.0): candidate `4276161`, refined afterwards (postponed, the four
revision states, the end of the summary; then approved with agreed and endorsed; then the
rewrite of a summary an earlier candidate wrote; then the result of Close Report in a dialog the
execution does not wait on). `Code.js` keeps its version line; a release adds it.

Template reports only have the two menu items (SA4 Report › 📄 Report). The functions exist in
CENTRAL and Legacy too, but nothing there calls them.

## What it does

**Update Disposition Summary** writes one section at the end of the report:

```
Disposition Summary                                   (Heading 1)
Written from the Status of the TDoc tables by SA4 Report > Report > Update Disposition
Summary, and kept up to date by every update. Text typed between this line and its last
table stops it from being updated; text below it is kept.
Agreed, Endorsed and Approved Documents (17)          (bold)
  TDoc | Title | Source | Status                      (table; "None." without TDocs)
Other Disposed Documents (12)
  TDoc | Title | Source | Status
Documents Not Yet Disposed (0)
  None.
```

- Every TDoc table of the report is one row in exactly one table, in the order of the report.
  A TDoc with two tables is one row.
- Title, Source and Status are those of its TDoc table; the TDoc number has the link of its
  TDoc table. Agreed, Endorsed and Approved keep their own status.
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
| Agreed, Endorsed and Approved | agreed, endorsed, approved | the positive decisions of the meeting |
| Other disposed | noted, revised, merged, withdrawn, replied | a final decision of the meeting on the document; a revised document is disposed and its revision has a row of its own |
| Not yet disposed | available, reserved, parked, postponed, Plenary, other | no decision yet (available, reserved), deferred within the meeting (parked), no final substantive decision at this meeting (postponed), left to the plenary (Plenary), or no meaning (other) |
| Not yet disposed, "not a known status" | anything else: no status, a typed text such as "agreed, see minutes", a Portal status without an option ("not pursued"), two tables with different statuses | fail safe: nothing counts as disposed unless it is listed |

Case and whitespace are ignored ("  AGREED " is agreed).

Confirmed: **approved** is in the first table with agreed and endorsed (it was "other disposed"
in the first candidates); **Plenary** and **postponed** are "not yet disposed", so Close Report
shows them instead of letting the report look clean. The key of the first class is
`AGREED_ENDORSED_APPROVED`; its count reads "agreed/endorsed/approved".

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

**What it owns** is exactly what it writes, read from its heading on: the note, the three class
lines, each followed by its table or "None.", and empty paragraphs between them. It ends with
the table (or "None.") of the last class. That end is the boundary:

| Text typed … | What happens |
|---|---|
| below the summary, with or without a heading of its own (also a line "None.") | not the summary's; never touched. The summary is updated above it. |
| into it — between its heading and its last table, so that a part of it stands below the text | the summary is not written, nothing is removed; the dialog (or the update log) names the text and says to move it |
| above its heading | outside it |

A summary table deleted by hand is written again whole.

**A summary an earlier candidate wrote** is the summary's own too: the lines that candidate wrote
differently — the first class line "Agreed and Endorsed Documents (N)" (`4276161`, `00a50c5`) and the
note of `4276161` — are recognised by exactly those texts (`DISPOSITION_SUMMARY_EARLIER_TITLES_`,
`DISPOSITION_SUMMARY_EARLIER_NOTES_`), and the next update rewrites the summary in the current
form, in place. Without this, the 6G smoke copy (written by `00a50c5`) was refused by `ed43e37`: the
old first line was taken for text typed into the summary. The summary's own tables are never read
as data: every update classifies the TDoc tables again, with the current rules.

A summary that is not written — refused, or a status that cannot be read — is reported under the
title **"Disposition Summary NOT updated"**, never under the title of a success. Named ranges, the other way to mark a
boundary, are used nowhere in this project; the summary is bounded the way the Attendance
section and the registration table are, by what it writes.

It is rewritten only when what it would say differs from what it says (otherwise not touched:
no revision, no cursor jump for the people editing). The empty paragraphs after its end are
left as they are, so a rewrite adds none.

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
2. when the report has a TDoc as revised: downloads the TDoc list of the meeting (the same
   `downloadAndGroupTdocs_()` as an update; after the configuration guard, so never the
   built-in SA4#136 defaults). Without such a TDoc nothing is downloaded;
3. says **Report ready to close** or **Report NOT READY TO CLOSE**, with the counts.

**How the result is shown.** In a modal HTML dialog: the counts, then every list in full, as
tables (TDoc | Title | Source | Current Status; Original TDoc | Revision | Source | Report state |
3GU state | Action), scrollable, with a Close button. Text from the report is escaped. The same
result is written to the execution log as text. The dialog is opened after the document lock is
released, and the execution ends with it: unlike an alert, a modal HTML dialog does not keep the
execution waiting. (With an alert, the 6G smoke copy — 114 TDocs, 15 not yet disposed, 9 revision
findings — computed and logged the result at 00:01:29 and then waited on the alert until "Exceeded
maximum execution time" at 00:06:12.)

Not ready while any of these holds — each is listed with what to do:

| Check | Listed as |
|---|---|
| a TDoc not yet disposed (incl. an unknown status) | TDoc \| Title \| Source \| Current Status (the dialog shows 40; the summary table has all) |
| revision not reserved in 3GU (below) | Original TDoc \| Revision \| Source \| Report state \| 3GU state, and the action — a list of its own |
| revision reserved but not uploaded in 3GU | the same, a list of its own |
| revision state could not be verified in 3GU | the same, with the reason, a list of its own |
| something could not be checked: a status, the summary (text typed into it) | the reason |

The dialog counts each of them separately (`Revision not reserved in 3GU: 1.` …).

### A TDoc the report has as revised

**What "revised" means — the evidence.** The check rests on one invariant: *in 3GU a TDoc has the
status revised when, and only when, a revision is reserved as "revision of" it, and that
revision is then its "Revised to".*

- Real TDoc list, meeting 85916 (`tests/fixtures/meeting-85916-tdoc-list-values.json`): 15
  revisions of TDocs of the same list. All 15 "Is revision of" are mirrored by the original's
  "Revised to", all 15 originals have the status revised, and no TDoc is revised without a
  "Revised to" (statuses: 77 available, 15 revised, 2 reserved, 1 withdrawn).
- Real TDoc list, meeting 86172 (MBS): 5 reserved TDocs are revisions of main-meeting TDocs
  (S4-…) — a revision can be in another meeting's list. Such a case is "could not be verified",
  never "not reserved".
- The code: the agenda resolver (`selectAgendaCandidate_()`) reads revised the same way — the
  TDoc superseded by its "revision of" (meeting 60778). The report's Status becomes revised in
  only two ways: the Portal status sync (`applyTdocStatusUpdate_()` and its dropdown variant, from
  a 3GU status revised — consistent by construction), or a person choosing it. Draft revisions
  in the Inbox/Drafts folder never set a Status (their tables are inserted as Available).
- 3GPP practice: a revision of a TDoc is a new TDoc number; drafts (r1, r2) are not.

So a person choosing revised before the revision is reserved is exactly the case to report;
anything the list does not show reliably is "could not be verified", not a guess.

For a TDoc whose **Status in the report is revised** (classified as above; the text "revised"
anywhere else — Disposition, Minutes — decides nothing):

| TDoc list (3GU) | State | Close Report says | Action |
|---|---|---|---|
| "Revised to" B, and B has its hyperlink in the list (the upload sign of the upload completion, `tdocListLinkUrl_()`) | complete | nothing | — |
| no "Revised to", and the TDoc is still available or reserved | not reserved | Revision not reserved in 3GU. | Author needs to reserve the revision in 3GU. |
| "Revised to" B, and B has no hyperlink in the list | reserved, not uploaded | Revision reserved but not uploaded in 3GU. | Author needs to upload the reserved revision in 3GU. |
| the list could not be read (or the meeting is not saved); it has no "Revised to" column; it does not have the TDoc; it has it as revised without naming a revision; it has another decision (agreed, noted, …) or no status; the "Revised to" is not a TDoc number; B is not in the list; the links of the list cannot be read | not verified | Revision state could not be verified in 3GU. | Check the TDoc and its revision in 3GU (with the reason). |

Each of the last three blocks the close. Nothing is ever written to 3GU.

The Revision column shows "Revised to" from 3GU where there is one, else the TDoc numbers the
report's Disposition and Revisions cells name (for information only; they decide nothing).

## Tests

`tests/disposition-summary-close-report.test.js`: the classification of every option
(postponed not yet disposed) and of unknown statuses; the summary (every TDoc once, counts, order, links, title,
source, status, "None.", text statuses, duplicates); idempotence (5 runs and 5 updates
element-for-element identical; seven status changes back and forth, each followed by an
update: no element and no empty paragraph more); text typed into it never removed; prose typed below it (also "None.") kept through updates;
a summary table deleted by hand; an unreadable status; Close Report (clean, available, reserved,
parked, postponed, several, unknown, 45 open, no revised TDoc: no list read, TDoc list unreadable,
no saved meeting, master template, nothing changed but the summary, nothing stored); the four
revision states, each in its own list, on the real update and as pure rules; placement (ad-hoc with and without Attendance, Attendance
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
4. Close Report. Read the lists: not yet disposed; revision not reserved; reserved but not
   uploaded; could not be verified.
5. Correct the statuses in the minutes; have the authors reserve, or upload, their revisions in
   3GU; check by hand what could not be verified.
6. Close Report again until "Report ready to close".
7. Update Report Now once more: the summary says "unchanged" in the log; nothing else changed.
