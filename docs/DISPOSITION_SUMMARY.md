# Disposition Summary and Close Report

**Status: RELEASED AND DEPLOYED. Template release T-2026.10.16 (`Code.js` 2.26.0), release commit
`5d5699d68083b8677f4a28c63a19b24be932f4b8`, tag `template-release/T-2026.10.16`. Deployed to the master
template, MBS and 6G on 2026-10-30; the template's Document Status dropdown has "partially agreed"
(§ Release and deployment). Not yet used in a live report.** Developed on branch
`feature/disposition-summary-close-report` from `e2ceddc` (T-2026.10.15): candidates `4276161` …
`261106c`, smoke-tested on a Drive copy of the 6G report.

**T-2026.10.17 (`Code.js` 2.26.1), release commit `24c583a`, tag `template-release/T-2026.10.17`: RELEASED AND
DEPLOYED** to the master template, MBS and 6G on 2026-10-30 (docs/TDOC_LIST_IDENTITY.md § Release and deployment;
branch `fix/disposition-meeting-scope`). The live 6G summary of 150 TDocs is kept as evidence, not regenerated. The first Update Disposition Summary in live 6G (T-2026.10.16, 2026-10-30) listed 150 TDocs:
every TDoc table of the report, also those of earlier calls that the report keeps. The candidate reads
the TDocs of the configured meeting only (§ Which TDocs). Until it is released, T-2026.10.16 stays
frozen in production: no Update Disposition Summary and no Close Report in live 6G.
The same candidate refuses a TDoc list that is not shown to be the meeting's before any change (docs/TDOC_LIST_IDENTITY.md):
the 36 "earlier-call" TDocs in live 6G came from the Portal list of 86178 on 2026-10-30, not from the series.

**T-2026.10.18 (`Code.js` 2.26.2), release commit `c34a122`, tag `template-release/T-2026.10.18`: RELEASED AND
DEPLOYED** to the master template, MBS and 6G on 2026-10-30 (§ Release and deployment of T-2026.10.18)
(branch `fix/disposition-report-local-scope`, accepted candidate `b55e67e`; read-only release gate on the
exports of 2026-10-30: 6G 115 TDocs in scope, no S4aP260008…043; MBS 34, S4aI260080…113). **The summary is
report-local.** It reads the TDocs the report has registered, not the Portal TDoc list (§ Which TDocs), so
Update Disposition Summary works while the Portal returns a list that is not the meeting's (6G, 86178, since
2026-10-30). Close Report checks the TDoc list on top and is never ready without it (§ Close Report). The
update and its TDoc list identity are unchanged (docs/TDOC_LIST_IDENTITY.md).

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

- Every TDoc table **of the meeting** (§ Which TDocs) is one row in exactly one table, in the order of
  the report. A TDoc with two tables is one row.
- Title, Source and Status are those of its TDoc table; the TDoc number has the link of its
  TDoc table. Agreed, Endorsed and Approved keep their own status.
- Nothing else is written: no status, no revision, nothing in 3GU, no minutes.

**Close Report** writes the same summary (creating it where there is none), checks the report
and says whether it is ready to close. See § Close Report.

### Which TDocs: those the report has registered (after T-2026.10.17, candidate)

**Report → Disposition Summary; Portal → synchronisation; Close Report → summary + external check.**
The summary is derived from the report alone. It does not read the TDoc list, so a Portal list that
cannot be read or is not the meeting's neither stops it nor gets into it.

The scope (`reportTdocScope_(body)`) is the set of canonical TDoc numbers that

1. have a row in the report's own **Registration of Documents** table (header `TDoc | Title | Source |
   Agenda Item`, with or without `Session`; every such table of the report), and
2. are not reallocated to removed / withdrawn / n/a in the report's own **Document Reallocations**
   table (`interpretReallocationTarget_()`).

A TDoc table whose number is not in the scope is left out — of the summary, of the not-yet-disposed
check and of the revision check — and counted: "N TDoc tables not registered in this report (not in
its Registration of Documents table, or reallocated to Removed) left out." Its table is not touched.

Why the registration table is the report's scope:

- **Every build writes it** from the list it builds from: main meetings, ad-hoc meetings, the 6G
  branch of `buildSkeletonWithTdocTables()`, and the legacy `downloadAndProcessFromWeb()`
  (`createSummaryTable_()`).
- **Every update appends to it.** The TDocs of its list that the table does not have are added
  (`updateRegisteredDocumentsTable_()`), in the same run that inserts their TDoc tables and before the
  summary is kept up to date.
- **Nothing removes a row.** No code deletes one, so the table is the report's record of every TDoc
  it took on. That covers every call of a continuing report: in 6G, the call of S4aP260098 and the
  later one of S4aP260149. "Only the latest agenda" would be wrong.
- **It is written only after the list passed** the TDoc list identity (T-2026.10.17), by the same
  writes that insert the TDoc tables. It needs no agenda TDoc (MBS works the same), and it follows the
  agenda selector of a main meeting (only the selected TDocs are registered).
- **Tables pasted by hand are not registered.** A TDoc table pasted from another report is in no
  registration row, so it is left out.

What it does **not** protect against: a report contaminated before T-2026.10.17 (an update from a
wrong list wrote both the tables and their registration rows). Such a report must be repaired in the
document (live 6G was restored from its version history).

Where the scope comes from:

| Entry | Scope |
|---|---|
| Report › Update Disposition Summary | `reportTdocScope_()`; no TDoc list is read |
| an update (keeping a summary that is there) | `reportTdocScope_()`, after the update has registered the TDocs of its list |
| Report › Close Report | `reportTdocScope_()` for the summary and the TDocs checked; the TDoc list (`readMeetingTdocScope_()`) for the external check |

**Fail closed.** There is no scope, and nothing falls back to every TDoc table of the document, when:

- the report has no Registration of Documents table;
- the table has no TDoc;
- the table has **none of the TDocs of the report's TDoc tables**, so it does not describe them.

In those cases no summary of nothing replaces the one that is there. A report without TDoc tables
leaves nothing out and is no failure.

- Update Disposition Summary writes nothing: "Disposition Summary NOT updated", with the reason;
- an update leaves the summary as it is (`dispositionSummary: 'no-scope'`, logged);
- Close Report writes nothing and says NOT READY TO CLOSE: "The report could not be checked and the
  Disposition Summary was not written: … Check the Registration of Documents table of the report and
  run Close Report again."

`writeDispositionSummary_()` itself refuses without a scope (`no-scope`), and
`readReportDispositions_()` throws without one, so no caller can read the whole document.

**Before (T-2026.10.17):** the scope was the meeting's TDoc list, read for the summary. A list that
could not be read or was refused by its identity (6G, 2026-10-30: the Portal list of 86178 without
S4aP260098) stopped Update Disposition Summary.

## The classification

The disposition of a TDoc is the Status of its TDoc table, read as every reader reads it
(`detectTdocTablesInDocument_()` → `readTdocStatus_()`: the text of the cell, else the value
selected in its dropdown). The vocabulary is the 14 options of the "Document Status" dropdown
of the master template (`docs/STATUS_DROPDOWN_RELEASE_CANDIDATE.md`), plus the Portal wording
"replied to" (`STATUS_DROPDOWN_ALIASES_`). One constant, `DISPOSITION_OF_STATUS_`, classifies
them, and the summary, Close Report and its dialog all use it:

| Class | Statuses | Why |
|---|---|---|
| Agreed, Endorsed and Approved | agreed, partially agreed, endorsed, approved | the positive decisions of the meeting; "partially agreed" is a form of agreement here and keeps its own Status in the table |
| Other disposed | noted, revised, merged, withdrawn, replied | a final decision of the meeting on the document; a revised document is disposed and its revision has a row of its own |
| Not yet disposed | available, reserved, parked, postponed, Plenary, other | no decision yet (available, reserved), deferred within the meeting (parked), no final substantive decision at this meeting (postponed), left to the plenary (Plenary), or no meaning (other) |
| Not yet disposed, "not a known status" | anything else: no status, a typed text such as "agreed, see minutes", a Portal status without an option ("not pursued"), two tables with different statuses | fail safe: nothing counts as disposed unless it is listed |

Case and whitespace are ignored ("  AGREED " is agreed).

"partially agreed" is no option of the template's dropdown (the template holds that definition,
maintained by hand; the code has no option list): it comes typed, or as an option added to a
report's own dropdown (6G: S4aP260082). It is classified in the code like the options. Adding it to
the template's dropdown is a change of the master template document, for new reports only (a
report keeps the definition it was built with); not part of this candidate.

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

It (after T-2026.10.17, candidate):
1. reads the TDocs the report has registered (§ Which TDocs) and writes the summary from them;
2. takes the TDocs not yet disposed from the same entries;
3. **the external check**, always and once: it downloads the TDoc list of the meeting
   (`readMeetingTdocScope_()`). That is the same `downloadAndGroupTdocs_()` as an update, after the
   configuration guard (never the built-in SA4#136 defaults) and the TDoc list identity, and it stores
   nothing. It then checks the revisions of the TDocs of step 1 against that list.
   - If the list cannot be read, or is not shown to be the meeting's, the summary of step 1 still
     stands.
   - Every TDoc revised in the report is then "Revision state could not be verified in 3GU".
   - The dialog says "External verification (TDoc list of this meeting, 3GU): NOT completed",
     followed by "External verification could not be completed: <reason> …".
   - The report is **never** ready without this check, also when no TDoc is revised.
4. says **Report ready to close** only when all of these hold: every TDoc is disposed, no revision
   finding, nothing could not be checked, and the external check is completed. Otherwise it says
   **Report NOT READY TO CLOSE**, with the counts.

Before, under T-2026.10.17, the list was the scope of everything: without it, nothing was written or checked.

#### The second source of the external check: the Portal's TDoc list page (candidate after T-2026.10.18)

Section CLOSE REPORT: THE PORTAL TDOC LIST PAGE AS FALLBACK EVIDENCE of `Code.js`;
`tests/close-report-portal-page-fallback.test.js`. Branch `fix/close-report-html-fallback`.

**Why.** 6G, 2026-10-30: the Excel export of the meeting's TDoc list,
`GenerateDocumentList.aspx?meetingId=86178` — the configured source, the same URL the Portal's "Download
Tdocs to Excel" button offers — served, from the afternoon on and still on 2026-11-01, the 36 TDocs of the
earlier calls (S4aP260008…043, reserved Feb–Mar 2026), while the Portal's TDoc list **page**
`TdocList.aspx?meetingId=86178` showed the 116 of the meeting (S4aP260067…182, "116 contributions found,
displaying 1 to 116"). The request makes no difference (cache-busting parameter, parameter case, a browser's
User-Agent and Referer, a session that opened the page first: byte-identical answers). The identity guard
refused the export, as it must; Close Report could verify none of the 39 revised TDocs. The cause is inside the
Portal and not known. It is not tied to a closed meeting: the same day, 86172 (MBS, running) exported 34 = the
page's 34; 85916 (ended 28 Sept) exported 95 while its page showed 101 (the six newest missing).

**What.** For Close Report's external check — and for nothing else: not the update, not the build, not the
legacy import, not the Disposition Summary — the page is a second source of evidence
(`readMeetingTdocScope_()` → `readPortalTdocListPageScope_()`):

| The configured list (Excel) | What Close Report does |
|---|---|
| downloaded, shown to be the meeting's | used as before; the page is not fetched. Source: Excel |
| downloaded, refused by its identity | the page is fetched for the saved `MEETING_ID` |
| cannot be had (download, conversion) | the page is fetched |
| shown to be the meeting's but with no TDoc of the meeting | as before: not completed; no page |

The page is read by its own adapter (`parsePortalTdocListPage_()`, not the agenda-discovery parser, which
skips the disabled links and keeps the first revision only) and is used only when it is **one complete
list**: the grid `rgTdocList_ctl00` is there; its pager says "N contributions found, displaying 1 to N"; there
are exactly N rows; every row has its twelve cells and a TDoc number; no number twice; every TDoc cell is a
link to the document (uploaded) or a disabled link (reserved, not uploaded). A login page, an error page, an
empty answer, a paged or truncated list, a dropped cell, a mangled number: refused, with the reason. Read of a
row: TDoc, uploaded or not (with the document's link), Title, Source, Type, For, Status, Agenda item and its
title, "Revision of", and the **whole** "Revised To" cell ("S4aP260174 , S4aP260175"). Not read, not
invented: Contact, reservation and upload times (the page has none; Close Report needs none).

Then the **same chain** as for a downloaded list: the rows are put in the shape of the Excel rows (the header
of the export's names; the link cell known for every TDoc) → `assertTdocListIdentity_()` on the raw rows
(`'read'`: nothing stored; the agenda TDoc, else the anchor, else every TDoc table of the report) → the agenda
selector and the Document Reallocations (`portalTdocListPageRecords_()`, the filter of
`downloadAndGroupTdocs_()`) → `checkRevisionsIn3gu_()`. The rows of the two sources are never combined. A
count proves nothing: a page of 116 other TDocs parses and is refused for want of S4aP260098; a wrong page
with TDocs in common with the report is refused too.

When the page cannot be used either, the report is NOT READY as before, and the dialog names both: "External
verification could not be completed: the TDoc list of this meeting could not be read: <reason>; the Portal
TDoc list page (TdocList.aspx) could not be used either: <reason>." When it is used, the dialog says "External
verification … completed. Source: the Portal TDoc list page (TdocList.aspx); the configured TDoc list could not
be used.", and the log "External TDoc verification source: Portal HTML fallback (…)" (else "… source: Excel").

**On the real page of 86178** (`tests/fixtures/meeting-86178-tdoclist-page-2026-10-30.html`, the captured
page with the view state blanked and the meeting combo trimmed; `…-tdoc-list-wrong-2026-10-30.json`, the
wrong export of the same evening): 116 rows, the 10 disabled ones among them (S4aP260099, 172, 173, 174, 175
…), S4aP260140 revised to "S4aP260174 , S4aP260175", identity by S4aP260098. On a report built from them, with
099 reallocated to removed as live 6G has it: S4aP260083 → 172 withdrawn: not reserved; S4aP260140 → 174
withdrawn, 175 reserved not uploaded: not uploaded, naming 175; S4aP260074 → 099 (not in the list) and 148
agreed, uploaded: complete. None "TDoc list not read".

**Residual risk, not addressed here: an export that passes the identity but is stale.** 85916 on 2026-11-01:
the export has 95 TDocs, the page 101, and the export passes the identity (its agenda TDoc is in it). Close
Report then uses the export and never looks at the page. A revision reserved after the export's state is "not
in the TDoc list" (could not be verified — fail closed); a TDoc withdrawn after it could still be judged live.
Proving freshness would need an invariant the Portal does not give (nothing in either source dates itself);
none is invented. This change solves "the export cannot be shown to be this meeting's list", nothing more.

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
| every live revision of "Revised to" has its hyperlink in the list (the upload sign of the upload completion, `tdocListLinkUrl_()`); withdrawn ones aside | complete | nothing | — |
| no "Revised to", and the TDoc is still available or reserved; **or** every revision of "Revised to" is withdrawn | not reserved | Revision not reserved in 3GU (with "No live revision is reserved in 3GU; revision … is withdrawn" when so) | Author needs to reserve the revision in 3GU. |
| a live revision of "Revised to" has no hyperlink in the list | reserved, not uploaded (naming the live, unuploaded revisions only) | Revision reserved but not uploaded in 3GU. | Author needs to upload the reserved revision in 3GU. |
| the list could not be read (or the meeting is not saved); it has no "Revised to" column; it does not have the TDoc; it has it as revised without naming a revision; it has another decision (agreed, noted, …) or no status; the "Revised to" is not a TDoc number; a revision it names is not in the list — unless a **later** live revision is uploaded and every live revision is (below); the links of a live revision cannot be read | not verified | Revision state could not be verified in 3GU. | Check the TDoc and its revision in 3GU (with the reason). |

"Revised to" can name several revisions: 3GU keeps every revision ever reserved, also those withdrawn
since (6G: S4aP260140 "revised to S4aP260174, S4aP260175", 174 withdrawn, 175 reserved). Every number
is read (`findSA4DocumentIdsInText_()`); a withdrawn revision is an abandoned reservation, not a live
one, and can no longer be uploaded. The older readers of "Revised to" (revision placement, the
auto-written disposition) still read the first number only: BUGS_AND_ISSUES.md, out of scope here.

A number of "Revised to" that the report's TDoc list does not have is set aside only when the rest
proves the revision complete: every live revision the list has is uploaded (links readable), and one
of them comes after every missing number in "Revised to" (3GU appends reservations in their order).
6G: S4aP260074 "revised to S4aP260099, S4aP260148" -- 099 is withdrawn in 3GU and was removed from
the report by a reallocation, so the list the report reads does not have it; 148 is agreed and
uploaded: complete. A missing number alone, with a live revision not uploaded or unreadable, with
only withdrawn revisions, or listed after the uploaded one: could not verify.

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
parked, postponed, several, unknown, 45 open, no revised TDoc: the list read once, TDoc list unreadable,
no saved meeting, master template, nothing changed but the summary, nothing stored); the four
revision states, each in its own list, on the real update and as pure rules; placement (ad-hoc with and without Attendance, Attendance
imported later, no close, text typed at the end, an annex heading after it, main meeting); the
MBS fixture end to end; a report without a summary is never given one; Build Report from
Scratch; the meeting scope (section M: an accumulated report with TDoc tables of earlier calls —
available, reserved, parked, postponed, revised, agreed — and one reallocated to Removed: only the
meeting's TDocs, once each; the earlier ones neither block nor are revision-checked; the meeting's
revised TDocs fully checked; idempotent; the list followed; fail closed for an unreadable and an empty
list on all three entries); and that without this change `Code.js` and `ReportCreator.js` are those
of T-2026.10.15, byte for byte, and that against T-2026.10.16 only `continuousUpdateCore_` changed
outside the section. The ledger view `codeWithoutDispositionSummary()` in
`tests/helpers/before-structure-fix.js` takes it out for the older ledgers. Section M is, after
T-2026.10.17, a continuing report with two calls, tables pasted by hand and one reallocated to Removed.

`tests/disposition-report-local-scope.test.js` (after T-2026.10.17) runs 17 scenarios, once on this
`Code.js` and once on the `Code.js` of T-2026.10.17 (from git):

- **Normal reports:** the correct list, the Portal unavailable (HTTP 503), and the list of another
  meeting (with and without TDocs in common).
- **Continuing and MBS-style reports:** a continuing report of two calls with two agenda TDocs; an
  MBS-style report without an agenda TDoc (anchor), with a wrong list that overlaps it.
- **The summary's own section:** idempotence; a deleted summary or summary table written again; prose
  below it kept; text typed into it refused.
- **Update Report Now** still refuses a wrong list, with zero mutation.
- **Close Report** writes the summary but is not ready without the external check, and ready with it.
- **Fail closed:** no registration table; a registration table without the report's TDocs; tables
  that are not registered.

`tests/close-report-portal-page-fallback.test.js` (candidate after T-2026.10.18): the parser on the real page
of 86178 and what it refuses (login, empty, no grid, no count, none found, paged, truncated, a TDoc twice,
eleven cells, a mangled number, a cell that is no link; a count without the agenda TDoc); the real case (the
wrong export, the real page, a report of the 116: the fallback, 083 / 140 / 074, nothing but the summary
written, the two other entries never fetch the page); the conditions (Excel valid: no page; refused: page;
unavailable: page; both wrong: not ready with both reasons and zero mutation; HTTP 503, login, paged; a
plausible count without the agenda TDoc; a wrong page with TDocs in common); MBS without an agenda TDoc (the
anchor accepts and refuses the page as it does a list; the report bootstrap; a reader stores nothing); the
update never fetches the page and refuses the wrong list as before; the summary fetches nothing; and the
ledger `codeWithoutCloseReportPageFallback()`, which gives the `Code.js` of T-2026.10.18 byte for byte.

The released code fails exactly the 11 scenarios this change is for and passes the other six. The
ledger `codeWithoutReportLocalScope()` gives the `Code.js` of T-2026.10.17 byte for byte. Against it:
the identity section, `downloadAndGroupTdocs_`, `downloadAndProcessFromWeb`, the classification and
the summary's renderer are unchanged; `continuousUpdateCore_` changed in its one call of the sync only.

## Release and deployment

Release T-2026.10.16, 2026-10-30.

| Step | Result |
|---|---|
| Release commit | `5d5699d68083b8677f4a28c63a19b24be932f4b8` on the accepted candidate `261106c`: the version line 2.26.0 (2026-10-30) and its changelog entry; without them, `Code.js` is the candidate's. The plain version checks of the tests say 2.26.0; the ledger views take the header out with the feature |
| Tests on the release commit | 111 files, 7,573 checks, 0 failures; the release-integrity suites pass |
| Tag and publication | Lightweight tag `template-release/T-2026.10.16` on the release commit. `origin/master` fast-forwarded `e2ceddc` → `5d5699d` and the tag pushed in one atomic push, without force. Every older tag is where it was |
| Release bundle | `tools/template-release.js --release T-2026.10.16` from the tagged commit: the five production files; the four source files equal to the tag (line ends aside); `Release.js` T-2026.10.16, template, 2.26.0, the release commit, the tag, the template's document and script |
| Master template runtime | Container and current runtime (T-2026.10.15) proven, one normal push, no `--force`; read back: exactly five files, four byte-identical, the manifest identical as JSON, runs T-2026.10.16 |
| MBS runtime | The adoption bundle from the tag; proven (its container "MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e", T-2026.10.15); pushed and read back as above. Script files only: no report content, Document Property, trigger or sharing was touched |
| 6G runtime | The adoption bundle with `--allow-legacy-target`, authorised for this release only; proven immediately before the write (its container "6G Media Minutes – SA4-e (AH) on FS_6G_MED", T-2026.10.15); pushed and read back as above. Script files only |
| Template dropdown | "partially agreed" added to the template's one "Document Status" definition with the colours of "agreed" (`updateDropdownDefinitionProperties`, options only, tied to the revision read). By a one-off standalone helper project, not a production file. Rehearsal on a Drive copy of the template: passed (the copy kept its definition). The template: passed. A separate read-only verification: one definition, 15 options, "partially agreed" once with its own option ID and the colours of "agreed", the revision of the apply. As run and reported by the person who started the helper's functions in the script editor |

Not changed: the dropdowns of existing reports (a report keeps the definition it was built with:
6G and MBS keep theirs); no report was built, updated or closed by the release. No Update
Disposition Summary and no Close Report has run in a live report yet. The runtime to roll back to
is T-2026.10.15 (its bundles in `dist/`).

Left for cleanup, not deleted: the helper project "ONE-OFF Dropdown Helper T-2026.10.16 – DELETE
AFTER" and the rehearsal copy "DROPDOWN REHEARSAL – SA4 Report Template copy – T-2026.10.16 –
DELETE AFTER" (with the copy of the template's script it carries), in the releasing user's Drive.

## Release and deployment of T-2026.10.18

Release T-2026.10.18, 2026-10-30: the report-local Disposition Summary (§ Which TDocs) and Close Report's
external check (§ Close Report).

| Step | Result |
|---|---|
| Release commit | `c34a122fcb3414c98c0cd6a6f54467be14ac3efa`, on the accepted candidate `b55e67e`. It adds the version line 2.26.2 (2026-10-30) and its changelog entry; without them, `Code.js` is the candidate's. The ledger (`codeWithoutReportLocalScope()`) takes the 2.26.2 header out with the change and gives the `Code.js` of T-2026.10.17 byte for byte; the version-pinned tests say 2.26.2 |
| Release gate (read-only) | On the exports of 2026-10-30, with the candidate's own functions. 6G (20:35): 115 TDoc tables, 116 registered (S4aP260067…182; 099 reallocated to removed, no table), scope 115, none of S4aP260008…043, S4aP260098 and S4aP260149 included. MBS (19:47): 34 tables = 34 registered = S4aI260080…113, scope 34 |
| Tests on the release commit | 113 files, 7,708 checks, 0 failures. Among them: disposition-report-local-scope 50, disposition 223, tdoc-list-identity 52, release-file-set 48, template-release 26 |
| TDoc list identity | Not weakened: against T-2026.10.17 the identity section, `downloadAndGroupTdocs_`, `downloadAndProcessFromWeb` and `buildSkeletonWithTdocTables` are byte-identical; `continuousUpdateCore_` changed in its one call of the summary sync only |
| Tag and publication | Lightweight tag `template-release/T-2026.10.18` on the release commit. One atomic push, without force: `origin/master` fast-forwarded `d19e52d` → `c34a122`, and the tag published. Every older tag is unchanged (T-2026.10.16 `5d5699d`, T-2026.10.17 `24c583a`) |
| Release bundle | `tools/template-release.js --release T-2026.10.18` from the tagged, clean commit: the five production files. The four source files are equal to the tag (line ends aside). `Release.js` names T-2026.10.18, template, 2.26.2, the release commit and the tag |
| Adoption bundles | `--adopt T-2026.10.18` for `mbs-live`, and for `6g-live` with `--allow-legacy-target`, authorised for T-2026.10.18 only. Both are byte-identical to the release bundle |
| Before the pushes | Each target was pulled read-only into its own scratch folder: all three ran T-2026.10.17 (`24c583a`). Each container was proven with the Apps Script API (`projects.get`, read-only): "SA4 Report Template", "MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e", "6G Media Minutes – SA4-e (AH) on FS_6G_MED". Their last update was the T-2026.10.17 push |
| Master template runtime | Pushed at 21:48:50 (CET) |
| MBS runtime | Pushed at 21:49:14. Script files only |
| 6G runtime | Pushed at 21:49:26. Script files only |
| Read-back | A fresh pull of each of the three targets: exactly five files, four byte-identical (line ends aside), the manifest identical as JSON. `Release.js`: T-2026.10.18, 2.26.2, `c34a122`, tag `template-release/T-2026.10.18` |

Every clasp command used `--user haudiobe.ts@gmail.com` and ran from its own scratch folder. None ran
from the repository root, and none used `--force`.

**The deployment ran nothing in the reports:** no update, no Update Disposition Summary, no Close
Report, no build, no session action. No trigger was created, started or stopped; the 6G automatic
update stays off. 3GU was not touched. The runtime to roll back to is T-2026.10.17 (its bundles in
`dist/`).

**The one live smoke, by Thomas:** SA4 Report › Report › Update Disposition Summary in live 6G. The
Portal still returns a list that is not the meeting's; the summary does not read it. Expected:

- "Disposition Summary written at the end of the report."
- 115 TDocs: 16 / 84 / 15, unless a status changed after the 20:35 export.
- None of S4aP260008…043.
- Nothing outside the summary changed.

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
