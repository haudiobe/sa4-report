# Ad-hoc sessions and attendance (release candidate)

2026-10-02. Branch `feature/adhoc-sessions-attendance` on master `48edcc1`
(T-2026.10.4, `Code.js` 2.17.4).

**Status: RELEASE CANDIDATE, G2 COMPLETE (2026-10-05, §10): the smoke test and every visual
check after it passed; the production code that was accepted is that of commit `33ded25`.
`Code.js` 2.18.0, proposed template release T-2026.10.5. Not tagged, not pushed, not deployed: the tag, the squash into
master and the deployment to the master template are separate, explicitly authorized steps.**

The design and the reasons are in `ADHOC_SESSIONS_ATTENDANCE_DESIGN.md`. This document records
what was built, where it departs from that design, and how it is released.

---

## 1. Decisions

| # | Decision |
|---|---|
| 1 | The feature is for **ad-hoc reports only**, and only once sessions are configured. A main-meeting report and an ad-hoc report without sessions behave as in T-2026.10.4. |
| 2 | No existing report gains a table, a column or text because the code changes. The Session column of an existing report is added only by an explicit action. |
| 3 | A session is **one logical meeting** and may span several calendar days (`date` to `endDate`). It has a **stable id**; label, dates and times can change. Everything else refers to sessions by id. An id is never reused. |
| 4 | **Teams' summary values are authoritative.** Duration, average attendance and the number of attendance records are shown as Teams states them and never calculated. |
| 5 | Of a Teams export only the meeting summary and, per attendee, name, company and e-mail are kept. Participant ID (UPN), dial-in numbers, join and leave times, individual durations, roles and the file itself are never stored, shown or logged. |
| 6 | The company comes from a leading `[Company] Name` or from what the rapporteur types into the attendee table. Nothing is inferred from an e-mail domain. |
| 7 | A TDoc belongs to the first session whose cut-off — its planned end on its **last** day, or the end of that day — is not before its `Uploaded` time. The Portal's `Uploaded` is **UTC** (checked against the public file server) and is converted to the time zone of the script. A report can be told that its list is on the sessions' clock instead; a stored choice is kept. |
| 8 | Manual TDoc sessions are additions or replacements per TDoc; "no session" can be set explicitly. |
| 9 | Opening details are free text per session: Chair, minute taker(s), an administrative note. Nothing is inferred or matched against the attendance. |
| 10 | Status and statistics are derived views: counts and Teams' values, read-only, nothing stored. They are shown in dialogs, not written into the report. (The design proposed a "Series statistics" block in the document; that was not built.) |
| 11 | There is no AI-generated summary of a meeting or a session. |
| 12 | Teams attendance has **one record per day** of a session: an export is accepted when its date is one of the session's days, and importing a day again replaces that day only. |
| 13 | The Session column of the registration table shows a **code** per session — `A01`, `A02`, … by the chronological order of the sessions — because a label does not fit there. The code is derived every time and **never stored**; sessions are stored and referred to by their ids as before. The Session administration section names each session with its code and its label; everywhere else a session is called by its label. |
| 14 | Report-level and session-level information are separate. With sessions, a build writes an **Online information** block (the meeting as a whole) above the Session administration section (each session). It contains only what the report is configured with; nothing is looked up or made up. |

---

## 2. The menu

A report whose meeting type is ad-hoc gets one more submenu, between *Document Reallocation*
and *Automatic Updates…*. It appears the next time the document is opened after the report
became ad-hoc.

| 🗓 Sessions and Attendance | Function |
|---|---|
| Configure Sessions… | `configureAdhocSessions` |
| Import Teams Attendance… | `importTeamsAttendance` |
| Refresh Attendance Section | `refreshAdhocAttendanceSection` |
| Assign TDoc Sessions… | `assignAdhocTdocSessions` |
| Edit Opening Details… | `editAdhocOpeningDetails` |
| Post-meeting Statistics… | `showAdhocSessionStatistics` |

The rest of the report menu, the master template's menu and the CENTRAL / Legacy menu are
unchanged. Every action refuses in the master template.

---

## 3. What a report gets

Generated parts, each at most once, each replaced as a whole and never touching the text
around it:

| Part | Where | Written by |
|---|---|---|
| **Online information** (heading, then one line each, only where the report has the value: Meeting name, Start Date, End Date, Portal meeting URL, TDoc List URL, Excel Docs URL, Docs Folder, Drafts Folder) | In the "Opening of the session" sub-section, above Session administration | **Build from Scratch only.** Afterwards an existing block is kept up to date when the sessions or the opening details are saved; a report without the block never gets one from a save or an update |
| **Session administration** (heading, then per session: its code and when it is, Chair, minute taker(s), note) | End of the "Opening of the session" sub-section, before "Registration of Documents" | Build from Scratch; saving the opening details; saving the sessions when the section or details exist |
| **Session** column of the registration table: the codes of the TDoc's sessions (`A01`; `A01, A02`; `–` for none) | Fifth column; the table then has explicit column widths | Build from Scratch with sessions; the explicit action in *Assign TDoc Sessions…*; then kept up to date by every update |
| **Attendance** (heading, then per session — and per day of a session of several days — Teams' statistics and a Name / Company / Email table) | End of the closing section | An import; *Refresh Attendance Section*; saving the sessions; Build from Scratch |

With sessions configured, Build from Scratch no longer writes the line
"<Chair> opens the session on … at <start> CEST."; the Session administration section says who
chaired and when. Build from Scratch keeps the imported attendance (with the Company cells) and
the opening details, and its confirmation says so.

The Online information block ends with its own lines: text typed below it — between it and
the Session administration section, where the opening of a meeting is minuted — is not part of
it and is never replaced. Start Date and End Date are the first day of the first session and
the last day of the last. The two Portal addresses are written only when the TDoc list URL of
the report is a Portal list that names the meeting id; a Drafts Folder only when one is
configured for the report (the folder a main meeting derives is not assumed). There is no
line for shared minutes or scribes: the report has no such value (§10.3).

An update writes neither the Online information, nor the Session administration, nor the Attendance section. A new TDoc
table is placed at a numbered heading or at the end of the document, so it can follow a
generated section but is never inside one.

**Report Status Summary** ends with a "Sessions and attendance" block for such a report;
**Post-meeting Statistics…** shows the same and the figures of each session. For this the
status summary downloads the TDoc list; if that fails, the rest of the status is shown.

---

## 4. Stored state

| Property | Store | Content | Limit |
|---|---|---|---|
| `ADHOC_SESSIONS` | report state (adopted) | `{ v, nextId, sessions: [{ id, label, date, endDate?, start, end }] }` — `endDate` only for a session of several days | 8,000 characters |
| `ADHOC_TDOC_SESSIONS` | report state (adopted) | `{ v, clock, add, set }`: TDoc numbers grouped by session ids | 8,000 characters |
| `ADHOC_SESSION_ATTENDANCE`, `ADHOC_SESSION_ATTENDANCE_<id>_<generation>_<n>` | Document Properties | an index, and per attendance record the summary and `[name, company, e-mail]` rows in chunks. A record is the session id for a session of one day, `<id>d<YYYYMMDD>` for a day of a session of several days; manual companies are under the id `companies` | 2,500 characters per chunk |
| `ADHOC_SESSION_OPENING` | Document Properties | `{ v, sessions: { <id>: { chair, minuteTakers, note } } }` | 8,500 bytes |

Every value has schema version 1. A value that cannot be read counts as "nothing stored" for
its part and is reported; a value written by a newer release is never overwritten. A session
that has attendance, manual TDoc assignments or opening details is not removed until they are
cleared. Nothing is stored for the status or the statistics.

---

## 5. Production diff against T-2026.10.4

Two files change. `HyperLink.js`, the manifest and both HTML files are unchanged, and the
release tooling is unchanged. The bundle is the same seven files.

### `Code.js` (2.17.4 → 2.18.0)

| Function | Change |
|---|---|
| header | version line, changelog entry |
| new sections before `ARCH-009` | sessions, the Teams parser, attendance, TDoc sessions, opening details, status: 143 new functions |
| `ADDON003_ADOPTION_FIXED_KEYS_` | `ADHOC_SESSIONS`, `ADHOC_TDOC_SESSIONS` |
| `downloadAndGroupTdocs_()` | adds the upload time to each TDoc — only for an ad-hoc report with sessions |
| `createSummaryTable_()` | a fifth column (the session codes) when it is given the sessions; that table gets explicit column widths |
| `updateRegisteredDocumentsTable_()` | recognises the table with four or five columns; a new row gets as many cells as the table has; a five-column table gets its column widths |
| `continuousUpdateCore_()` | passes the sessions on; keeps an existing Session column up to date; reports a failure of that step with the collector's |
| `buildSkeletonWithTdocTables()` | keeps the Company cells before clearing; builds the registration table with the Session column; leaves out the "<Chair> opens …" line; writes the two generated sections at the end — each only with sessions |
| `analyzeReportStatus()` | one call that adds the block of an ad-hoc report with sessions |

No other function of `Code.js` differs (tested against the tag).

### `template/ReportCreator.js`

| Function | Change |
|---|---|
| `buildTemplateReportMenu_()` | the submenu, for an ad-hoc report |
| `isAdhocReportForMenu_()` (new) | reads the meeting type from the document's own properties |
| `confirmTemplateBuildFromScratch_()` | names the attendance and the opening details among what is kept, when there are any |

### CENTRAL and Legacy

`Code.js` is shared, so the code is in every runtime; the menu is not. The CENTRAL / Legacy
menu is unchanged. A report without the properties of §4 runs the same code paths as before:
each hook returns at once when the report is not ad-hoc or has no sessions. One thing is
visible outside the template runtime: *Analyze Report Status* of an ad-hoc report that has
sessions shows the new block as well.

---

## 6. Departures from the design

| Design | Built |
|---|---|
| A "Series statistics" block in the report (§14) | Status block and statistics dialog; nothing in the document |
| Join / leave times and roles stored per attendee | Not stored |
| Company corrections in `ADHOC_COMPANY_MAP` | Kept with the attendance, read back from the attendee tables |
| Upload clock to be established | UTC by default, compatibility setting for other lists |
| Records without times excludable in the preview | Listed and flagged; not excludable |
| A session is one call on one date | A session is one logical meeting and may span several days; attendance has one record per day (design §3.1, added before release) |
| The Session column shows the label of a session | It shows a derived code (`A01`); the label is in the Session administration section beside the code |
| The opening: one paragraph per session, written by the build as prose | A generated Online information block for the meeting and the generated Session administration section for the sessions |

---

## 7. Tests

92 test files, 6,014 checks at the release candidate. For this feature:

| File | Covers |
|---|---|
| `adhoc-sessions-config.test.js` | session model, dialog, what did not change against T-2026.10.4 |
| `teams-attendance-parser.test.js` | the parser, on a synthetic export |
| `adhoc-attendance-import.test.js` | preview, import, storage in chunks, rendering, company cells, rebuild |
| `adhoc-tdoc-sessions.test.js` | upload times, the clock, assignment, manual sessions, the column, updates |
| `adhoc-tdoc-uploaded-display.test.js` | the read-only **Uploaded** column of *Assign TDoc Sessions…*: the clock of the assignment, shown only, never stored |
| `adhoc-registration-table-widths.test.js` | the column widths of the registration table with the Session column: build, update, adding the column, repetition; nothing else is touched |
| `adhoc-session-codes-opening-info.test.js` | the session codes (order, renaming, nothing stored) and where the label stays; the Online information block: lines, links, place, range of the meeting, missing values, typed text, reports without sessions |
| `adhoc-opening.test.js` | opening details, the section, session changes, rebuild |
| `adhoc-session-status.test.js` | status block, statistics, failures, privacy |
| `adhoc-integration.test.js` | all generated parts of one report through build, updates and rebuild |
| `adhoc-multiday-sessions.test.js` | sessions of several days: the model, the cut-off, attendance per day, rendering, companies, opening, status |
| `adhoc-legacy-attendance.test.js` | a report that already has an attendance list is never changed |
| `adhoc-attendance-format.test.js` | the look of the generated attendee tables |
| `template003-menu.test.js` | the menu tree and the list of new functions |

Every test uses synthetic data. The mutation checks of each stage were repeated on the
integrated code.

Not provable locally, and therefore the subject of §9: reading a file in the browser, the
display values of the real TDoc list, heading / bold / link formatting in a real document, and
the size of the alert and the dialogs.

---

## 8. Not in this release

No AI summary. No statistics written into the report. No attendance from other sources than a
Teams export in English. No change to main-meeting reports, the build of a report without
sessions, the collector, the trigger, the release tooling or the CENTRAL add-on.

**Sessions of several days** (added after the smoke test, before the tag). A stored session of one day is unchanged: it has no `endDate`, its attendance is one record under the session id, and its texts are what they were. A session with an `endDate` has one cut-off on its last day, one set of opening details, and up to one attendance record per day; the Attendance section then says which day each block is. The days of a session cannot be changed so that a day with imported attendance falls outside them: such a save is refused and changes nothing. A manually made closing subsection of a report (such as a dated sub-heading under the closing item) is not generated by this feature and is never changed, renamed or removed by it.

Known and accepted: the submenu needs a reopen; Report Status Summary of a session report
needs the TDoc list and is slower than before; a sessions value with very many non-Latin labels
could exceed one property and would then be refused visibly.

**A report that already has an attendance list** (found in the smoke test). A report can
contain an attendance block that this feature did not write — made by hand or by another tool:
a statistics table and a Name / Company / Email table under the closing section. No code of
this repository ever generated such a block, so nothing proves where one came from, and it is
**never changed or removed**. An import writes the generated Attendance section next to it and
its result says that the other table exists and was not changed; the user removes it by hand
once. The generated section reads back the Company cells of its own tables only. A table of
the same shape elsewhere is read once, at the first import, so that companies entered before
are taken over; after that it is not read, so it cannot override what is typed into the
generated table.

---

## 9. Release

- `Code.js` **2.18.0**; proposed template release **T-2026.10.5**.
- No tag `template-release/T-2026.10.5` exists. It is created on the release commit only after
  the smoke test below has passed. T-2026.10.0 to .4 are not moved.
- The feature enters master as one squashed commit; the feature branch is not pushed.
- Reports already created or migrated stay on T-2026.10.4 until they are deliberately updated.

### Smoke test before the tag

On a **disposable copy** of an ad-hoc report, with a candidate bundle built from the release
candidate (the seven files; its `Release.js` names the candidate commit). Never on a live
report, the master template, CENTRAL or Legacy.

| # | Action | Expected |
|---|---|---|
| A | Reopen; *Configure Sessions…*: enter the sessions, save, open again | Submenu with six items; the sessions are shown as saved |
| B | *Edit Opening Details…*: a Chair for one session, save | "Session administration" heading directly above "Registration of Documents"; session lines bold |
| C | *Import Teams Attendance…* with a real export: first for a session whose days do not include its date, then for its own | Preview refuses the wrong date; then shows the Summary of the file; after Import an "Attendance" section at the end of the closing section, with a `mailto:` link; no telephone number, no Participant ID |
| D | *Assign TDoc Sessions…* | Opens with the real TDoc list; no uploaded TDoc is "cannot be read"; the **Uploaded** column shows each upload time in the time zone of the report; three TDocs checked against the Portal's upload time plus the UTC offset |
| E | *Add the Session column* in that dialog | Five cells in every row, same number of rows, no TDoc twice |
| F | *Report Status Summary*; *Post-meeting Statistics…* | Both readable; the Teams values equal those of the Attendance section |
| G | *Update Report Now* | Completes; both generated sections unchanged; the column still correct |
| H | *Build Report from Scratch…* (the copy only) | One Session administration section with the Chair; no "<Chair> opens the session …" line; Session column; one Attendance section with the attendees |
| I | Look through the document | Nothing lost or moved; each generated part once; nothing of the export that is not meant to be shown |

**Pass** only if all of A–I hold and no dialog or execution fails. **Stop** on any lost or
moved content, an upload time that is unreadable or off by the offset, a discarded Teams field
that appears, a generated part that appears twice, or an exception that could also affect
existing reports. After a stop nothing is tagged or deployed.

---

## 10. Closeout

### 10.1 Smoke test: passed

2026-10-05, on the disposable copy of an ad-hoc report, with the candidate built from commit
`3e97021` (the seven files, read back from the script project and equal to the candidate).
Findings of earlier candidates were fixed before this one and are described in §8; each fix
was installed and tested again. What was found in the visual reviews after it was changed in `Code.js` only (§10.2, §10.3),
installed and looked at again; the production code of the release is that of commit `33ded25`.

| Step | Result |
|---|---|
| Configure Sessions | **Pass.** One session of one day with start and end, and one session of three days with a start on its first day and a cut-off on its last. The range was stored, shown again on reopening, and rendered as a range in the Session administration section. |
| Assign TDoc Sessions | **Pass.** The real TDoc list loaded. The Uploaded column showed the upload times in the time zone of the report, and the automatic session of every TDoc could be understood from it, including TDocs uploaded shortly before and after the first session. No TDoc was assigned by hand; no current TDoc was without a session. |
| Session column | **Pass.** The existing four-column registration table was upgraded by the explicit action. Each TDoc showed its session; a reserved row of the report that is not in the current list showed "–". The table therefore has one row more than the dialog counts TDocs. |
| Attendance refresh | **Pass.** The attendance imported earlier was kept; the Attendance section appeared once; statistics and session line correct; table header bold and shaded, attendee rows not bold, compact, column widths as designed; Company values kept. An older attendance list that had been removed by hand did not come back. |
| Attendance of a session of several days | **Pass (dialog only).** The import dialog named the session with its range and explained that one report is imported per day and each is kept as its own day. No attendance was imported for that session: its days had not taken place. An import for a day of such a session is covered by tests only. |
| Post-meeting Statistics | **Pass.** Sessions and overall range; opening details and attendance each for one of two sessions; every TDoc with a session, all by upload time, none by hand, none without; Session column present; per session its dates, times, attendance state and number of TDocs. |
| Update Report Now (step G) | **Pass.** Ran without an error and returned to the report. The sessions, the opening details, the attendance and the Session column were as before. |
| Build Report from Scratch | **Pass.** The rebuild completed. No "<Chair> opens the session …" line; Session administration once, with the opening details; the range of the second session kept; registration table with the Session column; Attendance once, formatted correctly, with the Company values. A dated subsection that had been made by hand under the closing item was gone, as everything hand-made in a rebuilt section is; the removed older attendance list did not return. |

Steps B and C of §9 (opening details, an import with a real export) were done with the earlier
candidates of this test; what they produced is what the rows above found kept.

### 10.2 Found in the visual review, fixed after the test

With the Session column the registration table was hard to read: the Title column was far
too narrow, the titles wrapped heavily and the table became long. No width was ever set on
that table, so Google Docs gave each column the same share of the page — a quarter with the
four columns of every release, a fifth with the Session column.

A registration table **with the Session column** now gets explicit widths, as shares of the
page width of the report (468 pt unless configured otherwise). The first allocation (15 / 40 /
20 / 12 / 13 %) was installed on the disposable copy and looked at there: the widths were
right, but a label such as "AHG Call 1" still wrapped in the Session column. With the codes
of §10.3 that column needs little room. A second allocation (15 / 45 / 20 / 12 / 8 %) was looked
at as well: codes, Session column and titles were right, but with 70 pt a TDoc number broke
before its last digit. The TDoc column therefore gets 19 %, and the allocation is now:

| TDoc | Title | Source | Agenda Item | Session |
|---|---|---|---|---|
| 19 % (89 pt) | 42 % (197 pt) | 19 % (89 pt) | 12 % (56 pt) | 8 % (37 pt) |

They are set by Build Report from Scratch, when the column is added to an existing table, by
every update and when assignments are saved (`applyRegistrationTableWidths_()`); only a cell
whose width differs is written, so repeating it changes nothing, and a table built before
this fix gets its widths at its next update. A problem with the widths is logged and never
fails a build or an update. The four-column table of a report without sessions is not
touched, nor are the tables of the TDocs or the attendee tables.

**Verification: passed in the real document.** The candidate built from commit `33ded25` was
installed on the disposable copy (one normal `clasp push`, no `--force`; the seven files read
back and equal to the candidate) and *Update Report Now* was run. TDoc numbers stay on one
line; Title has room; Source and Agenda Item are usable; the codes in the Session column are
clearly readable; the wrapping of the header texts of the two narrow columns is accepted. The
allocation 19 / 42 / 19 / 12 / 8 % is final: no further change of this table is asked for.
Automated: `adhoc-registration-table-widths.test.js`, with mutation checks.

### 10.3 Second visual review: session codes and the Online information block

**Session codes.** See decision 13. `adhocSessionCode_()` / `adhocSessionCodes_()` derive the
codes; `adhocTdocSessionCellText_()` writes them into the Session column; the Session
administration section puts the code before each session ("A01: AHG Call 1, September 22, 2026,
15:00–18:00"). A table that still shows labels gets the codes at its next update. A code is a
position, not an identity: a session added before the others takes `A01` and the others move
on — in the table and in the Session administration section together, at the next update.

**Where the opening information of a main-meeting report comes from.** Nothing of the "Online
information" of an established report is generated by `Code.js`:

| Item | Provenance in an established (main-meeting) report | In an ad-hoc report with sessions now |
|---|---|---|
| Meeting name, Start Date, End Date | Text of the report template (section X.1), copied by the build and filled in by hand | Generated: the configured meeting name; the first day of the first session and the last day of the last |
| Portal meeting URL, TDoc List URL | The same | Generated from the meeting id that the configured TDoc list URL names; left out otherwise |
| Excel Docs URL | The same | The configured TDoc list URL |
| Docs Folder | The same | The configured docs folder |
| Drafts Folder | The same | Only when a drafts folder is configured for the report |
| Link to the shared minutes | The same | Not written: no such value exists for a report |
| Scribes | The same; typed by hand, stored nowhere | Not written: see below |

The opening of an ad-hoc report has been generated, not copied from the template, since before
this feature, and consisted of the one line "<Chair> opens the session on … at <start> CEST.".
That path never wrote meeting information; the session feature removed that one line (in favour
of the Session administration section) and nothing else. The information was therefore never
there for this report family; it is added now, for reports with sessions.

**Scribes.** No code produces a scribes list and nothing stores one: in an established report
it is template text edited by hand. For an ad-hoc report with sessions, the people who took the
minutes are recorded per session ("Minute taker(s)") in the Session administration section. A
second, report-level list would have no source and would duplicate that, so none is generated,
and a session's minute taker is not turned into a "scribe". A report-level scribes field can be
added later if it is wanted as its own piece of data.

**Verification:** automated tests (`adhoc-session-codes-opening-info.test.js` and the updated
tests of the feature, with mutation checks), and the visual check of 2026-10-05 on the
disposable copy with the candidate built from `4fa9214`: passed, apart from the TDoc column
width (§10.2).

### 10.4 Found, not part of this release

After the rebuild the closing section reads "… was closed on August 27, 2026 at <end> CEST."
(two sentences). This text is not generated: Build Report from Scratch copies the closing
section word for word from the report template document, replacing only the section number and
the mailing list (`copySectionContentWithReplacement_()`). That code is the same, byte for
byte, in this candidate, on master and in T-2026.10.4, and no commit of this feature touches
it. It is a date left in the template, so every report built from it shows it — before this
release as after it. **Not a regression and not a release blocker; no code is changed for it.**

Follow-up, outside this release: in the template document, replace the date in those two
sentences by a placeholder such as `<date>`. `<end>` stays: it is the placeholder the minute
taker fills in, like `<start>` in the opening.

### 10.5 Open items

**G2 is complete.** Nothing blocks the release. The date in the closing text of the report
template (§10.4) remains a matter of that document's content, separate from this release. The decisions listed as open in the design (§22 there) are
settled by §1 and §6 of this document. The production code has no TODO or FIXME of this
feature. What is known and accepted is in §8.

### 10.6 Still to be done, each on its own explicit authorization

| | Step | Changes |
|---|---|---|
| A | Squash the feature onto master as one commit (the release commit); the feature branch is kept and never pushed | the local repository |
| B | Tag `template-release/T-2026.10.5` on that commit; build the seven-file bundle from it with the release tool | the local repository, `dist/` |
| C | Push master and the tag (no force) | the public repository |
| D | Deploy the bundle to the master template's script project; read it back | the master template: reports created from then on |
| E | Update an existing report — only if that is wanted for a particular report | that report |

No step includes the next. Reports that exist — the live reports among them — keep the
release they have through A to D. What was tagged, deployed and is live is recorded here
afterwards.
