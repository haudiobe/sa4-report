# Ad-hoc sessions and Teams attendance: design

2026-10-02. Branch `feature/adhoc-sessions-attendance`, based on master `48edcc1`
(`Code.js` 2.17.4, template release T-2026.10.4). Design and discovery only: no runtime, test,
tooling or release file was changed, nothing was pushed or deployed, and no Google document or
Apps Script project was opened.

Line numbers refer to `Code.js` and `template/ReportCreator.js` at `48edcc1`.

**Scope.** Everything here applies to ad-hoc reports (`MEETING_TYPE = adhoc`) that have
sessions configured. A main-meeting report, and an ad-hoc report without session
configuration, behave exactly as they do today.

**What this design rests on, and what it does not.**

| Input | State |
|---|---|
| The code at `48edcc1` | Read. The 79 test files pass on the branch base. |
| A real Teams attendance export of one call | Read locally, read-only, for its **structure** only. Every figure, time and example in this document is synthetic (the values of the test fixture) or generic; nothing of the file is recorded here. The file contains personal data and must never be committed. |
| The manually produced 6G ad-hoc report | **Not opened** (no Google document was touched). Its presentation is taken from the task description. Points that depend on its exact form are listed in §22. |
| A real ad-hoc TDoc-list workbook with upload timestamps | **Not available in the repository.** The test fixtures carry a column subset without them. See §6.1; this must be checked before stage B. |

---

## 1. Existing architecture findings

### 1.1 Runtimes and where a feature can live

`Code.js` is the single runtime source for three deployments: the CENTRAL add-on, Legacy bound
copies, and the SA4 Report Template with the reports created from it. Template-specific
behaviour is switched at call time by `templateRuntimeRelease_()` (709). The template bundle is
exactly seven files (`tools/template-release.js` 86–93); the adoption mode refuses a bundle
that does not list exactly those seven. A new runtime file would therefore break adoption of
the existing releases. **The feature adds no file**: logic goes into `Code.js`, menu wiring
into `template/ReportCreator.js`.

A report is pinned to the release it was created or migrated with. New code reaches an
existing report only through a deliberate push to its Script ID.

### 1.2 Meeting type and configuration

- `MEETING_TYPE` is normalized by `normalizeMeetingType_()` (2843): blank means `main`, an
  unknown value throws.
- `getMeetingContext_()` (3041) returns `meeting.type`, `meeting.name`, the resolved sources
  (`mailingList`, `revisionsUrl`, …) and the agenda selector (`all` for ad-hoc). It has no
  notion of sessions, dates or times.
- `getReportConfig_()` (2616) carries `MEETING_DATE` (2666), a display string such as
  `September 22, 2026`, derived from the Portal start date only
  (`computeMeetingDateFromStartDate_()`, 12986). The Portal's end date and start time are shown
  in the Configure Meeting preview (`startDateRaw` / `endDateRaw`, 13103) but are not stored.
- The Portal describes an ad-hoc series as one meeting with one date range (fixture: meeting
  86172, 1 October to 5 November). Individual calls are not in any source the code reads.
  **Sessions can only be user-configured.**

### 1.3 State

- All report state is flat string properties behind `getReportStateStore_()` (2274): Document
  Properties by default, a document-namespaced Script Properties copy for a document adopted by
  the CENTRAL scheduler.
- Limits stated in the code (2166–2174): 9 KB per value, 500 KB per store.
- `ADDON003_ADOPTION_FIXED_KEYS_` (2489) lists what a background run needs;
  `CONFIG_DIALOG_MANAGED_KEYS_` (8071) lists what Configure Meeting writes and mirrors.
- `persistConfigurationSettings_()` (8107) is the one writer for meeting configuration. It is
  also used by the report creator's bootstrap, whose payload is whitelisted
  (`TEMPLATE_BOOTSTRAP_CONFIG_KEYS_`, ReportCreator 55), string-only and limited to 4,000
  characters in a Drive file description.
- `clearAllCaches()` (8885) deletes only `DISCUSS_`, `REVIS_` and `A1_EMPTY_CACHE_` keys.

### 1.4 Building the report

`buildSkeletonWithTdocTables()` (9766) reads the Document Reallocations table, parses the
agenda and downloads the TDoc list **before** it clears the body (9849), then writes the whole
report. For an ad-hoc meeting:

- The administrative block is placed by `getAdministrativeAgendaAnchors_()` (10641) from the
  real agenda: opening sub-section, Registration of Documents, Document Reallocations,
  Documents, IPR.
- The opening is one generated line (9984–9986):
  `<Chair> opens the session on <MEETING_DATE> at <start> CEST.` Chair and start time are
  literal placeholders by design (SA4-PROD-007A). `CEST` is a literal.
- The registration table is `createSummaryTable_()` (4190): four columns, TDoc, Title, Source,
  Agenda Item.
- The closing item copies the `X.Z` section of the shared template document (10111–10116).
  There is no attendance or statistics content.
- Only the Document Reallocations table survives a rebuild: it is read before the clear and
  written back at the end (10212–10227). Everything else typed into the document is lost, and
  the build confirmation says so (ReportCreator 923).

`runFullReportBuildCore_()` (9656) runs the phases skeleton, e-mail, revisions, abstracts and
formatting without any UI call after the document was changed (2.17.3).

### 1.5 Updating the report

`continuousUpdateCore_()` (915) is shared by the trigger, the CENTRAL scheduler and Update
Report Now (ReportCreator 959). It inserts tables for new TDocs, updates statuses, appends rows
to the registration table **only when a TDoc was added** (998), re-arranges revisions, then
runs the collector.

`updateRegisteredDocumentsTable_()` (2020) finds the registration table by an exact test: the
first row has **exactly four cells** reading TDoc, Title, Source, Agenda Item. If no such table
is found it returns silently. It appends rows and never rewrites one.

### 1.6 TDocs, revisions, dispositions

- `downloadAndGroupTdocs_()` (10250) converts the workbook to a Sheet and looks up ten columns
  by header name. It does not read `Uploaded` or `Reservation date`, but each TDoc object keeps
  its whole `row`, so further columns are available once their index is looked up.
- A **formal revision is a new TDoc number.** The list's `Revised to` column drives
  `getRevisedTo_()` (5540), `orderTdocsByRevision_()` (5555), the `REVISION_MAP` property (5600)
  and the "Revised to …" text in the Disposition row. A revision has its own row in the list,
  its own table in the report and its own row in the registration table.
- A **draft file** in the revisions folder (for example `<TDoc>_<company>.docx`) is a different
  thing. `updateRevisions_()` (5123) lists them in the Revisions row of the TDoc they belong to
  and keeps them in `REVIS_<tdoc>`, as file name and link. No timestamp is recorded.
- Minutes and Disposition are free text typed by the rapporteur. Status comes from the list's
  `TDoc Status`; `applyTdocStatusUpdate_()` (1961) overwrites a status in the document only
  while it is reserved or available, or when the new one is a revision.
- There is no record of when, or in which call, a TDoc was considered.

### 1.7 E-mail discussions

`checkRSSFeed_()` (4665) keeps one `DISCUSS_<tdoc>` property per TDoc: an object keyed by
message, each with author and date.

### 1.8 Menus and formatting

- Template reports get the "SA4 Report" menu from `buildTemplateReportMenu_()` (ReportCreator
  572). `onOpen()` is a simple trigger; reading the document's own properties there is already
  done (`addTemplateReportMenuHead_()`, 615).
- The CENTRAL / Legacy menu in `onOpen()` (581) is unchanged since before the template runtime,
  and `tests/template003-menu.test.js` and `tests/no-duplicates.test.js` pin both menus.
- `removeRowHeightAndSpacing()` (7089) formats **every** table: spacing, font, and the first
  row shaded as a header. `isConfigTable_()` (7270) treats any table whose first row reads
  Key / Value as configuration. Tables are identified by their header cells throughout
  (`isTDocTable_`, `isReallocationTable_`, `isDeadlineExtensionTable_`).

### 1.9 Existing building blocks to reuse

| Need | Existing code |
|---|---|
| Quote-aware row splitting | `splitCsvRows_()` (12489), comma only |
| Size limits for client-supplied text | `ADHOC_AGENDA_CSV_MAX_BYTES_` / `_MAX_ROWS_` (12447) |
| List address from a list name | `deriveEmailExportRecipientFromMailingList_()` (13872) |
| Portal meeting link | `portalMeetingUrl` form (12213) |
| Document link | `buildEmailExportDocumentUrl_()` (14135) |
| Read before clear, restore after | Document Reallocations in the build (9793, 10212) |
| Dialog with public RPC wrappers | `resolveMeetingForConfigDialog()` (13360) |
| Pure core, injected services, thin live wrapper | ReportCreator `*With_(deps)` |

---

## 2. Product requirements

1. An ad-hoc report can describe 1 to about 4 sessions (calls) of one series.
2. Sessions are metadata. The agenda remains the structure of the report; there is one
   registration table.
3. With sessions configured, the registration table has a compact Session column.
4. A Teams attendance export can be imported for one session. The report then shows, per
   session, a statistics block and an attendee table (Name, Company, Email).
5. The opening section carries more of what the 6G report contains by hand today.
6. A compact summary covers the whole series.
7. Nothing changes for main-meeting reports, for ad-hoc reports without sessions, or for any
   existing document merely because its code is updated.
8. The Teams participant ID (UPN) never reaches the report or the stored state.
9. Names are not rewritten beyond removing Teams' own markers. Companies are not invented.

---

## 3. Proposed session data model

A session is a small record. All times are **wall-clock times in the report's time zone**
(`Session.getScriptTimeZone()`; `Europe/Berlin` in the manifest) and are stored as text, never
as instants. Nothing is converted between zones.

| Field | Type | Meaning |
|---|---|---|
| `id` | `s1`, `s2`, … | Stable. Issued from a counter that is never reused. Not derived from the date, so a date can be corrected without losing the attendance import or the overrides attached to the session. |
| `label` | short text | What the Session column and headings show. Default: day and month of the date (`22 Sep`). Editable. Unique within the report. |
| `date` | `YYYY-MM-DD` | Required. The first calendar day of the session. |
| `endDate` | `YYYY-MM-DD` or absent | The last calendar day, for a session of several days (§3.1). Absent for a session of one day. |
| `start` | `HH:mm` or blank | Planned start, on the first day. |
| `end` | `HH:mm` or blank | Planned end, on the last day: the cut-off for the TDocs of the session (§6). |

Not stored in the session record:

- **Order.** Always derived from date, then start. A stored sequence number could contradict the
  dates.
- **Actual start, end and statistics.** They belong to the attendance import of that session
  (§4) and exist only once a file was imported.
- **Kind (call, offline, …).** No behaviour depends on it. The label can say "23 Sep offline".

Validation, applied before anything is saved:

- at least one session; at most 12;
- valid date; valid times; `start` before `end` when both are given;
- labels unique and at most 16 characters;
- the cut-off times of §6 are strictly increasing. Two sessions on one date therefore both need
  an end time for the earlier one.

### 3.1 Sessions of several days

*Added before the first release, after the live smoke test: the first real series has a
session of three days.*

**A session is one logical meeting, not a calendar day.** It has one id, one label, one set of
opening details and one place for its TDocs, and it may last several days: "AHG Call 1" on
22 September, "AHG Session 2" from 26 to 28 October.

- `date` is the first day; `endDate`, when present, the last. `endDate` is never before
  `date`. A session whose last day is its first day has no `endDate`: it is stored exactly as a
  session was before `endDate` existed, so nothing stored needs migrating and the schema
  version stays 1.
- `start` is on the first day and `end` on the last. On one day the end must be later than the
  start; over several days it need not be.
- The days of two sessions must not overlap. A session may begin on the last day of the one
  before only when the planned end of the first and the planned start of the second say so.
  Order is by first day, then planned start.
- **Display.** One helper writes a range of days: `September 22, 2026`,
  `October 26–28, 2026`, `October 30 – November 1, 2026`,
  `December 31, 2026 – January 2, 2027`. The default label is `22 Sep`, `26–28 Oct`,
  `30 Oct – 1 Nov`. Where the times of a session of several days are shown, they are said to
  be on the first and on the last day, never written as one day's `09:00–17:00`.
- **TDoc assignment is per logical session** (§6): the cut-off is the planned end on the
  **last** day, or the end of that day. Every TDoc uploaded after the previous session's
  cut-off and up to it belongs to the session, whichever of its days it was uploaded on.
- **Teams attendance has one record per day** (§4, §12.2). Teams exports one report per
  meeting, so a session of three days has up to three. A record is identified by the day of
  its own Teams start time; an export is accepted when that day is one of the session's days.
  Importing a day again replaces that day's record and no other. The Attendance section names
  the session with its days in every block, and says which day each block is. A record is
  stored under its own key: the session id for a session of one day (the only form there was
  before), `<session id>d<YYYYMMDD>` for a day of a session of several days. A record stored
  under the session id is read as the record of the day its own Teams start time says, so
  attendance imported before a session was given an `endDate` stays readable.
- **Attendance stays inside its session.** The days of a session cannot be changed in Configure
  Sessions so that a day with imported attendance falls outside them: the save is refused as a
  whole, names the session and the days, and changes nothing. No attendance is removed or
  moved and the session is not extended to fit. The user removes that attendance first, or
  keeps the day in the session. Extending a session, or shortening it to days that still
  contain all its attendance, is allowed.
- **Opening details** are one record per session, not per day.
- **Company corrections** apply to the attendee, not to a day: a Company typed for an attendee
  identified by e-mail (or by an exact name that stands for one person in its table) is that
  attendee's Company on every day and in every session. Two attendees of one name without
  e-mail never share one.

A single predicate decides whether any of this feature is active:

```
adhocSessionsEnabled_(context) :=
    getMeetingContext_(context).meeting.type === 'adhoc'
    AND the ADHOC_SESSIONS property holds a valid configuration with >= 1 session
```

A main-meeting document that somehow holds the property is not affected: the meeting type is
checked first.

---

## 4. Persistence model

| Key | Content | Store | Background-relevant |
|---|---|---|---|
| `ADHOC_SESSIONS` | `{ "v": 1, "nextId": 4, "sessions": [ … ] }` | report state store | yes |
| `ADHOC_SESSION_OVERRIDES` | `{ "<TDoc>": { "reg": "s2", "also": ["s3"] } }` | report state store | yes |
| `ADHOC_ATTENDANCE_<id>` | import header: file name, SHA-256 of the text, import time, summary values, chunk count, excluded attendee keys, the "Meeting" line | Document Properties | no |
| `ADHOC_ATTENDANCE_<id>_<n>` | attendee records, chunked below 8 KB | Document Properties | no |
| `ADHOC_COMPANY_MAP` (chunked the same way if needed) | `{ "<attendee key>": "<company>" }`, manual corrections only | Document Properties | no |

- `ADHOC_SESSIONS` and `ADHOC_SESSION_OVERRIDES` are read by the update path, so they are
  added to `ADDON003_ADOPTION_FIXED_KEYS_` and mirrored on save like the Configure Meeting keys.
  Both are small: four sessions are about 0.5 KB; an override is about 40 bytes.
- Attendance is interactive-only, like `PARSED_AGENDA` today. It is never copied to the CENTRAL
  store, which is shared by all reports.
- Size, as an order of magnitude: 70 records of about 130 bytes are about 9 KB, so two chunks per
  session; four such sessions are about 36 KB of the 500 KB.
- Stored per attendee: display name, company and its source, e-mail, role, first join, last
  leave, duration in seconds, the two Teams markers as flags, and how many raw records were
  merged. **The UPN is dropped while parsing and is never stored.**
- Sections 3 and 4 of the export are never parsed and never stored.
- `clearAllCaches()` does not touch any `ADHOC_` key.
- Writes are ordered so that an interrupted run is repeatable: chunks first, the header last. A
  header is the only thing that makes an import exist.

A Drive copy of a report starts with empty properties (TEMPLATE-001). The copy keeps the
rendered tables as plain content and has no sessions configured.

---

## 5. Session configuration UX

A separate dialog, **Configure Sessions…**, in the template report menu. Configure Meeting is
not changed, so the dialog a main-meeting report shows stays byte-identical.

- One row per session: Label, Date, Start, End, and a remove button. "Add session" appends a
  row.
- On first use the dialog proposes one session from `MEETING_DATE`, with blank times. Nothing
  else is proposed: the Portal knows the series, not the calls.
- A session that has an attendance import shows that, and removing it asks first. Removing a
  session removes its import and the overrides that name it.
- **Before saving, the dialog shows the effect:** how many TDocs change session, with the
  first ten listed (§6, case I). Save is a second click.
- Save writes the properties. It changes the document only if the report already has a
  Session column, and then only that column and the generated session blocks (§12).
- In a report that has no Session column yet, saving sessions changes nothing in the document.
  The column appears with the next Build Report from Scratch, or through the explicit item
  **Add Session Column to Registration Table…**.

New reports: the creator dialog and its bootstrap payload stay as they are. After Finish Report
Setup in an ad-hoc report, the summary names Configure Sessions… as the optional next step.
Carrying sessions in the bootstrap would need a new payload schema and space in a 4,000
character budget for no gain: the sessions are entered in the report in one dialog either way.

---

## 6. Session assignment algorithm

### 6.1 The timestamp

The only authoritative time is the TDoc list's **`Uploaded`** column. `docs/ADHOC_MEETING_ANALYSIS.md`
§3 records `Reservation date` and `Uploaded` in the header row of a real ad-hoc list. No code
reads them and no fixture contains them.

To be established on a real workbook before stage B is implemented:

1. Is `Uploaded` present in the workbook the report actually reads (the Portal document list
   for the meeting ID as well as the FTP list)?
2. What does `getValues()` return for it after the conversion to a Sheet: a Date, or text?
3. Which time zone do the values mean?

Rules that do not depend on the answers:

- Reserved, not yet uploaded: no timestamp, so **no session**. The reservation date is not
  used; reserving is not uploading.
- Column missing, or a value that cannot be read as a time: no automatic assignment for that
  TDoc. The cell stays blank and the status summary says why. Manual assignment still works.
- The time a TDoc was first seen by an update is **not** used. It depends on when the trigger
  happened to run.

### 6.2 The rule

Each session has a cut-off:

```
cutoff(session) = date + end        when an end time is configured
                = date + 23:59:59   otherwise
```

A TDoc with upload time `u` belongs to **the first session whose cut-off is not before `u`**.
If there is none, it has no session.

```
assignSession_(u, sessionsInOrder):
    if u is absent            -> none (reason: not uploaded)
    for s in sessionsInOrder:
        if u <= cutoff(s)     -> s
    -> none (reason: after the last session)
```

It is a pure function of the upload time, the session list and the overrides. It is recomputed
whenever it is needed and never stored.

This reads the product rule "uploaded before a session" as *before the session is over*. The
literal reading, before it *starts*, would give a document uploaded during a call to the next
session, which contradicts the requirement that a revision made in a session belongs to that
session. See decision 1 in §22.

The actual end from a Teams import is **not** used as a cut-off. Importing attendance must not
move documents between sessions as a side effect.

### 6.3 The cases

| | Case | Result |
|---|---|---|
| A | Uploaded before Session 1 | Session 1. |
| B | Uploaded exactly at a boundary | The cut-off is inclusive: `u = cutoff(s)` is session `s`. A time equal to a session's start lies inside that session's window. |
| C | Uploaded between two sessions | The later one. |
| D | Uploaded after the last cut-off | No session. The cell shows `–`. Not attributed to the last session; adding a later session picks it up. |
| E | Original in Session 1, revision in Session 2 | Two TDoc numbers, two rows. The original stays in Session 1 with "Revised to …"; the revision gets its own session from its own upload time. |
| F | Several revisions over several sessions | The same, per TDoc number. No row is duplicated. |
| G | Considered again later, no new revision | Nothing in the data says so. Recorded by hand as an additional session (§8); the cell reads `22 Sep, 24 Sep`. |
| H | Dates known, times not | The cut-off is the end of the session's day. Every upload on the day of a session belongs to it. |
| I | A session is added after documents exist | Assignments are recomputed. The dialog shows which TDocs move before the save. Overrides are kept. |
| J | The automatic result is wrong | Manual override, which always wins (§8). |

---

## 7. Revision and session semantics

- **Registration session.** One per TDoc number: the automatic result, or the override. This
  is what the Session column shows first.
- **A formal revision is its own TDoc** with its own registration session. That is the whole
  mechanism behind cases E and F; no revision-specific session logic is needed, and the
  existing revision ordering and "Revised to" text are untouched.
- **Additional sessions.** A TDoc can be marked by hand as also considered in further
  sessions. This is the only representation of "considered in a session", because the data has
  no other.
- **Draft files** in the revisions folder are not attributed to a session. Their upload time is
  not captured today, and they are not registered documents. Future work if wanted.
- **Disposition time** does not exist in the data and is not inferred.

TDoc detail tables are not changed. Session information appears in the registration table
only, so the discussion e-mail export, which renders TDoc tables, is unaffected.

---

## 8. Manual override model

`ADHOC_SESSION_OVERRIDES` maps a TDoc number to `reg` (registration session) and/or `also`
(additional sessions). An entry without `reg` keeps the automatic registration session and only
adds sessions.

- **Assign TDocs to Sessions…** lists the TDocs with their automatic session and lets the user
  set or clear an override and tick additional sessions.
- An override naming a session that no longer exists is dropped when that session is removed,
  and the dialog says so before the save.
- **Cells edited by hand.** Before the code rewrites a Session cell it compares the cell with
  what it last would have rendered. If the cell differs and consists of valid session labels,
  it becomes an override. If it differs and cannot be read, the cell is left alone and reported
  in the status summary. A hand edit is never silently overwritten.
- Overrides survive Build Report from Scratch: they are in properties, not in the document.

---

## 9. Teams CSV parser design

### 9.1 What the real export looks like

The structure of a real export. The examples are synthetic:

- UTF-16 little-endian with a byte-order mark (`FF FE`). CR LF line ends.
- **Tab-separated**, despite the `.csv` extension.
- Four sections, each introduced by a line of its own: `1. Summary`, `2. Participants`,
  `3. In-Meeting Activities`, `4. Meeting Engagement`. A blank line ends a section.
- Summary: six label/value rows. The meeting title is a **quoted field containing a line
  break**.
- Participants: a header row of 15 columns (seven identity and time columns, then eight
  `Engagement: …` columns) and one row per attendance record. Some rows have only 12 cells.
- Dates read `9/22/26, 2:48:05 PM`: United States order, two-digit year, 12-hour clock, **no
  time zone**.
- Durations read `3h 23m 20s`. Components that are zero are omitted (`1h 7s`, `43m 12s`,
  `2h 5m`).
- Roles: Organizer, Presenter, Attendee.

### 9.2 Reading the file

- The dialog reads the chosen file in the browser as bytes, detects the encoding from the
  byte-order mark (UTF-16 LE, UTF-16 BE, UTF-8; without a mark, UTF-16 LE if every second byte
  of the first kilobyte is zero, else UTF-8), decodes it and sends the **text** to the server.
- Limits before anything is parsed: 2 MB of text, 5,000 lines.
- All parsing is one pure server function, `parseTeamsAttendanceReport_(text)`, so there is one
  implementation and it runs under Node.
- Rows are split by a new delimiter-aware splitter with the quote rules of `splitCsvRows_()`.
  `splitCsvRows_()` itself is not changed. The delimiter is a tab if the Participants header
  line contains one, else a comma.

### 9.3 Sections

- A section starts at a row with a single cell matching `<digits>. <title>`.
- Required: `Summary` and `Participants`, by their English titles. If either is missing the
  import is refused with a message naming what was found.
- Sections 3 and 4 are skipped without being parsed.
- An export from a Teams client in another language has other titles and headers. It is
  refused with a clear message; nothing is guessed. Support needs a real file of that kind.

### 9.4 Summary

Label/value rows, matched by label: `Meeting title`, `Attended participants`, `Start time`,
`End time`, `Meeting duration`, `Average attendance time`.

- Title: white space, including the embedded line break, collapsed to single spaces.
- Each value is kept **as Teams states it**. Duration and average are parsed for display as
  `H:MM:SS` and never recomputed.
- A missing or unreadable Summary value is a blocking error: the statistics block cannot be
  produced without inventing it.

### 9.5 Participants

- Columns are found by header name: `Name` (required), `First Join`, `Last Leave`,
  `In-Meeting Duration`, `Email`, `Role`. `Participant ID (UPN)` is recognized only so that it
  is known to be discarded. `Engagement: …` columns are ignored.
- A row shorter than the header is padded with blanks. Extra cells are ignored.
- A row without a name is skipped and counted as a warning.
- A row whose join, leave and duration are all blank is valid; it is flagged (§10).
- A time or duration that is present but unreadable is a warning on that row; the row is kept.
- If the number of rows differs from `Attended participants`, that is a warning.

### 9.6 Dates and times

- Accepted forms: `M/D/YY, h:mm:ss AM|PM` as observed, and the same with a four-digit year or a
  24-hour clock. Anything else is refused.
- Month-first or day-first cannot be told from the text when both numbers are at most 12. The
  parser tries both and accepts the reading under which **the Summary start date equals the
  date of the session being imported**. If neither or both readings fit, the import is refused.
- The export has no time zone. Times are taken as wall-clock times of the report's zone, shown
  in the preview for confirmation, and stored as text. They are never shifted.

---

## 10. Attendance normalization and deduplication

### 10.1 What the export means

The relations between the quantities of an export, with the figures of the synthetic test
fixture (`tests/fixtures/teams-attendance-synthetic.js`):

| Quantity | Value | How it is obtained |
|---|---|---|
| Attendance records | 25 | Summary `Attended participants`; equals the number of Participants rows |
| Records with complete join and leave times | 21 | The others lack a join time, a leave time, or both |
| Attendees after normalization | 23 | Two pairs of records are one attendee each; two guests who share a display name stay separate |
| Average attendance, as Teams states it | 2:31:07 | Summary |
| Mean of the listed durations | 1:55:35 | Computed |
| Meeting duration, as Teams states it | 3:23:20 | Summary |
| End minus start | 3:23:21 | Computed |

Consequences:

- **An attendance record is not a person.** The number of records, of records with any
  attendance, and of distinct people all differ.
- **A person who leaves and rejoins is one Participants row.** Teams has already merged the
  intervals: in such a row the time between first join and last leave exceeds the duration.
  The single intervals are in section 3 and are out of scope.
- **Duplicates arise from a second identity**, not from rejoining: a guest who joins twice
  under the same name, or one person on two devices.
- **Teams' average cannot be reproduced** from the rows by any of the obvious divisors, and its
  duration differs from end minus start by one second. Both are shown as Teams states them, and
  the report does not present a computed figure under Teams' label.
- What the rows without times mean is not stated by the export. They are flagged, not
  interpreted.

### 10.2 Names

Applied in this order, and nothing else:

1. remove the delimiter-level quotes; trim; collapse runs of white space; Unicode NFC;
2. remove a trailing `(External)` or `(Unverified)`, repeatedly, case-insensitively, and record
   each as a flag.

Everything else is kept as the person wrote it: `Doe, Jane`, `[ExampleCorp] Alex Kim`,
`Alex Kim (ExampleCorp)`, `Alex Kim - ExampleCorp`, names in other scripts, and a second
spelling in brackets. Names with non-ASCII characters are common.

In the export studied, the `(Unverified)` records have no e-mail and the `(External)` records
have one.

One record's name is a telephone number (a dial-in). A name consisting only of digits, spaces
and `+` is flagged as a dial-in, and the number is not written into the report (decision 6).

### 10.3 Attendee key and merging

```
key = lower-case e-mail          when the record has a valid e-mail
    = case-folded normalized name otherwise
```

- Records with the same key are merged: earliest join, latest leave, durations added, and the
  number of merged records kept.
- A record with an e-mail and a record without one are never merged automatically, even with
  the same name. The preview lists such pairs as possible duplicates.
- The preview lets the user exclude single records. Exclusions are stored with the import, so
  importing the same file again gives the same table.

---

## 11. Company and e-mail policy

### 11.1 What Teams supplies

Teams has **no company field**. After removing the Teams markers, a display name can carry:

| Evidence in the display name | How often |
|---|---|
| `[Org] Name` prefix | rare |
| `Name (Org)` suffix | a few |
| `Name - Org` suffix | a few |
| `Name Org/Team` suffix | rare |
| Bracketed text that is a name in another script, not an organization | a few |
| Telephone number (a dial-in) | rare |
| Nothing | the large majority |

Most records have an e-mail. The domains are those of organizations, not of public mail
providers, and a good part of them are sub-domains.

### 11.2 Policy

1. **A manual correction** for that attendee, from `ADHOC_COMPANY_MAP`.
2. **The `[Org] Name` prefix**, the one form that is unambiguous. The name itself is left as
   written.
3. **Otherwise blank.**

Not done:

- **No company from the e-mail domain.** A domain is not a company name: sub-domains,
  subsidiaries, universities and consultants do not map to the name a delegate is registered
  under, and a wrong company in an official attendance list is worse than a blank one.
- **No company from a trailing bracket or dash.** The same shape holds a second spelling of the
  name in real exports.

So a first import fills very few Company cells automatically. The user completes them in the
table:

- Before an attendee table is rendered again (re-import, another session's import, a rebuild),
  the Company cells of the existing tables are read. A cell that is not empty and differs from
  the stored value becomes a manual correction for that attendee key.
- A correction applies to that person in every session of the report. A series has largely the
  same participants in each call, so the work is done once.

How Company was filled in the manual 6G report is not known here (decision 7). If it came from
the 3GPP meeting registration, that list would be the reliable source, and importing it would
be a later, separate feature.

### 11.3 E-mail

- A value is an e-mail if it is one address without white space, with one `@` and a dot in the
  domain. It is shown as written and linked with `mailto:`.
- No e-mail: the cell is empty.
- **The UPN is never a fallback.** It regularly differs from the e-mail, so it is a different
  identifier, not a second copy.
- Teams-internal links do not occur in the Summary or Participants sections and none is
  generated.

First join and last leave are stored and not shown.

---

## 12. Report rendering design

### 12.1 Registration table

With sessions enabled, the build creates the registration table with a fifth column:

```
TDoc | Title | Source | Agenda Item | Session
```

The cell shows the label of the registration session, then any additional sessions
(`22 Sep, 24 Sep`), `–` for "after the last session", and nothing for "not uploaded yet".

- Table detection accepts the four-column header (as today) or the five-column one. No existing
  document has a five-column table with this header, so nothing changes for one.
- Appending rows writes four or five cells according to the table that is found, **not**
  according to the configuration. A report whose sessions were removed keeps its column and
  keeps receiving rows, with an empty Session cell.
- With sessions enabled and a five-column table, each update also refreshes the Session cells
  whose value changed, for example when a reserved TDoc is uploaded. These are text-only
  changes and do not trigger the formatting pass.
- A report with sessions configured and a four-column table is left as it is until the user
  rebuilds or runs Add Session Column to Registration Table….
- The main-meeting 6G branch (9927) and every other caller pass no sessions and produce four
  columns.

### 12.2 Attendance

One generated container at the end of the closing agenda item's section: a heading
**Attendance**, and inside it one block per session that has an import, in session order.

```
Attendance
  <label> – <long date>
    Statistics
      Meeting: …
      Session date: September 22, 2026
      Start: 14:48:05
      End: 18:11:26
      Duration: 3:23:20
      Attendance records: 25
      Average attendance: 2:31:07
    Attendees
      Name | Company | Email
```

- The statistics are paragraphs. A two-column table would have its first data row shaded as a
  header by the formatting pass, and a Key / Value header would be read as a configuration
  table.
- `Meeting:` is the Teams title with white space collapsed, editable in the import preview and
  stored with the import. The value in the task's example is neither the raw Teams title nor
  `MEETING_NAME`, so it is not derived (decision 9).
- Start, End, Duration, Attendance records and Average attendance are Teams' Summary values.
- The attendee table has one row per attendee key, in the order of the export.
- **The container is replaced as a whole** every time it is rendered. It is found by its
  heading, and it ends at the next heading of the same or a higher level. Apart from Company
  cells (§11.2), text typed inside it is not kept, and the import preview says so.
- The closing prose above the container (closing discussion, thanks, closing time) is manual
  and is never touched.
- If no closing section is found, the container is appended at the end and the result says so.

### 12.3 Rebuild

Build Report from Scratch renders the Session column, the opening (§13), the Attendance
container and the series statistics (§14) from the stored state. Before the body is cleared,
the Company cells of the existing attendee tables are read (§11.2), as the Document
Reallocations are today.

---

## 13. Rich opening design

A: known today. B: derived reliably. C: configured once per report. D: configured per
session. E: manual prose.

| Field | Class | Source and rendering |
|---|---|---|
| Meeting name | A | `MEETING_NAME` |
| Session label and date | D | session record |
| Opening time | D | planned `start`; blank gives the placeholder `<start>` as today |
| Time-zone abbreviation | B | from the session date and the report's zone, replacing the literal `CEST` (wrong after the change to winter time, and series do cross it) |
| Chair | E | the placeholder `<Chair>` as today (decision 12) |
| Minute takers | E | a line `Minute takers: <names>` per session; always manual |
| Link to these minutes | B | the document's own URL |
| Portal meeting reference | B | `MEETING_ID` in the Portal meeting URL form the resolver already uses |
| Attendance registration | B + E | the same Portal link in a fixed sentence; the wording is to be confirmed |
| Reflector | B | the effective mailing list as its `list.etsi.org` address |
| Schedule | B | a table Session / Date / Time from the session records |
| Agenda approval | A + E | the agenda TDoc is named; the decision is manual |
| Registration approval | E | manual sentence above the registration table |

No configuration field is added for the opening. Everything in class B comes from values that
exist already or from the session records.

Layout under the existing "Opening of the session" sub-section:

1. a general block, once: minutes link, Portal reference with the attendance sentence,
   reflector, schedule table;
2. one paragraph per session:
   `<label> (<date>): <Chair> opens the session at <start> <zone>.` and
   `Minute takers: <names>`.

Rules:

- With sessions disabled the opening is the one line of today, with the literal `CEST`.
  Unchanged.
- The opening is written by Build Report from Scratch only. It is prose the rapporteur edits,
  so no update rewrites it.
- The schedule table is the exception: it is found by its header row and refreshed with the
  other session blocks.
- A session added later does not get an opening paragraph automatically. **Insert Opening
  Paragraph for Session…** adds one at the end of the opening sub-section on request.
- The actual start from Teams is not written into the opening after the fact.

---

## 14. Post-meeting statistics design

One generated block, **Series statistics**, at the top of the Attendance container: a table
with one row per session and a short list of totals. Produced by the build and by Update
Session Sections; never by a background update.

| Metric | Source | Reliable |
|---|---|---|
| Sessions | `ADHOC_SESSIONS` | yes |
| Sessions with imported attendance | `ADHOC_ATTENDANCE_<id>` headers | yes |
| Registered TDocs | TDoc tables in the report, identified as in `analyzeReportStatus()` (8598) | yes |
| TDocs per session | §6 and §8 | yes, where upload times exist; the rest is shown as "no session" |
| Status counts | the Status row of each TDoc table, grouped by its normalized value; every value that occurs is listed, none is predefined | yes; the values are the list's statuses plus whatever the rapporteur wrote |
| Revised documents | `REVISION_MAP` (document → revised to) | yes |
| TDocs with e-mail discussion; messages | `DISCUSS_<tdoc>` | yes, as collected: from the collection start date, within the feed's 2,000 newest items |
| Per session: start, end, duration, attendance records, average | attendance import, Teams' Summary | yes |
| Per session: distinct attendees | attendee keys (§10.3) | yes, under the stated key rule |
| Distinct attendees over the series | union of attendee keys | **optional**: exact for people with an e-mail, by name otherwise |
| Documents considered, in total or per session | none | **no**. Not derivable. Possible only from the manual additional sessions of §8, and then it counts what was marked, not what happened. Future. |
| Dispositions by kind | Disposition is free text | **no**. The status counts cover the decided outcomes. |
| Draft revisions per session | `REVIS_<tdoc>` has no time | **no**. Future. |

The existing Report Status Summary dialog is unchanged.

---

## 15. Menu and UI changes

Template report menu only. One new sub-menu after Document Reallocation:

```
🗓 Sessions and Attendance
    Configure Sessions…
    Assign TDocs to Sessions…
    Import Teams Attendance…
    Update Session Sections
    Sessions and Attendance Status
    ──
    Add Session Column to Registration Table…
    Insert Opening Paragraph for Session…
```

- The sub-menu is added only when the document's `MEETING_TYPE` property is `adhoc`, read in
  `onOpen()` like the bootstrap state today. A main-meeting report's menu is unchanged. A report
  configured as ad-hoc in the current sitting shows it after the document is opened again.
- Every item also checks the meeting type itself and refuses in a main-meeting report and in
  the master template.
- **Import Teams Attendance…**: choose the session, choose the file, Preview, Import. The
  preview shows the Summary values, the record, attendee and merge counts, records without
  times, possible duplicates, dial-ins, warnings, the editable Meeting line, and whether an
  earlier import for the session is replaced. Removing a session's import is in the same
  dialog.
- **Update Session Sections** re-renders the Session column, the schedule table, the Attendance
  container and the series statistics from the stored state.
- The CENTRAL / Legacy menu is not changed.

---

## 16. Backward compatibility

| Document | Behaviour |
|---|---|
| Main-meeting report | No property is read beyond `MEETING_TYPE`, which is read today. No menu item, column, table or text changes. The build, the update, the configuration dialog and the creator are byte-identical in effect. |
| Ad-hoc report without `ADHOC_SESSIONS` | As today. The predicate of §3 is false and every new code path is skipped. In the template runtime the new sub-menu is visible; its items change nothing until sessions are saved. |
| Migrated 6G and MBS reports | Pinned to T-2026.10.4. Nothing changes without a deliberate push. After one, they are in the row above. |
| New ad-hoc report | Sessions are offered after setup and are optional. |
| Existing report whose code is updated | No column, table or paragraph appears. The column needs an explicit action or a rebuild; attendance needs an import. |

Invariants that the tests of §20 pin:

1. With the predicate false, `buildSkeletonWithTdocTables()`, `continuousUpdateCore_()`,
   `createSummaryTable_()` and `updateRegisteredDocumentsTable_()` produce what they produce
   at `48edcc1`.
2. The update path adds a column to no table.
3. No background run writes attendance, statistics or opening content.

One hazard to state: a report with a Session column must not be updated by an **older**
runtime. Its registration-table test wants exactly four cells, would not find the table, and
would stop appending rows without an error. A report only gets the column from a runtime that
knows it, and reports are pinned, so this needs the CENTRAL menu to be used in a migrated
document, which `docs/EXISTING_REPORT_MIGRATION.md` §5.7 already rules out.

---

## 17. Migration behaviour

- No data migration. All keys are new; an absent key means the feature is off.
- `ADHOC_SESSIONS` carries a version. An unknown version is treated as "not configured" and
  reported in the status dialog, never guessed at.
- The 6G report already has attendance written by hand. An import adds the generated container;
  it does not find or replace the manual text. The user removes the manual version once.
  Nothing happens unless the user imports.
- A report updated to the new runtime needs no step unless sessions are wanted.
- Rollback to an older release: the properties are ignored by old code. If a Session column
  exists, remove the column by hand first (§16).

---

## 18. Error handling and idempotency

**Import**

- Everything is validated before anything is written: encoding, sections, Summary, header,
  date agreement with the session, limits. Preview and Import run the same pure function on the
  same text; Import is refused if the text's hash differs from the previewed one.
- A file whose start date is not the session's date is refused.
- The same file for the same session again: nothing is written; the container is rendered again
  only if it is missing.
- Another file for the same session: replaces the import, after the preview said so.
- The same file for another session: refused, naming the session that has it.
- The import and Update Session Sections take the document lock that the updates take. If an
  update is running, nothing is done and the dialog says so.
- No pop-up after the document was changed (2.17.3). The result goes back to the open dialog.
- Properties first, document second. If rendering fails the import exists, and Update Session
  Sections completes it.

**Sessions**

- An invalid configuration is refused in the dialog and nothing is saved.
- A stored configuration that cannot be read counts as "not configured"; the update goes on as
  for a report without sessions, and logs it once.
- Session logic never makes an update fail. An error in the Session column refresh is caught,
  logged and returned with the collector failures, so Update Report Now shows it and the
  trigger carries on.

**Rendering**

- Each generated element is found by its heading or header row and replaced, never appended a
  second time.
- A Session cell that was edited by hand and cannot be read is left alone (§8).

---

## 19. Exact code extension points

`Code.js`:

| Location | Change |
|---|---|
| New section after the reallocation helpers | Session model: read, validate, order, cut-off, `adhocSessionsEnabled_()`, `assignSession_()`, override handling, label rendering. Pure. |
| `ADDON003_ADOPTION_FIXED_KEYS_` (2489) | Add `ADHOC_SESSIONS`, `ADHOC_SESSION_OVERRIDES`. |
| `downloadAndGroupTdocs_()` (10279–10344) | Look up `Uploaded` and pass `uploadedCol` in each TDoc object. One header lookup and one field; nothing else. |
| `createSummaryTable_()` (4190) | Optional trailing argument with a session resolver; with it, a fifth header cell and a fifth cell per row. Callers that omit it are unchanged. |
| `buildSkeletonWithTdocTables()` — before the clear (9793) | Read the Company cells of existing attendee tables. |
| — `emitOpeningAdminBlock` (9975–10003) | With sessions enabled: the opening of §13, and the session resolver passed to `createSummaryTable_()`. The existing branch stays for the disabled case. |
| — closing item (10111–10116) and synthetic close (10201–10210) | After the template content: the Attendance container and the series statistics. |
| `updateRegisteredDocumentsTable_()` (2028–2033, 2053–2066) | Accept the five-column header; append a fifth cell when the table has one. |
| `continuousUpdateCore_()` after 1000 | With sessions enabled and a five-column table: refresh changed Session cells. Caught and reported, never thrown. |
| New section near `splitCsvRows_()` (12489) | `splitDelimitedRows_()`, `parseTeamsAttendanceReport_()`, the date and duration parsers, name normalization, key, merge, company resolution. Pure. |
| New section after the e-mail export | Attendance persistence (chunking), the three renderers, the statistics collection, and the public dialog and RPC functions. |

`template/ReportCreator.js`:

| Location | Change |
|---|---|
| `buildTemplateReportMenu_()` (572), after the Document Reallocation sub-menu (588–592) | The Sessions and Attendance sub-menu, when `MEETING_TYPE` is `adhoc`. |
| `finishReportSetup()` (893) | One sentence naming Configure Sessions… for an ad-hoc report. |
| `confirmTemplateBuildFromScratch_()` (923) | With sessions configured: say that sessions, overrides and imported attendance are kept. |

Not changed: `onOpen()`'s CENTRAL / Legacy menu, `configureMeetingSettings()`,
`persistConfigurationSettings_()`, `getMeetingContext_()`, `getReportConfig_()`, the creator
and its bootstrap payload, `splitCsvRows_()`, the revision ordering, the collector, the
discussion e-mail export, `tools/template-release.js`, the manifest and the bundle's file list.

---

## 20. Proposed automated tests

Plain Node scripts in the existing style, loading `Code.js` through `tests/helpers/load-code.js`.

**Fixtures.** A **synthetic** Teams export in UTF-16 LE with a byte-order mark, written for the
tests with invented names. It reproduces every structural feature of the real file: tabs, the
quoted two-line title, the four sections, rows of 12 and 15 cells, rows without times, both
markers, a bracket prefix, a bracketed second spelling in another script, a duplicated guest
name, a dial-in, and a UPN that differs from the e-mail. Its Summary and row values are chosen
so that the documented relations of §10.1 hold. The real export is never added to the
repository. Further variants are derived in the tests: UTF-8, comma-separated, day-first dates,
a missing section, a translated section title.

| File | Covers |
|---|---|
| `adhoc-sessions-model.test.js` | Validation; stable ids; derived order; cut-offs; the predicate for main, ad-hoc without sessions, a damaged property, an unknown version. |
| `adhoc-session-assignment.test.js` | Cases A–J of §6.3 by name; absent and unreadable upload times; overrides; additional sessions; dropping overrides of a removed session. |
| `adhoc-session-column.test.js` | Five-column build with sessions; four-column build without, compared with today's output; update appends four or five cells according to the table; refresh of changed cells; a hand-edited cell is adopted or left alone; no column is added by an update. |
| `adhoc-sessions-compat.test.js` | The invariants of §16 for a main meeting and for an ad-hoc meeting without sessions, against output captured at `48edcc1`. |
| `teams-attendance-parser.test.js` | Decoding; delimiter detection; sections; Summary; padding of short rows; date forms and the month/day decision; durations; refusal of other languages; limits; the UPN is absent from the result. |
| `teams-attendance-normalize.test.js` | Marker removal and nothing else; NFC; key; merging; no merge across e-mail and no e-mail; dial-in; the four quantities of §10.1 kept apart. |
| `teams-attendance-company.test.js` | The hierarchy; bracket prefix only; never from a domain; never from a trailing bracket; corrections read back from tables and applied across sessions. |
| `teams-attendance-import.test.js` | Preview writes nothing; hash check; same file twice; another file; same file for another session; date mismatch; chunking round trip; header written last. |
| `adhoc-attendance-rendering.test.js` | Container found and replaced, not duplicated; statistics lines; `mailto:` links only for valid addresses; empty e-mail cell; neither UPN nor join and leave times in the output; rebuild keeps sessions, overrides, attendance and corrections. |
| `adhoc-opening.test.js` | Disabled: today's line exactly. Enabled: general block, one paragraph per session, placeholders, zone abbreviation in summer and winter. |
| `adhoc-series-statistics.test.js` | Each metric of §14 from its source; unreliable metrics are absent. |
| existing `template003-menu.test.js`, `no-duplicates.test.js` | Extended: the sub-menu appears for an ad-hoc report only, every target exists, the CENTRAL / Legacy menu is unchanged. |
| existing `addon003-adoption.test.js`, `addon005b-config-precedence-e2e.test.js` | Extended: the two new keys are adopted; attendance keys are not. |

`tests/helpers/fake-document.js` needs links per cell and heading levels for the rendering
tests. That is an addition to a test helper, not to the runtime.

What tests cannot show, and a rehearsal on a Drive copy must: the file input in an Apps Script
dialog and the transfer of about 20 KB of text; what `getValues()` returns for `Uploaded`; the
look of the rendered blocks.

---

## 21. Suggested staged implementation plan

Each stage is one reviewable change with its tests, leaves all existing tests passing, and is
inert for every existing document. The letters follow the task; the **recommended order is
A, C, D, B, E, F, G**: attendance is the larger saving of manual work, and stage B depends on a
check that has not been made.

| Stage | Content | Visible afterwards |
|---|---|---|
| **A** | Session model, validation, persistence, adoption keys, Configure Sessions…, the status dialog, the conditional sub-menu. | Sessions can be configured. The document does not change. |
| **C** | Splitter, parser, normalization, merging, company resolution, the synthetic fixture. Pure functions only. | Nothing. |
| **D** | Import dialog, preview, persistence, the Attendance container, corrections read back, rebuild preservation, Update Session Sections. | Attendance per session. |
| **B0** | Check `Uploaded` on a real workbook (§6.1). No code. | — |
| **B** | Upload time in the TDoc objects, assignment, the Session column in build and update, the explicit add-column item, the override dialog. | Session column. |
| **E** | Opening of §13, zone abbreviation, the schedule table, Insert Opening Paragraph…. | Richer opening in new builds. |
| **F** | Series statistics. | Summary block. |
| **G** | Build confirmation and setup texts, changelog (`Code.js` 2.18.0), user notes, a rehearsal on a Drive copy of an ad-hoc report, then a template release with a new release ID. Published tags are not touched. | Released. |

If B0 shows that the workbook has no usable upload time, stage B is reduced to manual
assignment and the automatic rule is dropped. Nothing else depends on it.

---

## 22. Open product decisions

Each has a recommendation, which the design above assumes.

At the release candidate none of these is open any more: what was decided and built is in
§1 and §6 of `ADHOC_SESSIONS_RELEASE_CANDIDATE.md`. The table is kept as the record of the
study.

| # | Decision | Recommendation |
|---|---|---|
| 1 | Does a document uploaded **during** a session belong to it or to the next? | To it: the window ends at the session's end (§6.2). The literal "before the session starts" contradicts the revision requirement. |
| 2 | Without a planned end, where does a session's window end? | At the end of its day. The alternative, the Teams end time, makes an attendance import move documents. |
| 3 | Documents uploaded after the last session | No session (`–`), not the last one. |
| 4 | What the Session column shows | The label, by default day and month (`22 Sep`). `S1`, `S2` would renumber when a session is inserted. |
| 5 | Are hand-edited Session cells honoured? | Yes, as overrides when readable (§8). The stricter alternative is dialog-only. |
| 6 | A dial-in whose name is a telephone number | Listed as "Dial-in participant"; the number is not written. This is the one place where a name is replaced. |
| 7 | **Company.** How was it filled in the manual 6G report? Is a mostly empty column after the first import acceptable? | Blank unless `[Org]` prefix or manual correction; corrections carry over between sessions. If the source was the 3GPP registration list, plan its import separately. |
| 8 | Records without join and leave times | Listed, flagged in the preview, excludable. Teams counts them as attended. |
| 9 | The `Meeting:` line | Teams title, editable in the preview. |
| 10 | Empty e-mail | Empty cell, not "Not provided": it would appear in a good part of the rows. |
| 11 | Show distinct attendees in the report next to Teams' "Attendance records"? | Per session in the series table only; the statistics block keeps the 6G wording. |
| 12 | A stored chair name | No. Keep the placeholder. One field per report would save typing a name once per session. |
| 13 | Wording of the attendance-registration and reflector sentences, and the exact form of the statistics block | To be taken from the 6G report, which was not opened for this study. |
| 14 | Teams exports in other languages | English only until a real file of another language is available. |
| 15 | Draft files and sessions | Not in this feature. |
| 16 | Where the Attendance container goes when a report has closing text per session | One container at the end of the closing section. Per-session placement would mix generated and manual text. |
