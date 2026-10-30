# TDoc list identity

**Status: CANDIDATE for T-2026.10.17 (`Code.js` 2.26.1), branch `fix/disposition-meeting-scope`. NOT RELEASED,
NOT DEPLOYED.** Section TDOC LIST IDENTITY of `Code.js`; tests `tests/tdoc-list-identity.test.js`.

A downloaded TDoc list is used only when it is shown to be the list of the configured meeting. Before, a
list was trusted because it could be downloaded.

## Why

6G, 2026-10-30. The report asks the Portal for
`https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178`.

| Read | Rows | TDocs | Agenda TDocs in it | Reservation dates |
|---|---|---|---|---|
| 06:09 UTC | 116 | S4aP260067 … 182 | S4aP260098 (22 Sept), S4aP260149 (26–28 Oct) | October 2026, as `05/10/2026 10:05:32` |
| later the same day | 36 | S4aP260008 … 043 | S4aP260020, S4aP260035, S4aP260038 (Feb–Mar calls) | Feb 19 – Mar 30 2026, as `2/19/2026 3:00:31 PM` |

The two have no TDoc in common. Both workbooks are the same Portal template (creator "3GPP Portal",
modified 2026-02-26, sheets `TDoc_List`, `CR_Packs_List`, `Parameters`); nothing in either names its
meeting, date or agenda. The update of the live report inserted the 36 TDocs of the earlier calls
(150 = 114 of this call + 36; the Disposition Summary's 16 / 93 / 41 against the smoke copy's
16 / 84 / 14 of the day before).

## The rules

On the **raw rows** of the list — before the agenda selector, the Document Reallocations and the
grouping, and before anything is changed (`assertTdocListIdentity_()`, called by
`downloadAndGroupTdocs_()` and `downloadAndProcessFromWeb()`):

1. **Agenda TDoc.** When the saved `AGENDA_TDOC` is a valid TDoc number, the list must have a row
   for it. Any status counts (reserved, not uploaded, withdrawn …): a row is a row. A Document
   Reallocation of it does not count. It is the meeting's own document, found in its own list
   (Discover Agenda / TDocs). With it nothing else is asked and no anchor is kept. 6G: S4aP260098.
2. **Anchor.** Without a valid agenda TDoc (MBS: agenda.csv; main meetings: an agenda document),
   `TDOC_LIST_ANCHOR` (report state, like `REVISION_MAP`) holds
   `{ tdoc, meetingId, tdocListUrl, acceptedAt }` of a list accepted before. A later list must have a
   row for `tdoc` (a Portal list keeps every TDoc it listed; a withdrawn one stays, as withdrawn). An
   anchor of another meeting ID or TDoc list URL — Meeting Setup changed them — is no anchor.
3. **No anchor yet.** It is established from the report: **every** TDoc table of the report must be
   in the list, and there must be one. An update writes only TDocs of its list, so a report that the
   list does not account for is either not this list's or not clean; it is refused either way, and
   nothing is guessed. A report already holding TDocs of another meeting gets no anchor from any list;
   it needs an agenda TDoc, a repair, or a rebuild. Build Report from Scratch, an explicit start on a
   meeting just set up, takes the list it is given and records the anchor (trust on first use, by the
   person building).

The anchor is written only by an update or a build, after the list passed. It is a TDoc in the list
and, unless this is a build, in the report: an uploaded one first (an uploaded document is not taken
back), the lowest number of those. It is kept until it stops applying (another meeting ID or URL).
Update Disposition Summary, Close Report and the session dialogs check the same way and store
nothing. `TDOC_LIST_ANCHOR` is not among the keys copied when a report is adopted for the add-on
background runs (`ADDON003_ADOPTION_FIXED_KEYS_`); there the report is the evidence again (rule 3).

## When a list is refused

The error starts "The TDoc list was not used: it could not be shown to be the list of this meeting
(…)" with the reason, and is logged as `TDoc list identity: REFUSED -- …`.

| Path | What happens |
|---|---|
| Update (manual and automatic) | `{ success: false, notUpdated: true }` before anything: no TDoc table or agenda heading, no status or upload metadata, no registration table or registered-documents summary, no Session column, no revision moved, no Disposition cell, no `REVISION_MAP`, no abstract request, no structure repair, no Disposition Summary sync |
| Build Report from Scratch | refused before the document is cleared |
| Update Disposition Summary | "Disposition Summary NOT updated", with the reason |
| Close Report | NOT READY TO CLOSE; nothing checked, no revision check, nothing written |
| Enable Session Column / Assign TDoc Sessions | nothing written; a saved assignment is kept, the Session column is not refreshed (the notice says so) |
| `downloadAndProcessFromWeb()` (legacy) | the error, the temporary file trashed, no table written |

There is no fallback to the TDoc tables of the document. The zero-overlap guard of the Disposition
Summary (`writeDispositionSummary_()`, `c48ba94`) stays as a second line.

## Tests

`tests/tdoc-list-identity.test.js`:

- **A**, the contamination case: report X already holds TDocs of meeting Y, and the Portal returns
  Y's list, which overlaps the report. The update, Update Disposition Summary, Close Report and Build
  all refuse, and the document and its properties stay identical. Against the code without the
  guard, the update inserts and synchronises from Y.
- **B**, the agenda TDoc: its row reserved, withdrawn or revised is accepted; reallocated to Removed is
  accepted; a list without it is refused.
- **C**, the anchor:
  - bootstrap and its log; a reader stores nothing;
  - accepted: appended lists, changed statuses, a withdrawn or reallocated anchor;
  - refused: a disjoint list, and a partly overlapping list without the anchor;
  - another meeting ID or URL: the old anchor is not used;
  - a contaminated report and a report without tables get no anchor;
  - `S4aA260090-manual` is no identity.
- **D**, the rules, pure, on the shapes of the two 6G lists.
- **E**, the ledger against `c48ba94`.

`tests/adhoc-tdoc-sessions.test.js` covers the Session column. The synthetic lists of the test
harnesses now have the row of their configured agenda TDoc, as real lists do (under no agenda item,
so no report places it).

## Evidence on the 6G smoke copy

The smoke copy is "SMOKE TEST 6G Media Minutes – SA4-e (AH) on FS_6G_MED", script
`1QuKJV6BCq…`, a Drive copy of live 6G from 2026-10-29.

- **19:18, 2026-10-30, still on candidate `3e820ea`, without this guard** (`d09649c` was not yet
  pushed; the runtime was read back to prove it).
  - Update Report Now read the Portal list of 86178 (36 TDocs of the earlier calls) into a report of
    116 TDocs.
  - The log said "Added 36 new, updated 0 statuses". The registered-documents structures were
    rewritten, the revision map had 15 revised documents, and 5 revision tables were moved. The
    collector and the formatting ran; the run took 237 s.
  - This reproduces the contamination of live 6G. The copy is kept as it is, as a contaminated
    fixture with an overlap that is not zero.
- **Later the same evening, on `d09649c`** (pushed once with the guarded procedure; read back: five
  files, byte-identical, the manifest identical as JSON).
  - Update Report Now, with the Portal still giving the 36-TDoc list, was refused:
    "The report was not updated: The TDoc list was not used: it could not be shown to be the list of
    this meeting (it has no row for S4aP260098, the agenda TDoc of this meeting -- it may be the list
    of another meeting)".
  - The agenda TDoc decided it. The 36 tables already in the report did not make the list
    acceptable.
  - **Passed.**

## MBS (86172), before a release: provisional

Read-only, 2026-10-30. Live MBS itself could not be read: the clasp token has no Docs or Drive read
scope, the Drive connector's authorisation has expired, and Document Properties can only be read by
running a script in the document.

- **Configuration, as recorded at its migration** (docs/EXISTING_REPORT_MIGRATION.md §5): ad-hoc,
  family MBS, meeting 86172, the Portal list `meetingId=86172`, agenda.csv, **no agenda TDoc**. So
  rule 3 (bootstrap from the report) decides its first update.
- **Raw list now:** 34 TDocs, S4aI260080 … 113.
  - The older list in tests/fixtures (10 TDocs) is entirely in it: the list has only grown.
  - S4aI260084 … 088 are revisions of main-meeting TDocs (S4-261749 … 783). Those originals are not
    in the list.
- **Report structure of 2026-10-13** (tests/fixtures/mbs-report-structure-damaged.json): 29 TDoc
  tables, all 29 in the raw list, and no table for a main-meeting TDoc. The 5 list TDocs it does not
  have were reserved on 10-14 … 10-20.
- **So if live MBS holds no TDoc table outside its list,** its first update bootstraps: anchor
  S4aI260080, uploaded and the lowest.
- **To be confirmed on live MBS before the release:** every TDoc table of the report is in the raw
  list.
