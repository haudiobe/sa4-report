# TDoc list identity

**Status: RELEASED AND DEPLOYED. Template release T-2026.10.17 (`Code.js` 2.26.1), release commit
`24c583a4a55c1ffdd219063b02c4dab10d09512e`, tag `template-release/T-2026.10.17`. Deployed to the master
template, MBS and 6G on 2026-10-30 (§ Release and deployment). Developed on branch
`fix/disposition-meeting-scope` (`d09649c`).** Section TDOC LIST IDENTITY of `Code.js`; tests `tests/tdoc-list-identity.test.js`.

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
Close Report and the session dialogs check the same way and store nothing. (Update Disposition
Summary read the list here too until the candidate after T-2026.10.17. It now reads no list: its scope
is the report's registration table, docs/DISPOSITION_SUMMARY.md § Which TDocs.) `TDOC_LIST_ANCHOR` is not among the keys copied when a report is adopted for the add-on
background runs (`ADDON003_ADOPTION_FIXED_KEYS_`); there the report is the evidence again (rule 3).

## When a list is refused

The error starts "The TDoc list was not used: it could not be shown to be the list of this meeting
(…)" with the reason, and is logged as `TDoc list identity: REFUSED -- …`.

| Path | What happens |
|---|---|
| Update (manual and automatic) | `{ success: false, notUpdated: true }` before anything: no TDoc table or agenda heading, no status or upload metadata, no registration table or registered-documents summary, no Session column, no revision moved, no Disposition cell, no `REVISION_MAP`, no abstract request, no structure repair, no Disposition Summary sync |
| Build Report from Scratch | refused before the document is cleared |
| Update Disposition Summary | after T-2026.10.17 (candidate): not affected. It reads no TDoc list; the summary is written from the report's registered TDocs. (T-2026.10.17: "Disposition Summary NOT updated", with the reason) |
| Close Report | NOT READY TO CLOSE. After T-2026.10.17 (candidate), the summary is written from the report; "External verification could not be completed", with the reason; every revised TDoc "could not be verified". (T-2026.10.17: nothing checked, nothing written) |
| Enable Session Column / Assign TDoc Sessions | nothing written; a saved assignment is kept, the Session column is not refreshed (the notice says so) |
| `downloadAndProcessFromWeb()` (legacy) | the error, the temporary file trashed, no table written |

There is no fallback to the TDoc tables of the document. The zero-overlap guard of the Disposition
Summary (`writeDispositionSummary_()`, `c48ba94`) stays as a second line.

## Tests

`tests/tdoc-list-identity.test.js`:

- **A**, the contamination case: report X already holds TDocs of meeting Y, and the Portal returns
  Y's list, which overlaps the report. The update, Close Report and Build all refuse, and the document
  and its properties stay identical. Against the code without the guard, the update inserts and
  synchronises from Y. (After T-2026.10.17, candidate: Update Disposition Summary reads no list and
  leaves the summary of the report as it is; Close Report is not ready, "External verification could
  not be completed".)
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

## Release and deployment

Release T-2026.10.17, 2026-10-30. It contains this guard and the meeting scope of the Disposition
Summary (docs/DISPOSITION_SUMMARY.md).

| Step | Result |
|---|---|
| Release commit | `24c583a4a55c1ffdd219063b02c4dab10d09512e`, on the accepted candidate `d09649c` (with the documentation commit `b0739c3`). It adds the version line 2.26.1 (2026-10-30) and its changelog entry; without them, `Code.js` is the candidate's. The ledger takes the 2.26.1 header out with the identity change, and the version-pinned tests say 2.26.1 |
| Tests on the release commit | 112 files, 7,657 checks, 0 failures. Among them: tdoc-list-identity 52, disposition 222, release-file-set 48, template-release 26, after-10-8-ledger 38, report-structure-fix 75, configuration-guard 32, template003-menu 116, abstract-cache-retry 36, meeting-86178-production 118 |
| Tag and publication | Lightweight tag `template-release/T-2026.10.17` on the release commit. One atomic push, without force: `origin/master` fast-forwarded `3b0638f` → `24c583a`, and the tag published. Every older tag is unchanged (T-2026.10.16 still `5d5699d`) |
| Release bundle | `tools/template-release.js --release T-2026.10.17` from the tagged, clean commit: the five production files; the four source files equal to the tag (line ends aside); `Release.js` naming T-2026.10.17, template, 2.26.1, the release commit and the tag |
| Master template runtime | Container "SA4 Report Template" and its runtime (T-2026.10.16, `5d5699d`) proven. One normal push, no `--force`, at 20:00:30. Read back: exactly five files, four byte-identical, the manifest identical as JSON; it runs T-2026.10.17 |
| MBS runtime | The adoption bundle from the tag. Proven: container "MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e", on T-2026.10.16. Pushed at 20:00:58 and read back as above. Script files only |
| 6G runtime | The adoption bundle with `--allow-legacy-target`, authorised for T-2026.10.17 only. Proven: container "6G Media Minutes – SA4-e (AH) on FS_6G_MED", on T-2026.10.16. Pushed at 20:01:26 and read back as above. Script files only |

All three pushes were made with `clasp --user haudiobe.ts@gmail.com`, each from its own scratch folder,
never from the repository root and never with `--force`. **The deployment ran nothing in the
reports:** no update, no Update Disposition Summary, no Close Report, no build, no repair. No trigger
was created, started or stopped, and 3GU was not touched.

**Gates before the release:**

- **MBS (no agenda TDoc, so its first update sets up the anchor from the report).** Checked on a DOCX
  export of live MBS from 2026-10-30 against the raw Portal list of `meetingId=86172` from the same
  day:
  - 34 TDoc tables against 34 list TDocs (S4aI260080 … 113), 34 in common, none on one side only.
    S4aI260084 … 088, revisions of main-meeting TDocs, are TDocs of its own list.
  - `evaluateTdocListIdentity_` accepts this data.
  - Expected at its next normal update: `TDoc list identity: the list of this meeting (report; anchor
    S4aI260080 recorded)`. A different list is refused.
- **6G, the real negative smoke** (§ Evidence on the 6G smoke copy). With `d09649c`, the wrong
  36-TDoc Portal list was refused for want of S4aP260098, although the copy already held the 36.

**Live 6G after the deployment:**

- **Its automatic updates are intentionally paused,** and stay so until the recovery is decided.
  The deployment did not change any trigger.
- **Its content is contaminated and not yet repaired:** the 36 TDocs of the earlier calls, inserted
  by updates under T-2026.10.16 from the wrong Portal list, with what those updates also wrote.
- **Its Disposition Summary of 150 TDocs is kept as evidence** and has not been regenerated.
- While the Portal returns the wrong list, any update of 6G is refused (no row for S4aP260098).
- The runtime to roll back to is T-2026.10.16 (its bundles in `dist/`).
