# Issue #28: architecture and implementation assessment

Written before any code for issue #28 ("Post-SA4#138 maintenance: titles, provisional revisions, AHG sessions, agenda
diffs and UI"). It records what the code does today, which existing models each requirement reuses, and how the work is
split into three PRs.

## Baseline

| | |
|---|---|
| Production (all five SA4#138 reports) | T-2026.10.33, Code.js 2.27.1, `b7299c1` (deployed 2026-11-20; evidence `dist/backups/sa4-138-deploy-T-2026.10.33/`) |
| Newest candidate | T-2026.10.34, Code.js 2.27.2, `21e0b8e` ("handed to the plenary"): dry runs only for all five reports, never written |
| Master template | T-2026.10.30 (`729ebfa`), still the head of `origin/master` |
| Base of issue #28 | `21e0b8e`, pushed as `origin/fix/close-report-handed-to-plenary` so each PR shows only its own diff |

`21e0b8e` was chosen over `b7299c1` because phase B extends Close Report, which T-2026.10.34 changed; building on .33
would make that conflict certain. If T-2026.10.34 is not released, phase A rebases onto `b7299c1` without conflict
(the two layers touch different functions; the ledger test proves each layer separately).

Tests on the base: 132 test files, all passing (`node tests/<file>`; about 5 minutes serially, 1 minute 6-way parallel).

## Overlapping work checked

- Open PRs #10 (`fix/bounded-update`) and #14 (`fix/reviewer-run-breaker`) are 44–45 commits behind master and touch
  the update budget and the Reviewer, not titles, dispositions, sessions or agendas. No overlap; left alone.
- Issue #27 (re-allocation only into the destination table) is explicitly deferred ("do not implement now"); nothing
  here implements it, and phase C's agenda diffs only *propose* reallocations.
- No other branch or worktree touches title sync, provisional revisions, carry-forward or the plenary Agenda/Schedule
  folders.

## What the code does today (per requirement)

### Titles (phase A)
- A title is written only when a TDoc table or a registration row is **created** (`insertNewTdoc_`,
  `buildSkeletonWithTdocTables`, `createSummaryTable_`, `updateRegisteredDocumentsTable_` appends only).
- No update rewrites a title; there is no "title changed" detection. The generated tables (Disposition Summary, the
  four plenary tables, Undisposed Documents) copy the title of the TDoc table, so a change made there propagates.
- Measured on the fake document: after the list title of S4-263008 changes, the report still says the old title in
  all three report types (text statuses, dropdowns, Goes to Plenary).

### Final disposition headings (phase A)
- The summary has one heading, "Disposition Summary" (Heading 1). The class lines ("Agreed, Endorsed and Approved
  Documents – Goes to Plenary (12)" …) are bold Normal paragraphs; each document is a table row. There is no
  per-document heading. Decision (2026-11-23, the user): the class lines become Heading 2.
- Several readers end the summary at the first heading after its own (`findDispositionSummarySection_`,
  `plenarySummarySectionOfContent_`, `nextPlenaryLandscapeStep_`). Without care, Heading 2 class lines would truncate
  the section, make the update refuse to write it, or insert a portrait section break above the first class line.

### Performance (phase A)
Stage-attributed DocumentApp calls of an **unchanged** update, 250 TDocs, fake document (relative measure):

| Stage | Text statuses | Dropdowns |
|---|---|---|
| Status checks (`updateTdocStatus_`) | 17,750 | 0 (decided from the snapshot) |
| Existing-TDoc scan (two passes over every table) | 5,057 | 541 |
| Structure check (no snapshot) | 4,148 | 0 |
| Upload completion | 2,500 | 0 |
| Registration links | 2,293 | 0 |
| Looking for a Disposition Summary that does not exist | 1,649 | 1,649 (75% of the update) |
| Reallocation map | 1,041 | 0 |
| **Total** | **35,263** | **2,190** |

Goes to Plenary reports were already lean (552). Remaining costs that phase A does not address are listed in
docs/ISSUE28_PHASE_A.md § Not done.

### Provisional revisions (phase B)
- Status is a dropdown or text in the Status row; no per-TDoc property holds it. The Plenary Yes/No lives in the
  compact Status row; the plenary agenda item in `PLENARY_AGENDA_AUTO/MANUAL:<TDoc>` document properties.
- "Revised to" comes **only** from the TDoc list (`getRevisedTo_`, `buildRevisionMapFromTdocList_` → `REVISION_MAP`),
  and `setDispositionRevisedTo_` writes "Revised to X" into the Disposition row, never over an existing one.
- Local files `S4-261234r01.docx` / `_r1` map to their parent TDoc (`parseSA4DocumentId_` family match) and never
  create a TDoc table; the local Docs folder uses an exact match. A draft file named with *another* TDoc number does
  create a minimal table (`insertRevisedDocTablesAfter_`) — a risk phase B must not widen.
- Close Report: `result.problems` blocks; `result.informationNotes` is the hook for information; warnings exist only
  for revisions.

### AHG sessions (phase C)
- `ADHOC_SESSIONS` (`{v:1,nextId,sessions:[{id:'sN',label,date,endDate?,start,end}]}`) with stable ids; codes A01…
  are derived. `ADHOC_TDOC_SESSIONS` holds manual overrides (`add`/`set`, `clock`).
- Assignment today uses **Uploaded** only, cutoff = planned end (or 23:59:59), inclusive; reserved-not-uploaded gets no
  session. `tests/adhoc-tdoc-sessions.test.js` asserts that "Reservation date" is not read. Issue #28 changes the rule
  to reservation time with a +6 h cutoff — a deliberate rule change, made opt-in per report (phase C).
- No carry-forward exists; "considered again later" is a manual `add` override.
- Document Reallocations are a table in the document; only the dialog adds rows. Nothing reads the plenary
  Agenda/Schedule folders.

### Plenary Agenda/Schedule directories (phase C), observed 2026-11-23
- `…/Inbox/Drafts/Plenary/Agenda/`: "SA4#138 Tdoc allocation rev0…rev9.docx". rev3–rev9 all show the same listing
  time (2026/11/23 7:28, re-uploaded), so **the listing time cannot pick the newest**; the `revN` in the name must.
  Each file holds per-agenda-item allocations as short numbers with disposition suffixes (`2229app`, `1858n`) and
  chains (`1859->2253->2260app`); a legend colour marks "moved to a different A.I.". The AHG telcos are free text under
  "Telcos", with real errors: "CEST" for December/February dates, "9:00:10:00", a duplicated RTC line.
- `…/Plenary/Schedule/`: four `.xlsx` (Draft, -MBSChair, -v2, plain) with no clean version order; a room × slot grid
  whose day headers read "Monday 20 … Friday 24" for a 16–20 November meeting. Usable as a low-confidence hint only.

### Local listing agent
Apps Script never fetches 10.10.10.10; the agent writes JSON to a Drive folder, read only when
`LOCAL_LISTING_DRIVE_FOLDER_ID` is set. Nothing in issue #28 sets it, installs the task, or adds a reference.

## Split into PRs

| PR | Branch (base) | Content |
|---|---|---|
| A | `feature/i28-a-titles-headings-perf` (`fix/close-report-handed-to-plenary`) | title sync + review dialog, Heading 2 class lines, update costs |
| B | `feature/i28-b-provisional-revisions` (A) | "Revision expected" state, transfer once on registration, conflicts, Close Report check |
| C | `feature/i28-c-ahg-sessions-agenda` (B) | reservation/+6 h assignment (opt-in), carry-forward, agenda/schedule diffs, Meeting Control view |

Stacked because B and C both use the update hooks and title data of A, and C's Meeting Control view lists B's pending
decisions. Each phase is a separately removable ledger layer (`tests/helpers/before-structure-fix.js`).

## Constraints kept by every phase
No write to the completed SA4#138 reports, the master template, CENTRAL, Legacy or any bound script; no deployment; no
merge without approval; no reactivation of the local listing agent; manual minutes, dispositions and overrides are
never overwritten; issue #27 not implemented.
