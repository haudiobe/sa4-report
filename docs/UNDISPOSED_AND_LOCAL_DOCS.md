# Undisposed Documents and local document availability (T-2026.10.32, Code.js 2.27.0)

Built on T-2026.10.31 (Code.js 2.26.15, the concurrent-edit hotfix). Two changes, one release.
The code is the section "UNDISPOSED DOCUMENTS AND LOCAL DOCUMENT AVAILABILITY" of `Code.js`, right
after Goes to Plenary, plus seven marked changes to released functions (`tests/helpers/before-structure-fix.js`,
`UNDISPOSED_CHANGES`). Without them `Code.js` is that of T-2026.10.31 byte for byte.

## A. The Undisposed Documents section

**What it lists.** Every TDoc of the report that is registered in it (the TDocs of its Registration
of Documents table, as the Disposition Summary counts them) whose Status is **Available**, **Parked**
or **Reserved** and whose Plenary is **No or blank**. They are listed in the order of the report,
with the columns Document, Title, Source, Status and Agenda Item, and a link on each number. A count
line says "Not yet disposed, not going to plenary (N)". With none it says "None." and has no table.
A TDoc whose tables disagree about its Status is not listed; the Disposition Summary shows it as
unresolved.

**Where it is.** The section goes at the end of the opening item, under the next free number. That
number is never lower than .4, because .3 is kept for Document Reallocations.

| Report | Opening item has | Section |
|---|---|---|
| Standard (Audio 7, MBS 8, Video 9, RTC 10) | X.1.1 Opening, X.1.2 Registration of Documents (X.1.3 Document Reallocations when needed) | **X.1.4 Undisposed Documents** |
| Standard with TDocs allocated before the agenda sections | … X.1.4 Documents | **X.1.5 Undisposed Documents** |
| 6G Media (11) | 11.0.1 Opening, 11.0.2 Registration, 11.0.3 Document Reallocations, **11.0.4 IPR and antitrust reminder** | **11.0.5 Undisposed Documents** |

No heading of the report is renamed or moved for it. In 6G Media, 11.0.4 is the IPR reminder, so
the section is 11.0.5. If a heading appears later that takes its number (an "X.1.4 Documents", say),
the next update renumbers the section.

**When it is written.** It is written only in main-meeting reports; ad-hoc reports are left as they
are. Every update writes it, and only when what it says changes:
- In a report without Goes to Plenary, it is written right after the Disposition Summary, from the
  same read. With no summary it uses the update's own Docs API read, which a report with Status
  dropdowns has. Without either (text statuses) it makes one Docs API read of its own. Only in a
  runtime without the Docs API does it read the TDoc tables through DocumentApp, and then at most
  every 15 minutes (always while the report has no section yet).
- In a report with Goes to Plenary, it is written after the four-table summary, from that summary's
  read of the saved document.

The update after a Build from Scratch creates it. A change touches only the rows that differ: a new
TDoc is one row inserted, not a new table.

**What it owns.** It owns its heading, its note, its count line and its table (or "None."):
- **Text typed below its table** is kept.
- **Text typed between its heading and its table** stops it from being written. The update then
  says so: "move it below its table".
- **Nothing else is touched:** not the four disposition tables, the minutes, the statuses or the
  Plenary values.

Its table's first column is "Document", not "TDoc". The rest of the code reads every table whose
first cell is "TDoc" as a TDoc, registration, reallocation or summary table, so the section's table
is never one of them.

**Timing.** A report without Goes to Plenary and without a summary takes its statuses from a Docs
API read of the saved document (the update's own, or the section's), not from this update's writes. A TDoc table inserted by that same update,
or a status the update writes as text, appears in the section at the next update. The revisions and
the e-mail discussions of an update are compared with that same read.

**Cost** (DocumentApp calls on the fake document, 75 TDocs; the real time is measured in the pilot):

| Update | T-2026.10.31 | T-2026.10.32 |
|---|---|---|
| Goes to Plenary, unchanged | 198 | 202 |
| Goes to Plenary, one status and one Plenary changed | 309 | 313 |
| Goes to Plenary, a new TDoc and a Portal status | 1811 | 2031 |
| Status dropdowns, no Goes to Plenary, unchanged | 770 | 790 |
| Status dropdowns, no Goes to Plenary, a new TDoc | 2113 | 2158 (+ about 210 at the next update) |
| Text statuses, no summary, unchanged (+ one Docs API read) | 10803 | 10863 |
| No Docs API, no summary, unchanged (the 15-minute guard; smaller test report) | 3552 | 3702 |

A first version read every TDoc table through DocumentApp at every update in the text-status case
(3552 -> 14938 calls in the smaller report); the read of its own and the guard replace that.

## B. Local document availability

The Docs folder of the meeting-room server (`http://10.10.10.10/ftp/SA/SA4/Docs/`) holds the TDocs
uploaded in the room. Google cannot reach that private address, so the local listing agent lists
it (`"docsUrl"`, docs/LOCAL_LISTING_AGENT.md) into `Docs.listing.json` in its Drive folder.

**What a report reads.** On every update and every Build from Scratch, before the TDoc list is
applied, the report reads that listing, with the checks of the revisions' listings:
- it was written by the agent and is complete (checksum, link count);
- it is a listing of the Docs folder of this report's meeting-room server (the address before
  `/Inbox` in its Local Revisions URL, plus `/Docs/`);
- it is at most 30 minutes old.

**What counts.** A file is a TDoc only when its name, up to the first ".", is exactly a TDoc number,
and it sits directly in that folder:

| File | TDoc |
|---|---|
| `S4-261234.zip`, `s4-261234.ZIP` | S4-261234 |
| `S4-261234_r1.zip`, `S4-261234r1.docx`, `S4-2612345.zip`, `sub/S4-261234.zip` | none |

**Reserved → Available.** Such a TDoc is remembered (`LOCAL_DOCS_CONFIRMED`). While the TDoc list
still says "reserved" for it, the update takes it as "available". The existing Status sync then
moves the report's Status from Reserved to Available by its own rule, which only ever moves a Status
from Reserved or Available.
- **Never changed:** a final disposition (Agreed, Noted, Revised, Withdrawn, …), Parked, the Minutes
  and the Plenary values.
- **Only the list row of the run changes:** the TDoc list itself is not altered.

**Outages.** These change nothing and revert nothing:
- a missing listing;
- a stale, incomplete or foreign listing;
- an unreadable Drive folder;
- a meeting-room server that is down.

A TDoc confirmed once stays Available. *Check Connections* shows the last check: its result, its
time and how many TDocs have been confirmed so far. The Portal's own statuses and the public 3GPP
folders are read as before.

## Tests

- `tests/undisposed-local-docs.test.js`:
  - numbering: standard, with X.1.4 Documents, 6G Media 11.0.5, renumbering, a numbered heading
    with another style;
  - what it lists;
  - the writer: creation, unchanged, row by row, text typed below and above, "None." and back,
    nothing else touched;
  - the Docs API read: no DocumentApp call when unchanged, a read that does not fit is not trusted;
  - the row plan;
  - the Docs folder: the exact rule, every listing check, nothing taken back;
  - Reserved → Available: a final disposition and the Minutes untouched;
  - the ledger: T-2026.10.31 byte for byte.
- `tests/local-listing-agent-docs.test.js`: the real agent against a local server: the Docs file,
  accepted by a report, `-Group Docs`, without `docsUrl`, the Docs folder failing.
- `tests/goes-to-plenary.test.js` M: the DocumentApp calls with and without Goes to Plenary,
  compared with T-2026.10.31.
