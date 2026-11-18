# Goes to Plenary, the compact Status row, the four disposition tables and the revision relationship

**Status: CANDIDATE, not released and not deployed.** Branch `feature/plenary-dispositions`
(PR #26) is based on `af3779f` (T-2026.10.29, `Code.js` 2.26.13). The release candidate is
T-2026.10.30 (`release/T-2026.10.30`). Nothing has been pushed to the master template or to the
five production reports. The candidate runs only in the disposable MBS copy.

Code: the section `GOES TO PLENARY AND THE FOUR DISPOSITION TABLES (after T-2026.10.29)` of
`Code.js`, small changes to 21 released functions, and one menu item in
`template/ReportCreator.js`. Tests: `tests/goes-to-plenary.test.js`. The ledger layer is
`codeWithoutGoesToPlenary()` in `tests/helpers/before-structure-fix.js`; taking it out gives the
T-2026.10.29 files byte for byte.

**A report changes only when its rapporteur runs Report › Set Up Goes to Plenary….** A report
that has not been set up behaves exactly as under T-2026.10.29.

## 1. The compact Status row

```
| TDoc   | S4-262013 (revision of S4-262012)                       |
| …      |                                                         |
| Status | [agreed ▾]  Plenary [☑ Yes ▾]  A.I.: 15.4               |
```

- **No extra rows.** The Status cell holds the status, exactly as before (a dropdown, or text
  for a status without an option), then the code's own part:
  - "Plenary": a native dropdown, ☐ No / ☑ Yes, ☐ No by default.
  - "A.I.: X.X": shown only when Plenary is ☑ Yes and the SA4 plenary agenda item is known.
- **Why a dropdown.** A Google Docs checklist box cannot be read by Apps Script or by the Docs
  API (the API's `Bullet` has no checked state), so it could never decide the tables.
- **The A.I. refreshes on the next update.** Google Docs cannot show or hide text when a
  dropdown changes, so A.I. appears or disappears at the next update, not at once.
- **The selection belongs to the rapporteur.** No code clears it, a change of Status never
  touches it, and it is kept by every update, TDoc-list refresh, Update TDoc Revisions, Import
  Local Inbox Listing and Build Report from Scratch.
- **How the code finds its part.** The code's part begins with an invisible mark: a zero-width
  space and an em space before "Plenary". Every reader of the Status reads only what stands
  before it. That covers the status sync, the Disposition Summary, Close Report, Report Status
  Summary, the discussion e-mails and the Docs API index. The changes are in
  `findStatusInDocTable_`, `readTdocStatus_`, `findStatusText_` and `statusDocsCellEntry_`.
- **How the part is written.**
  - The code's part is written only through the Docs API, at the positions of one read, in
    batches tied to its revision (`finalizePlenarySelection_()`).
  - DocumentApp never rewrites such a cell. A status that is text is replaced in front of the
    mark (`setStatusCellText_()`), and it still becomes a dropdown as soon as it has an option.
- **The Status "Plenary" is retired.** A TDoc cannot be both "agreed" and "Plenary", so (the
  rapporteur's decision, 2026-11-18):
  - a Status "Plenary" becomes Plenary ☑ Yes;
  - the Status becomes the TDoc list's status, or "available" until the next update moves it
    on;
  - the option leaves the report's "Document Status" dropdown.

## 2. The SA4 plenary agenda item

The value lives in document properties, so that it is kept while A.I. is hidden:

| Property | Holds |
|---|---|
| `PLENARY_AGENDA_AUTO:<TDoc>` | what the update derived (§ Mapping) |
| `PLENARY_AGENDA_MANUAL:<TDoc>` | what the rapporteur typed over "A.I.: …" |
| `PLENARY_AGENDA_SHOWN:<TDoc>` | what the code last wrote there, to tell an edit from its own text |

What the rapporteur can do with "A.I.: …":

| Action | Result |
|---|---|
| Type over the value | Manual override; never overwritten, kept while Plenary is ☐ No |
| `-` | Blank on purpose; no A.I. shown |
| Clear it to "A.I.:" | Automatic again |
| A work item changes | Only the automatic value moves |

The Disposition Summary shows the item for every TDoc, also for those not going to plenary.

### Mapping

The sources (nothing is invented):

- **The meeting's agenda TDoc.** For SA4#138 this is S4-261857. It is read once and kept in
  the report.
- **The work items:**
  - the TDoc list's **Related WIs**;
  - the title tag (`[AMD_ARCH_Ph2-MED] …`);
  - the acronym of the SWG agenda item and of the items above it.
- **A CR's Release.** This is the TDoc list's **Release** column.

The plenary parts are found by their titles:

| Part | SA4#138 |
|---|---|
| maintenance: top-level "Release N and earlier matters" | **14** |
| the items under "Release N Features" | 15.1–15.7 |
| the items under "Study Items" | 16.1–16.7 |

Rules, first match wins:

| # | Rule | SA4#138 |
|---|---|---|
| 1 | A CR to the maintenance release or earlier | → 14 |
| 2 | Under an SWG "Release N and earlier matters" item (7.4 / **8.4** / 9.4 / 10.4) | → 14 |
| 3 | The work items name exactly one 15.x/16.x item; more than one → blank | 8.5 → 15.4, 11.x → 16.7 |
| 4 | TEI of the maintenance release or earlier (TEI19) | → 14 |
| 5 | The SWG item has the title of exactly one 15.x/16.x item | 8.8 → 15.7 |
| 6 | Otherwise blank, for the rapporteur | 8.3, 8.9 |

> **To confirm:** "Maintenance: use 8.4". In S4-261857, 8.4 is the MBS SWG item; the plenary
> item of that name is 14. The candidate writes 14 (a one-line change for 8.4).

An ad-hoc meeting has no SA4 plenary agenda: no A.I., and no agenda is read.

## 3. The four disposition tables, on landscape pages

**Content.** The Disposition Summary at the end of the report has four tables, in this order:

1. Agreed, Endorsed and Approved Documents – Goes to Plenary
2. Other Dispositions – Goes to Plenary
3. Agreed, Endorsed and Approved Documents – Not Going to Plenary
4. Other Dispositions – Not Going to Plenary

- Columns: TDoc (the number, with the link of its TDoc table), Title, Source, Status, SWG Agenda
  Item, SA4 Plenary Agenda Item.
- Each contribution is in exactly one table, in the order of the report.
- "Positive" means the statuses agreed, partially agreed, endorsed and approved.
  Everything else is "other", including a missing, unknown or conflicting status. The taxonomy
  is the existing one (`DISPOSITION_OF_STATUS_`).

**Regeneration.** Every update writes the summary, creating it where there is none:
- It rewrites the summary only when the content changes; a repeated update makes no API write.
- It never writes over text typed into the summary. Text typed below it is kept.
- It writes nothing when a status or a selection cannot be read.
- *Update Disposition Summary* stays, and Close Report writes the same tables. Close Report's
  counts and its "not yet disposed" list are unchanged.

**Landscape.** The summary stands in a section of its own:

- **The section:** a NEXT_PAGE section break before its heading, with
  `SectionStyle.flipPageOrientation`. The rest of the report (agenda, TDoc tables, minutes) stays
  portrait. Text typed below the summary gets a portrait section again.
- **Page size:** the document's. The Docs API can change only the orientation of a section, so
  an A4 report gets A4 landscape and a Letter report gets Letter landscape.
- **Column widths:** taken from the real page (its height less the side margins), with the Title
  widest. The tables flow over as many landscape pages as they need.
- **Repeated updates** add no break and change no orientation: each step is checked against a
  fresh read.
- **When it is set:** after every update, Build from Scratch, Update Disposition Summary and Close
  Report.

## 4. The revision relationship in the TDoc row

`S4-26XXXX (revision of S4-26YYYY)`, with both numbers linked where the TDoc list has their
links.

- **Source.** Only the TDoc list's own metadata counts: the predecessor's "Revised to" names the
  revision, or the revision's "Is revision of" names the predecessor. It never comes from
  numbering, titles, or a newer file in a revisions folder.
- **Several generations.** Each revision names its immediate predecessor (001 ← 012 ← 013).
- **Shown without a link:** a predecessor of another meeting.
- **Nothing shown:** two different predecessors, or a cycle.
- **Kept as it was:** the Revisions row and the minutes.
- **Updates.** Every update with the TDoc list refreshes the relationship, and removes it when
  the list no longer says so.
- **Readers.** Every reader of the TDoc number reads the part before " (revision of …)"
  (`safeCellText_`, the Docs API index, the upload completion).

## Migration

**Reports not set up yet:** run *Report › Set Up Goes to Plenary…*. It asks first, makes every
Status row compact (☐ No), retires the Status "Plenary", and is safe to run again. The next
update fills in A.I., the revision relationships and the four landscape tables.

**The disposable MBS copy (the first candidate's two-row layout):** the next update or Set Up
migrates it automatically:

1. The selection of each "Goes to Plenary" row is kept in `PLENARY_CHOICE_PENDING:<TDoc>`, and a
   manual "Plenary Agenda Item" becomes `PLENARY_AGENDA_MANUAL`.
2. Then the two rows are removed and the compact row is written.
3. A selection stays pending until its dropdown stands. A failed end of the run never loses one;
   the next update completes it.

**Rollback for one report:** restore the version named before the set-up.

## Tests

`node tests/goes-to-plenary.test.js` runs the real build, update, set-up, Update TDoc Revisions,
summary and Close Report on the fake document and fake Docs API. The fake now models cells that
mix text and several dropdowns, `insertText`, and section breaks.

| | What is tested |
|---|---|
| A | Every mapping rule on S4-261857; the cell parsers; predecessors (chain, conflict, cycle, other meeting) |
| B | A report that is not set up: unchanged; an ad-hoc report |
| C | Set-up: compact rows, exact layout, status readers, retired "Plenary", writes, run again |
| D | A.I. only for ☑ Yes; the four tables; landscape section and widths; revision relationships and links |
| E | Repeated updates: same document, zero API writes, nothing duplicated |
| F | Status changes, the Portal sync of a text status in a compact row, new TDocs, Update TDoc Revisions |
| G | Manual A.I., hidden while ☐ No, "-", cleared, work-item changes, agenda unreadable |
| H | Text below (portrait again) and into the summary; unreadable selection; Update Disposition Summary |
| I | Build from Scratch keeps selections, manual values, revisions; landscape after the next update |
| J | Close Report |
| L | Migration of the two-row layout; a failed end of run keeps the selection pending |
| K | Ledger, never throws, nothing fetched but the agenda, no `setText` on a Status cell, local listing agent untouched |

## What only the disposable copy can show

The fake API models what the code relies on, but not Google Docs' rendering:

1. The compact row's look: the spacing of the em spaces, and the two dropdowns side by side.
2. The landscape section:
   - Is the Disposition Summary on landscape pages, with the rest portrait?
   - Is there no blank page? (Google Docs puts an empty paragraph before a section break.)
   - Is the section kept when the summary is rewritten?
3. Removing "Plenary" from the Status dropdown keeps every other selection.
4. Does the SA4#138 TDoc list have "Related WIs", "Release" and "Is revision of" under those
   names? Without "Is revision of", only the predecessor's "Revised to" gives the relationship.
