# Goes to Plenary, the four disposition tables and the SA4 plenary agenda item

**Status: CANDIDATE, not released and not deployed.** Developed on branch
`feature/plenary-dispositions` from `af3779f` (T-2026.10.29, `Code.js` 2.26.13). It is
planned for template release T-2026.10.30. Nothing has been pushed to the master template or
to the five production reports.

Code: the section `GOES TO PLENARY AND THE FOUR DISPOSITION TABLES (after T-2026.10.29)` of
`Code.js`, twelve small changes to released functions, and one menu item in
`template/ReportCreator.js`. Tests: `tests/goes-to-plenary.test.js`. The ledger layer is
`codeWithoutGoesToPlenary()` in `tests/helpers/before-structure-fix.js`; taking it out gives the
T-2026.10.29 files byte for byte.

## What it does

1. **Goes to Plenary.** Every TDoc table of a report that is set up gets two rows right below
   its Status:

   ```
   | Disposition         | Revised to S4-26xxxx |
   | Status              | [agreed ▾]           |
   | Goes to Plenary     | [☐ No ▾]             |   ← new: "☐ No" / "☑ Yes"
   | Plenary Agenda Item | 15.4                 |   ← new: automatic, or typed by hand
   ```

   - The selection is independent of the disposition. It is "☐ No" for every existing and
     every new contribution.
   - The selection belongs to the rapporteur. No code clears it, and a change of Status never
     touches it.
   - It is kept by every update, every TDoc-list refresh, Update TDoc Revisions, Import Local
     Inbox Listing and Build Report from Scratch.
2. **Four disposition tables.** The Disposition Summary at the end of the report has four
   tables, in this order:
   1. Agreed, Endorsed and Approved Documents – Goes to Plenary
   2. Other Dispositions – Goes to Plenary
   3. Agreed, Endorsed and Approved Documents – Not Going to Plenary
   4. Other Dispositions – Not Going to Plenary

   The columns are TDoc (with the link of its TDoc table), Title, Source, Status, SWG Agenda
   Item and SA4 Plenary Agenda Item. Each contribution is in exactly one table, in the order of
   the report. Close Report writes the same tables, and its counts and its "not yet disposed"
   list are unchanged.
3. **Automatic regeneration.** Every Update Report Now and every automatic update writes the
   summary, creating it where there is none. The menu command **Report › Update Disposition
   Summary** stays. The update behaves as follows:
   - It rewrites the summary only when the summary would say something else. A second update
     with nothing new writes nothing at all, not even an API call.
   - It never writes over text typed into the summary. The dialog or the log names the text.
     Text typed below the summary is kept.
   - It writes nothing when a Status or a selection cannot be read. It never guesses.
4. **SA4 plenary agenda item.** The update fills in "Plenary Agenda Item" from the work item and
   the meeting's agenda (§ Mapping). A value typed by hand is a manual override and is never
   overwritten.

## Why a dropdown and not a checkbox

A Google Docs checklist box cannot be read by Apps Script or by the Docs API. The API's
`Bullet` has only `listId`, `nestingLevel` and `textStyle`, and nothing else in the document
records whether a box is ticked. A real checkbox could therefore never be authoritative for
the tables.

"Goes to Plenary" is instead a native two-option dropdown, "☐ No" / "☑ Yes". It is read exactly
like the Status dropdown: the text of the cell, else the value selected in it, through the same
single API read per execution (`buildStatusDocsIndex_()`). The rapporteur chose this option on
2026-11-18.

## The status taxonomy (unchanged)

The classes are those of `DISPOSITION_OF_STATUS_`, the classification the Disposition Summary
and Close Report already use (docs/DISPOSITION_SUMMARY.md). No status value was added.

| Table | Statuses |
|---|---|
| positive (tables 1 and 3) | agreed, partially agreed, endorsed, approved |
| other (tables 2 and 4) | noted, revised, merged, withdrawn, replied (replied to); available, reserved, parked, postponed, other; no status; a status that is not an option ("(not a known status)"); two tables of one TDoc with different statuses ("conflicting: …") |

Missing and unknown dispositions are safe: they are never counted as positive, and they show
as they are written.

### The Status "Plenary" is retired

The manual-only Status option "Plenary" (left to the plenary) said what "Goes to Plenary" says
now, and a TDoc cannot be both "agreed" and "Plenary". The rapporteur decided on 2026-11-18 that
"Plenary" leaves the Status dropdown:

- **Set Up Goes to Plenary…** gives a TDoc whose Status is "Plenary" (a dropdown or typed) the
  value "☑ Yes" and sets its Status to "available". The next update moves the Status on to the
  TDoc list's status, by the normal status sync. "Plenary" is then removed from the report's
  "Document Status" dropdown (`updateDropdownDefinitionProperties`).
- **An update** of a set-up report handles a "Plenary" that appears later the same way, using
  the TDoc list's status directly.
- **Build Report from Scratch** keeps it as "☑ Yes".
- The master template's dropdown still has the option until the template is changed
  (docs/ROLLOUT_T-2026.10.30.md). New reports built from it lose the option at their set-up.

## Mapping: the SA4 plenary agenda item

Sources (nothing is invented):

- **The meeting's agenda TDoc.** For SA4#138 this is S4-261857, which the build already
  downloads. The update reads it once and keeps it in the report (`PLENARY_AGENDA_ITEMS`).
- **The contribution's work items:**
  - the TDoc list's **Related WIs** column;
  - the tag of its title (`[AMD_ARCH_Ph2-MED] …`);
  - the acronym of its SWG agenda item and of the items above it (`8.5 AMD_ARCH_Ph2-MED (…)`,
    `11 FS_6G_MED (…)`).
- **A CR's Release.** This is the TDoc list's **Release** column.

The parts of the plenary agenda are found by their titles, never by a number in the code. At
SA4#138 they are:

| Part | SA4#138 |
|---|---|
| maintenance: top-level "Release N and earlier matters" | **14** "Release 19 and earlier matters" |
| work items: the items under "Release N Features" | 15.1–15.7 |
| study items: the items under "Study Items" | 16.1–16.7 |

Rules, first match wins:

| # | Rule | SA4#138 example |
|---|---|---|
| 1 | A CR to the maintenance release or earlier → maintenance item | CR to Rel-19 under 8.5 → 14 |
| 2 | Under an SWG item "Release N and earlier matters" → maintenance item | anything under 7.4 / **8.4** / 9.4 / 10.4 → 14 |
| 3 | The work items name exactly one 15.x/16.x item → that item; more than one → blank (ambiguous) | 8.5 → 15.4; 8.7 → 16.3; 11.3.1 → 16.7; "AMD_ARCH_Ph2-MED, Energy_PRO-MED" → blank |
| 4 | A TEI of the maintenance release or earlier → maintenance item | TEI19 under 8.8 → 14 |
| 5 | The SWG item has the title of exactly one 15.x/16.x item → that item | 8.8 "Other Rel-20 matters including TEI" → 15.7 |
| 6 | Otherwise blank, for the rapporteur | 8.3 Reports/Liaisons, 8.9 AOB → blank |

> **To confirm:** the request said "Maintenance: use 8.4". In the SA4#138 agenda, 8.4 is the
> **MBS SWG's** item "Release 19 and earlier matters"; the plenary item of that name is **14**.
> The candidate therefore maps every SWG maintenance item (7.4, 8.4, 9.4, 10.4) to 14. If the
> plenary column is to show 8.4 literally, that is a one-line change to rule 2.

An ad-hoc meeting has no SA4 plenary agenda. Its column stays empty and no agenda is read.

### Manual values

| The cell holds … | The update … |
|---|---|
| nothing | writes the automatic value |
| exactly what the update last wrote there (`PLENARY_AGENDA_AUTO:<TDoc>`) | treats it as automatic, and moves it when the work item changes |
| anything else (typed by hand) | keeps it: a manual override, never overwritten |
| `-` | keeps it, and the summary shows the cell as blank on purpose |

Emptying a manual cell makes it automatic again.

## Migration of existing reports

**A report changes only when the rapporteur runs Report › Set Up Goes to Plenary….** Before
that, a report under this release behaves exactly as under T-2026.10.29: three-table summary, no
rows, no agenda download.

The set-up:

- Asks first, with the numbers.
- Adds the two rows to every TDoc table that does not have them.
- Retires the Status "Plenary" (above).
- Makes the dropdowns, in batches tied to the document revision they were planned on, as the
  Status dropdowns are made.
- Changes no status, no minutes and no other row.
- Is safe to run again ("set up already").

The next update fills in the plenary agenda items and writes the four tables. If the report
has a three-table summary, that summary is rewritten in place. The existing presentation
decision that is recovered is the Status "Plenary" (→ "☑ Yes"); all others start "☐ No".

Rollback for one report:

- Restore the document from its version history (taken before the set-up).
- Or remove the two rows, delete the document property `GOES_TO_PLENARY`, and push the
  T-2026.10.29 bundle.

## Tests

`node tests/goes-to-plenary.test.js`. It runs the real build, update, set-up, summary and Close
Report on the fake document and fake Docs API, as an MBS report of SA4#138 with the real
agenda.

| | What is tested |
|---|---|
| A | Every mapping rule on S4-261857, precedence, missing and ambiguous mappings, the taxonomy |
| B | A report that is not set up: unchanged, and no agenda read; an ad-hoc report |
| C | The set-up: rows, dropdowns, Status "Plenary", option removed, writes tied to revisions, run again |
| D | The four tables: order, counts, columns, each TDoc once, links |
| E | Two more updates: same document element by element, zero API writes |
| F | The selection survives status changes, new TDocs, the Portal sync and Update TDoc Revisions |
| G | Manual values, `-`, an emptied cell, work-item changes, an agenda that cannot be read |
| H | Minutes, text below and into the summary, an unreadable selection |
| I | Build from Scratch keeps selections and manual values; "Plenary" does not come back |
| J | Close Report |
| K | The ledger: what changed and nothing else; never throws; nothing fetched but the agenda; the local listing agent is untouched |

Nine older test files read their source checks through the new ledger view, the menu test has
the new item, and one field list has the two new list columns. Every other test is unchanged.

## Open points for the disposable-copy test

These cannot be shown with the fake API and are checked in the disposable MBS copy:

1. `updateDropdownDefinitionProperties` with the existing `optionId`s keeps every other Status
   selection.
2. The real TDoc list of SA4#138 has the columns "Related WIs" and "Release" under those names.
   Without them, rules 1 and 4 and the Related-WIs part of rule 3 do not apply; the title tag
   and the SWG item still do.
3. The look of the two new rows next to the Status row.
