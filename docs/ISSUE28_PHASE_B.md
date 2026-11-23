# Issue #28 phase B: provisional revision decisions (Code.js 2.29.0, candidate)

**Status: candidate, not released, not deployed.** Branch `feature/i28-b-provisional-revisions`, stacked on phase A
(docs/ISSUE28_PHASE_A.md). Code: section "ISSUE 28 B" at the end of `Code.js`, two one-hook changes
(`continuousUpdateCore_`, `closeReportCore_`), one menu item. Ledger: `codeWithoutIssue28B` → phase A byte for byte.

## The problem
In the meeting a contribution is agreed "with changes"; its revision is a new TDoc that 3GU registers later (minutes
or hours). Until now the rapporteur had to remember to set the revision's Status, Plenary choice and plenary agenda
item once it appears, and was tempted to set the *original* to "agreed", which is wrong.

## What the rapporteur does
*Report › Revision Expected…*: TDoc (autocomplete from the Registration of Documents table), decision (agreed,
endorsed, approved, noted, merged, replied, postponed; "partially agreed" is no option of the template's dropdown), Goes to Plenary (Yes/No/–), plenary agenda
item, optionally the expected TDoc number. **The original's Status is not changed.** The dialog lists every record with
its state, a search box, **Close…** (while waiting or in conflict, see below) and Undo transfer, and the log.

## What an update does (only an interactive update, as for the summary and Goes to Plenary)

| Situation (from the TDoc list of this update) | Result |
|---|---|
| No "Revised to" / "Is revision of" for the original | waits ("no revision is registered in 3GU yet") |
| "Revised to" names a TDoc the list does not have yet | waits, naming it |
| One registered successor with a TDoc table, Status reserved/available | **transfer**: its Status gets the decision; the original becomes "revised" (only if undecided); recorded `transferred` |
| The successor shows the decision already | recorded `transferred`, nothing written |
| The successor has another decision (set by hand) | **conflict**, nothing written |
| Two or more successors | **conflict**, nothing written |
| An expected number was given and the list makes another TDoc the revision | **conflict**, nothing written |
| The successor is revised further (A → B → C) | followed to the newest registered revision (≤ 5 generations); a cycle is a conflict |

- **Exactly once:** a `transferred` record is never processed again; a later manual change of the revision stays.
- **Source of truth:** only the TDoc list's "Revised to" of the original and "Is revision of" of the successor. Never a
  number pattern, a title, or a file such as `S4-261234r01.docx` / `_r1` in a revisions folder (the local `r01` files
  stay in the Revisions cell of their own TDoc as before; the dialog refuses them as an expected successor).
- **Status write:** the update's own status path — text, or the Status dropdown at the end of the run
  (`finalizeStatusDropdowns_`), only while the revision still shows what was read (reserved/available).
- **Goes to Plenary:** when the revision's compact Status row is made in that update (the normal case: a new
  revision), it gets the recorded choice (`PLENARY_CHOICE_PENDING:<tdoc>`) and the agenda item
  (`PLENARY_AGENDA_MANUAL:<tdoc>`, unless one was typed for it). A revision that already has a Plenary dropdown is
  **never changed** (the code never changes an existing Plenary dropdown); if its choice differs, the record says
  "set it by hand" and Close Report blocks until then. Without Goes to Plenary the choice is kept in the record only.
- **Undo:** "Undo transfer" is done by the next update, through the same path, only if the revision still shows the
  transferred decision (else refused, said); it does not undo a Plenary dropdown already made (the dialog says so).

## Close Report
Blocking (in `result.problems`): a record waiting for 3GU, a conflict, an undo still to be done, a Plenary choice to be
set by hand. Rationale: the decision is not yet in any registered TDoc, so the report is not complete.

**Decision 2026-11-23: blocking, with an explicit, auditable way out.** When a successor is never registered, the
rapporteur closes the record in the dialog ("Close a record without a transfer"):

| Outcome | Meaning |
|---|---|
| `not-submitted` | the revision was never submitted or registered |
| `decided-on-original` | decided on the original TDoc instead (its Status is then set by hand, as any decision) |
| `withdrawn` | the revision was withdrawn |
| `other` | resolved otherwise |
| `cancelled` | the record was made by mistake or is no longer wanted |

A reason of at least 10 characters is required. The record keeps `resolution: { outcome, reason, by, at, before }`
(`by` the e-mail Apps Script gives for the user, else `dialog`; `before` the state and reason it was closed from) and the
log gets an entry with the same. Such a record no longer blocks: Close Report lists it under **Information** with outcome,
reason, who and when (`appendProvisionalRevisionInformation_`, a hook after the classification). It is never transferred
afterwards, even if a revision is registered later (test F2). Nothing is written into the document by closing a record.

## State
`PROVREV_<original>` (JSON, ~300 bytes) and `PROVREV_LOG` (≤ 8.5 KB) in the report state; the Plenary keys of the
existing Goes to Plenary model. No new document structure; nothing is written into the TDoc tables when a record is made.

## Configuration
None needed. A report without records behaves exactly as phase A (test G: same DocumentApp calls, same result).

## Migration
None: the feature starts empty. Old reports are unaffected until the rapporteur records something.

## Rollback
Code rollback is safe at any time: older code ignores `PROVREV_*`. Records waiting at that moment are simply not
transferred (set those revisions by hand); transfers already made stay as written Status values.

## Tests
`tests/i28b-provisional-revisions.test.js` (A–H): validation incl. `r01` names; every decision of the pure function;
recorded/waiting/transferred-once on an MBS report of SA4#138 with Status dropdowns; conflicts (hand decision, two
successors, unexpected successor) and their resolution; Goes to Plenary for a new and an existing revision row; a typed
plenary agenda item kept; cancel; undo (done, and refused after a hand change); F2: closing a record without a transfer
(outcome and reason required, audit fields, information in Close Report, no transfer afterwards); no cost without
records; ledger.

## Open questions
1. ~~Should a waiting record block Close Report or only warn?~~ Decided 2026-11-23: it blocks; an explicit resolution
   with a reason ends the block (above).
2. Should "Plenary to be set by hand" be resolved automatically by replacing the existing dropdown (delete + insert
   through the Docs API)? Not done: the code has never changed an existing Plenary dropdown, and a live session may be
   editing it.

## Independent review (2026-11-23) and what changed

| # | Finding | Outcome |
|---|---|---|
| 1 | A queued dropdown write (transfer or undo) was sent at the end of the run without checking the fresh read: a delegate's choice made during the update was overwritten | **Fixed.** Each provisional write carries what it expects (`run.expect`); `planStatusDropdownRequests_` sends it only if the fresh read still shows that, else leaves it (one guard line, ledgered). Test R |
| 2 | "Set the Plenary choice by hand" blocked Close Report for ever | **Fixed.** Every update re-checks such records and clears the flag once the revision's dropdown shows the choice. Test E |
| 3 | A record was marked transferred before the write was confirmed; a failed write was lost | **Fixed.** New states `transferring` / `undoing`; the next update confirms against the revision itself: shows the decision → transferred; still undecided → written again; decided otherwise → transferred with a note "changed by hand after the transfer" (kept, not blocking). Close Report asks for the confirming update. Test R |
| 4 | Two records naming one revision: the Plenary choice was last-wins | **Fixed.** Different choices → both conflicts; identical → one transfers, the other is "shared". Test R |
| 5 | A chain stopped at a middle revision whose own "Revised to" was not registered yet | **Fixed:** waits. Test R |
| 6 | Plenary rows not readable (no status index) were treated as "pending" | **Fixed:** treated as "to be set by hand" (verified by the next update) |
| 7 | "partially agreed" offered but no template dropdown option | **Fixed:** not offered |
| 8 | Undo offered where nothing was written; the agenda item not restored | **Fixed:** Undo only when this code wrote the Status; the agenda item it wrote is removed on undo. The original stays "revised" (as the list says) |
| 9 | An expected number anywhere in the chain is accepted (B expected, A→B→C written to C) | By design: the decision belongs to the newest revision; documented here |
| 10 | Close Report reads the state of the last update | Messages now say "as of the last update … run an update" |
