# Status dropdowns for an existing report (T-2026.10.8)

Developed 2026-10-06 on branch `feature/status-dropdown-migration`, on master `a289628`
(T-2026.10.7, `Code.js` 2.19.0). This document was written as the release candidate; §10 records
the release.

**Status: RELEASED AND DEPLOYED. Template release T-2026.10.8, `Code.js` 2.20.0, release commit
`7a6fcc4d6d34ffbb8b8763f1da501f669f3fc338`, tag `template-release/T-2026.10.8`. Published and
deployed to the master template on 2026-10-06 (§10). T-2026.10.7 and its tag were not changed.
No existing report was updated or converted.**

## 1. Why

T-2026.10.7 gives a report status dropdowns with its first Build Report from Scratch
(`docs/STATUS_DROPDOWN_RELEASE_CANDIDATE.md`). A report that existed before keeps its text
statuses: an update never converts them, and a rebuild would — but a rebuild writes the report
again, and the minutes typed into it are gone.

This release adds one menu action that converts the Status fields of an existing report **in
place**.

## 2. What it does

`SA4 Report > Report > Convert Status Fields to Dropdowns…` (template reports only; not in the
CENTRAL or Legacy menu; refused in the master template).

1. It reads the document once through the Docs API and looks at the Status cell of every TDoc
   table. Nothing is written.
2. It shows what it found — how many tables, how many are dropdowns already, how many it would
   convert, which stay text and why — and asks. It says that the fields are converted in place,
   that every field keeps its status, that minutes and everything else stay, and that the report
   is **not** rebuilt and the TDoc list is not read. Without the answer Yes nothing happens.
3. On Yes it takes the lock an update of the report holds, makes sure the report has its one
   "Document Status" definition, replaces the text of each eligible Status cell by a dropdown
   with the same value, reads the document back, checks every cell, and reports.

It changes no status, asks the Portal nothing, and changes nothing through DocumentApp.

### What is converted

A Status cell is converted when all of this holds:

- its table is a TDoc table whose TDoc number is a valid one and occurs once in the report;
- the cell is one paragraph of plain text;
- the text is a status of the dropdown.

| The cell holds | Result |
|---|---|
| A "Document Status" dropdown | unchanged ("already a dropdown") |
| One of the fourteen statuses, in any case, with any outer whitespace | that option |
| "replied to" (the Portal's wording) | replied |
| "parked", "Plenary" or "other", typed as the status | **that option** |
| Any other text ("not treated", "agreed, see minutes", …) | stays that text |
| Nothing | stays empty |
| A dropdown of another definition; more than one paragraph | skipped, named in the dialog |
| A TDoc number with two tables; a "TDoc" that is no TDoc number | skipped, named in the dialog |

Two rules meet at "parked", "Plenary" and "other". A status **that comes from the Portal** never
selects them (T-2026.10.7, unchanged): the Portal does not have them, they are decisions of the
meeting. A status **that is already written in the report** as exactly one of these words is
such a decision, and the conversion keeps it. Nothing becomes "other" that does not say "other".

## 3. How it reuses T-2026.10.7

Nothing of the T-2026.10.7 mechanism was duplicated. The conversion uses the same read
(`statusDocsGet_`, `buildStatusDocsIndex_`), the same table identification, the same definition
discovery and creation (`readStatusDropdownSourceDefinition_`,
`createStatusDropdownDefinition_`), the same mapping (`mapPortalStatusToDropdownOption_`,
`statusDropdownOptionByName_`), and the same request planning (`planStatusDropdownRequests_`:
one delete-and-insert pair per cell, from the end of the document to its start, a cell skipped
when its text is not exactly what was planned).

New: the plan of a conversion (`planStatusDropdownMigration_`), the comparison of everything
that is not a Status value (`statusDocsOtherContent_`), the run (`migrateStatusFieldsToDropdowns_`),
its two dialog texts and the menu function (`convertStatusFieldsToDropdowns`).

### One change to T-2026.10.7 behaviour: the revision guard

The writes that replace text by a dropdown address positions in the document. T-2026.10.7 sent
them right after reading; a change made by somebody else between the read and the write would
have moved the positions.

All such writes — of a build, of an update and of the conversion — now go through one sender
(`sendStatusDropdownInsertPairs_`). Every batch names the revision of the document its positions
were read at (`writeControl.requiredRevisionId`), and each following batch the revision the
previous one left. Google Docs refuses a batch when the document is at another revision, and
applies a batch as a whole or not at all. Without a known revision nothing is sent.

For a build or update this means: when somebody edits the document in that moment, the dropdowns
of that run are not written, the statuses stay text, and the next update writes them (the
self-heal of T-2026.10.7). Changes of the value of an existing dropdown name the dropdown and no
position; they need no revision and are sent as before.

## 4. The mark

`STATUS_DROPDOWNS = 1` (document property) is what makes a report a report with status
dropdowns: updates give new TDocs a dropdown and heal text statuses they wrote themselves.

The conversion sets it at exactly one point: **after the read back, when the run is complete** —

- no write failed,
- every cell that was to be converted is a dropdown of the report's definition with the option
  its text stood for, and no text is left beside it,
- every other Status cell is what it was,
- and the document does not differ outside its Status values from what was read before.

It is not set by the preview, not before the writes, and not by a run that ends in any other
way. Until then the report stays what it was: an unmarked report whose updates convert nothing.
A report in which some cells are dropdowns and the mark is missing works — dropdowns are read
and updated wherever they are — and the next run of the action completes and marks it. The
action therefore still asks and runs on an unmarked report in which nothing is left to convert.

## 5. When something fails

| What happens | Result |
|---|---|
| The report or the template cannot be read; the template has no "Document Status" | An error message. Nothing was written |
| The definition cannot be created | An error message. No Status was changed |
| A batch fails | The batches before it are applied and complete, the failed one is not applied at all, later ones are not sent. Every Status is either its dropdown or its text. Not marked |
| The document is edited while it runs | The batch is refused by its revision; as above |
| The reply does not name a revision | The run stops before the next batch |
| A failure is reported for a batch that was in fact applied | Everything is converted, the run says it is not complete, not marked; the next run finds nothing to convert, checks, and marks |
| A cell is not what was asked for afterwards | Named in the dialog ("to be looked at"). Not marked |
| The document differs outside its Status values afterwards | Said so. Not marked; the next run checks again |
| An update of the report is running | "Nothing was changed … Try again in a minute." |

In every case the action can simply be run again. A cell that has its dropdown is left alone, the
report's definition is found and used again, so no second definition and no second dropdown can
come of it. Batches hold at most 50 cells.

## 6. Compatibility

- A report that is not converted behaves as with T-2026.10.7: text statuses, no API call by its
  updates, no definition, no mark. Installing the release changes nothing by itself.
- A new report, and a rebuilt one, get their dropdowns as with T-2026.10.7; the rule for what a
  rebuild keeps is unchanged.
- A converted report is a T-2026.10.7 dropdown report: the same definition, the same mark, the
  same reader, the same update rule. That rule, unchanged: an update moves a selected value on
  only **from** reserved or available, or **to** revised. A Portal status "revised" therefore
  always applies — also over a value that was selected or typed by hand, parked, Plenary and
  other included. Every other Portal status leaves such a value alone. The conversion keeps a
  status as it is written; it does not protect it from later updates beyond this rule.
- The manifest is unchanged from T-2026.10.7 (Docs v1 advanced service, same scopes). The bundle
  has the same seven files.
- CENTRAL and Legacy do not have the action and do not use the feature.

## 7. Tests

| | |
|---|---|
| `tests/status-dropdown-migration.test.js` | 14 sections, the real conversion and the real update against the fake document and fake Docs API |
| Mutation checks of the conversion and the revision guard | 72 of 72 caught |
| Mutation checks of T-2026.10.7, re-run | 75 of 75 and 35 of 35 caught |
| The full suite | 95 files, all pass |

The fake Docs API got a revision model for this: every read names a revision, a write with
`requiredRevisionId` is refused at another revision, every reply names the new one. The smoke
test (§9) showed that Google Docs behaves the same way.

## 8. Risks that remain

After the smoke test (§9):

1. **Menu dialogs.** The smoke test ran the conversion function the menu action calls, under the
   same lock, and logged the texts of both dialogs; the dialogs themselves — the question, the
   answer No, the report — were not clicked through in the document. They are covered by the
   automated tests only. The No path writes nothing by construction: it ends after the preview,
   and the preview was shown to change nothing.
2. **The revision guard and other people.** The conflict of the smoke test was an edit made
   through the API. An edit typed by a person is a new revision in the same way, but a report
   that is being edited all the time may refuse the write of a run repeatedly. That is safe —
   nothing is lost, the run says so, it can be repeated — and applies to builds and updates too:
   their dropdowns are then written by a later update. Best run when nobody is typing.
3. **Size.** The largest run had 95 TDoc tables and two batches. A report with several hundred
   TDocs means more batches of the same kind and a larger read; a run that is cut off, by the
   time limit or otherwise, is finished by the next.
4. **Suggestions and unusual cells.** A Status cell with a pending suggestion, a comment anchor,
   or formatting split over several runs was not probed. Such a cell is either skipped by the
   plan or refused as part of its batch; neither loses content.
5. **The comparison of "everything else" is of text.** It proves that no text outside the Status
   values changed, not that no formatting did. The requests touch nothing but the Status cells.
6. **Text statuses that stay.** "agreed, see minutes" stays text by design; the dialog lists
   them. They can be corrected by hand and the action run again.
7. **A converted status follows the update rule from then on** (§6): the Portal's "revised"
   replaces a value set by hand. This is the behaviour of T-2026.10.7, seen in the smoke test,
   and not changed here.

## 9. Smoke test in a real Google document: passed

2026-10-06, on the disposable test report, with the candidate of commit `5d5abcf` installed in
its script project (seven files; the manifest identical to T-2026.10.7; one normal `clasp push`
without a prompt). No other report or project was touched.

The steps were run by a temporary helper file in that script project, one execution per run. It
called the conversion function of the candidate, unmodified, under the document lock — that is,
everything the menu action does after the answer Yes — and compared the document before and
after through the Docs API: every Status cell, a digest of all text that is not a Status value,
a digest of the Minutes cells, the definitions, the mark, the revision. Running a function of a
document-bound script remotely was not possible with the existing tooling, so each run was
started by hand in the script editor. The helper was removed afterwards; the project was read
back with exactly the seven files of the candidate. It is not part of the repository or of any
bundle.

### Run 1 — a controlled existing report

The report was made an existing one: 35 TDoc tables, not marked; 33 text statuses, 1 empty,
1 dropdown left from before; statuses typed in for ten tables; markers typed into Minutes cells.

| | Result |
|---|---|
| The read names a revision | yes |
| Preview | 35 inspected, 1 already a dropdown, 32 to convert, 2 left as text (1 of them empty), 0 skipped |
| After the preview (the state of the answer No) | no Status changed; other text identical; Minutes identical; definitions, mark and revision unchanged |
| Conversion | 32 converted, 1 already a dropdown, 2 left as text; marked |
| available, agreed, revised | the dropdown value of the same name |
| parked, Plenary, other | the dropdown value of the same name |
| "Replied to" | replied |
| "  NOTED " | noted |
| "agreed, see minutes" | still that text |
| empty | still empty |
| The dropdown that was there | unchanged |
| Everything else | other text identical; Minutes identical; exactly one "Document Status" definition |
| Second conversion | 0 converted; no Status changed; revision unchanged; the preview finds nothing to convert |

### Run 2 — Update Report Now on the converted report

The update completed. The report stayed marked, with one definition; the Minutes were identical.
One Status changed: a TDoc whose value had been set to "other" for the test, and which the Portal
lists as revised, became "revised". That is the update rule (§6), not a defect. Every other
controlled value, the text status and the empty one stayed.

### Run 3 — more than one batch

Every Status was made text again, the definition deleted, the mark removed; 60 synthetic TDoc
tables were appended, since the report had only 33 eligible cells.

| | Result |
|---|---|
| Conversion | 95 inspected, 93 to convert, **93 converted in two batches**, 0 not converted |
| Revision chain | the second batch was accepted with the revision the first one returned |
| Definition | exactly one, created by this run |
| Mark | set, after the read back |
| Everything else | other text identical; Minutes identical |
| Second conversion | 93 already dropdowns, 0 converted; content identical; revision unchanged |

### Run 4 — the document is edited in the middle of a run

Prepared as for run 3: 93 cells to convert, no definition, not marked. The unmodified conversion
was run; right before its second batch was sent, a few characters were inserted at the very
start of the document through the API — an edit that moves every position after it.

| | Result |
|---|---|
| First batch | 50 cells converted |
| Second batch | **refused by Google Docs**: the required revision does not match the latest revision |
| The run | not complete: 50 of 93 converted, 43 not converted; **not marked** |
| Writes at stale positions | none |
| After taking the inserted characters out again | other text identical to before the run; Minutes identical |
| The 50 converted cells | each a dropdown with the value of its text |
| The 43 others | each as it was |
| The next run | 50 already dropdowns, 43 converted, 0 not converted; the one definition used again; marked after the read back |
| Final state | 93 dropdowns, 1 text status, 1 empty; one definition; marked; Minutes identical |

The synthetic tables were removed afterwards.

### Not exercised

The dialogs of the menu action in the document (§8, 1), and an edit typed by a person during a
run (§8, 2).

### One thing learned about the tooling

`clasp push` (3.3.0) reports "Script is already up to date" and sends nothing when every local
file equals its remote counterpart — also when the project holds a file more than the local
folder. Removing the helper therefore took a push that differed in one file, followed by the push
of the exact candidate. A deployment is verified by reading the project back and counting its
files, not by the message of the push.

## 10. Release: complete

Each step was authorized on its own and verified before the next.

| Step | Result |
|---|---|
| Smoke test in a real Google document (§9) | **Pass** |
| Release commit | **Complete.** One squashed commit on master, `7a6fcc4d6d34ffbb8b8763f1da501f669f3fc338`, with `a289628f1bfefe064f916e93bdd635e64e3a38cf` (T-2026.10.7 and its closeout) as its only parent and the tree of the accepted branch, `abe9a30e85925da1c27aa30498d1cf8f4bab5b0b` |
| Automated tests and mutation checks, on the release commit | **Pass** (see below) |
| Tag and publication | **Complete.** The lightweight tag `template-release/T-2026.10.8` on the release commit; master and the tag pushed together in one atomic push, fast-forward, without force. Every older release tag is where it was, locally and published |
| The release bundle | Built by the release tool from the tagged commit: the seven production files. Its six source files are byte-identical to the candidate of the smoke test; `Release.js` differs from the candidate's only in the commit it names and the time it was built |
| Deployment to the master template's script project | **Pass** (see below) |

### Verification on the release commit

| | |
|---|---|
| Conversion tests (`tests/status-dropdown-migration.test.js`) | 131 checks, 0 failures |
| Status dropdown tests (`tests/status-dropdown.test.js`) | 190 checks, 0 failures |
| Mutation checks of the conversion and the revision guard | 72 of 72 caught |
| Mutation checks of T-2026.10.7 | 75 of 75 and 35 of 35 caught |
| The full suite | 95 files, 0 failing |

The full suite counts 6,451 checks in a checkout that has the repository's own, git-ignored
`.clasp.json`, and 6,449 in one that does not: two checks — that this very target is refused by
the release tool and by the adoption tool — can only run where that file exists. The suite was
run on the release commit in a checkout without the file (6,449, none failing); the two test
files concerned were then run on the same commit in the checkout that has it, with all their
checks passing. Nothing failed and nothing was skipped that could have run.

### Deployment to the master template

| | |
|---|---|
| Before | The project held T-2026.10.7, `Code.js` 2.19.0: seven files, each identical to that release's bundle, no other file |
| The bundle, checked again right before | The tag on the release commit; every file matching the bundle's own manifest; the ignore file admitting exactly the seven production files; no helper file |
| The push | One normal `clasp push` from the release bundle. Seven files pushed. No prompt (the manifest is the one of T-2026.10.7). No `--force` |
| Read-back | **Exactly seven files, each byte-identical to the release bundle** |
| `Release.js` as deployed | T-2026.10.8, template, 2.20.0, the release commit, the tag |
| The manifest as deployed | Unchanged from T-2026.10.7: Europe/Berlin, V8, Drive v3 and Docs v1, the same six OAuth scopes |

The deployment replaced the files of the script project and nothing else:

- the master template **document** was not modified, and its "Document Status" definition
  (`docs/STATUS_DROPDOWN_RELEASE_CANDIDATE.md`, §8) was not touched;
- no conversion was run on it, and no Build Report from Scratch;
- no trigger was changed;
- the disposable test report, the live reports, CENTRAL and Legacy were not touched by it.

### What is live

- **The master template: T-2026.10.8.** A report created from it has the action "Convert Status
  Fields to Dropdowns…" in its menu; as before, it gets its dropdowns with its first Build
  Report from Scratch and has nothing to convert.
- **The live 6G report: still T-2026.10.7**, its Status fields still text. Not converted.
- **The live MBS report: not yet on T-2026.10.8**, and not converted.
- CENTRAL and Legacy were not deployed and do not use the feature.

Next, and only for the 6G report: install T-2026.10.8 in its script project, verify by reading
the project back, and then run the conversion from its menu — in place, without rebuilding the
report. Each of these is a separate, explicitly authorized step.
