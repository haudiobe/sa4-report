# Status dropdowns (T-2026.10.7)

Developed 2026-10-05 on branch `feature/status-dropdown`, on master `3399cf3` (T-2026.10.6,
`Code.js` 2.18.1). This document was written as the release candidate; §8 records the release.

**Status: RELEASED AND DEPLOYED. Template release T-2026.10.7, `Code.js` 2.19.0, release commit
`cc28bf4ab808a2f30a66cfb8ca5392491f25509c`, tag `template-release/T-2026.10.7`. Published on
2026-10-05; deployed to the master template on 2026-10-06 (§8). T-2026.10.5 and T-2026.10.6 were
not changed. Existing reports were not updated.**

The Status of a TDoc table can be a native Google Docs dropdown, "Document Status", instead of
text: a decision is set in the Minutes with one click and shows in the colour of its option.

---

## 1. What was established before any code

Probes on disposable documents, 2026-10-05:

| Question | Result |
|---|---|
| Does a dropdown survive a copy made by DocumentApp within its document? | Yes |
| Into another document, or in a Drive copy of the whole document? | No: a dead chip remains. A definition belongs to its document |
| Can DocumentApp read, create or set a dropdown? | No: it is an `UNSUPPORTED` element, its text is `''` |
| Can the Google Docs API? | Yes: `createDropdownDefinition`, `insertDropdown`, `updateDropdownProperties`; a read returns the selected value, also after a change by hand |
| Can one execution change the document with DocumentApp, save it, and then use the API on it? | Yes (bound document, run from its menu) |

So nothing is copied. The definition is created in each report through the API, and the
dropdowns are inserted through the API.

## 2. Decisions

| # | Decision |
|---|---|
| 1 | The **master template** holds the definition "Document Status": option names, their order, their foreground and background colours. It is maintained there by hand. The code contains no option list and no colour. |
| 2 | Build Report from Scratch reads that definition through the API and creates an equal one in the report, unless the report has one of that title. **A report keeps the definition it was built with**; later changes of the template do not reach it. |
| 3 | Every path that creates a TDoc table writes the Status as **text**, as before. At the very end of a build or an update the document is saved, read once through the API, and each new TDoc whose Portal status has an option gets a dropdown in place of that text. |
| 4 | **Mapping.** A Portal status selects the option of the same name, case and outer whitespace aside; "replied to" selects *replied*. *parked*, *Plenary* and *other* are never selected automatically. Any other status, and an empty one, stays as the Portal's own text. Nothing is mapped to *other*. |
| 5 | The status summary, the discussion e-mails and the Portal sync read the value selected in a dropdown through the API. A text status is read as before. |
| 6 | **The sync keeps its rule**: a status moves on only from *reserved* or *available*, or to a "revised" status. A decision set in the Minutes is kept. |
| 7 | A Portal status **without an option**, for a TDoc that has a dropdown: the dropdown and its value stay; nothing is replaced by text; the status is logged and returned with the result of the update (`statusNotes`). |
| 8 | **Existing reports are not converted.** A report gets dropdowns only by a Build Report from Scratch of this release, which marks it (`STATUS_DROPDOWNS`) as soon as it has its definition. An update makes dropdowns only in a report that is so marked. |
| 9 | Template reports only. CENTRAL and Legacy do not use the feature. |
| 10 | Whatever is missing or fails — no Docs API service, the template not readable, no definition in it, a failed call, a failed save — the statuses are the text already written, and the build or update completes. |
| 11 | **A dropdown that could not be made is made later.** In a marked report every update looks for Status cells that are still plain text although their text is a status with an option, and makes their dropdowns. A status without an option, a cell that already has a dropdown and a table whose number occurs twice are never touched. A report that is not marked is not read for this. |
| 12 | **Build Report from Scratch keeps what was selected.** Before the document is cleared, the values of the dropdowns of a marked report are read, by TDoc number. After the rebuild a TDoc that is in the report again gets its previous value where the sync would have kept it: the value is an option of the definition, it is not *reserved* or *available* (those follow the Portal), and the Portal does not say "revised". Otherwise it gets the normal result. New TDocs get the option of their Portal status; a TDoc that is gone stays gone; a text status is written from the Portal as always. The values are held for that one build only; nothing is stored. |

## 3. How it works

- `noteStatusDropdownCandidate_()` — called by `buildSkeletonWithTdocTables()`,
  `appendTdocDetailTable_()`, `insertNewTdoc_()`, `insertRevisedDocTablesAfter_()` and, where
  the sync writes a text status, `applyTdocStatusUpdate_()`: remembers the TDoc number and its
  Portal status for this execution.
- `captureStatusDropdownValuesForRebuild_()` — in `buildSkeletonWithTdocTables()`, directly
  before the document is cleared (decision 12).
- `finalizeStatusDropdowns_()` — the last step of `runFullReportBuildCore_()` (as the phase
  "status dropdowns"), of `continuousUpdateCore_()` and of `collectRevisionsOnly()`:
  1. `saveAndClose()` of the active document;
  2. one `Docs.Documents.get()` with a field mask;
  3. the Status cells are found by structure — a table whose first cell reads "TDoc", its
     number beside it, the row labelled "Status" — never by a position remembered from before;
  4. the definition is created if the report has none, and the report is marked;
  5. per TDoc one pair `deleteContentRange` + `insertDropdown` (an empty cell: the insert
     alone), from the end of the document to its beginning, fifty pairs per `batchUpdate`; a
     batch is applied as a whole or not at all.
  A text that is no longer the status that was written is left alone.
- `readTdocStatus_()` — the one way to read a status: the cell's text when it has any, else the
  value selected in its dropdown. Used by `analyzeReportStatus()`,
  `detectTdocTablesInDocument_()` and `generateTdocDiscussionEmails()`.
- `applyTdocStatusUpdateToDropdown_()` — the rule of `applyTdocStatusUpdate_()` on the value of
  a dropdown. It writes nothing itself; the change is sent in the last step, resolved by TDoc
  number against the document as it is then (so a table that was moved in the same run is found).
- A Status cell that holds a dropdown is never written with `setText()`, in any path, including
  the old sheet-based one.

**API calls.** A report that is not marked: none, in any operation. A marked report: a build 1
read before the clear (a rebuild), then 1 read and 1 or more writes, plus 1 read of the template
and 1 write the first time; an update 1 read, plus 1 read and 1 write when it adds a TDoc, moves
a status on or makes up a dropdown; the status summary and the e-mail list 1 read each.

## 4. Production diff against T-2026.10.6

| File | Change |
|---|---|
| `appsscript.json` | the advanced service `Docs` (Google Docs API v1) beside `Drive`. No new OAuth scope: `documents` is already there |
| `Code.js` 2.18.1 → 2.19.0 | header; a new section "TDOC STATUS DROPDOWNS" (24 functions) before the upload-completion section; one line each in the four creation paths; the dropdown branch and one note in `applyTdocStatusUpdate_()`, a guard in `updateStatusWithStrictRules_()`; the status read in the three readers; the read before the clear in `buildSkeletonWithTdocTables()`; the last step in `runFullReportBuildCore_()`, `continuousUpdateCore_()` and `collectRevisionsOnly()` |

`template/ReportCreator.js`, `HyperLink.js`, both HTML files and the release tooling are unchanged.

### The manifest and the Legacy project

`appsscript.json` is one tracked file with two consumers:

- the **template bundle**: `tools/template-release.js` copies it into every release bundle;
- the **repository root as a clasp project**: the root `.clasp.json` (not tracked) points at the
  Legacy project with the root as its directory, and the root `.claspignore` lets through
  `appsscript.json`, `Code.js`, `HyperLink.js` and the two HTML files. A `clasp push` from the
  repository root would therefore send this same manifest, and this `Code.js`, to Legacy.

So the Docs service reaches Legacy only if somebody pushes from the repository root. Nothing in
the release procedure does that. If it happened, the effect of the manifest entry alone would be
the `Docs` name being available in that project; the feature would still not run there, because
it requires the template runtime (`Release.js`), and no scope would be added. Whether Legacy's
project accepts the service without an API enablement of its own is not known and not tested.

It can be confined to the template: the release tool could add the service to the bundle's
manifest and leave the tracked file as it was. That is a change of the release tooling and its
tests, and was not made; it is the option to take if Legacy's manifest must stay byte-identical.

## 5. Tests

- `tests/status-dropdown.test.js` — the real build, update, readers and sync against the fake
  document and a fake Docs API (`tests/helpers/fake-docs-api.js`): the mapping; the definition
  as an exact copy; reading a document; the build; the order save → read → write, and nothing
  after; new, revision and registration-section TDocs; the readers; the sync and the protection
  of decisions; unmapped statuses; 130 TDocs in three batches; moved and reallocated tables;
  every failure; reports with text statuses and their zero API calls; a dropdown that could not
  be made and is made by a later update; what a rebuild keeps and what it does not.
- The fake Docs API applies each request against the positions as they are after the requests
  before it, applies a batch as a whole or not at all, and refuses a write to a document with
  unsaved DocumentApp changes. The fake document refuses every change after it was saved and
  closed in an execution.
- Existing tests: the checks that pin "nothing else changed" and the version were extended for
  this release; no other expectation changed.

## 6. What the evidence is

Three kinds, kept apart.

**Automated tests** (§5) cover every rule and every failure path against a fake document and a
fake Docs API. They cannot show how Google Docs itself behaves.

**The smoke test in a real Google document** (§7) shows that: the advanced service installed
through the manifest, a definition created from a real source definition, native dropdowns with
their options, order and colours, editing by hand, and what an update and a rebuild do with a
value set by hand.

**Not exercised by hand**, and covered by automated tests only:

- a reallocation and *Update TDoc Revisions* with tables that are moved;
- *Prepare Discussion E-mails…* and *Report Status Summary* against a value set by hand;
- a time-driven update (the save at its end);
- a report of the size of a main meeting (the save at the end of a large build; several batches);
- a TDoc without a Portal status whose dropdown had a value before a rebuild (the insert into an
  empty cell);
- the field masks of the two reads were used in the smoke test but not looked at separately: a
  refused mask falls back to an unmasked read and would not have been noticed.

These were left out on purpose: they do not justify further changes by hand to the disposable
report before the release.

## 7. Smoke test: passed

2026-10-05, on the disposable copy of an ad-hoc report, with the candidate built from commit
`6b777f2` (the production files of the release candidate are those of that commit).

**Installation.** The target was verified as the disposable project before the push. One normal,
interactive `clasp push`: clasp reported that the manifest had changed and asked whether to
push and overwrite it; that question was answered with yes by hand. Seven files were pushed. No
`--force` was used. The report was given the document property `STATUS_DROPDOWN_SOURCE_DOC_ID`,
naming the document that has the real "Document Status" definition.

| Test | Expected | Seen | |
|---|---|---|---|
| Build Report from Scratch | Completes; every Status with an option is a native dropdown; none is lost or empty | Completed; mapped statuses (available, revised among them) were native, coloured, clickable dropdowns; no Status lost or blank | **Pass** |
| The definition | The options of the source, in its order, with its colours; a live dropdown, not a dead chip | Opening a dropdown showed the option set in the intended order (available, noted, agreed, revised, parked, merged, …) with the custom colours; it was editable | **Pass** |
| A "revised" Portal status | It applies, whatever was selected | A TDoc the Portal has as revised stayed revised | **Pass** |
| A value set by hand, then Update Report Now | Kept against a Portal status that is not "revised" | A TDoc the Portal has as available was set to parked by hand; after the update it was parked | **Pass** |
| The same, then Build Report from Scratch | Kept, by the rule of the sync | After the complete rebuild the same TDoc was parked | **Pass** |
| A dropdown replaced by plain text, then Update Report Now | The text becomes a dropdown again, with its value | The dropdown of that TDoc was replaced by the text "noted"; after the update it was a native dropdown showing noted | **Pass** |

What this settles of the questions the tests could not: the Docs advanced service works after a
push of the manifest, with no further enablement; a dropdown inserted where coloured Status text
stood looks as it should; the save at the end of a build and of an update, followed by the API
calls, works in a real report.

## 8. Release: complete

Each step was authorized on its own and verified before the next.

| Step | Result |
|---|---|
| Automated tests and mutation checks, on the release commit | **Pass.** Status dropdowns 190 checks; mutations 99 of 99 caught; the full suite 94 files, 6,320 checks |
| Smoke test in a real Google document (§7) | **Pass** |
| Release commit, tag, publication | **Complete.** One squashed commit on master, `cc28bf4ab808a2f30a66cfb8ca5392491f25509c`, with T-2026.10.6 as its only parent and the tree of the accepted branch; the lightweight tag `template-release/T-2026.10.7` on it; master and the tag pushed together, without force |
| The release bundle | Built by the release tool from the tagged commit: the seven production files. Its six source files are byte-identical to the candidate of the smoke test; `Release.js` names the release commit and the tag |
| "Document Status" definition in the master template document | **Pass** (see below) |
| Deployment to the master template's script project | **Pass.** The project held T-2026.10.6 before. One normal `clasp push`; clasp asked whether to push and overwrite the changed manifest, which was answered with yes by hand; seven files pushed; no `--force` |
| Read-back after the deployment | **Pass.** Seven files, each byte-identical to the release bundle; `Release.js` says T-2026.10.7, template, 2.19.0, the release commit and the tag; the manifest has Europe/Berlin, V8, Drive v3, Docs v1 and the six OAuth scopes it had |

### The definition in the master template

New reports read the definition from the master template document (decision 1), so it had to be
there before the deployment. It was not: the document had no dropdown definition at all.

It was created with one `createDropdownDefinition` request of the Docs API, from the accepted
definition of the document that was the source in the smoke test — the same request the released
code sends when it builds a report. No dropdown was inserted into the template's text. The
request was tied to the revision of the template that had just been read.

Verified afterwards, by a read through the API: exactly one definition titled "Document Status";
fourteen options, in this order and spelling — available, noted, agreed, revised, parked, merged,
approved, reserved, endorsed, withdrawn, other, replied, Plenary, postponed; the foreground and
background colour of every option equal to the source; no dropdown visible in the template; and
the rest of the document unchanged.

**A first attempt was stopped and undone by its own check, wrongly.** The one-off helper that
did the seeding compared what the API returned with `JSON.stringify()`, for which the order of
an object's keys counts; the API returns the keys of a colour (red, green, blue), and of other
objects, in no fixed order. It therefore reported fourteen colour differences and a changed
document, removed the definition it had created, and stopped. No colour value differed: all
fourteen foreground and background colours were numerically identical. The second version of the
helper compares a colour as its three numbers and the document serialized with its keys sorted;
with it the same seeding passed every check.

This was a fault of that diagnostic helper only. The released code never compares or hashes what
the API returns — it passes the colours on as it reads them — and was not changed.

### What is live

- The master template: T-2026.10.7. Reports created from it from now on get status dropdowns
  with their first Build Report from Scratch.
- Reports that existed before keep the release they have and their text statuses. None was
  updated as part of this release; updating one is a separate step for that report.
- CENTRAL and Legacy were not deployed and do not use the feature.
