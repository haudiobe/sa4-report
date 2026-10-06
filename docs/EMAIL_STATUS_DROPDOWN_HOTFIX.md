# The Status of a TDoc in discussion e-mails (T-2026.10.10)

A hotfix on T-2026.10.9, developed and released on 2026-10-06. This document is the record of
the release; it was written after it.

**Status: RELEASED AND DEPLOYED. Template release T-2026.10.10, `Code.js` 2.21.1, release commit
`03ec9fc92c9227adfe3342c26ef4dc6b568162c7`, tag `template-release/T-2026.10.10`. Published and
deployed to the master template on 2026-10-06 (§5). T-2026.10.9 and its tag were not changed. No
existing report was updated.**

## 1. The regression

In a report whose Status fields are native dropdowns (new template reports since T-2026.10.7,
converted reports since T-2026.10.8), a generated discussion e-mail showed an **empty Status**
for the TDoc.

The e-mail copies the TDoc table cell by cell through DocumentApp. DocumentApp has no text for a
dropdown: it sees an element of type UNSUPPORTED, and the text of the cell is empty. The table
renderer of the e-mail export, `docTableToHtml_()`, therefore wrote an empty paragraph
(`&nbsp;`) into the Status cell.

Only the e-mail was affected. Generating an e-mail writes nothing to the report, so no report
was damaged: the dropdown and its value were always intact. The eligibility check of Generate
(reserved, agreed and approved TDocs are refused) was not affected either; it has read the
status through `readTdocStatus_()` since T-2026.10.7.

## 2. The fix

One function changed: `docTableToHtml_()` in `Code.js`. Nothing else in the production files
changed besides the version line and the changelog entry.

- The Status row is the first row labelled "Status" or "TDoc Status" — the row
  `findStatusInDocTable_()` finds.
- For the value cell of that row, and only when the cell holds a dropdown and no text, the
  renderer writes the value `readTdocStatus_()` reads, HTML-escaped, in the paragraph an
  ordinary cell has.
- `readTdocStatus_()` stays the one way a status is read. The renderer has no dropdown logic of
  its own.

Unchanged:

- a text Status, known or not, is rendered as before;
- text typed beside a dropdown is the status (the rule of `readTdocStatus_()`), rendered as text
  always was;
- a dropdown that cannot be read (no Docs API, a failing read) leaves the cell empty; nothing
  throws;
- every other cell, the subject, the file name and the rest of the e-mail;
- the colours of the dropdown options are not reproduced in the e-mail.

The added lines only read. `readTdocStatus_()` uses the one Docs API read per execution that
Generate already made for its eligibility check. No write, no save, no property and no Portal
request was added.

## 3. Tests

`tests/email-status-dropdown.test.js` (40 checks) runs the real build of a report, the real
renderer, both e-mail builders and the real Generate function on the fake document and the fake
Docs API. Among its checks:

- every one of the 14 Document Status values, "Plenary" with its capital;
- the e-mail table of a TDoc with a dropdown is, byte for byte, that of the same TDoc with the
  same Status as text;
- text beside a dropdown, an unreadable dropdown, a value with HTML in it;
- a dropdown in another row, and in a second row labelled Status, is not given the value;
- Generate changes nothing in the report, reads the Docs API once and never writes; generating
  again gives the same e-mails; an agreed and a reserved TDoc are still refused.

The fake cell paragraph of the test helpers gained `editAsText()` (an addition). The tests that
state the version of `Code.js` name 2.21.1, and the ledgers that list the changed functions name
`docTableToHtml_()`.

| | Result |
|---|---|
| Full suite | 104 files, 6974 checks, 0 failing |
| Mutations of the fix | 10 made; 9 killed by behaviour or code checks; 1 equivalent |

The mutations: the renderer not using `readTdocStatus_()` (two variants), the value given to any
row, to the label cell, to another row, to every row labelled Status, no HTML escaping, the fix
taken out, text beside a dropdown replaced. The equivalent one removes the "holds a dropdown"
guard of the renderer; `readTdocStatus_()` makes the same check itself, so the output is the
same, and only the ledger notices it.

## 4. Smoke test in a real Google document: passed

2026-10-06, on the disposable test report, with the candidate of hotfix commit `0f2c5f3`
installed in its script project (the seven files, plus one temporary helper file). The target
was verified as the disposable project before each push; every push was a normal `clasp push`
with the account named explicitly.

The helper had no rendering and no dropdown logic of its own. In one execution it called the
candidate's `docTableToHtml_()` and `generateTdocDiscussionEmails()` — the function behind the
Generate button — and read the document through the Docs API before and after.

| | Result |
|---|---|
| The TDoc | S4aP260098, a native Document Status dropdown, selected value "available" |
| DocumentApp alone | empty text and an UNSUPPORTED element: the regression, reproduced |
| `docTableToHtml_()` | `<p style="margin:0 0 4px 0;">available</p>` in the Status cell |
| The whole table (11 rows) | identical to the rendering with the Status as plain text |
| The generated `.eml` | contains the table with the Status value, not the one with a blank Status |
| The dropdown afterwards | still there, still "available" |
| The definition and its 14 options | unchanged |
| Document content and properties | unchanged (digests) |
| Document revision | the same before and after: nothing was written |
| Result | **PASS** |

The generated `.eml` and `.zip` were moved to the trash. The project was then restored and read
back: exactly the seven files it had before, identical to them. The helper is not part of the
repository or of any bundle.

## 5. Release: complete

| | |
|---|---|
| Release commit | `03ec9fc92c9227adfe3342c26ef4dc6b568162c7`, parent `66dc4d4` (the closeout of T-2026.10.9) |
| Relation to the tested hotfix | the squash of hotfix commit `0f2c5f3`; the same tree |
| Tag | `template-release/T-2026.10.10`, lightweight, on the release commit |
| Push | master and the tag in one atomic, normal push; the ten earlier release tags unchanged |
| Bundle | built by the release tool from the tagged commit: the seven production files |
| Against T-2026.10.9 | `Code.js` and `Release.js` differ; the other five files are identical |
| Against the smoke candidate | only `Release.js` differs (the commit it names and the time it was built) |
| `Release.js` | T-2026.10.10, template, `Code.js` 2.21.1, the release commit, the tag |
| Master template before | T-2026.10.9, `Code.js` 2.21.0, exactly seven files, identical to that bundle |
| Deployment | one normal `clasp push` of the seven files, without a prompt |
| Read back | exactly seven files, identical to the bundle; manifest scopes and services unchanged |

The master-template document itself was not changed. The 6G and MBS reports, CENTRAL and Legacy
were not touched; no existing report was rebuilt, updated or moved to this release.

## 6. Follow-ups

- Existing reports keep the code they have. A report running T-2026.10.7, .8 or .9 shows the
  empty Status in its e-mails until it is moved to this release; that is a separate step.
- The follow-ups of T-2026.10.9 (`docs/ADHOC_AUTOMATION_AND_REVIEWER_TOKEN.md`, §7) stand.
