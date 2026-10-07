# Shared Minutes

**Status: RELEASED AND LIVE. Template release T-2026.10.11 (`Code.js` 2.22.0), commit
`3f737ff0f5e2e80d8d18ed3ba9065cf14a944020`, tag `template-release/T-2026.10.11`. Deployed to
the master template and adopted by the 6G and MBS live reports on 2026-10-07 (§ Deployment).**

Template reports only. CENTRAL and Legacy have no such section and no such code.

## What it does

Configure Meeting has a section **Shared Minutes** with one button, **Create Shared Minutes**.
Pressing it:

1. creates one separate Google Doc for the report;
2. lets **anyone with the link edit** it (Drive permission `type: anyone`, `role: writer`,
   `allowFileDiscovery: false`);
3. reads that permission back from Drive;
4. writes `Link to the shared minutes: <document name>` into the report.

Afterwards the section shows the document name, "Anyone with the link can edit ✓", **Open
Shared Minutes**, **Verify** and **Forget**.

The document is created only by that button. Nothing is created or shared by opening a report,
by opening or saving Configure Meeting, by an update (by hand or by a trigger), by a build, or
by the discussion e-mails, and none of them asks Drive about the shared minutes. Build Report
from Scratch only writes the link again, from a valid record that is stored already.

## The document

| | |
|---|---|
| Name | The title a build gives the report, with its first "Minutes" replaced by "Shared Minutes": `Video SWG Shared Minutes SA4#137-e`, `6G Media Shared Minutes – SA4-e (AH) on FS_6G_MED`. Taken from the configuration, not from the file name. |
| Folder | The folder the report is in, when the user can add files to it. Otherwise the user's My Drive, and the message says so. |
| Content | Its title, and one line with the meeting name and date. Nothing else. |
| Marker | Drive `appProperties` set in the request that creates it: `sa4Purpose = shared-minutes`, `sa4ReportId`, `sa4MeetingId`. |

## What is stored

One Document Property of the report, `SHARED_DOCUMENT`:

```json
{ "v": 1, "fileId": "…", "url": "https://docs.google.com/document/d/…/edit", "name": "…",
  "reportDocumentId": "…", "meetingId": "…", "createdAt": "…", "permissionVerifiedAt": "…" }
```

- It is used only in the document it names and for the meeting it names.
- `permissionVerifiedAt` is `null` until the permission was read back. The link is written into
  the report only when it is set.
- `SHARED_DOCUMENT_PENDING` (`{ "startedAt": <ms> }`) exists while a creation is under way.
- Neither key is a Configure Meeting key or an adoption key: nothing reaches CENTRAL.

## One document, also after a failure

The button is idempotent and goes on from where an earlier press stopped.

| State when the button is pressed | What it does |
|---|---|
| Nothing recorded, Drive has no marked document | Creates, records, shares, verifies, links. |
| Nothing recorded, Drive has exactly one marked document | Uses it; creates nothing. |
| Nothing recorded, Drive has more than one | **Nothing.** Lists them; the user trashes all but one. |
| Nothing recorded, none found, a creation was started under two minutes ago | Nothing; asks to press again later. Drive's search can lag behind a creation. |
| Recorded, not shared | Shares the same document, verifies, links. The button reads "Finish sharing". |
| Recorded and shared | Re-reads the permission; changes nothing. |
| Recorded, the document is in the trash or cannot be opened | Nothing. The user restores it, or uses Forget. |
| Recorded for another meeting than the report is configured for now | Nothing until the user uses Forget. |

If Drive refuses or drops the permission, the result is "SHARING FAILED": the document is kept
and recorded, nothing is linked, and the setup is not reported as done.

**Verify** reads the document and its permissions and records the answer. It never changes a
permission. **Forget** removes the record and, when the report line links to that document, the
line; it asks Drive nothing and does not change the document.

## The line in the report

- **Main meeting:** the `Link to the shared minutes:` line the build copies from the meeting
  report template is filled in where it stands. Without it, a new line goes directly below the
  "Opening of the session" heading.
- **Ad-hoc report with sessions:** directly below the last line of the Online information
  block. The block is replaced without it and the line stays.
- **Ad-hoc report without sessions:** directly below the "Opening of the session" heading.

The label is text and the document name is the link. A line that already says this is not
written again.

## Safety

- **`reportDocumentId` guard.** A `SHARED_DOCUMENT` record is used only in the document it
  names. A record that arrives in another document (a copied value) is read as "not created":
  it is not shown, not verified, and not linked by a build.
- **`meetingId` guard.** A record is used only for the meeting it names. When the report is
  configured for another meeting, nothing is created or linked until the user uses Forget.
- **No adoption by a copy.** A document is taken over only when its marker names this report
  and this meeting. A record that points at another report's document is refused by Create
  and by Verify: the document is neither shared again nor linked.
- **A stale link in a copied report is presentation only.** A copy carries the link line of
  the report it was copied from as text. Configure Meeting flags it; it is replaced by Create
  or Verify, and a Build Report from Scratch does not write it again. Nothing reads it.
- **Sharing is believed only when read back.** After the permission is created, the
  permissions of the document are listed (`Permissions.list`), and the setup counts as done
  only when an entry with `type: anyone` and `role: writer` is there.
- **`appProperties` and other script projects.** In the live smoke the marker written by one
  report's script project was not visible to another bound script project, neither by reading
  the file nor by search. That is an observation. Correctness does not depend on it: the two
  guards above decide.

## Live validation

Run on real Google Drive against commit `73a7cfc` — the tree of which is the tree of the
release commit — on disposable documents with their own bound script projects. Everything the
smoke made was moved to the trash, and the smoke project was restored to its previous seven
files (hash-verified).

| Check | Result |
|---|---|
| Create: exactly one Google Doc, with the expected name, folder, content and `appProperties` | Pass |
| Permission, read back independently with `Permissions.list`: `anyone`, `writer`, `allowFileDiscovery: false` | Pass |
| Editing in a private browser window, signed out (by hand) | Pass |
| Exactly one link line in the report, linked to the document | Pass |
| Second Create: no second document, same file, no second permission, one link line | Pass |
| Verify: succeeds, `permissionVerifiedAt` set, no permission changed | Pass |
| Build Report from Scratch: link restored exactly once, record unchanged, no Shared Minutes request to Drive | Pass |
| Copy safety in a second document with its own script project: `reportDocumentId` guard, `meetingId` guard, a forged record refused, nothing adopted, the first document and its permission unchanged | Pass |
| Create New SA4 Report's copy path (`makeCopy` into the folder of the document): the copy keeps its bound script project with the release candidate's code | Pass |
| Cleanup | Pass |

**Covered by the automated tests only, not by the live smoke:**

- the main-meeting placement (filling the template's existing `Link to the shared minutes:` line);
- the placement in an ad-hoc report with sessions;
- sharing that Drive refuses or drops ("SHARING FAILED");
- a real Drive copy of a report with shared minutes: the second document of the smoke was a
  separate document given the first one's text and record, not a Drive copy. That a Drive copy
  starts without Document Properties was verified live earlier (TEMPLATE-001).

One copy made during the preparation of the smoke arrived without its bound script project.
It could not be reproduced: the copy test above, with the release candidate, kept the script.
The cause is not established.

## Deployment

T-2026.10.11 on 2026-10-07. Each project was read back after the push: exactly seven files,
each identical to the release bundle; `appsscript.json` unchanged from T-2026.10.10, so no new
scope and no new authorization.

| Project | Before | After |
|---|---|---|
| Master template | T-2026.10.10 (`Code.js` 2.21.1) | T-2026.10.11 (`Code.js` 2.22.0) |
| 6G live report | T-2026.10.10 | T-2026.10.11 |
| MBS live report | T-2026.10.10 | T-2026.10.11 |

Against T-2026.10.10 the bundle differs in `Code.js`, `ReportCreator.js` and `Release.js`; the
other four files are the same.

No document migration is needed: a report has no shared minutes until somebody presses the
button. The adoption pushed script files only; no report body, Document Property, trigger or
Drive permission was touched, and no Shared Minutes document was created by it.

## Code

- `template/ReportCreator.js`: the section "Shared Minutes" at the end (30 functions, 8
  constants). The Drive calls are in `sharedMinutesDrive_()`; only `createSharedMinutesWith_()`
  creates or shares, and only the button `createSharedMinutes()` calls it.
- `Code.js`: two call sites, in the template runtime only: `configureMeetingSettings()` (the
  section and its script) and the end of `buildSkeletonWithTdocTables()` (the link).
- No new OAuth scope. `appsscript.json` is unchanged.

## Tests

`node tests/shared-minutes.test.js` — the rules, every state of the table above and the press
after it, Verify and Forget, the three placements, a real build/update/trigger/e-mail/rebuild,
a copied report, the dialog, and what did not change. The complete suite at the release
commit: 105 files, 7158 checks, no failure.
