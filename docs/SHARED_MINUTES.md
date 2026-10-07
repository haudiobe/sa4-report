# Shared Minutes

**Status: implemented and tested locally (`Code.js` 2.22.0). Not released, not deployed, not
smoke-tested against Google Drive.** Template reports only; CENTRAL and Legacy have no such
section and no such code.

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

Nothing is created or shared by opening a report, by opening or saving Configure Meeting, by an
update (by hand or by a trigger), by a build, or by the discussion e-mails. None of them asks
Drive about the shared minutes. A build writes the link again from what is stored.

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

## A copied report

A copy starts without Document Properties, and a `SHARED_DOCUMENT` value that names another
document is ignored. The marker of the original's document is private to the original's script
project, so the copy cannot find or adopt it: Create in a copy makes the copy's own document.
The link text the copy inherited in its body is shown as stale in Configure Meeting; it is
replaced by Create or Verify, and a Build Report from Scratch does not write it again.

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
a copied report, the dialog, and what did not change.

## Not verified yet

Everything Drive does is tested against a fake. One live smoke is still required for:

- the HTML upload being converted to a Google Doc with name, folder and `appProperties`;
- `anyone` / `writer` being accepted for the account, and what Drive answers when it is not;
- `files.list` finding the document by its `appProperties` (and how soon after its creation);
- a copy of the report not seeing the original's `appProperties`.
