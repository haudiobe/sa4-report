# Agenda Word conversion (candidate after T-2026.10.19)

Status: **candidate, not released, not deployed.** `Code.js` stays 2.26.3 until the release.

## Incident: MBS SA4#138, 2026-11-11

The first Full Build of a new MBS report for SA4#138 (main meeting, Calgary), created from the
template at T-2026.10.19, failed twice in a row in the skeleton phase:

```
Full report build failed: Unexpected error while getting the method or property
openById on object DocumentApp.
✗ Skeleton + TDoc tables: FAILED
```

The failure came about 4.5 s after the start, and the report stayed empty. The agenda is read
before the document is cleared, so the report was not changed.

## Where it fails

```
runFullReportBuild → runFullReportBuildCore_ (phase "skeleton")
  → buildSkeletonWithTdocTables
    → parseAgendaForReport_            AGENDA_TDOC = S4-261857 (the one agenda-typed TDoc of SA4#138)
      → downloadMeetingAgenda_         …/TSGS4_138_Calgary/Docs/S4-261857.zip
        → extractWordFromZip_          "S4-261857 Meeting agenda for SA4#138.docx"
        → parseAgendaWordBlob_
          → convertWordBlobToGoogleDoc_   Drive.Files.create(…, mimeType Google Doc) → new ID
          → DocumentApp.openById(<that new ID>)   ← fails
```

The ID passed to `openById` is the temporary Google Doc that Drive created a moment earlier
from the agenda's Word file. It is not the report, the master template or the preamble source
document. The code did not log that ID and trashed the document in `finally`.

The other `openById` calls in the build can be ruled out:

- **The preamble source.** It is opened only after the document is cleared and titled, and the
  report is still empty.
- **The Google-Doc agenda fallback.** It is used only when `AGENDA_TDOC` is empty, and here it
  is S4-261857.
- **The context document.** The menu path passes no context.

## What was checked

| Question | Finding |
|---|---|
| Is the agenda file unusual? | No. The ZIP holds one valid OOXML `.docx`, the same in structure as SA4#136's S4-260868, down to the `SA4#NNN` in its name. |
| Invalid ID? | No. The ID is the one Drive's own `create` call returned. |
| Permissions? | No. The temporary document is created by, and owned by, the account that runs the build. |
| Authorization or scope? | No. `appsscript.json` has `documents` and `drive`, and the Drive call just before succeeded. |
| Temporary failure? | Not on its own. It happened twice in a row with the same input. |

## What was wrong in the code

1. **The Word file went to Drive with whatever content type `Utilities.unzip()` gave it.**
   Drive converts based on that type. The code never set it and never checked what Drive
   returned.
2. **The conversion used the v2 metadata field `title` against the v3 service.** v3 ignores it.
3. **Nothing checked that the result was a native Google Doc** (`mimeType`).
4. **There was no retry** for a converted document that is not readable the moment Drive
   returns it.
5. **No ID, type or step was logged,** and only the message of the original error survived.
   Its stack and the failing step were lost.
6. **A failure to trash the temporary document was swallowed silently.**

Which of 1, 3 or 4 caused the SA4#138 failure cannot be decided offline: it depends on what
Google's services return. The candidate handles all three and logs enough to tell them apart
on the first live run.

## The change

All of it is in the `AGENDA WORD CONVERSION` section of `Code.js`. Only `parseAgendaWordBlob_`
and `convertWordBlobToGoogleDoc_` change; how the agenda is parsed does not.

1. **The upload.** The blob is sent with the content type of a Word file, chosen by its
   extension: `.docx` → `application/vnd.openxmlformats-officedocument.wordprocessingml.document`,
   `.doc` → `application/msword`. The metadata uses v3 `name`, and Drive is asked for
   `fields: id,name,mimeType`.
2. **The result check.** If Drive returns anything other than
   `application/vnd.google-apps.document`, the result is refused at once with an error that
   names the type, and the file is trashed.
3. **The open.** `openById` is retried with waits of 0 / 1 / 2 / 4 s: four attempts, at most
   7 s of waiting.
4. **The log.** Every step is logged with an `[AGENDA]` prefix:
   - the source name, its size, the content type it came with and the one it was sent as;
   - the ID and type Drive returned;
   - each open attempt, with the original error and its stack when it fails;
   - the trash.

   The error thrown names the step, the document ID and the original message.
5. **Cleanup.** The temporary document is trashed in `finally`, in every case. A trash failure
   is logged and never replaces the real error.
6. **`diagnoseAgendaConversion()`.** You run it from the Apps Script editor. It runs the same
   download, conversion, open and parse for the configured agenda TDoc, then trashes the
   temporary document. It writes no property and does not touch the report.

Tests are in `tests/agenda-word-conversion.test.js`. They cover the upload, success, delayed
availability, permanent failure, a wrong type from Drive, a Drive error, a failed trash, the
whole path through `diagnoseAgendaConversion()`, and the ledger
(`codeWithoutAgendaWordConversion()`), which gives back the `Code.js` of T-2026.10.19 byte for
byte.

## Deployment procedure (after approval)

1. **Merge and release.** Merge the PR. Make the release commit: version 2.26.4, the changelog
   entry, and the version-pinned tests. Then tag `template-release/T-2026.10.20` and build the
   release:

   ```
   node tools/template-release.js --release T-2026.10.20
   node tools/template-release.js --release T-2026.10.20 --write
   ```
2. **Push to the master template** from the bundle, as for T-2026.10.19.
3. **Adopt the release in the SA4#138 MBS report's own bound script.** That report was created
   at T-2026.10.19 and keeps that code until it is pushed. Dry run first, then write:

   ```
   node tools/template-release.js --adopt T-2026.10.20 --target-script <SA4#138 MBS script ID> --label MBS-138
   node tools/template-release.js --adopt T-2026.10.20 --target-script <SA4#138 MBS script ID> --label MBS-138 --write
   ```

   Then push from the adoption bundle and read the result back: five files, and `Release.js`
   showing T-2026.10.20.
4. **Verify without a build.** In the SA4#138 report's Apps Script editor, run
   `diagnoseAgendaConversion` and read the `[AGENDA]` lines in the execution log:
   - `converted, opened and parsed -- N agenda item(s)` means it is fixed. Only then run Full
     Report Build once.
   - `Drive did not convert … has the type …` means the conversion itself is refused. The next
     step is to change the conversion, not to retry.
   - `could not be opened after 4 attempts` together with `Drive.Files.create returned <id>
     (application/vnd.google-apps.document …)` means Drive made a Google Doc that DocumentApp
     cannot open. Open that ID in Drive's trash by hand to see whether it opens in the editor.
5. **Other reports.** The other reports (the 6G report and the 86172 MBS report) get the
   release only by a separate, explicit decision. Their builds use this path only when they
   have an agenda TDoc. No trigger is touched.
