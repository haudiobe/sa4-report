# Agenda Word conversion (T-2026.10.20, Code.js 2.26.4)

Status: **released as T-2026.10.20 (Code.js 2.26.4).** Candidate commit `0b7fd91`, merged by
PR #1 (`3b62504`). The deployment record is at the end of this document.

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

## Release and deployment of T-2026.10.20 (2026-11-11)

| Step | Result |
|---|---|
| Review and merge | PR #1 reviewed: the change, temporary-file cleanup, security (no new scope, manifest unchanged, no new URL) and release compatibility (the five-file set). Merged as `3b62504` (candidate `0b7fd91`). |
| Release commit | `fc3ef7d9401c43c3b7ec6c14fe5eae569bfb1672`. It adds the version line 2.26.4 (2026-11-11) and its changelog entry; the version-pinned tests say 2.26.4. Without the header, `Code.js` is the `Code.js` of `3b62504`; the ledger gives the `Code.js` of T-2026.10.19 byte for byte. |
| Tests | 115 files, 7,812 checks, 0 failures, on both the PR head and the release commit. |
| Tag and publication | Lightweight tag `template-release/T-2026.10.20` on the release commit. `origin/master` fast-forwarded `3b62504` → `fc3ef7d`, without force, and the tag was pushed. |
| Release bundle | `--release T-2026.10.20 --write`: the five production files. The four source files equal the tag; `Release.js` reads T-2026.10.20 / 2.26.4 / `fc3ef7d`. |
| Adoption bundle | `--adopt T-2026.10.20 --label mbs-138`: dry run, then `--write`. The guardrails passed (not the template, not Legacy, not CENTRAL). The five files are byte-identical to the release bundle. |
| clasp identity | `clasp --user haudiobe.ts@gmail.com show-authorized-user`: the template owner's account. |
| **Master template: before** | Container proven with the Apps Script API (`projects.get`): its parent is the "SA4 Report Template" document; last update 2026-11-01 17:40:07 UTC (the T-2026.10.19 push). A read-only pull gave five files equal to T-2026.10.19 (line ends aside), with `Release.js` reading T-2026.10.19 / 2.26.3 / `277e5a5`. |
| **Master template: push** | 2026-11-11 12:46:39 UTC (13:46:39 CET), from its own scratch folder, without `--force`. The staged files matched the hashes in `RELEASE-MANIFEST.json`. |
| **Master template: read-back** | A fresh pull: exactly five files, all byte-identical to the bundle. `Release.js` reads T-2026.10.20 / 2.26.4 / `fc3ef7d`, and `Code.js` 2.26.4. The project's update time is 2026-11-11 12:46:37 UTC. |
| **SA4#138 MBS report: before** | Container proven: its parent is "MBS SWG Minutes SA4#138"; the project was created 2026-11-11 11:28 UTC by the template copy and not updated since. A read-only pull: `Code.js`, `HyperLink.js`, `ReportCreator.js` and `Release.js` are T-2026.10.19. **Unexpected difference:** `appsscript.json` has the same content as T-2026.10.19 and T-2026.10.20, but with the key `runtimeVersion` moved to the end, probably re-serialized by the copy. |
| **SA4#138 MBS report: push** | Approved by Thomas despite the key order. `clasp push` (without `--force`) detected the manifest difference, asked to overwrite it, and skipped the push ("Skipping push."). As the release tool requires, nothing was forced. |
| **SA4#138 MBS report: state after the skipped push** | Unchanged: update time still 11:28:18 UTC, and all five files byte-identical to the pull taken before. It still ran **T-2026.10.19 / 2.26.3**. |
| **SA4#138 MBS report: manifest overwrite** | Approved by Thomas. Thomas ran the push and confirmed the manifest overwrite; the project's update time is 2026-11-11 12:54:24.783 UTC. |
| **SA4#138 MBS report: read-back** | A fresh pull into a new scratch folder: exactly five files, all byte-identical to the adoption bundle and to the master template's read-back. The four source files equal the tag; `appsscript.json` now has the release's key order. `Release.js` reads T-2026.10.20 / 2.26.4 / `fc3ef7d`, and `Code.js` 2.26.4. |
| **SA4#138 MBS report: configuration and triggers** | A push replaces script files only: it writes no Document Property and creates no trigger. The report document's last modification is 12:54:24.783 UTC, the same instant as the push, by the owner; nothing changed the document after it. Neither the Document Properties nor the triggers can be listed without running code (and the clasp credential has no `script.processes` scope), so they were not read. No report function was run. |

**Deployed:** the master template and the SA4#138 MBS report both run T-2026.10.20 / 2.26.4.
`diagnoseAgendaConversion` is to be run next, from the SA4#138 Apps Script editor. Full Report
Build was not run.

**Not touched:** MBS 86172, 6G (its automatic update stays off), every trigger, every report
configuration, and the repository-root `.clasp.json`. Every clasp command used
`--user haudiobe.ts@gmail.com` and ran from its own scratch folder.

## Second failure, with T-2026.10.20: DocumentApp cannot open the converted agenda (T-2026.10.21, Code.js 2.26.5)

Status: **released as T-2026.10.21 (Code.js 2.26.5).** Candidate commits `9559d17` and `dc2da01`,
merged by PR #2 (`aca8ea3`). The deployment record is at the end of this document.

### Evidence

The live run of `diagnoseAgendaConversion()` in the SA4#138 MBS report on T-2026.10.20 showed:

- S4-261857 downloads.
- `Drive.Files.create` returns `application/vnd.google-apps.document`.
- `DocumentApp.openById(...).getBody()` fails on all four attempts, the last one 7 s after the
  first, with "Unexpected error while getting the method or property openById on object
  DocumentApp."
- The temporary document is trashed.

The same Word file was converted again outside Apps Script, by the template owner's account
through the Drive API, and the result read back. The temporary files were deleted afterwards.

| Check | SA4#138 (S4-261857) | SA4#136 (S4-260868) |
|---|---|---|
| Converted type | native Google Doc, owned, editable | native Google Doc, owned, editable |
| Drive HTML export (rendered by the Docs backend) | 1 table, 101 rows, header `A.I.#`; 228 paragraphs; 7 list items | 1 table, 103 rows, header `A.I.#` |
| Drive plain-text export | 8,204 characters | 7,574 characters |
| Docs API `documents.get` | not testable this way: the API is disabled in clasp's OAuth client project (HTTP 403 `SERVICE_DISABLED`). This is a property of that client, not of the document. | — |

The Word files are structurally alike: the same parts, the same elements and characters, and
valid UTF-8.

### Root cause, as far as the evidence goes

- **The conversion is sound.** The Google Doc that Drive makes from S4-261857 is complete and
  readable by the Docs backend.
- **The failure is DocumentApp's own loading of that document.** It is deterministic: four
  attempts over 7 s, and two builds before. Longer waits will not help.
- **Why DocumentApp refuses this document is not visible from outside Google.** The error is
  Apps Script's generic wrapper for an internal error.
- **Neither the content nor the permissions explain it.** The account owns the document, and the
  manifest has the `documents` and `drive` scopes. There is no evidence that this conversion path
  has ever worked live through DocumentApp in a template report.

### The fix: three readers, one parser

The conversion, the retries and the cleanup are unchanged. The agenda is then read by the first
of three readers that succeeds:

1. **DocumentApp**, as before.
2. **The Docs API** (`Docs.Documents.get`) on the same converted document, before it is trashed.
   The Docs advanced service is already in the manifest and in live use by the status dropdowns.
3. **The Word file itself**: `word/document.xml` and `word/styles.xml`, with no Google Doc.

Each reader gives `parseAgendaStructureWithText_()` the same small Body view. That view holds
only what the parser reads:

- the top-level paragraphs (a numbered paragraph is a LIST_ITEM, as in DocumentApp);
- each paragraph's text and heading;
- the tables, where a cell's text is its paragraphs joined by "\n".

The parser is unchanged. So the agenda items, their order, and the TDocs grouped under them by
agenda-item number are the same whichever reader read them.

**An empty or malformed agenda is not a read.** A reader counts as successful only if
`agendaItemsProblem_()` accepts its items: at least one item, every item with a number, a title
and a level, and at least one number like an agenda item's ("8", "8.1"). Otherwise the reader
has failed and the next one is asked. If none succeeds, the build fails with an error that names
each reader; it never continues with an empty agenda.

**Logging.** Each reader's failure is logged with its original error and stack, and the log names
the reader that read the agenda. The log holds counts only, never agenda text. If all three
readers fail, the error names each one and its reason.

**`diagnoseAgendaReaders()`.** Run it from the editor. It converts the configured agenda once,
runs all three readers, logs for each whether it read the agenda, how many items, and whether
they equal the others', then trashes the temporary document. It writes nothing.

### Equivalence with Google's conversion

The Word-file reader is checked against Google's own conversion of the same file: the Drive
export of the converted Google Doc, read with the parser's rules (fixtures
`tests/fixtures/agenda-S4-261857*` and `agenda-S4-260868*`).

| Agenda | Items from Google's conversion | Items from the Word file | Result |
|---|---|---|---|
| S4-261857 (SA4#138) | 100 | 100 | identical in number, title, level and order |
| S4-260868 (SA4#136) | 102 | 102 | identical in number, title, level and order |

- **TDoc association.** All 388 TDocs of the SA4#138 TDoc list fall under an agenda item read
  from the Word file. The MBS projection, 8.1–8.10, holds 46 of them, exactly as with Google's
  items.
- **One export detail.** The HTML export writes some spaces as `&nbsp;`, while the plain-text
  export of the same document has none. The reference therefore normalises them to spaces. The
  Word file's agenda cells hold no no-break space.

### Not changed

- the parser;
- the Drive conversion and its checks;
- the DocumentApp open and its retries;
- the cleanup;
- the build, the agenda choice and the download;
- the TDoc lists, the update and Close Report.

Nothing else in the report changes: no trigger, no configuration and no draft revision.

### Another diagnostic run

This is recommended after deployment and before any build: run `diagnoseAgendaReaders()` in the
SA4#138 report.

- **Expected:** DocumentApp FAILED; the Docs API either reads the 100 items or fails; the Word
  file reads the 100 items, the same as the Docs API where both read them.
- **Then:** if so, one Full Report Build. Its log should say "read by Docs API" or "read by Word
  file".

## Release and deployment of T-2026.10.21 (2026-11-11)

| Step | Result |
|---|---|
| Review | The PR #2 diff was reviewed against five checks. DocumentApp stays the first reader. The Docs API and Word-file readers feed the unchanged parser, so hierarchy, order and TDoc association are unchanged. The temporary document is trashed in a `finally` on every path, including in `diagnoseAgendaReaders()`. The diagnostic writes no property and nothing in the report. **One gap was found:** a reader whose agenda came back empty counted as successful, which stopped the fallback. Fixed in `dc2da01` (`agendaItemsProblem_()`, tests section K, 7 checks). |
| Merge | PR #2 merged as `aca8ea3` (candidates `9559d17` and `dc2da01`). |
| Release commit | `692c07861aa76c7907cc181687acdd258863f135`. It adds the version line 2.26.5 (2026-11-11) and its changelog entry; the version-pinned tests say 2.26.5. Without the header, `Code.js` is the `Code.js` of `aca8ea3`; the ledger gives the `Code.js` of T-2026.10.20 byte for byte. |
| Tests | 116 files, 7,869 checks, 0 failures, on both the merged PR and the release commit. |
| Tag and publication | Lightweight tag `template-release/T-2026.10.21` on the release commit. `origin/master` fast-forwarded `aca8ea3` → `692c078`, without force, and the tag was pushed. |
| Release bundle | `--release T-2026.10.21`: dry run, then `--write`. The four source files equal the tag; `Release.js` reads T-2026.10.21 / 2.26.5 / `692c078`. The staged files matched the hashes in `RELEASE-MANIFEST.json`. |
| clasp identity | `clasp --user haudiobe.ts@gmail.com show-authorized-user`: the template owner's account. |
| **Master template: before** | Container proven (`projects.get`): its parent is the "SA4 Report Template" document; last update 2026-11-11 12:46:37 UTC (the T-2026.10.20 push). A read-only pull: five files byte-identical to the T-2026.10.20 read-back. No unexpected difference. |
| **Master template: push** | 2026-11-11 19:13:57 UTC, from its own scratch folder, without `--force`. |
| **Master template: read-back** | A fresh pull: exactly five files, all byte-identical to the release bundle. `Release.js` reads T-2026.10.21 / 2.26.5 / `692c078`, and `Code.js` 2.26.5. |
| **SA4#138 MBS report: before** | Container proven: its parent is "MBS SWG Minutes SA4#138"; last update 12:54:24 UTC (the T-2026.10.20 push), with nothing changed since. A read-only pull: five files byte-identical to the T-2026.10.20 read-back, including the manifest (the release's key order). No unexpected difference. |
| **SA4#138 MBS report: adoption** | `--adopt T-2026.10.21 --label mbs-138`: dry run, then `--write`. The guardrails passed. The five files are byte-identical to the release bundle. The live manifest was already byte-identical to the bundle's, so clasp asked no overwrite question. |
| **SA4#138 MBS report: push** | 2026-11-11 19:14:53 UTC, from its own scratch folder, without `--force`. |
| **SA4#138 MBS report: read-back** | A fresh pull: exactly five files, all byte-identical to the adoption bundle and to the master template's read-back. `Release.js` reads T-2026.10.21 / 2.26.5 / `692c078`, and `Code.js` 2.26.5. |
| **SA4#138 MBS report: configuration and triggers** | A push replaces script files only: it writes no Document Property and creates no trigger. The document's last modification is the push instant, 19:14:53 UTC, by the owner. The Document Properties and triggers cannot be listed without running code, so they were not read. No report function was run. |

**Deployed:** the master template and the SA4#138 MBS report both run T-2026.10.21 / 2.26.5.
`diagnoseAgendaReaders()` is to be run next, by hand, from the SA4#138 Apps Script editor. Full
Report Build was not run.

**Not touched:** MBS 86172, 6G (its automatic update stays off), every trigger, every report
configuration, every draft revision, and the repository-root `.clasp.json`. Every clasp command
used `--user haudiobe.ts@gmail.com` and ran from its own scratch folder. The TDoc revision
feature is not part of this release.
