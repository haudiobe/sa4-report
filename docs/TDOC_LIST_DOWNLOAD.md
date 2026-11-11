# TDoc-list download (T-2026.10.22, Code.js 2.26.6)

Status: **released as T-2026.10.22 (Code.js 2.26.6).** Candidate commit `e9a9ac8`, merged by PR #3
(`8497d01`). The deployment record is at the end of this document.

## Incident: MBS SA4#138, 2026-11-11

With T-2026.10.21, Full Report Build of the SA4#138 MBS report read the 100 agenda items through the Docs API
fallback, then stopped in the skeleton phase:

```
Could not download TDOC list: HTTP 403
```

The configured TDoc list is

```
https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx
```

A diagnostic function added by hand in the report's Apps Script editor (`diagnoseTdocListDownload`, not part of
any release) fetched the same path from both hosts. The `www.3gpp.org/ftp` address answered HTTP 200 with
110,999 bytes starting `504b0304` (the XLSX/ZIP signature):

```
https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx
```

## Root cause, in the code

`downloadAndGroupTdocs_()` (used by the build, the update and Close Report's external check) did this:

```js
const response = UrlFetchApp.fetch(meetingUrl, { muteHttpExceptions: true });
if (response.getResponseCode() >= 400) {
  throw new Error('Could not download TDOC list: HTTP ' + response.getResponseCode());
}
const blob = response.getBlob();   // then DriveApp.createFile + SpreadsheetApp.open
```

1. **One address only.** It fetched `cfg.TDOC_LIST_URL` as configured. `ftp.3gpp.org` refuses Apps Script with
   403, and the code had no other address to try. The same files are served under `https://www.3gpp.org/ftp/`.
   The rest of the report (the agenda ZIP and the documents) already uses that host by default.
2. **No check of the answer.** Any status below 400 was taken as the workbook, including an HTML page sent with
   HTTP 200. Such a page would then fail later, and obscurely, in the xlsx → Sheet conversion.

`readTdocListWorkbook_()` (ad-hoc agenda discovery) already checked the HTTP status and the XLSX signature.
`downloadAndGroupTdocs_()` did not.

## The fix

There is a new section, TDOC LIST DOWNLOAD, and one changed line in `downloadAndGroupTdocs_()`.

- `tdocListDownloadCandidates_(url)` lists the configured address first. Only for an address on `ftp.3gpp.org`
  does it add the same path under `https://www.3gpp.org/ftp/`, the alternative the diagnostic proved. Look-alike
  hosts (`ftp.3gpp.org.example.org`, `ftp.3gpp.org@example.org`) get no alternative.
- `tdocListDownloadAttempt_(url)` makes one request. It accepts the answer only with HTTP 2xx, a non-empty body
  and the XLSX signature (`isXlsxSignature_`, the check `readTdocListWorkbook_` already uses). Otherwise it gives
  the reason: `HTTP 403`, `HTTP 200 with an empty body`, or `HTTP 200 but not an Excel workbook (N bytes
  starting 3c21444f, text/html)`. It never throws, and it reads nothing of the content beyond the first four
  bytes.
- `downloadTdocListXlsx_(url)` tries the addresses in order and returns the first accepted one. It logs each
  failure, and logs when the alternative was used. If every address fails, it throws `Could not download TDOC
  list: …` naming each one. With a single address, the error is the same as before (`… HTTP 500`, or the fetch's
  own error if the fetch threw).
- `diagnoseTdocListDownload()` is the released version of the hand-written diagnostic. It reads the configured
  address, asks every candidate with the same check the build uses, and writes nothing. Production and
  diagnostic can no longer diverge.

**What does not change:**
- the configured address, which stays in the configuration and stays the list's identity (the identity anchor
  records the configured address, not the one that answered);
- the xlsx → Sheet conversion, the column lookup, `assertTdocListIdentity_`, the grouping, and every TDoc
  association;
- the agenda readers, the other list readers (`readTdocListWorkbook_`, the Portal page fallback), the legacy
  menu tools, triggers and configuration.

## Tests

`tests/tdoc-list-download.test.js` has 43 checks. It runs the real build, update and Close Report of
`tests/helpers/template-report.js` and compares the result with the released T-2026.10.21 `Code.js`:

| Case | Result |
|---|---|
| The configured address works | One request. The report is identical to T-2026.10.21's. |
| The configured address answers 403, the alternative works | The build succeeds with the same report. T-2026.10.21 fails here with the production error. The configured address is unchanged and is the identity. |
| HTTP 200 with an HTML page | Refused, and the alternative is asked. T-2026.10.21 converted the HTML page. |
| Both fail (403/404, or both throw) | One error naming each address. Nothing is converted or written. |
| The alternative is not a workbook (HTML, empty, `.xls` bytes, a truncated signature) | Refused. The error names both reasons. |
| Not on `ftp.3gpp.org` (`www.3gpp.org/ftp`, Portal) | No alternative. The error text is as before. |
| `diagnoseTdocListDownload()` | Every address with the build's check. Nothing written. |
| Close Report, the update | They download the same way. |
| The `Code.js` diff | One section plus one line in `downloadAndGroupTdocs_`. With the section taken out, `Code.js` is T-2026.10.21's byte for byte. From the blob on, `downloadAndGroupTdocs_` is unchanged. |

The fakes of two legacy-parity tests now return a blob with the XLSX signature, as a real download does (the
build now checks it). Two release-scope tests compare against the code without this later section, as for every
earlier layer.

## Release plan (after review and approval)

1. Merge the PR. Release T-2026.10.22 / `Code.js` 2.26.6: the version line, the changelog entry, the ledger
   header pair, and the version-pinned tests.
2. Master template: the guarded release procedure, with a read-back.
3. SA4#138 MBS report: the guarded adoption, a dry run, a push, and a read-back.
   - **Note:** the report's `Code.js` was edited by hand in the editor after T-2026.10.21 was deployed (the
     ad-hoc `diagnoseTdocListDownload` was added at the end). The adoption push replaces it. The release has its
     own `diagnoseTdocListDownload()`, so the function stays available under the same name.
4. In SA4#138, run `diagnoseTdocListDownload()` by hand. Expected: `configured …ftp.3gpp.org…: refused -- HTTP
   403; alternative …www.3gpp.org/ftp…: accepted -- HTTP 200, 110999 bytes starting 504b0304`. Then run one Full
   Report Build. Its log should show `[TDOC LIST] downloaded from https://www.3gpp.org/ftp/…`.

MBS 86172 and 6G receive the release only by a separate decision.

## Release and deployment of T-2026.10.22 (2026-11-11)

| Step | Result |
|---|---|
| Re-check | PR #3 at head `e9a9ac8`, based on the then-current master: 8 files, as reviewed. Tests: 117 files, 7,912 checks, 0 failures. |
| Merge | PR #3 merged as `8497d01`. |
| Release commit | `ae1d1ca985111e702347ae17617687ef0a79d7c7`. It adds the version line 2.26.6 (2026-11-11) and its changelog entry; the version-pinned tests say 2.26.6. Without the header, `Code.js` is the `Code.js` of `8497d01`; the ledger gives the `Code.js` of T-2026.10.21 byte for byte. Tests: 117 files, 7,912 checks, 0 failures. |
| Tag and publication | Lightweight tag `template-release/T-2026.10.22` on the release commit. `origin/master` fast-forwarded `8497d01` → `ae1d1ca`, without force, and the tag was pushed. |
| Release bundle | `--release T-2026.10.22`: dry run, then `--write`. The four source files equal the tag; `Release.js` reads T-2026.10.22 / 2.26.6 / `ae1d1ca`. Only `Code.js` and `Release.js` differ from T-2026.10.21; the manifest, `HyperLink.js` and `ReportCreator.js` are unchanged. |
| clasp identity | `clasp --user haudiobe.ts@gmail.com show-authorized-user`: the template owner's account. |
| **Master template: before** | Container proven: its parent is the "SA4 Report Template" document; last update 19:13:57 UTC (the T-2026.10.21 push). A read-only pull: five files byte-identical to the T-2026.10.21 read-back. |
| **Master template: push** | 2026-11-11 20:25:31 UTC, from its own scratch folder, without `--force`. The staged files matched the SHA-256 hashes in `RELEASE-MANIFEST.json`. |
| **Master template: read-back** | A fresh pull: exactly five files, each matching its SHA-256 in `RELEASE-MANIFEST.json` and byte-identical to the bundle. `Release.js` reads T-2026.10.22 / 2.26.6 / `ae1d1ca` / `template-release/T-2026.10.22`, and `Code.js` 2.26.6. |
| **SA4#138 MBS report: before** | Container proven: its parent is "MBS SWG Minutes SA4#138"; last update 19:54:18 UTC. A read-only pull: `appsscript.json`, `HyperLink.js`, `ReportCreator.js` and `Release.js` are byte-identical to T-2026.10.21. `Code.js` is the `Code.js` of T-2026.10.21 byte for byte, followed by 47 appended lines with one function, `diagnoseTdocListDownload` (added by hand in the editor). That is the only deviation. |
| **SA4#138 MBS report: backup** | The five pulled files, with their SHA-256 sums and the hand-written function on its own, are kept locally (outside the repository; `dist/` is not committed). The live `Code.js` had SHA-256 `b6ec4ad0…`. |
| **SA4#138 MBS report: adoption** | `--adopt T-2026.10.22 --label mbs-138`: dry run, then `--write`. The guardrails passed. The five files are byte-identical to the release bundle; the live manifest was already byte-identical, so there was no overwrite question. |
| **SA4#138 MBS report: push** | 2026-11-11 20:26:52 UTC, from its own scratch folder, without `--force`. The hand-written function is replaced by the released `diagnoseTdocListDownload()`, under the same name. |
| **SA4#138 MBS report: read-back** | A fresh pull: exactly five files, each matching its SHA-256 in `RELEASE-MANIFEST.json`, byte-identical to the adoption bundle and to the master template's read-back. `Release.js` reads T-2026.10.22 / 2.26.6 / `ae1d1ca`, and `Code.js` 2.26.6. |
| **SA4#138 MBS report: configuration and triggers** | A push replaces script files only: it writes no Document Property and creates no trigger. The document's last modification is the push instant, by the owner. Properties and triggers cannot be listed without running code, so they were not read. No report function was run. |

**Deployed:** the master template and the SA4#138 MBS report both run T-2026.10.22 / 2.26.6.
`diagnoseTdocListDownload()` is to be run next, by hand, from the SA4#138 Apps Script editor. Full Report
Build was not run.

**Not touched:** MBS 86172, 6G (its automatic update stays off), every trigger, every report configuration,
every draft revision, and the repository-root `.clasp.json`. Every clasp command used
`--user haudiobe.ts@gmail.com` and ran from its own scratch folder. The TDoc revision feature is not part of
this release.
