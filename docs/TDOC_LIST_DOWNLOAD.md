# TDoc-list download (candidate after T-2026.10.21)

Status: **candidate, not released, not deployed.** `Code.js` stays 2.26.5 until the release. The proposed
release is T-2026.10.22 (`Code.js` 2.26.6).

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
