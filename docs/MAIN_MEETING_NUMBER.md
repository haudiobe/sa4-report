# Main-meeting number and the TDoc list (T-2026.10.24, Code.js 2.26.8)

Status: **part of release T-2026.10.24 (`Code.js` 2.26.8), together with the meeting time zones.** Made on `fix/main-meeting-number-from-folder` (`db3030a`), taken into `release/T-2026.10.24` (decision of 2026-11-13, option A). GitHub issue #9. Not yet merged into master, tagged or deployed.

## What happened

In the integration test of T-2026.10.24 (a disposable copy of Audio SWG SA4#138, meeting 60780), Configure Meeting was saved after Resolve. Update Report Now then failed:

```
Could not download TDOC list: HTTP 403 (https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23136.xlsx);
HTTP 403 (https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23136.xlsx)
```

## Cause

The cause is a code defect, present since the first commit and not introduced by 2.26.8.

| Source | Finding |
|---|---|
| Portal metadata | Correct. Resolve takes the document folder from the Portal's `MtgDocURL`: `…/TSGS4_138_Calgary/Docs/`. |
| Saved configuration | **Wrong.** `MEETING_FOLDER = TSGS4_136_Montreal` and `MEETING_NUMBER = 136` were saved next to `FTP_BASE = …/TSGS4_138_Calgary/Docs/`. |
| Cache | Not involved. The TDoc list URL is built anew on each run (`getReportConfig_()`), and nothing caches it. |
| Code | **Defect.** Resolve fills `FTP_BASE` but never the Meeting Folder and Meeting Number fields. For a **main** meeting, `persistConfigurationSettings_()` wrote a blank field as the historical default, `TSGS4_136_Montreal` and `136`. `getReportConfig_()` then built `FTP_BASE + 'TDoc_List_Meeting_SA4%23' + MEETING_NUMBER + '.xlsx'`. The dialog warned "Enter the meeting number, or paste the TDoc list URL" but did not block the save. The template creator already avoided this (`deriveMainMeetingFolderFromFtpBase_()` in ReportCreator.js); Configure Meeting did not. |

The file that exists is `TDoc_List_Meeting_SA4#138.xlsx`. The directory listing of `https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/` (2026-11-13) has exactly one TDoc list, and no `SA4#136` file.

**HTTP 403 means the file does not exist.** 3GPP answers a missing file with 403, not 404. A made-up file name in the same folder gets 403. `…%23138.xlsx` gets HTTP 200/206 with the XLSX content type, with a browser, a `Google-Apps-Script` or a curl user agent alike. Apps Script is not treated differently.

`ftp.3gpp.org` redirects a browser-like client to `www.3gpp.org/ftp/…`, but with `%23` decoded to `#`. The rest of the name then becomes a fragment, so the redirect fails even for the existing list. The `www.3gpp.org/ftp` twin that the download tries second (since T-2026.10.22) covers this.

## Was the report changed by the failed update?

**No.** `continuousUpdateCore_()` runs the readiness check, then downloads and checks the list, and only then makes its first change (`repairReportStructure_()`). A failed download throws before that. Writes only a temporary Drive file, which is trashed.

The message was misleading, though. Only an identity refusal was marked "not updated", so Update Report Now said *"The report may have been updated in part"*.

## The fix (section MAIN MEETING NUMBER of Code.js)

1. **Configure Meeting, main meeting.** A blank Meeting Folder or Number is taken from the document folder:
   - `…/WG4_CODEC/TSGS4_138_Calgary/Docs/` → `TSGS4_138_Calgary`, `138`
   - `…/TSGS4_137-e/Docs/` → `TSGS4_137-e`, `137-e`

   The SA4#136 defaults are written only when no document folder is known, as before. A folder or number that contradicts the document folder is refused before anything is written, unless a TDoc list URL is typed.
2. **`getReportConfig_()`.** A number that is not saved is the document folder's number, before falling back to 136.
3. **Update and build.** A saved main-meeting number that contradicts the saved document folder, with no saved TDoc list URL, is refused before any download and before the report is touched. The message names the number, the folder and the list name, and says what to do.
4. **Update.** A failed TDoc-list download is reported as *"The report was not updated: …"*. A failure after the download still says "may have been updated in part".

Unchanged:
- ad-hoc meetings;
- a typed TDoc list URL;
- a main meeting without a document folder;
- ReportCreator.js, HyperLink.js and appsscript.json;
- the version line (2.26.8); the changelog of 2.26.8 has a second entry, "Fixed: …".

## Tests

`tests/main-meeting-number.test.js` (sections A–H) covers:
- the reproduction with the candidate's code;
- the corrected save;
- refusal of contradictions;
- unchanged legacy and ad-hoc paths;
- the preflight for update, build and trigger;
- download failure versus a later failure;
- the ledger: without the section and its changes, `Code.js` is the candidate's byte for byte.

Changed expectations:
- `tests/template003-menu.test.js`: a failed list download now reads "not updated";
- `tests/meeting-time-zones.test.js` section J: compares the layer below this one.

## A report already saved with the wrong number

The copy of issue #9 has `MEETING_FOLDER = TSGS4_136_Montreal` and `MEETING_NUMBER = 136` saved. With the fix:
- an update is refused with the reason;
- re-saving Configure Meeting as shown (136) is refused;
- clearing both fields, or typing `TSGS4_138_Calgary` / `138`, saves the correct values.

Without the fix, typing `138` into Meeting Number (and `TSGS4_138_Calgary` into Meeting Folder) and saving has the same effect.
