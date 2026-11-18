# Rollout of T-2026.10.30 (Code.js 2.26.14): plan, not yet approved

T-2026.10.30 is T-2026.10.29 plus (docs/GOES_TO_PLENARY.md, PR #26):

- **The compact Status row:** `Status | [status ▾] Plenary [☐ No / ☑ Yes ▾] A.I.: X.X`, with no
  extra rows.
- **The SA4 plenary agenda item:** automatic, or typed by hand.
- **The four disposition tables,** in a landscape section of their own.
- **The revision relationship** in the TDoc row: `S4-26XXXX (revision of S4-26YYYY)` -- in **every** report, set up or not.
- **A faster update** (the summary, the plenary data and the status sync read the document once through the Docs API), and *Advanced and Repair › Check Report Integrity* (read-only).

T-2026.10.29 (`af3779f`, Code.js `cc32cbb6…`) is in production on all five SA4#138 reports,
deployed 2026-11-17.

Release files: Code.js, ReportCreator.js and Release.js change. HyperLink.js and appsscript.json
are those of T-2026.10.29.

There is **no automatic migration** of the Plenary workflow. Until its rapporteur runs *Report ›
Set Up Goes to Plenary…*, a report keeps:
- the three-table summary in portrait;
- its Status rows;
- its workflow (no agenda download).

What every report gets at once:
- the revision relationship in its TDoc rows;
- the faster status sync and formatting;
- the read-only integrity check.

The release can therefore go to all five scripts first, and each report can be switched on
separately.

Every step below needs approval. Nothing here has been done in production. The local listing
agent is not part of this release and is not touched.

## 0. Validation on the disposable MBS copy (before any production step)

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138". It holds the compact Status row
(migrated by the second candidate), and possibly a partial update: the update of 2026-11-18
stopped after about 361 s.

**Done by Claude:** the third candidate bundle (`dist/release-candidate/T-2026.10.30/`) was pushed
into the copy's bound script. The push was guarded: the copy had to run the second candidate
(Code.js `ec7e3d17…`). The read-back matches `RELEASE-MANIFEST.json`. No function was run in the
document.

**By you, in the copy (menu SA4 Report):**

1. *File › Version history › Name current version*: "before third candidate". This is the
   backup to roll back to.
2. *ℹ️ About This Report* shows T-2026.10.30 / 2.26.14.
3. *🛠 Advanced and Repair › Check Report Integrity*. It changes nothing. Note what it lists:
   - "Status dropdowns missing" (expected: S4-262302, perhaps others, left by the interrupted
     run);
   - TDocs with two tables (expected: none);
   - summary headings (expected: 1);
   - section breaks;
   - text typed into the summary (expected: none).
4. *Update Report Now*.
   - Then open *Extensions › Apps Script › Executions* and note the duration of that run. It
     must end without "Exceeded maximum execution time".
   - The [PERF] lines in its log show the time of each part.
5. *Check Report Integrity* again: "Status dropdowns missing: none". S4-262302 shows
   `[reserved ▾] Plenary [☐ No ▾]`.
6. *Update Report Now* a second time with nothing changed:
   - its duration is well below the first;
   - the version history shows no new change by the script.
7. **The Status rows:** every TDoc reads `[status ▾] Plenary [☐ No/☑ Yes ▾]`, including reserved
   ones, with no extra rows. Pick ☑ Yes on one agreed TDoc and change the Status of another. Run
   *Update Report Now*: the first shows `A.I.: X.X` and moves to table 1; the second moves table.
8. **The four tables**, on landscape pages after a new page break:
   - tables 1 and 2 (Goes to Plenary) have **six** columns;
   - tables 3 and 4 (Not Going) have **five**, without SA4 Plenary Agenda Item, and a wider
     Title;
   - TDoc numbers are linked;
   - there is no blank page around the section break, and the report before it is portrait.
9. **Revision relationships:**
   - a revision reads `S4-26XXXX (revision of S4-26YYYY)`, and both numbers open their TDocs;
   - the Revisions row and the minutes are as before.
10. **Manual values:** an A.I. typed earlier is still there, and minutes typed earlier are
    unchanged.
11. If anything differs: restore the named version. Claude can push the second candidate back
    from `dist/integration-test/rc-T-2026.10.30-<UTC>/before/`.

Please send the integrity results of steps 3 and 5 and the two durations of steps 4 and 6.

## A. Master template

1. Push the bundle into the master template's script (as for T-2026.10.29).
2. **Template document, separate approval:** remove the option "Plenary" from the master
   template's "Document Status" dropdown. Only reports built later are affected; existing
   reports lose the option at their set-up.

## B. Scripts, one report at a time (MBS, Video, 6G Media, RTC, Audio)

For each report:

1. Verify identity: the script is bound to the inventory document, the document name is the
   production report's, and the script runs T-2026.10.29 (Code.js `cc32cbb6…`).
2. Take a fresh `clasp -u <user> pull` backup. It must equal that report's T-2026.10.29
   read-back.
3. Push the runtime files of the bundle byte for byte (never `--force`).
4. Pull back and compare with `RELEASE-MANIFEST.json`. Release.js must say T-2026.10.30 /
   2.26.14.
5. UI smoke test: *About This Report*, and the new menu item *Set Up Goes to Plenary…* is
   present. **Do not run it.**

Stop at the first mismatch.

## C. Switching on Goes to Plenary, per report (the rapporteur, after B)

Best between sessions, not while minutes are being typed:

1. *File › Version history › Name current version* ("before Goes to Plenary").
2. *Report › Set Up Goes to Plenary…* → Yes.
3. *Update Report Now* → check the compact Status rows, the four landscape tables, the plenary agenda items and the revision relationships.

Per-report rollback: restore the named version, and delete the document property
`GOES_TO_PLENARY` (or push T-2026.10.29 again).

## Risks

| Risk | Mitigation |
|---|---|
| Removing "Plenary" from the Status dropdown resets other selections | `optionId`s are kept, and the request names a replacement only for "Plenary"; this is verified in step 0.3.3 before production |
| The TDoc list has no "Related WIs"/"Release" | Rules 1 and 4 then do not apply. Title tags and SWG items still map; blank cells are completed by hand |
| "Maintenance → 8.4" is meant literally | The candidate writes 14 (the plenary item). Changing it is a one-line change before release |
| The compact Status row looks crowded | One line, no extra rows; the look is checked in step 0 |
| The landscape section break leaves a blank page, or is lost when the summary is rewritten | Checked in step 0.5; each update re-checks the section from a fresh read and repairs it |
| "Is revision of" is not in the TDoc list | The predecessor's "Revised to" alone still gives the relationship; nothing is inferred |
| A text status in a compact row is rewritten by DocumentApp | Never: it is replaced in front of the code's mark only (tested) |
| A rapporteur types into the summary | The update refuses and changes nothing; the log names the text |
| An update runs out of the six minutes (it did on 2026-11-18) | The summary, the plenary data and the status sync read the document once through the API; an unchanged update makes about 7x fewer DocumentApp calls than T-2026.10.29 (fake document); past 240/285 s the formatting and the summary wait for the next update, and the dropdowns are always completed. Real durations are measured in step 0 |
| The collection of mail/RSS and revisions (96 s on 2026-11-18) | Unchanged in this release; it stays the largest part of an update and depends on the network |
| A report is switched on during a session | Do it between sessions (step C). The set-up takes the document lock and refuses while an update runs |
