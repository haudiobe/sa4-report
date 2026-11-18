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

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138". It runs the fourth candidate. Its last
update took about 95.5 s. Its log said "summary: failed", and Check Report Integrity listed 28
TDocs with text Status values (S4-262281 to S4-262293, S4-262302 and others).

**Done by Claude:**

- The fifth candidate bundle (`dist/release-candidate/T-2026.10.30/`) was pushed into the copy's
  bound script.
- The push was guarded: the copy had to run the fourth candidate (Code.js `23c15120…`).
- The read-back matches `RELEASE-MANIFEST.json`.
- No function was run in the document.

**By you, in the copy (menu SA4 Report):**

1. **Back up.** *File › Version history › Name current version*: "before fifth candidate".
2. **Before.** Run *🛠 Advanced and Repair › Check Report Integrity*. Expected:
   - "Status dropdowns missing": the 28;
   - "Marked as a report with Status dropdowns: no (the next update marks it …)".

   If it says "Convert Status Fields to Dropdowns… makes them dropdowns" instead, the report has no
   Status dropdown to go by. Run that command (it checks before it changes anything), then
   continue with step 3.
3. **First update.** Run *Update Report Now*. In *Extensions › Apps Script › Executions*, note the
   duration and these log lines:
   - `Status dropdowns: 28 inserted …`;
   - "… it is marked now";
   - `Goes to Plenary (update): … summary: replaced; landscape: unchanged`;
   - `Found N existing TDOCs`;
   - a line "Disposition Summary: DocumentApp counts … before it than the Docs API", if one is
     there.

   If the dialog says "updated only in part", note the stages and reasons it names.
4. **After.** Run *Check Report Integrity* again. Expected:
   - "Status dropdowns missing: none";
   - "Marked … : yes";
   - "TDoc tables" equals N from step 3.
5. **Look at the report.**
   - S4-262281 and S4-262302 show `[status ▾] Plenary [☐ No ▾]`, with the same status value as
     before.
   - Any Plenary Yes and A.I. values are as they were.
   - The minutes are unchanged.
   - The Disposition Summary has four tables on landscape pages: six columns for the two
     "Goes to Plenary" tables, five for the two "Not Going" tables.
   - TDoc numbers are linked.
6. **Second update, nothing changed.** Run *Update Report Now* again. Expected:
   - "summary: unchanged", and no "Status dropdowns: N inserted" line;
   - no "updated only in part";
   - no new script edit in the version history.
7. **One real change.** Set one Plenary to ☑ Yes and run *Update Report Now*. The TDoc moves to a
   "Goes to Plenary" table, and the log says "summary: replaced".
8. **If anything is wrong:** restore the named version. Claude can push the fourth candidate back
   from `dist/integration-test/rc-T-2026.10.30-<UTC>/before/`.

Please send:

- the log lines of step 3 (above all, the `summary:` line);
- the integrity results of steps 2 and 4;
- the durations of steps 3 and 6.

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
| Status values left as text in a report with Goes to Plenary that was never marked (the copy had 28) | The update restores them when the report already has Status dropdowns, and marks it; with none, Check Report Integrity names Convert Status Fields to Dropdowns…. Only the status part is replaced; Plenary, A.I., minutes untouched. Checked in step 0 |
| The Disposition Summary is not rewritten ("summary: failed" in the copy) | Found in DocumentApp and checked element by element; the new summary is written before the old one is taken out; on failure the old one stays, the reason is logged, and the update says "updated only in part". Checked in step 0 |
| An update runs out of the six minutes (the first and the second update did, on 2026-11-18) | One Docs API read per update; revisions, discussions, existing TDocs and the structure check compare against it and write only cells that change (an unchanged update: 620 DocumentApp calls against 95,154 for the third candidate, fake document). Collectors have a 150/215 s budget, formatting 240 s, summary 285 s; a stage left over is reported as "updated only in part" and completed by the next update. Real durations are measured in step 0 |
| The collection of mail/RSS and revisions (96 s on 2026-11-18) | Unchanged in this release; it stays the largest part of an update and depends on the network |
| A report is switched on during a session | Do it between sessions (step C). The set-up takes the document lock and refuses while an update runs |
