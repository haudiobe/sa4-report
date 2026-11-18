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

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138". It runs the third candidate. Its second
update on 2026-11-18 ran past ten minutes and may have stopped part-way.

**Done by Claude:**

- The fourth candidate bundle (`dist/release-candidate/T-2026.10.30/`) was pushed into the copy's
  bound script.
- The push was guarded: the copy had to run the third candidate (Code.js `0b9eba40…`).
- The read-back matches `RELEASE-MANIFEST.json`.
- No function was run in the document.

**By you, in the copy (menu SA4 Report):**

1. **Back up.** *File › Version history › Name current version*: "before fourth candidate".
2. **Check the version.** *ℹ️ About This Report* shows T-2026.10.30 / 2.26.14.
3. **First update.** Run *Update Report Now*. This run may still have repairs to make after the
   interrupted run. In *Extensions › Apps Script › Executions*, note:
   - (a) its **duration**;
   - (b) that it ends without "Exceeded maximum execution time";
   - (c) from its log:
     - the line `[PERF] report snapshot … ms`;
     - the lines "Revisions: N cell(s) compared …, W written" and "E-mail discussions: …";
     - any `[PERF] left to the next update: …`;
     - any TDoc table "not in the TDoc list".

   If the dialog says "updated only in part", that is correct behaviour, not a failure. Run step 4.
4. **Second update, nothing changed.** Run *Update Report Now* again and note the same three
   items. Expected:
   - well under six minutes;
   - "Revisions: … 0 written";
   - "E-mail discussions: … 0 written";
   - no stage left over;
   - *Version history* shows no new edit by the script.
5. **Run the integrity check.** *🛠 Advanced and Repair › Check Report Integrity* (read-only).
   Expected:
   - "Status dropdowns missing: none";
   - one summary heading;
   - the TDoc table(s) "not in the TDoc list" that explain 75 against 76. Nothing has been
     deleted; you decide.
6. **Spot-check one real change.** Change one Status and set one Plenary to ☑ Yes, then run
   *Update Report Now*. Expected:
   - both move to the right disposition table;
   - the Revisions rows, the minutes and the manual A.I. values are unchanged.
7. **The look,** unchanged from the third candidate:
   - compact Status rows;
   - four tables on landscape pages: six columns for tables 1–2, five for tables 3–4;
   - `S4-26XXXX (revision of S4-26YYYY)` with both numbers linked.
8. **If anything is wrong:** restore the named version. Claude can push the third candidate back
   from `dist/integration-test/rc-T-2026.10.30-<UTC>/before/`.

Please send:

- the durations of steps 3 and 4;
- their "Revisions" and "E-mail discussions" lines and any "left to the next update" lines;
- the integrity result of step 5.

The timeout counts as fixed only when step 4 finishes well within six minutes.

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
| An update runs out of the six minutes (the first and the second update did, on 2026-11-18) | One Docs API read per update; revisions, discussions, existing TDocs and the structure check compare against it and write only cells that change (an unchanged update: 620 DocumentApp calls against 95,154 for the third candidate, fake document). Collectors have a 150/215 s budget, formatting 240 s, summary 285 s; a stage left over is reported as "updated only in part" and completed by the next update. Real durations are measured in step 0 |
| The collection of mail/RSS and revisions (96 s on 2026-11-18) | Unchanged in this release; it stays the largest part of an update and depends on the network |
| A report is switched on during a session | Do it between sessions (step C). The set-up takes the document lock and refuses while an update runs |
