# Rollout of T-2026.10.30 (Code.js 2.26.14): plan, not yet approved

T-2026.10.30 is T-2026.10.29 plus **Goes to Plenary, the four disposition tables and the SA4
plenary agenda item** (docs/GOES_TO_PLENARY.md, PR "Goes to Plenary").

T-2026.10.29 (`af3779f`, Code.js `cc32cbb6…`) is in production on all five SA4#138 reports,
deployed 2026-11-17.

Release files: Code.js, ReportCreator.js and Release.js change; HyperLink.js and appsscript.json
are those of T-2026.10.29.

There is **no automatic migration**. A report under T-2026.10.30 behaves exactly as under
T-2026.10.29 (three-table summary, no new rows, no agenda download) until its rapporteur runs
*Report › Set Up Goes to Plenary…*. The release can therefore go to all five scripts first, and
each report can be switched on separately.

Every step below needs approval. Nothing here has been done in production. The local listing
agent is not part of this release and is not touched.

## 0. Validation on the disposable MBS copy (before any production step)

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138".

1. Push the candidate bundle (`dist/release-candidate/T-2026.10.30/`) into the copy's bound
   script. The push is guarded: the copy must run T-2026.10.29 (Code.js `cc32cbb6…`) and the
   read-back must match `RELEASE-MANIFEST.json`.
2. In the copy, take a named version first (*File › Version history › Name current version*:
   "before Goes to Plenary").
3. In the copy, menu **SA4 Report**:
   1. *ℹ️ About This Report* shows T-2026.10.30 / 2.26.14.
   2. *📄 Report › Update Report Now*. Nothing new appears: no rows, and the summary is as it was.
   3. Set the Status of one TDoc to **Plenary** (if the copy's dropdown still has it). Then
      *📄 Report › Set Up Goes to Plenary…*. The dialog names the TDoc. Answer Yes, then check:
      - every TDoc table has "Goes to Plenary" (☐ No) and "Plenary Agenda Item" right below
        Status;
      - that TDoc shows ☑ Yes and Status "available";
      - the Status dropdown has no "Plenary" option any more;
      - no other Status changed (open-point 1 of docs/GOES_TO_PLENARY.md).
   4. *Update Report Now*:
      - the Disposition Summary at the end has the four tables in the order asked for;
      - "Plenary Agenda Item" is filled in: 8.5 → 15.4, 8.6 → 15.5, 8.7 → 16.3, 8.4 → 14 (see
        the "8.4" question), 8.3 → blank.
      - Note whether the TDoc list gave "Related WIs" / "Release" (open point 2): a CR to Rel-19
        under 8.5 shows 14 only with them.
   5. Run *Update Report Now* again. The summary is unchanged, and the document's version
      history shows no new edit by the script.
   6. Status changes, then *Update Report Now* after each:
      - pick ☑ Yes on an agreed TDoc → it moves to table 1;
      - change its Status to noted → table 2, and the selection stays ☑ Yes;
      - change its Status back to agreed → table 1.
   7. Manual values:
      - type `16.2` into one Plenary Agenda Item → it is kept by the update and shown in the
        table;
      - type `-` into another → it shows blank;
      - empty a manual cell → it becomes automatic again.
   8. Type a line below the summary → it is kept. Type a line between two of its tables → the
      update logs "not updated -- text was typed into it" and changes nothing; remove the line.
   9. *Update TDoc Revisions* and *Close Report*: no selection changes; Close Report writes the
      same four tables.
4. Roll the copy back to the named version and push T-2026.10.29 back if anything differs
   (`dist/integration-test/rc-T-2026.10.30-<UTC>/before/`).

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
3. *Update Report Now* → check the four tables and the plenary agenda items.

Per-report rollback: restore the named version, and delete the document property
`GOES_TO_PLENARY` (or push T-2026.10.29 again).

## Risks

| Risk | Mitigation |
|---|---|
| Removing "Plenary" from the Status dropdown resets other selections | `optionId`s are kept, and the request names a replacement only for "Plenary"; this is verified in step 0.3.3 before production |
| The TDoc list has no "Related WIs"/"Release" | Rules 1 and 4 then do not apply. Title tags and SWG items still map; blank cells are completed by hand |
| "Maintenance → 8.4" is meant literally | The candidate writes 14 (the plenary item). Changing it is a one-line change before release |
| Two new rows per TDoc table make the report longer | Two short rows below Status. The look is checked in step 0 |
| A rapporteur types into the summary | The update refuses and changes nothing; the log names the text |
| Updates take longer | One label scan per TDoc table, one property read per update, and the agenda is downloaded once per agenda TDoc. No API write when nothing changed |
| A report is switched on during a session | Do it between sessions (step C). The set-up takes the document lock and refuses while an update runs |
