# Rollout of T-2026.10.30 (Code.js 2.26.14): plan, not yet approved

T-2026.10.30 is T-2026.10.29 plus (docs/GOES_TO_PLENARY.md, PR #26):

- **The compact Status row:** `Status | [status ▾] Plenary [☐ No / ☑ Yes ▾] A.I.: X.X`, with no
  extra rows.
- **The SA4 plenary agenda item:** automatic, or typed by hand.
- **The four disposition tables,** in a landscape section of their own.
- **The revision relationship** in the TDoc row: `S4-26XXXX (revision of S4-26YYYY)`.

T-2026.10.29 (`af3779f`, Code.js `cc32cbb6…`) is in production on all five SA4#138 reports,
deployed 2026-11-17.

Release files: Code.js, ReportCreator.js and Release.js change. HyperLink.js and appsscript.json
are those of T-2026.10.29.

There is **no automatic migration**. A report under T-2026.10.30 behaves exactly as under
T-2026.10.29 until its rapporteur runs *Report › Set Up Goes to Plenary…*. Before that it has
the three-table summary in portrait, an unchanged Status row, no revision text and no agenda
download. The release can therefore go to all five scripts first, and each report can be
switched on separately.

Every step below needs approval. Nothing here has been done in production. The local listing
agent is not part of this release and is not touched.

## 0. Validation on the disposable MBS copy (before any production step)

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138". It holds the first candidate's
**two-row layout** ("Goes to Plenary" and "Plenary Agenda Item" below Status).

1. Push the candidate bundle (`dist/release-candidate/T-2026.10.30/`) into the copy's bound
   script. The push is guarded: the copy must run the first T-2026.10.30 candidate, and the
   read-back must match `RELEASE-MANIFEST.json`.
2. In the copy, take a named version first (*File › Version history › Name current version*:
   "before compact Status row"). Note the Goes to Plenary selections and any Plenary Agenda Item
   you typed by hand.
3. *ℹ️ About This Report* shows T-2026.10.30 / 2.26.14.
4. **Migration:** *📄 Report › Update Report Now*, then check:
   - the two rows are gone from every TDoc table;
   - the Status row reads `[status] Plenary [☐ No/☑ Yes]`, with the same selections as noted;
   - for ☑ Yes, `A.I.: X.X` follows (8.5 → 15.4, 8.6 → 15.5, 8.7 → 16.3, 8.4 → 14; see the 8.4
     question);
   - a value typed by hand in the old "Plenary Agenda Item" row is kept;
   - the minutes and all other rows are untouched.
5. **Landscape:** the Disposition Summary starts on a new page, in landscape. Check:
   - all four tables are there, in the requested order, with the wider Title column;
   - the report before it is still portrait;
   - there is no blank page around the section break;
   - type a line below the summary and *Update Report Now*: that line is on a portrait page
     again.
6. **Revision relationship:** a TDoc that the TDoc list names as a revision reads
   `S4-26XXXX (revision of S4-26YYYY)`.
   - Both numbers open their TDocs.
   - A TDoc that is not a revision shows its number alone.
   - The Revisions row is as before.
7. **Updates and edits.** Run *Update Report Now* after each step:
   1. Run it twice more. Nothing changes: no new section break, no duplicated text, the summary
      unchanged.
   2. Pick ☑ Yes on an agreed TDoc. Its A.I. appears and it moves to table 1. Change its Status
      to noted: it moves to table 2, and the selection and A.I. stay.
   3. Manual values:
      - type a different value after "A.I.:", which is kept;
      - pick ☐ No: A.I. is hidden, but the value stays in the summary;
      - pick ☑ Yes: the value is back;
      - type `-`: blank;
      - clear the value to "A.I.:": automatic again.
   4. *Update TDoc Revisions*, *Update Disposition Summary* and *Close Report*: no selection
      changes, the tables stay landscape.
   5. *Build Report from Scratch…* (on the copy only): selections and manual values come back.
      The next update writes the four landscape tables.
8. Roll the copy back to the named version and push the earlier candidate back if anything
   differs (`dist/integration-test/rc-T-2026.10.30-<UTC>/before/`).

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
| Updates take longer | One label scan per TDoc table, one property read per update, and the agenda is downloaded once per agenda TDoc. No API write when nothing changed |
| A report is switched on during a session | Do it between sessions (step C). The set-up takes the document lock and refuses while an update runs |
