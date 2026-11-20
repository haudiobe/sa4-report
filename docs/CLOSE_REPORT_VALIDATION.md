# Close Report validation (T-2026.10.33, Code.js 2.27.1)

Built on T-2026.10.32 (`144b6f8`, Code.js 2.27.0).

**Release files:** Code.js changes, and so does Release.js, which the build generates (it is not
kept in Git). HyperLink.js, ReportCreator.js and each report's appsscript.json stay as they are.

The code is in the section "CLOSE REPORT VALIDATION" of `Code.js`, after the Undisposed Documents
section. Three released functions change: `closeReportCore_` (the classification line, and its log line
reads the summary's count), and
`formatCloseReportResult_` and `closeReportResultHtml_`, which now hand over to the section. The
ledger is in `tests/helpers/before-structure-fix.js` (`CLOSE_VALIDATION_CHANGES`). Without this
layer, Code.js is that of T-2026.10.32 byte for byte.

## What changes

Before this release, any finding kept the report from closing, including:
- a reserved revision not yet uploaded;
- a revision 3GU could not confirm;
- a TDoc whose Status is "Plenary".

Close Report now sorts what it finds into three kinds, and only the first decides:

| Kind | What | Closing |
|---|---|---|
| **Blocking issues** | a TDoc not yet disposed (Available, Reserved, Parked, Postponed, Other, a status the classification does not know, tables that disagree), except Status "Plenary"; a TDoc the report has as revised whose revision is not reserved in 3GU; a check that could not be made: the meeting's TDoc list not read, or read without a "Revised to" column, a Status not readable, no Registration of Documents, text typed into the Disposition Summary | "Report NOT ready to close" |
| **Warnings** | a reserved revision not yet uploaded to 3GU; a revision whose state 3GU cannot confirm | the report can be closed |
| **Information** | every TDoc routed to the plenary: Status "Plenary", or "☑ Yes" in the Plenary dropdown of a report with Goes to Plenary | never blocks |

A TDoc can appear twice: one with "☑ Yes" and no final disposition blocks for its status, and is
also listed under Information, marked "also not yet disposed".

**What a revision warning says.** A revision 3GU cannot confirm may be on the Plenary agenda
instead. The dialog says so, and says that Close Report did not check the Plenary agenda. Close
Report does not read the Plenary agenda, and never says a revision was verified there.

The note is left out where the cause is the TDoc list itself: not read, no "Revised to" column,
or its links unreadable.
- **List not read, or no "Revised to" column:** the revisions were not checked at all. The
  Warnings part says how many, and the list problem is under Blocking issues.
- **Links unreadable:** the revision is reserved and only its upload is unknown. It stays a
  warning, with that reason.

**The dialog.**
- **Counts:** the same counts as before, plus a line "Blocking issues: N. Warnings: N.
  Information: N." when there is a finding.
- **Not yet disposed against blocking:** the count of TDocs not yet disposed is the Disposition
  Summary's. When some of them only have the Status "Plenary", the line says how many of them
  block, for example "4 not yet disposed in the Disposition Summary: 3 blocking, 1 with Status
  "Plenary" (routed to the plenary: information, not blocking)." Otherwise it reads as before
  ("4 not yet disposed."), since every one of them blocks. A TDoc with Plenary "☑ Yes" and no final
  disposition counts as blocking: its routing is not subtracted. The result carries
  `undisposedTotal`, `undisposedPlenaryStatus` and, in `unresolved`, the blocking ones.
- **The three parts:** each list under its kind, in full, with the same tables as before: the
  TDoc IDs, the revision relationship, the 3GU state and the recommended action.
- **Title:** "Report NOT ready to close" only when a blocking issue remains.
- **With warnings only:** the closing note says the report can be closed, that the warnings need
  no correction in the report, and that they are followed up in 3GU.

## What does not change

- **The Disposition Summary is written exactly as before.** Its "Documents Not Yet Disposed" still
  lists a TDoc whose Status is "Plenary". The dialog's counts of the three classes are the summary's.
- **Nothing in the report is changed** but the summary: no Status, disposition, Plenary choice or
  other content.
- **The rest of Close Report is as before:** the lock, the landscape step of Goes to Plenary and
  the modal dialog are unchanged, and `closeReport()` itself is unchanged. The log line keeps its
  text: its count of TDocs not yet disposed is the summary's.
- **The findings are the same**, and so is the external check: the classification and the 3GU
  check are untouched. Only the decision on top of them is new.
- **No new read:** in a report with Goes to Plenary, the Plenary choices come from the read the
  Disposition Summary is written from.

## Tests

`tests/disposition-summary-close-report.test.js`, part V:
- each scenario alone: a TDoc not yet disposed blocks; a revision not uploaded, a revision not
  verified and Status "Plenary" do not;
- a mixed report, with each TDoc in its own part;
- the wording about the Plenary agenda;
- closing with warnings, twice, with no change to the report;
- the same report with T-2026.10.32's code: the same summary, document, statuses, counts and
  findings;
- a check that could not be made still blocks;
- "☑ Yes" in a Goes to Plenary report;
- the ledger.

Four earlier checks of the same file changed with the rule: the TDoc with Status "Plenary" is now
under Information, and a revision not verified has the Plenary agenda note.
