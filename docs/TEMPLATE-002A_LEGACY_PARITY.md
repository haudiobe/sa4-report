# TEMPLATE-002A: Legacy parity gate

2026-10-01. Branch `template-002a/legacy-parity`, based on `sa4-report` master `d7c23bc`.
Local source convergence only: nothing was pushed, deployed or run against Google, and the
Legacy repository was only read.

**Question:** does the single `sa4-report` `Code.js` contain every accepted Legacy production
behaviour that the generic SA4 Report Template should have?

**Answer: yes, after this branch.** Three accepted Legacy fixes were missing and are ported.
Every other difference is either an intentional CENTRAL/template behaviour with a test, or
test mechanics. One of the intentional differences needs Thomas's confirmation (§4, item 3).

---

## 1. Baselines and method

| | |
|---|---|
| Main source | `sa4-report` master `d7c23bc` (`Code.js` 2.15.2) |
| Accepted Legacy | `sa4-report-legacy` master `b95e06e`; deployed, live-accepted source `4fe295a` (identical production files) |
| Legacy's own regression baseline | 888 checks in 10 suites |

1. **The complete list of Legacy production changes.** Legacy's first commit (`248e1c3`) is
   main's 2.12.0 commit (`e774b3e`) plus one change, the BUGFIX-LEGACY-001 readiness guard
   (54 differing lines, verified by diff). After that, 21 Legacy commits change `Code.js`.
   The integration commit `ee3e9c4` is the exact union of its two parents' changes. So the
   checklist is that guard plus 21 commits (§2).
2. **Legacy's regression suites run against main's `Code.js`.** Nine suites are imported as
   `tests/legacy-parity-*.test.js`. The tenth, the exporter suite, was already ported in
   ADDON-009 (`tests/addon009-legacy-exporter-port.test.js`).
3. **A function-by-function comparison** of both files (comments and whitespace ignored), to
   catch anything no test covers (§6).
4. Each observed difference is classified A, B, C or D. Nothing was ported for a D item
   before it was understood. No D item remains.

---

## 2. Matrix by Legacy change

| Legacy change | What it is | In main before this branch | Result |
|---|---|---|---|
| BUGFIX-LEGACY-001 (in baseline) | Ad-hoc build readiness guard | ADDON-007B3: one rule set, `assertMeetingReadyToBuild_(context)`; stricter (family, document folder) | Present, different implementation (B) |
| `9739085` LEGACY-UPGRADE-002 | Reviewer token never rendered or logged; keep/replace/clear; no invented SA4#136 defaults for ad-hoc | ADDON-007B1 (the original) | Present. 41 imported checks pass. |
| `b97b46c` LEGACY-UPGRADE-003 | Report-family inference, derived mailing list | ADDON-007B2 (the original) | Present. 63 imported checks pass. |
| `83de6cb` LEGACY-UPGRADE-004 | Ad-hoc administrative agenda anchoring | ADDON-007A (the original) | Present. 57 imported checks pass. |
| `32c8e2d` … `80ab081` (11 commits) LEGACY-UPGRADE-006 | Discussion e-mail (.eml) exporter | ADDON-009 port | Present, with the two decided differences (§4). 450 checks in the existing port test. |
| **`d1021cf` BUGFIX-LEGACY-002** | **No duplicate TDoc tables** | **Missing** | **Ported** (`a553200`) |
| `e564f70` | Document-relative links in exported e-mails | 2.15.2 | Present. 16 imported checks pass. |
| **`123369b` BUGFIX-LEGACY-003** | **Ad-hoc e-mail collection** | Partly (ADDON-008A2: family-aware table gate, full-ID scanner, short-ID fallback only for `S4-`) | **Remainder ported** (`e437581`) |
| **`cb1208c` LEGACY-0099** | **Removed / Withdrawn / N/A reallocations** | ADDON-008A1 / 008A1b (the originals); 9 of 10 touched functions already identical | **One missing site ported** (`ada9f97`) |
| `ee3e9c4` | Integration of the two fixes above | n/a | Union only, no additional code |

---

## 3. Category A: accepted Legacy fixes that were missing in main (all ported)

After test mechanics and intentional differences were set aside, exactly **26 imported checks
failed**, all for these three reasons.

### A1. Duplicate TDoc tables (BUGFIX-LEGACY-002): 4 checks

- `continuousUpdateCore_()`: the set of existing TDocs was filled only from the document scan
  before the loop. A TDoc number listed under two agenda groups in one download was inserted
  twice. A TDoc is now recorded as existing the moment it is inserted.
- `processWebDownloadedSheet_()` (the menu item "Legacy: Build Initial Report"): it never
  clears the document and never checked for an existing table, so each further run duplicated
  every table. It now scans existing TDoc tables once and skips a TDoc already present.

Generic: no TDoc number or meeting appears in the code. Duplicate tables already in a
document are left alone. The existing safe refusal for ambiguous duplicates in *Apply
Document Reallocations* ("the report contains N tables for this TDoc") is unchanged; that code
is identical in both files and is covered by `addon008a1-document-reallocation`.

### A2. Ad-hoc e-mail collection (BUGFIX-LEGACY-003): 21 checks

| Behaviour | Before | Now |
|---|---|---|
| Stored `EMAIL_START_DATE` | Ignored; always `2026-08-21` | Used when a valid `YYYY-MM-DD`; `2026-08-21` remains the fallback; a malformed value is logged |
| A1 archive list | Always the general `3GPP_TSG_SA_WG4` list | The list the collector resolved, the same one RSS reads |
| Collector Configuration table that only repeats the family default | Masked the meeting's Mailing List | Not treated as an override |
| Explicit table `LIST_NAME` | RSS URL stayed on the family list | RSS URLs follow it |
| Explicit table `RSS_URL_V2` | Archive list stayed on the family list | Archive list follows its `L=` |
| Mailing List saved as `<list>@list.etsi.org` | Rejected by the collector (but accepted by the exporter) | Read as that list; CR/LF and other domains stay invalid |

Already present through ADDON-008A2 and confirmed by the imported checks: family-aware exact
TDoc table gate, full-ID scanner across all SA4 TDoc families, canonicalisation and
deduplication, collision protection, explicit full IDs taking precedence, and the short-ID
fallback only for main `S4-` tables.

### A3. Reallocation suppression in the web-sheet import (LEGACY-0099): 1 check

`processWebDownloadedSheet_()` filed a TDoc reallocated to Removed / Withdrawn / N/A under a
literal "Removed" agenda item. It now uses the shared `interpretReallocationTarget_()` and
skips it, with a log line, as the build and Continuous Update already did.

The two differences the exploratory run showed in this suite were analysed before anything
was changed: one is this missing site (A), the other is a call signature (C, §5).

The live-accepted semantics hold on this `Code.js` (82 of 82 imported checks): Removed /
Withdrawn / N/A suppress the detail table; a later update does not recreate it; a Full Build
preserves valid reallocation state; an invalid or ambiguous application refuses before any
mutation.

---

## 4. Category B: intentional CENTRAL/template differences (kept, each tested)

| # | Difference | Decided by | Test |
|---|---|---|---|
| 1 | **Meeting-neutral outgoing e-mail introduction** ("the upcoming meeting", the meeting's own tag); no prefilled discussion deadline. Legacy: 6G wording, "the October meeting", 2026-10-15 15:00. | Thomas, 2026-10-01 | `template002a-intentional-differences` §1; `legacy-parity-email-doc-links` |
| 2 | **Subject tag derived from the meeting's mailing list**; the subject builder refuses to work without one. Legacy: constant `FS_6G_MED`. | Thomas, 2026-10-01 | `template002a-intentional-differences` §2; `legacy-parity-diagnose-6g-email-collection` |
| 3 | **No "Email Collection Start Date" field in the Meeting Configuration dialog**, and Save does not write `EMAIL_START_DATE`. Legacy has the field and saves it. | CENTRAL ADDON-007B1 (`0510461`), which removed it as a dead field | `template002a-intentional-differences` §3; `addon007b1-config-ui` |
| 4 | **Hourly trigger works; intervals that are not offered are refused before the running trigger is touched.** Legacy: "every hour" calls `everyMinutes(60)` and fails after deleting the trigger (a known Legacy defect, listed as "not included" in its release). | CENTRAL 2.15.1 / 2.15.2 | `template002a-intentional-differences` §4; `continuous-trigger-interval` |
| 5 | One readiness rule set instead of Legacy's `computeMeetingConfigReadiness_()`. | CENTRAL ADDON-007B3 | `addon007b3-readiness` |
| 6 | Two unused Legacy helpers are not ported: `extractTdocFromEmailExportSubject_()` (the collector's scanner does this) and the unwired diagnostic `diagnoseDuplicateTdocTables_()` with its helper. Neither is reachable from a menu or from other code in Legacy. | CENTRAL ADDON-009 port | `addon009-legacy-exporter-port` header; `legacy-parity-diagnose-6g-email-collection` |

**Item 3 needs Thomas's confirmation.** CENTRAL removed the field because nothing read the
saved value. That reason no longer holds: after A2 the collector does read it. As the code
stands, a template report can set the start date only through a Collector Configuration table
row, not through the dialog, so the ported behaviour is reachable for documents that already
carry the property but not from the dialog. Restoring the field is a small change (one input,
one save line, three assertions in `addon007b1-config-ui`). It was not made here because it
would reverse an accepted, tested CENTRAL decision that this task was told to preserve.

Item 4 is the add-on runtime's behaviour. The template runtime offers 15 and 30 minutes again
in TEMPLATE-002B; the "hourly works" part stays.

---

## 5. Category C: test-harness and environment differences

| Adaptation | Where | Why |
|---|---|---|
| Loader path | all nine suites | Legacy's `load-code.js` is kept as `tests/helpers/legacy-load-code.js`; it loads this repository's `Code.js` |
| Legacy git-history blocks skipped (13 checks) | `upgrade-002` (8), `-003` (3), `-004` (2) | They load Legacy commits (`git show 248e1c3:Code.js` and others) to show Legacy's own before/after or stage scope. Those commits exist only in the Legacy repository. |
| Legacy source-shape checks skipped (3 checks) | `upgrade-003` | They assert Legacy's source text (a guard without a context parameter, Legacy's comment, "no central registry"). The guard's behaviour is checked elsewhere in the same suite. |
| Fake DOM creates elements on demand | `upgrade-002`, `-003` | The CENTRAL dialog has more elements than Legacy's list of ids |
| Hidden `familyInfo` input handed to the fake DOM | `upgrade-003` | CENTRAL carries the per-family defaults in that input; Legacy inlined them |
| One call signature | `0099-reallocation-suppression` | `downloadAndGroupTdocs_(cfg, context, snapshot)` here, `(cfg, snapshot)` in Legacy |

Skipped checks print a `skip` line with the reason, so they stay visible in the output.
No assertion on behaviour was relaxed: an imported check either passes as written, or is
replaced by an explicit `[template]` check of the intended behaviour (12 such checks).

---

## 6. Category D: unknown

None remain. The function-by-function comparison after the ports shows:

- Legacy-only functions: the four named in §4 (items 5 and 6).
- Every other difference is one of: the execution-context parameter (ADDON-002/004), the
  reorganised dialog (ADDON-007B1), the readiness rules (ADDON-007B3), agenda and TDoc-list
  discovery (ADDON-008A), the exporter adaptations (ADDON-009, 2.15.0, 2.15.2), the trigger
  dialog (2.15.1/2.15.2), or the add-on menu. All are existing main features with their own tests.

---

## 7. Production `Code.js` changes on this branch

136 lines added, 21 removed against master. Three behavioural commits and one header commit:

| Commit | Change |
|---|---|
| `a553200` | `continuousUpdateCore_()`: `existingTdocs.add(tdocNumber)` after an insertion. `processWebDownloadedSheet_()`: scan existing TDoc tables once, skip a TDoc already present, record each inserted one. |
| `ada9f97` | `processWebDownloadedSheet_()`: `interpretReallocationTarget_()` for a reallocated TDoc; skip and log a removed one. |
| `e437581` | `getCollectorConfig_()`: start date from `resolveCollectorStartDate_(context)`; table-override precedence. New `buildCollectorRssUrl_()`, `resolveCollectorStartDate_(context)`. `normalizeEtsiListName_()` replaces `isValidEtsiListName_()`. `resolveCollectorListName_()` uses it. `buildArchiveIndexUrlsByDaysBack_()` uses the list it is given. |
| `75f641f` | Header only: version 2.16.0 and changelog entry. |

The ported code is Legacy's, line for line, except that it goes through this repository's
execution context: `resolveCollectorStartDate_(context)` reads the report state store, so a
background run reads the adopted document's central state. `EMAIL_START_DATE` is not an
adoption key, so CENTRAL background runs keep the default; the scheduler is untouched.

**One existing main assertion changed.** `addon008a2-email-association` expected
`3GPP_TSG_SA_WG4_MBS@LIST.ETSI.ORG` to be rejected as a Mailing List. It now expects the
accepted Legacy behaviour (read as that list), and checks another address as invalid instead.

**Duplicate functions.** Measured, not changed: the same five names are defined twice in both
files (`copySectionContentWithReplacement_`, `formatDeadline_`, `getMonthName_`,
`calculateTimeRemaining_`, `addTdocTablesOnly`). No ported fix touches any of them.

---

## 8. Test results

Complete suite on this branch: **72 test files, 3,652 checks, 0 failures.**

| Group | Files | Checks | Failures |
|---|---|---|---|
| Existing `sa4-report` tests | 62 | 3,224 | 0 |
| Imported Legacy suites (`legacy-parity-*`) | 9 | 408 (12 of them `[template]` checks) | 0 |
| `template002a-intentional-differences` | 1 | 20 | 0 |

Skipped with a printed reason: 16 Legacy checks (13 need Legacy git history, 3 assert Legacy
source shape). The Legacy exporter suite's 460 checks are represented by the existing port
test (450 checks; the 10 not carried over exercise Legacy-only code, as its header records).

Before the ports, 26 imported checks failed (A1: 4, A2: 21, A3: 1) and the existing 62 files
passed.

---

## 9. Template release gate

**Is this `Code.js` a safe behavioural source for building the first real SA4 Report
Template? Yes**, on this branch (not yet on master):

- all accepted generic Legacy fixes are represented;
- all existing main regressions are green;
- all relevant Legacy parity regressions are green;
- intentional differences are documented and tested;
- no unexplained parity difference remains.

Open before or during TEMPLATE-002B: Thomas's decision on the e-mail start date field (§4,
item 3), and merging this branch to master.
