# TEMPLATE-003: menu and UX review

2026-10-01. Branch `template-003/menu-review`, based on master `02c46fc`
(`Code.js` 2.17.3, live template release T-2026.10.3).

**Status: review only. No production code, test, version or release was changed. Nothing was
deployed or run.**

The subject is the menu a user sees in a report created from the template. The CENTRAL add-on
menu and the Legacy bound script are out of scope and must not change.

---

## 1. Current menu (report created from T-2026.10.3)

Built by `onOpen()` in `Code.js` (lines 553 to 665) with two hooks in
`template/ReportCreator.js` (`addTemplateReportMenuHead_`, `addTemplateReportMenuTail_`).

```
⚠️Scripts⚠️
├─ 🚀 Finish Report Setup                      (only until the report has been set up)
├─ ──────────
├─ 📝 INITIAL SETUP
│   ├─ ⚙️ Configure Meeting Settings
│   ├─ 🧪 Test All Connections
│   └─ 📋 Create Configuration Tables
├─ 🚀 REPORT OPERATIONS
│   ├─ ▶️ Run Full Report Build
│   ├─ ──────────
│   ├─ 0️⃣ Configure Meeting
│   ├─ 1️⃣+2️⃣ Build Skeleton + TDOC Tables
│   ├─ 3️⃣ Collect E-mail Discussion
│   ├─ 4️⃣ Collect Revisions
│   ├─ 5️⃣ Add Abstracts
│   ├─ ──────────
│   ├─ 🔄 Continuous Update (New TDOCs + Status)
│   ├─ ⏰ Manage Auto-Update Trigger
│   ├─ ──────────
│   ├─ 🔄 Update Report (During Meeting)
│   └─ 📊 Analyze Report Status
├─ 📋 DOCUMENT MANAGEMENT
│   ├─ ➕ Add Document Reallocation
│   ├─ 📊 View All Reallocations
│   ├─ 🗑️ Clear All Reallocations
│   ├─ 🔄 Apply Document Reallocations
│   ├─ 🔀 Re-arrange Revision Tables
│   ├─ ──────────
│   ├─ 📄 Parse Agenda Document
│   ├─ 🏗️ Auto-Create Report Structure
│   ├─ ──────────
│   ├─ 🧹 Clean Up Wrong Email Discussions
│   └─ 🔧 Remove Duplicate Email Entries
├─ 🔧 TOOLS & DIAGNOSTICS
│   ├─ 🧪 Test TDOC List URL
│   ├─ 🧪 Test Reviewer API
│   ├─ 🧪 Test Email Feeds
│   ├─ 🧪 Test Revisions Folder
│   ├─ ✅ Validate Configuration
│   └─ 🗑️ Clear All Caches
├─ 🎨 FORMATTING & FIXES
│   ├─ 🎨 Format Document
│   ├─ 🔗 Fix Links (portal→FTP)
│   └─ 📏 Fix Column Widths
├─ 📧 EMAIL EXPORT
│   └─ Prepare TDoc Discussion E-mails
├─ ──────────
├─ 🗄️ LEGACY (old workflow)
│   ├─ 📝 Legacy: Build Initial Report
│   └─ ⚠️ Legacy: Update All
└─ ℹ️ About This Report
```

36 items in 7 submenus (35 once the report is set up), calling 35 different functions:
*Configure Meeting* appears twice.

The master template document has its own two-item menu (*Create New SA4 Report*, *Template
Release Info*) and is not affected by this review, except for the menu name (§4.3).

---

## 2. Function map

Columns: **Doc** = changes the report document. **Net** = fetches from outside Google (3GPP,
ETSI list server, Reviewer). **Dialogs** = what the user sees. **Keep** = still needed in the
menu of a template report.

Purpose codes: SETUP first-time setup or build, MEET normal meeting operation, DISC discussion
workflow, AUTO automatic update, CONF configuration, REPAIR maintenance or repair, DIAG
diagnostic, OLD historical.

### 2.1 Top level and INITIAL SETUP

| Label | Function | What it does | Doc | Net | Dialogs | Purpose | Overlap | Keep |
|---|---|---|---|---|---|---|---|---|
| 🚀 Finish Report Setup | `finishReportSetup` | Reads the setup information the creator left in the file description, stores it as this report's meeting configuration, shows whether the report is ready to build. | no (properties and file description only) | no | 1 result | SETUP | *Configure Meeting* and *Run Full Report Build* do the same step silently on first use | yes, as today (disappears after setup) |
| ⚙️ Configure Meeting Settings | `configureMeetingSettings` | The meeting dialog: Portal meeting ID lookup, report family, sources, mailing list, e-mail start date, sender, Reviewer token. Saves to properties. | no | yes (lookup) | 1 form | CONF | identical to *0️⃣ Configure Meeting* | yes, once |
| 🧪 Test All Connections | `testAllConnections` | Fetches the TDoc list, the Reviewer API, the list RSS feed and the revisions folder; one result dialog. | no | yes | 1 result | DIAG | covers the four single tests | yes (Advanced) |
| 📋 Create Configuration Tables | `createConfigurationTables` | Writes a "Configuration" heading and key/value table at the top of the report (or refreshes it) and creates the Document Reallocations table. | **yes** | no | 1 result | OLD / DIAG | none | Advanced at most; see §3 |
| ℹ️ About This Report | `showTemplateInfo` | Origin, script release, meeting ID, Continuous Update interval, Script ID. | no | no | 1 info | DIAG / info | none | yes |

### 2.2 REPORT OPERATIONS

| Label | Function | What it does | Doc | Net | Dialogs | Purpose | Overlap | Keep |
|---|---|---|---|---|---|---|---|---|
| ▶️ Run Full Report Build | `runFullReportBuild` | **Clears the whole document**, then builds structure and TDoc tables, collects e-mail discussions and revisions, adds abstracts (if a token exists), formats. The Document Reallocations table is read first and restored. | **yes, replaces everything** | yes | 1 confirmation, nothing afterwards | SETUP | steps 1 to 5 below | yes |
| 0️⃣ Configure Meeting | `configureMeetingSettings` | Same function as above. | no | yes | 1 form | CONF | duplicate | no (duplicate) |
| 1️⃣+2️⃣ Build Skeleton + TDOC Tables | `buildSkeletonWithTdocTables` | **Clears the whole document** and builds structure and TDoc tables, then formats. No e-mail, revisions or abstracts. | **yes, replaces everything** | yes | **no confirmation**; 1 result | SETUP (partial) | first phase of Full Build | Advanced only, with a confirmation |
| 3️⃣ Collect E-mail Discussion | `collectEmailDiscussionOnly` | Reads the mailing list (RSS and archive) and writes the e-mail discussion cells. | yes | yes | 1 result | MEET / DISC | part of Full Build, Continuous Update, Update Report | yes |
| 4️⃣ Collect Revisions | `collectRevisionsOnly` | Reads the drafts/revisions folder and updates the revision entries. | yes | yes | 1 result | MEET | same | yes |
| 5️⃣ Add Abstracts | `addAbstractsOnly` | Asks the Reviewer API for every TDoc table without an abstract. Always runs, independent of the automatic-update switch. Without a token nothing is added. | yes | yes | 1 result | MEET | part of Full Build; optional part of Continuous Update | yes |
| 🔄 Continuous Update (New TDOCs + Status) | `continuousUpdate` | The complete incremental update, and the function the timer runs: downloads the TDoc list, **adds new TDocs, updates statuses**, updates the summary table, places revisions, optionally abstracts, then e-mail discussions and revisions, formats if tables changed. Holds the document lock. | yes | yes | **none at all**: no success message, and errors are only logged | MEET / AUTO | superset of *Update Report (During Meeting)* | yes; this is the real "update the report" |
| ⏰ Manage Auto-Update Trigger | `manageTriggers` | One dialog: active or not, interval (15 / 30 / 60 minutes), "fetch abstracts during each update" switch, Start or Stop. | no (trigger and properties) | no | 1 form | AUTO | *About This Report* also shows the interval | yes |
| 🔄 Update Report (During Meeting) | `updateReportIncremental` | E-mail discussions and revisions, then formatting. **Does not add new TDocs or update statuses.** Holds the document lock. | yes | yes | "Success", also when a collection step failed (errors are swallowed and logged) | MEET | = items 3 + 4 + Format; subset of Continuous Update | no (misleading name, covered) |
| 📊 Analyze Report Status | `analyzeReportStatus` | Counts TDocs by status and lists missing minutes and dispositions. Read-only. | no | no | 1 result | MEET | none | yes |

### 2.3 DOCUMENT MANAGEMENT

| Label | Function | What it does | Doc | Net | Dialogs | Purpose | Overlap | Keep |
|---|---|---|---|---|---|---|---|---|
| ➕ Add Document Reallocation | `addDocumentReallocation` | Form: TDoc, original and new agenda item, reason. Creates the reallocation table if missing, adds or updates one row. Does not move anything. | yes (table row) | no | 1 form | MEET / CONF | none | yes |
| 📊 View All Reallocations | `viewAllReallocations` | Lists the saved reallocations. | no | no | 1 list | MEET | none | yes |
| 🗑️ Clear All Reallocations | `clearAllReallocations` | Empties the reallocation table. | yes | no | confirmation + result | MEET | none | yes |
| 🔄 Apply Document Reallocations | `applyDocumentReallocations` | Moves TDoc tables to their new agenda item; **deletes the tables of TDocs reallocated to "removed" / "withdrawn"**. | yes | no | confirmation + result | MEET | builds and updates also honour the table | yes |
| 🔀 Re-arrange Revision Tables | `rearrangeRevisionTables` | Downloads the TDoc list, moves each revision under the document it revises, fills "Revised to". | yes | yes | confirmation + result | REPAIR | done by every Continuous Update | Advanced |
| 📄 Parse Agenda Document | `parseAgendaDocument` | Form; downloads the agenda TDoc (or reads a Google Doc) and stores the parsed agenda in a property. | no | yes | 1 form | OLD | the build parses the agenda itself | no |
| 🏗️ Auto-Create Report Structure | `autoCreateReportStructure` | Appends headings and placeholders from the stored parsed agenda. | yes (appends) | no | confirmation + result | OLD | superseded by the build; on a built report it adds a second set of headings | no |
| 🧹 Clean Up Wrong Email Discussions | `cleanUpWrongEmailDiscussions` | Removes e-mail entries that contain a time-zone pattern ("1600 CEST"), i.e. old false matches. | yes | no | confirmation + result | REPAIR | none | Advanced |
| 🔧 Remove Duplicate Email Entries | `removeDuplicateEmailEntries` | Removes duplicate e-mail entries (same author, date and time). | yes | no | confirmation + result | REPAIR | none | Advanced |

### 2.4 TOOLS & DIAGNOSTICS, FORMATTING & FIXES, EMAIL EXPORT

| Label | Function | What it does | Doc | Net | Dialogs | Purpose | Overlap | Keep |
|---|---|---|---|---|---|---|---|---|
| 🧪 Test TDOC List URL | `testTdocListUrl` | Fetches the TDoc list. | no | yes | 1 | DIAG | in *Test All Connections* | no |
| 🧪 Test Reviewer API | `testReviewerApi` | One Reviewer request with the stored token. | no | yes | 1 | DIAG | in *Test All Connections* | no |
| 🧪 Test Email Feeds | `testEmailFeeds` | Fetches the RSS feed of the **general SA4 list, hard-coded**, not the list this report reads. | no | yes | 1 | DIAG | in *Test All Connections* (same hard-coded list) | no |
| 🧪 Test Revisions Folder | `testRevisionsFolder` | Fetches the revisions folder. | no | yes | 1 | DIAG | in *Test All Connections* | no |
| ✅ Validate Configuration | `validateConfiguration` | Checks four settings (report type, TDoc list URL, token, revisions URL). | no | no | 1 | DIAG | Configure Meeting already shows readiness | no |
| 🗑️ Clear All Caches | `clearAllCaches` | Deletes the stored e-mail, revision and archive caches of this report. | no (properties) | no | confirmation + result | REPAIR | none | Advanced |
| 🎨 Format Document | `removeRowHeightAndSpacing` | Formats every table, removes empty paragraphs. | yes | no | none | REPAIR | last step of every build and update | Advanced |
| 🔗 Fix Links (portal→FTP) | `rewritePortalLinksInDoc_` | Rewrites Portal links of `S4-nnnnnn` documents to the main-meeting FTP folder. | yes | no | none | OLD | none | no; see §3 |
| 📏 Fix Column Widths | `fixColumnWidths` | Sets the TDoc table column widths. | yes | no | 1 result | REPAIR | part of *Format Document* | no |
| Prepare TDoc Discussion E-mails | `prepareTdocDiscussionEmails` | Dialog: choose TDocs, introduction and deadline; writes unsent `.eml` drafts (and a ZIP) to Drive. Sends nothing, does not change the report. | no | no | 1 form | DISC | none | yes |

### 2.5 LEGACY (old workflow)

| Label | Function | What it does | Doc | Net | Dialogs | Keep |
|---|---|---|---|---|---|---|
| 📝 Legacy: Build Initial Report | `buildInitialReport` | Confirmation, then the **old import path**: downloads the TDoc list into a temporary Drive spreadsheet and creates or merges TDoc tables from it (`downloadAndProcessFromWeb` → `processWebDownloadedSheet_`), then e-mail and revisions, then formatting. It does not build the agenda structure and does not clear the document. | yes | yes | confirmation; a pop-up in the middle ("TDOC list downloaded…"); "Success" at the end | no |
| ⚠️ Legacy: Update All | `updateAll` | No confirmation. Looks in the user's Drive for the spreadsheets of the old Colab workflow (`Template-Report-6G.xlsx`, hard-coded, and `Template-Report-<family>-Individual.xlsx`) and copies their tables in; then e-mail and revisions; then two optional hooks that do not exist in this code; then formatting. No lock. | yes | yes | none | no |

**Are these distinct old code or wrappers?** Both.

- The **import part is distinct old code** that nothing else in the template menu uses: the
  spreadsheet-driven table import (`processWebDownloadedSheet_`, `copyTableToDoc2`,
  `copyAllTablesToDoc_`, `copyTableToDoc_`).
- The **rest is the current code**: `collectorUpdate_()` and `removeRowHeightAndSpacing()` are
  the same functions the current build and updates use.
- In a report created from the template the Colab spreadsheets do not exist, so *Legacy: Update
  All* reduces to "e-mail + revisions + format": the same work as *Update Report (During
  Meeting)*, without its lock and without any message.
- *Legacy: Build Initial Report* still works on its own path, but it has a blocking pop-up in
  the middle and a "Success" pop-up after the document was changed. That is the pattern that
  timed out (T-2026.10.0) and failed (T-2026.10.2) in Full Build. If no TDoc list URL is
  configured it falls back to a hard-coded SA4#136 URL.

### 2.6 Useful, but not clearly exposed today

| Need | Today |
|---|---|
| "Update the report now" with a visible outcome | Only *Continuous Update (New TDOCs + Status)*, which shows nothing, not even an error |
| "Are automatic updates running, and how often?" | Inside the *Manage Auto-Update Trigger* dialog, or at the bottom of *About This Report* |
| Change the interval while updates are running | Not possible in one step: the dialog shows only *Stop* while active (the code behind it can replace a running trigger) |
| Abstracts during automatic updates on/off | A checkbox inside the trigger dialog |
| Reviewer API token | A field at the bottom of Configure Meeting |
| Result of the last Full Build | Only in the Apps Script execution log (`[FULLBUILD]` lines) |

Public functions that exist but are in no menu: `addReportSkeleton` and `addTdocTablesOnly`
(both just call the skeleton build), `updateAllFromWeb` (old import + collect + format),
`getTriggerStatus` (returns active / interval). None is needed in the menu.

---

## 3. Problems

**Names that hide what happens**

1. *Continuous Update (New TDOCs + Status)* is the real "update my report", but reads like a
   background setting. *Update Report (During Meeting)* reads like the real one, but does less:
   no new TDocs, no statuses.
2. *Run Full Report Build* does not say that it replaces the whole document, and neither does
   its confirmation. Minutes typed into the report are lost on a second build.
3. *Manage Auto-Update Trigger*, "Start Trigger", "Stop Trigger": Apps Script vocabulary.
4. *Collect* / *Add* for things the user thinks of as "update".
5. *Create Configuration Tables* sounds like a required setup step. It is not; it writes a
   technical table into the report.
6. The menu itself is called ⚠️Scripts⚠️.

**Duplicates and overlaps**

7. *Configure Meeting* twice.
8. Four ways to update: *Continuous Update*, *Update Report (During Meeting)*, *Legacy: Update
   All*, and steps 3 + 4. They differ in scope, locking and feedback (§2).
9. Two ways to build: *Run Full Report Build* and *Legacy: Build Initial Report*.
10. Four single connection tests next to *Test All Connections*; *Fix Column Widths* inside
    *Format Document*.

**Obsolete in a template report**

11. The whole LEGACY submenu (§2.5).
12. *Parse Agenda Document* and *Auto-Create Report Structure*: the manual two-step skeleton
    from before the build existed.
13. *INITIAL SETUP* as a group: a template report is configured by the creator.

**Dangerous or surprising**

14. *1️⃣+2️⃣ Build Skeleton + TDOC Tables* clears the document **without any confirmation**.
15. *Update Report (During Meeting)* reports "Success" even if e-mail or revision collection
    failed, because the shared collector logs its errors instead of raising them.
16. *Continuous Update* from the menu gives no sign of success or failure.
17. *Fix Links (portal→FTP)* calls a private function (name ending in `_`). The repository's
    own RESOLVER-HOTFIX note says menus can only call public functions, so this item is
    expected to fail with "Script function not found". Not verified live. It also only knows
    main-meeting document numbers.
18. *Test Email Feeds* and *Test All Connections* test the general SA4 list, not the list the
    report reads. For a 6G report that is the wrong list.
19. *Create Configuration Tables* writes a snapshot that can later act as an explicit override
    of the mailing list (the case described in the TEMPLATE-002C document, §9).
20. *Apply Document Reallocations* deletes TDoc tables for removed / withdrawn TDocs. The
    confirmation says so; the label does not.
21. Automatic-update triggers belong to the user who started them. Apps Script shows each user
    only their own triggers, so a second editor sees "Inactive" and can start a second one.
22. The manual single steps (3, 4, 5), the builds and the Legacy items do not take the document
    lock, so they can run while an automatic update is running.

---

## 4. Proposed menu for template reports

### 4.1 Tree

```
SA4 Report
├─ 🚀 Finish Report Setup                       (only until the report has been set up)
├─ ──────────
├─ 📄 Report
│   ├─ Build Report from Scratch…               runFullReportBuild
│   ├─ Update Report Now                        continuousUpdate            (see §6, B)
│   ├─ ──────────
│   ├─ Update E-mail Discussions                collectEmailDiscussionOnly
│   ├─ Update TDoc Revisions                    collectRevisionsOnly
│   ├─ Update Abstracts                         addAbstractsOnly
│   ├─ ──────────
│   └─ Report Status Summary                    analyzeReportStatus
├─ 💬 Prepare Discussion E-mails…               prepareTdocDiscussionEmails
├─ 🔀 Document Reallocation
│   ├─ Add or Change a Reallocation…            addDocumentReallocation
│   ├─ Show Reallocations                       viewAllReallocations
│   ├─ Apply Reallocations to Report…           applyDocumentReallocations
│   └─ Remove All Reallocations…                clearAllReallocations
├─ 🔄 Automatic Updates…                        manageTriggers
├─ ⚙️ Configure Meeting…                        configureMeetingSettings
├─ ──────────
├─ 🛠 Advanced and Repair
│   ├─ Check Connections                        testAllConnections
│   ├─ Format Report                            removeRowHeightAndSpacing
│   ├─ Re-arrange Revision Tables…              rearrangeRevisionTables
│   ├─ Remove Duplicate E-mail Entries…         removeDuplicateEmailEntries
│   ├─ Remove Wrong E-mail Matches…             cleanUpWrongEmailDiscussions
│   ├─ Clear Collection Caches…                 clearAllCaches
│   └─ Write Configuration Tables into Report   createConfigurationTables
└─ ℹ️ About This Report                         showTemplateInfo
```

22 items in 3 submenus (21 after setup), against 36 in 7. "…" marks an item that opens a
dialog or asks before acting.

### 4.2 Why this shape

- **Report** holds what a rapporteur does all meeting: build once, update, and the three
  explicit partial updates. The partial updates stay visible, because they are the answer to
  "the e-mails are stale" without a full update, and *Update Abstracts* is the only way to get
  abstracts when the automatic switch is off.
- **Prepare Discussion E-mails** is a single top-level item. A "Discussions" submenu would have
  one entry; *Update E-mail Discussions* belongs with the other updates.
- **Automatic Updates…** is one item, not Start / Stop / Status. The code has one dialog that
  shows the status and offers the one action that applies. Three menu items would need three
  new functions and two of them would always be the wrong one to click.
- **Configure Meeting…** once, at top level, not hidden in a submenu.
- **Build Report from Scratch…** rather than "Build / Rebuild": one name that is true before
  and after the first build, and says what happens to existing content.
- **Advanced and Repair** is everything that is only needed when something went wrong.

### 4.3 The menu name

*SA4 Report* in a report. The master template should use the same name for its two-item menu,
so that the creator's instructions ("in the report: SA4 Report menu…") match.

---

## 5. Changes: old → new

| Old path and label | New path and label |
|---|---|
| ⚠️Scripts⚠️ | SA4 Report |
| 🚀 Finish Report Setup | unchanged |
| INITIAL SETUP › ⚙️ Configure Meeting Settings | ⚙️ Configure Meeting… |
| REPORT OPERATIONS › 0️⃣ Configure Meeting | removed (duplicate) |
| INITIAL SETUP › 🧪 Test All Connections | Advanced and Repair › Check Connections |
| INITIAL SETUP › 📋 Create Configuration Tables | Advanced and Repair › Write Configuration Tables into Report |
| REPORT OPERATIONS › ▶️ Run Full Report Build | Report › Build Report from Scratch… |
| REPORT OPERATIONS › 1️⃣+2️⃣ Build Skeleton + TDOC Tables | hidden (§6) |
| REPORT OPERATIONS › 3️⃣ Collect E-mail Discussion | Report › Update E-mail Discussions |
| REPORT OPERATIONS › 4️⃣ Collect Revisions | Report › Update TDoc Revisions |
| REPORT OPERATIONS › 5️⃣ Add Abstracts | Report › Update Abstracts |
| REPORT OPERATIONS › 🔄 Continuous Update (New TDOCs + Status) | Report › Update Report Now |
| REPORT OPERATIONS › ⏰ Manage Auto-Update Trigger | 🔄 Automatic Updates… |
| REPORT OPERATIONS › 🔄 Update Report (During Meeting) | hidden (§6) |
| REPORT OPERATIONS › 📊 Analyze Report Status | Report › Report Status Summary |
| DOCUMENT MANAGEMENT › ➕ Add Document Reallocation | Document Reallocation › Add or Change a Reallocation… |
| DOCUMENT MANAGEMENT › 📊 View All Reallocations | Document Reallocation › Show Reallocations |
| DOCUMENT MANAGEMENT › 🔄 Apply Document Reallocations | Document Reallocation › Apply Reallocations to Report… |
| DOCUMENT MANAGEMENT › 🗑️ Clear All Reallocations | Document Reallocation › Remove All Reallocations… |
| DOCUMENT MANAGEMENT › 🔀 Re-arrange Revision Tables | Advanced and Repair › Re-arrange Revision Tables… |
| DOCUMENT MANAGEMENT › 📄 Parse Agenda Document | hidden |
| DOCUMENT MANAGEMENT › 🏗️ Auto-Create Report Structure | hidden |
| DOCUMENT MANAGEMENT › 🧹 Clean Up Wrong Email Discussions | Advanced and Repair › Remove Wrong E-mail Matches… |
| DOCUMENT MANAGEMENT › 🔧 Remove Duplicate Email Entries | Advanced and Repair › Remove Duplicate E-mail Entries… |
| TOOLS & DIAGNOSTICS › 🧪 Test TDOC List URL / Reviewer API / Email Feeds / Revisions Folder | hidden (covered by Check Connections) |
| TOOLS & DIAGNOSTICS › ✅ Validate Configuration | hidden |
| TOOLS & DIAGNOSTICS › 🗑️ Clear All Caches | Advanced and Repair › Clear Collection Caches… |
| FORMATTING & FIXES › 🎨 Format Document | Advanced and Repair › Format Report |
| FORMATTING & FIXES › 🔗 Fix Links (portal→FTP) | hidden |
| FORMATTING & FIXES › 📏 Fix Column Widths | hidden (part of Format Report) |
| EMAIL EXPORT › Prepare TDoc Discussion E-mails | 💬 Prepare Discussion E-mails… |
| LEGACY (old workflow) › 📝 Legacy: Build Initial Report | hidden |
| LEGACY (old workflow) › ⚠️ Legacy: Update All | hidden |
| ℹ️ About This Report | unchanged |

---

## 6. Hidden in template reports, kept in the code

Nothing is deleted from `Code.js`. Every function below stays, stays public, and stays in the
CENTRAL and Legacy menus exactly as today.

| Function | Why it leaves the template menu |
|---|---|
| `buildInitialReport`, `updateAll` (and the import code behind them) | Old workflow; pop-ups that conflict with the Full Build lessons; nothing to import in a template report |
| `updateReportIncremental` | Subset of *Update Report Now*; reports success after failures |
| `buildSkeletonWithTdocTables` | Clears the document without asking. It can return under *Advanced and Repair* once it has a confirmation (a small template-only wrapper) |
| `parseAgendaDocument`, `autoCreateReportStructure` | Manual skeleton path, superseded by the build |
| `testTdocListUrl`, `testReviewerApi`, `testEmailFeeds`, `testRevisionsFolder` | Covered by *Check Connections* |
| `validateConfiguration` | Configure Meeting shows readiness |
| `rewritePortalLinksInDoc_` | Expected not to be callable from a menu; main-meeting only |
| `fixColumnWidths` | Part of *Format Report* |
| second `configureMeetingSettings` entry | Duplicate |

All of them remain runnable from the Apps Script editor of a report if ever needed.

---

## 7. Risks and compatibility

### 7.1 Essentially zero-risk (menu only)

- Building the template report's menu in `template/ReportCreator.js` and leaving the existing
  menu code in `Code.js` as it is. `onOpen()` already branches on the template runtime, so one
  early return for template reports is the only change in `Code.js`. CENTRAL and Legacy do not
  contain `Release.js`, so their menu stays byte-for-byte what it is.
- Regrouping, hiding duplicates, hiding the LEGACY submenu, new labels: each item still calls
  the same function with no arguments.
- Reports already created from T-2026.10.0 to .3 keep their menu. They are pinned to their
  script.

### 7.2 Renames that need test and text changes

- **Tests.** `tests/template002b-runtime.test.js` asserts the template report's submenu list,
  the path of *Run Full Report Build*, the LEGACY submenu, and "no existing report operation
  was dropped" compared with the add-on menu. These must be rewritten on purpose. The add-on
  menu assertions (`addon005`, `addon009`, `no-duplicates`) stay untouched and keep guarding
  CENTRAL.
- **Texts that quote menu labels**, in `template/ReportCreator.js` (template only, safe to
  change): the file description written into each new report ("run ⚠️Scripts⚠️ > 🚀 Finish
  Report Setup"), the Finish Report Setup results ("⚠️Scripts⚠️ > 📝 INITIAL SETUP > …",
  "🚀 REPORT OPERATIONS > ▶️ Run Full Report Build"), the creator's closing text, and the About
  text.
- **Texts in `Code.js` shared with CENTRAL**: "The menu step 5️⃣ Add Abstracts…" in the trigger
  dialog, "⚙️ Configure Meeting Settings" in three diagnostics and three export errors, and the
  confirmation title "Run Full Report Build". Changing them changes CENTRAL's wording. Either
  leave them (slightly stale in a template report) or make them depend on the runtime.
- **Documentation**: the smoke-test tables in the TEMPLATE-002B/002C documents quote the old
  paths. They describe past releases and should stay as they are.
- No script calls a menu item by its label. Triggers call `continuousUpdate` by function name,
  which does not change.

### 7.3 Must not be consolidated without deeper work

- The four update paths (problem 8). They share the collector but differ in scope, locking and
  error handling. Hide the redundant ones; do not merge the functions.
- The two build paths. The Legacy import has its own table logic and parity tests.
- *Format Report* and *Fix Column Widths*; *Check Connections* and the single tests.

### 7.4 Not menu-only: each needs its own decision

| # | Issue | Smallest change | Touches |
|---|---|---|---|
| A | Build confirmation does not say the document is replaced | New confirmation text for the template runtime | `Code.js` (`runFullReportBuild`), TEMPLATE-002C tests of the text |
| B | *Update Report Now* is silent, also on failure | Template-only wrapper: document lock, "busy" message before any work, run the existing update core, no UI afterwards, throw if it reports failure (the core already returns success / error). The timer keeps calling `continuousUpdate` unchanged. | `template/ReportCreator.js`, new tests |
| C | Skeleton build clears without asking | Template-only wrapper with a confirmation; until then hidden | `template/ReportCreator.js` |
| D | Trigger dialog wording ("Trigger"), and no way to change the interval while active | Reword; offer "Change interval" when active | `Code.js` dialog, shared with CENTRAL |
| E | Connection test uses the general list | Use the report's list | `Code.js`, shared with CENTRAL |
| F | Triggers are per user (problem 21) | Say so in the dialog; a real fix needs a shared marker | design work |
| G | Manual steps do not take the document lock (problem 22) | Lock in the wrappers | `Code.js`, shared |

B follows the rule learned in 2.17.3: no UI call after the document was changed.

### 7.5 Can the menu adapt to the report's state?

Only cheaply in one way, which exists already: `onOpen()` may read the document's own
properties, and that is how *Finish Report Setup* disappears.

Adapting to "built / not built" is not recommended:

- there is no "built" marker today, so one would have to be added, or the document scanned on
  every open;
- a menu is only rebuilt when the document is opened, so after a build the old menu stays
  until reload, and rebuilding it from the script would be a UI call after the document was
  changed;
- a marker can be wrong (a report cleared by hand).

Wording that is true in both states (*Build Report from Scratch…*) solves the same problem
without state.

---

## 8. Recommended TEMPLATE-003 scope

**Stage 1: menu only (recommended as TEMPLATE-003).**

1. Template report menu as in §4, built in `template/ReportCreator.js`; one early return in
   `onOpen()`; CENTRAL and Legacy menus untouched.
2. Menu name *SA4 Report* for report and master template.
3. Template-only texts in `ReportCreator.js` updated to the new labels.
4. Tests: new assertions for the exact template tree, that every item calls an existing public
   function, that no LEGACY entry and no hidden function is in the template menu, that every
   hidden function still exists, and that the add-on menu is unchanged.

**Stage 1 additions, small and worth taking with it:** A (honest build confirmation) and B
(*Update Report Now* that fails visibly). Both are about the two most-used items.

**Later, separately:** C, D, E, F, G. D, E and G change code shared with CENTRAL and should be
decided with that in mind.

A stage 1 release changes `Code.js` (the `onOpen()` branch, and A if taken), so it needs a
version number and a new template release at that time. Not now.

### Decisions needed

1. Menu name: *SA4 Report*, or keep ⚠️Scripts⚠️?
2. *Automatic Updates…* as one dialog item (recommended), or separate Start / Stop / Status
   items (needs new functions)?
3. Take A and B into stage 1?
4. *Build Skeleton + TDOC Tables*: hide now and bring back with a confirmation later (C), or
   drop from the template menu for good?
5. *Write Configuration Tables into Report*: keep under Advanced, or hide (problem 19)?
6. Shared `Code.js` texts that quote old labels: leave, or make them runtime-dependent?
