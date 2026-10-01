# TEMPLATE-003 stage 1: the "SA4 Report" menu (release candidate)

2026-10-01. Branch `template-003/stage1-menu`, based on the menu review
(`template-003/menu-review`, `a00012e`) on master `02c46fc`.

**Status: COMPLETE / LIVE. Template release T-2026.10.4 (`Code.js` 2.17.4, commit `962fb7d`)
is deployed to the master template and accepted live. No further manual acceptance is
required for TEMPLATE-003 stage 1. See §10.**

Sections 1 to 9 are the release candidate as reviewed; §10 is the closeout record.

The review and the reasons are in `TEMPLATE-003_MENU_REVIEW.md`. This document records what
was decided and what was built.

---

## 1. Decisions

| # | Decision |
|---|---|
| 1 | The menu of a template report is called **SA4 Report**. The master template uses the same name. CENTRAL and Legacy menus stay as they are. |
| 2 | Automatic updates stay one item, **🔄 Automatic Updates…**, opening the existing dialog (`manageTriggers`). No separate Start / Stop / Status functions. |
| 3 | Review items **A** (honest build confirmation) and **B** (Update Report Now that fails visibly) are part of this stage. |
| 4 | *Build Skeleton + TDOC Tables* is removed from the template menu, not moved to Advanced. The function stays. |
| 5 | *Create Configuration Tables* is removed from the template menu. The function stays. |
| 6 | Texts in `Code.js` that CENTRAL and Legacy also show are not rewritten. Template-only texts follow the new labels. Shared texts that are now stale in a template report are listed in §6. |
| 7 | Final adjustment: the one shared sentence that named a menu item ("The menu step 5️⃣ Add Abstracts…") is replaced by neutral wording, the same in every runtime (§6). |
| – | *Re-arrange Revision Tables* is not shown either: the complete update already places revisions. |

---

## 2. The menus

### 2.1 Master template

```
SA4 Report
├─ 🆕 Create New SA4 Report            showCreateReportDialog
├─ ──────────
└─ ℹ️ Template Release Info            showTemplateInfo
```

### 2.2 Report

```
SA4 Report
├─ 🚀 Finish Report Setup               finishReportSetup        (only until setup is complete)
├─ ──────────                                                    (only until setup is complete)
├─ 📄 Report
│   ├─ Build Report from Scratch…       runFullReportBuild
│   ├─ Update Report Now                updateReportNow          (new)
│   ├─ ──────────
│   ├─ Update E-mail Discussions        collectEmailDiscussionOnly
│   ├─ Update TDoc Revisions            collectRevisionsOnly
│   ├─ Update Abstracts                 addAbstractsOnly
│   ├─ ──────────
│   └─ Report Status Summary            analyzeReportStatus
├─ 💬 Prepare Discussion E-mails…       prepareTdocDiscussionEmails
├─ 🔀 Document Reallocation
│   ├─ Add or Change a Reallocation…    addDocumentReallocation
│   ├─ Show Reallocations               viewAllReallocations
│   ├─ Apply Reallocations to Report…   applyDocumentReallocations
│   └─ Remove All Reallocations…        clearAllReallocations
├─ 🔄 Automatic Updates…                manageTriggers
├─ ⚙️ Configure Meeting…                configureMeetingSettings
├─ ──────────
├─ 🛠 Advanced and Repair
│   ├─ Check Connections                testAllConnections
│   ├─ Format Report                    removeRowHeightAndSpacing
│   ├─ Remove Duplicate E-mail Entries… removeDuplicateEmailEntries
│   ├─ Remove Wrong E-mail Matches…     cleanUpWrongEmailDiscussions
│   └─ Clear Collection Caches…         clearAllCaches
└─ ℹ️ About This Report                 showTemplateInfo
```

19 items in 3 submenus after setup (20 before), against 35 / 36 in 7. Separators inside
*Report* and before *Advanced and Repair* are the only addition to the agreed structure.

### 2.3 No longer shown in a template report, still in the code

`buildInitialReport`, `updateAll` (the LEGACY submenu), `updateReportIncremental`,
`buildSkeletonWithTdocTables`, `createConfigurationTables`, `rearrangeRevisionTables`,
`parseAgendaDocument`, `autoCreateReportStructure`, `testTdocListUrl`, `testReviewerApi`,
`testEmailFeeds`, `testRevisionsFolder`, `validateConfiguration`, `rewritePortalLinksInDoc_`,
`fixColumnWidths`, the second *Configure Meeting* entry, and `continuousUpdate` as a menu
item (it remains the function the timer calls).

No function was removed from `Code.js` or `ReportCreator.js`.

---

## 3. A: Build Report from Scratch

The item calls `runFullReportBuild()`, as before. In the template runtime its one question is
now:

> **Build Report from Scratch**
>
> This replaces the content of this document with a newly generated report.
>
> ⚠️ Everything that is in the document now is removed first. Meeting minutes and any other
> content entered by hand will be lost. Only the Document Reallocations table is kept.
>
> To bring an existing report up to date without losing anything, answer No and use
> Report > Update Report Now.
>
> The build then runs without further questions:
> • report structure and TDoc tables
> • e-mail discussions
> • TDoc revisions
> • abstracts (only if a Reviewer API token is configured)
> • formatting
>
> There is no completion message: the build has finished when "Running script" disappears.
>
> Replace the content of this document and build the report from scratch?

- Asked before anything is changed. *No* does nothing.
- Nothing else about the build changed: no dialog between the phases, no dialog at the end, a
  failed phase throws (2.17.1 and 2.17.3).
- In the CENTRAL add-on and in a Legacy copy the question is the old one, word for word.

---

## 4. B: Update Report Now

`updateReportNow()` (new, `ReportCreator.js`) is the manual complete update:

1. refuses in the master template;
2. takes the same document lock as `continuousUpdate()`; if an update is already running it
   does nothing and throws ("another update of this report is running right now");
3. runs `continuousUpdateCore_()`, the same work the timer does: new TDocs, statuses, summary
   table, revision placement, abstracts if switched on, e-mail discussions, revisions,
   formatting when tables changed;
4. releases the lock;
5. a completed update ends normally, with no dialog;
6. a failed update throws, so Docs shows the error and the execution is "Failed".

| Outcome | What the user sees |
|---|---|
| Completed | "Running script" disappears. No dialog. |
| The update itself failed (for example the TDoc list could not be read) | Error: "The report update failed: …" |
| The update ran, but e-mail or revision collection failed | Error: "The report update finished, but one part failed: • e-mail discussions: …". New TDocs and statuses were still updated. |
| Another update is running | Error: "The report was not updated: another update of this report is running right now." |

**The timer is untouched.** It still calls `continuousUpdate()`, which logs a failure and ends
normally, skips silently when the lock is busy, and returns nothing.

**One change in shared code was needed for this.** The collector (`collectorUpdate_()`) catches
the errors of its two steps and only logs them, so the update core reported success even when
e-mail collection had failed. Now:

- `collectorUpdate_()` also *returns* `{ failures: [...] }`. It returned nothing before, and no
  caller reads the value except the update core.
- `continuousUpdateCore_()` adds `collectorFailures` to its result when there are any.
  `success` and `error` are as before, and a clean run returns exactly `{ success: true,
  error: null }`.

Logging, the order of the steps, and "one failed step does not stop the other" are unchanged.
The add-on scheduler reads only `success` and `error`.

Limit: an error that `checkRSSFeed_()` or `updateRevisions_()` handle inside themselves is
still only logged. This stage does not look into those functions.

---

## 5. Production diff against T-2026.10.3

Two files change. `HyperLink.js`, the manifest and both HTML files are unchanged.

### `Code.js` (2.17.3 → 2.17.4)

| Function | Change | Part of |
|---|---|---|
| header | version line, changelog entry | release |
| `onOpen()` | The template runtime builds its menus in `ReportCreator.js` and returns. The menu code below it lost its template branches and is again line for line what it was before the template work (compared with `d7c23bc`). | menu |
| `runFullReportBuild()` | In the template runtime the question comes from `confirmTemplateBuildFromScratch_()`; otherwise the old `ui.alert` | A |
| `collectorUpdate_()` | also returns the failures it logs | B |
| `continuousUpdateCore_()` | passes those failures on in its result | B |
| `manageTriggers()` | one sentence of the dialog (§6); nothing else | decision 7 |

No other function of `Code.js` differs.

### `template/ReportCreator.js`

| Function | Change | Part of |
|---|---|---|
| `TEMPLATE_MENU_NAME_` (new) | `'SA4 Report'` | menu |
| `buildTemplateMasterMenu_()` | menu name | menu |
| `buildTemplateReportMenu_()` (new) | the report menu | menu |
| `confirmTemplateBuildFromScratch_()` (new) | the build question | A |
| `updateReportNow()`, `describeUpdateReportNowFailure_()` (new) | the manual update | B |
| `serializeBootstrapDescription_()`, `finishReportSetup()`, `showCreateReportDialog()`, `describeTemplateRuntime_()` | texts that named the old menu now name the new one; About says "Automatic updates" instead of "Continuous Update" | menu texts |

### Outside menu construction, A and B

Nothing in behaviour. Two things are worth naming because they are not menu lines:

1. The additive return values in `collectorUpdate_()` / `continuousUpdateCore_()` (§4). They
   are part of B, but they are in code CENTRAL also runs. Nothing there reads them.
2. Template-only text changes (file description of new reports, Finish Report Setup results,
   the creator's closing text, About).
3. One sentence in the trigger dialog (§6). This is the only change CENTRAL users can see.

### CENTRAL and Legacy

- Without `Release.js` the menu is identical, item by item, to the menu the T-2026.10.3
  `Code.js` builds (tested against that release, and against a literal copy of the tree).
- `continuousUpdate`, `continuousUpdateForDocument_`, `runAddonScheduler_`,
  `runAddonSchedulerTrigger`, `createContinuousTrigger`, `deleteContinuousTrigger`,
  `updateReportIncremental`, `updateAll`, `updateAllFromWeb`, `buildInitialReport` are
  byte-for-byte the T-2026.10.3 functions (tested).
- `manageTriggers` differs from T-2026.10.3 in one sentence of its dialog and nothing else
  (tested).
- The new `collectorFailures` value is read by `updateReportNow()` only; the scheduler reads
  `success` and `error`, and the other callers of the collector ignore its result (tested).
- The Legacy parity suites pass unchanged: 9 files, 409 checks.
- The Legacy repository was not touched.

---

## 6. Shared texts that are now stale in a template report

**Changed (decision 7).** The hint under "Fetch abstracts during each update" in the trigger
dialog said `The menu step "5️⃣ Add Abstracts" always works regardless of this setting.` A
template report has no such item. It now reads, in every runtime:

> Abstracts can also be updated manually at any time.

The sentence names no menu item, so it is correct for the template menu (*Update Abstracts*)
and for the CENTRAL / Legacy menu (*5️⃣ Add Abstracts*).

**Not changed (decision 6):**

| Where | Text | In a template report |
|---|---|---|
| Automatic Updates dialog | title "Manage Continuous Update Trigger", heading "Continuous Update Trigger", buttons "Start Trigger" / "Stop Trigger", messages "Trigger started / stopped" | Different vocabulary from the menu item that opens it. Understandable, not wrong. |
| Discussion e-mail export errors | "Open Configure Meeting Settings …" | The item is *Configure Meeting…*. Close enough. |
| Partial updates | "E-mail discussion collection completed." etc. | Fine. |

Messages inside functions that are no longer in the template menu (the single tests,
*Validate Configuration*, *Auto-Create Report Structure*) also quote old labels, but cannot be
reached from the template menu.

---

## 7. Tests

Complete suite: **78 test files, 4,143 checks, 0 failures** in the main worktree. One of those
checks runs only where a root `.clasp.json` exists (`tests/template-release.test.js`), so a
clean checkout has 4,142. Legacy parity: 9 files, 409 checks, 0 failures.

| | Files | Checks |
|---|---|---|
| Before (T-2026.10.3 closeout) | 77 | 4,030 |
| New: `tests/template003-menu.test.js` | +1 | +112 |
| Changed: `tests/template002b-runtime.test.js` | | 4 checks rewritten, count unchanged (125) |

`template003-menu.test.js` covers, in the order of the task:

1. exact master menu tree; 2. exact report tree before setup; 3. after setup;
4. no LEGACY submenu or item; 5. every hidden function still exists (template and add-on
runtime); 6. the add-on menu equals a literal tree and the menu of the T-2026.10.3 `Code.js`;
7. Legacy: the same menu code path, plus the unchanged parity suites;
8. *Build Report from Scratch* → `runFullReportBuild`; 9. the question says the content is
replaced and hand-entered minutes are lost; 10. no UI call after the build, also with a UI
that fails after the work;
11. *Update Report Now* runs the complete update (real `continuousUpdateCore_()` and
`collectorUpdate_()`), not `updateReportIncremental`; 12. it returns normally on success, shows
no dialog, and throws for a failed update, a failed collector step, and a busy lock;
13. `continuousUpdate()` still swallows, skips and returns nothing, and is byte-identical;
14. the three partial updates run their own steps with their existing messages;
15. *Automatic Updates* opens the existing trigger dialog;
16. no function of T-2026.10.3 was removed, and exactly four were added;
17. the trigger dialog's abstracts hint is the neutral sentence in both runtimes, names no
menu item, and is the only difference of `manageTriggers()` from T-2026.10.3.

The comparisons with T-2026.10.3 read that release from git and are skipped, with a note, in a
checkout that does not have the tag.

The four rewritten checks in `template002b-runtime.test.js` asserted the old template menu
(submenu list, path of the build, LEGACY submenu, "no operation dropped"). They now assert the
new submenus, the new path, the absence of the old import, and that every operation of the
add-on menu still exists as a function.

---

## 8. Not in this stage

No change to: trigger ownership, changing the interval while updates run, the mailing list
the connection test uses, locking of the manual steps, the build and update functions
themselves, exported file names and the list tag, the default discussion introduction.

---

## 9. Release

- `Code.js` **2.17.4**; proposed template release **T-2026.10.4**.
- At the time of the review no tag `template-release/T-2026.10.4` existed; it was created on
  the reviewed commit at deployment (§10). T-2026.10.0 to .3 are not moved.
- Reports already created keep their menu: each is pinned to the release it was created from.

### Acceptance after deployment (for later)

| # | Action | Expected |
|---|---|---|
| 1 | Open the master template | Menu *SA4 Report* with two items; *Template Release Info* shows T-2026.10.4 (Code.js 2.17.4) |
| 2 | Create a new report, open it | Menu *SA4 Report* as in §2.2, starting with *Finish Report Setup* |
| 3 | Report › Build Report from Scratch… | The question of §3; *Yes* builds; no dialog afterwards; execution Completed |
| 4 | Reload | *Finish Report Setup* is gone |
| 5 | Report › Update Report Now | No dialog; `continuousUpdate`-style log; execution `updateReportNow` Completed |
| 6 | Automatic Updates… → Start, 30 minutes | Trigger for `continuousUpdate`; About shows "Automatic updates: every 30 minutes" |

---

## 10. Closeout: T-2026.10.4 deployed and accepted live

### 10.1 Release

| | |
|---|---|
| Template release | **T-2026.10.4** |
| `Code.js` | 2.17.4 |
| Release commit | `962fb7de1f58302b54c51eeafa3c5c57d50f12c3` (tag `template-release/T-2026.10.4`, local) |
| Deployed | 2026-10-01, to the master template "SA4 Report Template", by a normal `clasp push` of the release bundle (seven files, no `--force`) |
| Verified | Before the push the template held exactly T-2026.10.3. A fresh pull after the push showed seven files, each byte-identical to the bundle. |
| Status | **Accepted. Stage 1 is COMPLETE / LIVE.** |

Template releases so far. The tags are local and are not moved.

| Release | `Code.js` | Commit | |
|---|---|---|---|
| T-2026.10.0 | 2.17.0 | `d8cb075` | TEMPLATE-002B |
| T-2026.10.1 | 2.17.1 | `6b65ebd` | TEMPLATE-002C |
| T-2026.10.2 | 2.17.2 | `77a23f3` | TEMPLATE-002C |
| T-2026.10.3 | 2.17.3 | `39dbfa1` | TEMPLATE-002C, accepted |
| **T-2026.10.4** | **2.17.4** | **`962fb7d`** | **TEMPLATE-003 stage 1, accepted** |

### 10.2 Live acceptance

| # | Step | Result |
|---|---|---|
| 1 | Reload the master template | Reloaded successfully. |
| 2 | Template Release Info | "SA4 Report Template (master document)", T-2026.10.4, Code.js 2.17.4, commit `962fb7d`, the correct template Script ID. |
| 3 | Create a new report for meeting 86178 | Created successfully. The normal one-time authorization prompt of the newly copied bound script appeared and was completed. |
| 4 | Inspect the *Build Report from Scratch* question | It clearly stated that the document content is replaced, that meeting minutes and content entered by hand are lost, that the Document Reallocations table is preserved, that *Update Report Now* is the non-destructive alternative, which phases the build runs, and that there is no completion pop-up. |
| 5 | Run *Build Report from Scratch* on the new report | Completed successfully, in about one minute. |
| 6 | Run *Update Report Now* on the built report | Completed normally. No error was reported, and no success or completion pop-up interfered with the workflow. |

**No further manual acceptance is required for TEMPLATE-003 stage 1.**

### 10.3 What is live now

- A report created from the template has the **SA4 Report** menu of §2.2; the master template
  has the two-item menu of §2.1.
- *Build Report from Scratch…* asks once, with the warning of §3, and shows nothing afterwards.
- *Update Report Now* runs the complete update and shows nothing when it completes; a failure
  is an error and a Failed execution (§4).
- The timer of *Automatic Updates…* still runs `continuousUpdate()`, unchanged.
- Reports created from T-2026.10.0 to .3 keep the menu of their release.
- The CENTRAL add-on and Legacy were not deployed. In `Code.js` 2.17.4 their only visible
  difference is the neutral abstracts sentence in the trigger dialog (§6), and it reaches
  CENTRAL only with a future CENTRAL deployment.

### 10.4 Still open, not part of stage 1

The review items that were deferred (§8, and items D to G of the review): the trigger dialog's
vocabulary and changing the interval while updates run, the mailing list the connection test
uses, trigger ownership per user, and the document lock for the manual partial updates.

### 10.5 Automated tests at closeout

Complete suite: **78 test files, 4,143 checks, 0 failures** (main worktree). Legacy parity:
9 files, 409 checks, 0 failures. The closeout changed this document only: no production
code, no test and no version number.
