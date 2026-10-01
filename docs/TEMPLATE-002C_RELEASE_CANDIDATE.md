# TEMPLATE-002C: Full Build orchestration, creator mailing list, discussion e-mail headers

2026-10-01. Branch `template-002c/full-build` (worktree `sa4-report-template-002c`), based on
the TEMPLATE-002B release candidate `d8cb075`. Local implementation and automated tests only.
Nothing was pushed or deployed; the deployed release T-2026.10.0 and its tag are unchanged.

**Status: ready for deployment as the next template release (proposed T-2026.10.1). Not deployed.**

---

## 1. Why this stage exists

The first report created from the template (meeting 86178, template T-2026.10.0, `Code.js`
2.17.0) ended its Full Build with "Exceeded maximum execution time". Creation and setup had
worked. The diagnosis (`diagnose/template-002b-fullbuild-timeout`, `a3021a7`, tests only):

- **Run Full Report Build called the single-step menu functions**, and each of those ends with
  its own blocking pop-up. Inside its one 6-minute execution the build stopped four times and
  waited for a click. The logged work added up to about a minute.
- The document was formatted twice.
- The optional Reviewer token was looked up and reported missing once per TDoc table, twice.
  This was noise, not the cause: without a token no request is made.

Two product decisions were added to the same stage: an optional mailing-list override in the
creator, and two changes to generated discussion e-mails (Reply-To, subject prefix).

---

## 2. Full Build

### 2.1 Architecture

`runFullReportBuild()` (menu) asks once, then calls `runFullReportBuildCore_()`, which has no
UI at all, then shows one summary.

| Phase | Core function | Notes |
|---|---|---|
| skeleton | `buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true })` | Same function as the menu item; the options suppress its pop-ups and its own formatting pass. Abstracts are not fetched here. |
| e-mail | `collectEmailDiscussionCore_()` | |
| revisions | `collectRevisionsCore_()` | |
| abstracts | `addAbstractsForTables_()` | Skipped as a whole when no Reviewer token is configured. |
| formatting | `removeRowHeightAndSpacing()` | The only formatting pass. |

The first phase that fails stops the build; later phases are not started; the failure is
reported once, in the final dialog, together with what did not run.

### 2.2 Dialogs before and after

| | Before (2.17.0) | After (2.17.1) |
|---|---|---|
| Before the work | "Continue?" | "Continue?" (unchanged) |
| After the skeleton | "Done: Built skeleton…" (blocking) | none |
| After e-mail | "E-mail discussion collection completed." (blocking) | none |
| After revisions | "Revision collection completed." (blocking) | none |
| After abstracts | "Abstract step completed…" (blocking) | none |
| At the end | "Full report build completed." | One summary: each phase with its result and duration, and the total |
| On failure | One error dialog, but only after the pop-ups before it had been clicked | One error dialog: the failed phase, its message, and the phases not run |

The single-step menu items (*Build Skeleton + TDOC Tables*, *Collect E-mail Discussion*,
*Collect Revisions*, *Add Abstracts*) keep their own completion pop-ups, with the same texts.

### 2.3 Reviewer token

- `REVIEWER_API_TOKEN` is optional. It is read **once per execution**
  (`getReviewerApiTokenForRun_()`); `getReportConfig_()` and the abstract fetch share that read.
- Without a token, Full Build skips the abstracts phase: no table scan, no request, one log
  line (`[FULLBUILD] abstracts: skipped … no Reviewer API token configured`).
- With a token, the tables are built without abstracts and the abstracts phase fetches them:
  one request per TDoc table, in its own timed phase.
- Outside Full Build (Continuous Update, the single-step item) a missing token is logged once
  per execution instead of once per TDoc table.
- Saving a token change in Configure Meeting resets the cached value.
- No token is stored in the template or in the setup information. A report created from the
  template has no token until one is entered in that report; provisioning is a separate concern.

### 2.4 Formatting: one pass

Kept: the final pass. Removed from Full Build: the pass inside the skeleton step.

`removeRowHeightAndSpacing()` formats every table of the document from scratch each time it
runs. The enrichment phases add content that needs formatting afterwards (e-mail lines,
revision tables, abstract rows), so the final pass is required, and a pass before them is
repeated work that the final pass redoes. No phase depends on formatting having happened.
The single-step *Build Skeleton + TDOC Tables* still formats, because nothing follows it.

### 2.5 Phase timing

```
[FULLBUILD] skeleton: start
[FULLBUILD] skeleton: done in 17210 ms -- 14 agenda items, 34 TDocs
[FULLBUILD] e-mail: start
[FULLBUILD] e-mail: done in 27040 ms
[FULLBUILD] revisions: start
[FULLBUILD] revisions: done in 3100 ms
[FULLBUILD] abstracts: start
[FULLBUILD] abstracts: skipped in 2 ms -- no Reviewer API token configured
[FULLBUILD] formatting: start
[FULLBUILD] formatting: done in 15020 ms
[FULLBUILD] total: 62380 ms -- skeleton done 17210 ms, e-mail done 27040 ms, …
```

Eleven lines per build. After a timeout the last `start` line names the phase that was running.

---

## 3. Create New SA4 Report: optional Mailing List

After *Look up* the dialog shows a **Mailing list** field with the list derived from the report
family. It can be replaced before *Create Report*; it does not have to be.

| Input | Result in the new report |
|---|---|
| Left as derived (or the derived name typed again, any case) | No `MAILING_LIST` property; the report derives the same list from its family, as before |
| Another list, as `<list>` or `<list>@list.etsi.org` | `MAILING_LIST = <list>` (plain name, as the collector normalizes it) |
| Anything else (spaces, another domain, a URL) | Refused with an explanation; *Create* stays disabled; the server refuses too, before anything is copied |
| Main meeting | The field shows the family list and is disabled: a main-meeting report always reads its family list. A forced override is refused. |

The value travels in the existing `mailingList` / `mailingListMode` fields of the setup
information and is stored by the same code as a manual Save, in the existing `MAILING_LIST`
property. Configure Meeting then shows the same value, and the collector, the recipient and
Reply-To all use it. Nothing in the code knows a particular meeting: for 86178 the lookup
derives `3GPP_TSG_SA_WG4`, and the user can type `3GPP_TSG_SA4_FS_6G_MED`.

The Email Collection Start Date decision is unchanged (new report: meeting start date; a
stored value wins; one property).

---

## 4. Generated discussion e-mails

### 4.1 Headers

| Header | Before | After |
|---|---|---|
| From | `DISCUSSION_EMAIL_SENDER` | unchanged |
| To | `<effective list>@list.etsi.org` (lower case) | unchanged |
| Reply-To | none (replies went to the sender) | `<effective list>@list.etsi.org` |
| Subject | `[<tag>,<agenda item>,<deadline>][<TDoc>] Discussion: <title>` | `[<agenda item>][<deadline>][<TDoc>] Discussion: <title>` |

Header order: `MIME-Version`, `X-Unsent`, `From`, `To`, `Reply-To`, `Subject`, `Content-Type`,
`Content-Transfer-Encoding`. The `.eml` body, quoted-printable encoding, CRLF line endings and
the ZIP are unchanged.

### 4.2 Reply-To rules

1. The effective mailing list is the one the recipient already comes from
   (`getMeetingContext_().sources.mailingList`): an ad-hoc meeting's Mailing List if set, else
   the report-family list; a main meeting always the family list.
2. It is normalized with the collector's rule (`normalizeEtsiListName_()`): `<list>` or
   `<list>@list.etsi.org`; the list name keeps its spelling.
3. Reply-To is `<list>@list.etsi.org`. It is never the sender.
4. A list that cannot be normalized refuses the export with an explanation. No file is
   written, and there is no fallback.

Example: Mailing List `3GPP_TSG_SA4_FS_6G_MED` gives
`Reply-To: 3GPP_TSG_SA4_FS_6G_MED@list.etsi.org`.

### 4.3 Subject

Old grammar: `[<tag>,<agenda item>,<deadline>][<TDoc>] Discussion: <title>`, where `<tag>` was
the mailing-list name without its SA4 prefix (`MBS`, `FS_6G_MED`) or the SWG name.

New grammar: `[<agenda item>][<deadline>][<TDoc>] Discussion: <title>`. The tag is gone and
nothing replaces it (no work item code). Agenda item, deadline token, TDoc bracket and
"Discussion: <title>" are exactly as before; a missing agenda item or deadline is omitted.

| Report | Old | New |
|---|---|---|
| MBS | `[MBS,2.5,26-10-15-1500CEST][S4aI260082] Discussion: …` | `[2.5][26-10-15-1500CEST][S4aI260082] Discussion: …` |
| 6G | `[FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260091] Discussion: …` | `[5.6.1][26-10-15-1500CEST][S4aP260091] Discussion: …` |
| Audio | `[AUDIO,3.1,26-10-15-1500CEST][S4aA260090] Discussion: …` | `[3.1][26-10-15-1500CEST][S4aA260090] Discussion: …` |
| Main SA4 | `[AUDIO,7.3,26-10-15-1500CEST][S4-261234] Discussion: …` | `[7.3][26-10-15-1500CEST][S4-261234] Discussion: …` |

The tag is still derived from the mailing list and still names the exported files
(`MBS_S4aI260082.eml`, the ZIP) and the default introduction. That was not part of the
decision; say so if it should go too.

Replies are still collected: the collector looks for the TDoc identifier anywhere in the
subject, so replies to the new form and to e-mails already sent in the old form are both
associated (tested).

This supersedes the earlier decision "subject tag derived from the mailing list". The
TEMPLATE-002A checks that asserted the tag were updated to the new form, not removed.

---

## 5. Version and release

- `Code.js` **2.17.1**: by the repository's convention a fix raises the patch number. The
  build no longer times out; the e-mail header and subject changes are small behaviour changes
  recorded in the same changelog entry.
- Proposed template release **T-2026.10.1**. `template-release/T-2026.10.0` is not moved; no
  new tag exists yet. The tag is created on the reviewed commit at deployment, as before.

---

## 6. Tests

Complete suite: **77 test files, 4,003 checks, 0 failures.**

| Group | Files | Checks |
|---|---|---|
| Existing `sa4-report` tests | 62 | 3,225 |
| Legacy parity suites (TEMPLATE-002A) | 9 | 408 |
| TEMPLATE-002A intentional differences | 1 | 20 |
| Template: bootstrap, release tool, probe | 3 | 122 |
| Template runtime (TEMPLATE-002B) | 1 | 125 |
| TEMPLATE-002C (new) | 1 | 103 |

Existing expectations that changed, all for the decided behaviour and with the number of
checks per file unchanged:

- Subject and header expectations in the exporter tests: 47 checks in 8 files
  (`addon009-discussion-email-export` 15, `addon009-legacy-exporter-port` 24, `email-export-doc-links` 1,
  `email-export-live-path` 1, `legacy-parity-adhoc-email-collection` 1,
  `legacy-parity-diagnose-6g-email-collection` 2, `legacy-parity-email-doc-links` 1,
  `template002a-intentional-differences` 2). The three Legacy parity checks among them are
  marked as intentional differences.
- `perf003-formatting-skip`: one more formatting call site is allowed to be conditional (the
  skeleton build's Full Build option); one check added.
- `template002b-runtime`: the creator dialog now has two inputs (the mailing list is hidden
  until lookup); the version check reads the version from `Code.js`.
- The diagnostic test from `a3021a7` described the old behaviour; it is replaced by the
  TEMPLATE-002C suite.

---

## 7. The existing smoke-test report

The report created from T-2026.10.0 during the smoke test is pinned to `Code.js` 2.17.0. It
must not be used to validate this fix, and it is not updated. After T-2026.10.1 is installed
in the master template, a **new** report is created for the final smoke test.

---

## 8. Final smoke test (after deployment of T-2026.10.1)

| # | Action | Expected |
|---|---|---|
| 1 | Template → Template Release Info | T-2026.10.3 (Code.js 2.17.3) |
| 2 | Create New SA4 Report → 86178 → Look up | Mailing list field shows `3GPP_TSG_SA4_FS_6G_MED` (with T-2026.10.2; T-2026.10.1 showed the general list, see §9) |
| 3 | Leave it as it is → Create Report → open it | New report |
| 4 | Run Full Report Build → confirm once | No pop-up at all after the confirmation (2.17.3, §10); the "Running script" notice ends; the report is built |
| 5 | Configure Meeting | Mailing List `3GPP_TSG_SA4_FS_6G_MED`; start date = meeting start date |
| 6 | Set the Discussion E-mail Sender, then Prepare TDoc Discussion E-mails for one TDoc; open the `.eml` | `Reply-To: 3GPP_TSG_SA4_FS_6G_MED@list.etsi.org`; subject `[<agenda item>][<deadline>][<TDoc>] Discussion: …` |
| 7 | Extensions → Apps Script → Executions | `runFullReportBuild` is **Completed**; `[FULLBUILD]` lines with the phase durations |

Still to be seen live for the first time: a complete Full Build in a copy, the creator dialog
with the new field, and the first trigger execution in a copy.

---

## 9. Patch after the T-2026.10.1 creator smoke test: 6G family mailing list (Code.js 2.17.2)

T-2026.10.1 was deployed and its creator was tried live with meeting 86178. The Mailing list
field was pre-filled with `3GPP_TSG_SA_WG4`. That is wrong for the 6G report family.

**Root cause.** The one family → list table, `MAILING_LISTS` in `Code.js`, mapped `6G` to the
general SA4 list. Every consumer derives from that table, so all of them showed or used the
general list for a 6G report without a saved Mailing List. Legacy has the same table entry;
its 86178 report works because a Mailing List override was saved in it.

**Fix.** One table entry: `'6G': '3GPP_TSG_SA4_FS_6G_MED'`. No other code changed, and nothing
is keyed on a meeting ID.

| Consumer | How it gets the list | After the fix, 6G without an override |
|---|---|---|
| Creator (lookup) | `MAILING_LISTS[family]` | field shows `3GPP_TSG_SA4_FS_6G_MED` |
| Configure Meeting | `buildReportFamilyInfo_()` | derived default `3GPP_TSG_SA4_FS_6G_MED` |
| Collector, RSS | `getReportConfig_().LIST_NAME` → `getCollectorConfig_()` | `…RSS&L=3GPP_TSG_SA4_FS_6G_MED…` |
| Collector, A1 archive | the same resolved list | `…A1=…&L=3GPP_TSG_SA4_FS_6G_MED` |
| Discussion e-mail To | `getMeetingContext_().sources.mailingList` | `3gpp_tsg_sa4_fs_6g_med@list.etsi.org` |
| Discussion e-mail Reply-To | the same value | `3GPP_TSG_SA4_FS_6G_MED@list.etsi.org` |

A saved Mailing List override still wins, and the creator's override field works as before.
Audio, Video, MBS, RTC, Liaison and New are unchanged.

Consequences to know:

- **Main-meeting 6G reports** read the 6G list too, because a main-meeting report always uses
  its family list. Before, they read the general SA4 list.
- **An unconfigured document** defaults to family 6G, so its default list is now the 6G list.
- **A "Create Configuration Tables" snapshot written before 2.17.2** in a 6G document holds
  the old default (the general list). It now differs from the family default, so the
  collector treats it as an explicit table override and reads the general list. This cannot
  happen in a report newly created from the template; it matters only if an existing 6G
  document with such a table is moved to this code.
- This is one more intentional difference from Legacy (marked in the parity suite).

**Release.** `Code.js` 2.17.2, proposed template release **T-2026.10.2**.
`template-release/T-2026.10.1` stays on `6b65ebd` as the record of what was live-tested.

**Tests.** Complete suite 77 files, 4,016 checks, 0 failures. Seven existing expectations of
the old 6G default were updated (golden fixture, `pure-logic`, `meeting-context` ×2,
`addon008a2`, `addon009-discussion-email-export`, and the Legacy parity e-mail suite, where
the changed checks are marked as intentional differences). The TEMPLATE-002C suite gained the
family-table checks, the collector RSS / A1 checks, the To / Reply-To checks for an unchanged
default, a second 6G meeting with a different ID, and keeps the override checks with another
list.

---

## 10. Patch after the T-2026.10.2 smoke test: no UI call after the build (Code.js 2.17.3)

**Live evidence (report for 86178, T-2026.10.2).** `runFullReportBuild` started 11:40:53, ran
170.6 s and ended FAILED. All five phases had finished at 11:42:08 (64.6 s). The final
"Success" pop-up never appeared. At 11:43:44 Apps Script reported "Service Documents failed
while accessing document with id …". After a reload the report was complete: content, e-mail
discussions and formatting were all there (abstracts absent as intended, no token).

**Cause.** After `[FULLBUILD] total` the code did one thing: `ui.alert('Success', summary)`.
That call hung for about 96 seconds and then threw. The build's changes were saved; only the
notification failed, and it turned a completed build into a failed execution. Why Google's
`ui.alert` fails after a build that has rewritten the whole document is not something the code
can show; the fix does not depend on the reason.

**Fix.** `runFullReportBuild()` makes no UI call after the build has changed the document.

| Outcome | Before (2.17.1 / 2.17.2) | After (2.17.3) |
|---|---|---|
| Build completed | `ui.alert('Success', summary)` | the function ends; the result is in the `[FULLBUILD]` log lines; the execution is "Completed" |
| A phase failed | `ui.alert('Error', summary)` | the function throws an error carrying the summary (failed phase, its message, phases not run); Docs shows it and the execution is "Failed" |
| Before the build | "Continue?" | unchanged (it runs before any change to the document, and worked live) |

Only `runFullReportBuild()` changed. No phase, no collector, exporter, formatter, token, trigger
or creator code is touched.

**What the user sees.** After confirming, Docs shows its "Running script" notice; when it goes
away the report is built. There is no completion pop-up any more. The project has no
non-blocking notification in use (no toast exists for Docs; sidebars and modeless dialogs are
UI calls of the same kind as the one that failed), so none was added.

**Why not the alternatives.**

- *Keep the pop-up but catch its failure:* the call did not fail quickly, it hung for about
  96 seconds first. Catching it would still leave every build waiting that long.
- *`saveAndClose()` between phases:* unproven for a bound script that keeps editing its own
  document, and it would change how the successful build writes. There is no evidence it is needed.
- *One execution per phase:* a redesign of a build that works.

**Not proven.** Whether the end of the script (where Google writes pending changes) can raise
the same error without the pop-up. In the live run the changes were saved, so the write itself
worked; one build with this release confirms it.

**Release.** `Code.js` 2.17.3, proposed template release **T-2026.10.3**.
`template-release/T-2026.10.2` stays on `77a23f3`.

**Tests.** Complete suite 77 files, 4,030 checks, 0 failures. The Full Build checks of the
TEMPLATE-002C suite were rewritten for this behaviour and extended: a completed build makes no
UI call after the confirmation and ends normally, also when the UI is made to fail the way it
did live; the five phases, their order, the total line, one formatting pass and the token
handling are asserted as before; a failure in each of the five phases makes the menu function
throw, with the failed phase and the phases not run in the message, also when the UI cannot
show anything.
