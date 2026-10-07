# Ad-hoc meeting automation and the personal Reviewer token (T-2026.10.9)

Developed 2026-10-06 on master `a960718` (T-2026.10.8, `Code.js` 2.20.0), in two batches that
were accepted one after the other and released together. This document is the record of the
release; it was written after it.

**Status: RELEASED AND DEPLOYED. Template release T-2026.10.9, `Code.js` 2.21.0, release commit
`0eb8bedcc04a3552a16bd5af3e6e7463d6624a4c`, tag `template-release/T-2026.10.9`. Published and
deployed to the master template on 2026-10-06 (§6). T-2026.10.8 and its tag were not changed.
No existing report was updated.**

The release changes template reports going forward. Nothing happens to an existing report
because the code is installed, and no report was migrated.

## 1. Batch 1: four improvements for ad-hoc template reports

### 1.1 The opening sentence

A new ad-hoc report without sessions has one generated sentence in its "Opening of the
session" sub-section. It used to be

    <Chair> opens the session on October 1, 2026 at <start> CEST.

with the start time left as a placeholder although the Portal has it, and "CEST" written into
the code. It is now, for the same meeting,

    <Chair> opens the session on October 1, 2026 at 15:30 UTC+2.

- The date and the time of day come from the Portal's `StartDate`; the zone is the offset of
  the Portal's `StartTimeZone`, written as `UTC+2`, `UTC-3:30` or `UTC`. No zone name is
  derived, and none is assumed.
- A value the Portal does not give stays a visible placeholder: `<start>`, `<time zone>`,
  `<meeting date>`. Midnight (`00:00`) counts as no time.
- `<Chair>` is always the placeholder. There is no field for the chair anywhere; it is filled
  in by hand.
- The two values are stored as Document Properties `MEETING_START_TIME` and
  `MEETING_TIME_ZONE`, together with `MEETING_START_BASIS` (the meeting id and the meeting date
  they belong to). When the report has since been given another meeting id or date, the stored
  start is not used.
- They are stored the same way whichever way the meeting was chosen: by *Create New SA4
  Report*, or later by *Configure Meeting* (its Save sends what its Resolve got from the
  Portal; no further Portal request). A report changed from meeting A to meeting B never keeps
  A's start.
- Unchanged: main-meeting reports (their opening is copied from the report template) and
  ad-hoc reports with sessions (the Session administration section says who chaired and when).

Functions: `buildMeetingOpeningSentence_()`, `computeMeetingStartTimeFromStartDate_()`,
`computeMeetingTimeZoneLabel_()`, `storeMeetingStart_()`, `meetingStartFromResolved_()`,
`getMeetingStartForOpening_()`.

### 1.2 The drafts / revisions folder of a new ad-hoc report

`REVISIONS_URL` is where the revisions of a report's TDocs are found and where its discussion
e-mails send them. A main meeting derives it by formula. An ad-hoc series has its own folder,
`<meeting root>/inbox/drafts/`, which the meeting lookup derives from the Portal's document
folder but does not ask for, so a new ad-hoc report had none until it was typed in.

- *Finish Report Setup* now asks for the derived folder **once**, after the report is marked as
  set up, and stores it as `REVISIONS_URL` when the answer is a folder listing (HTTP 200 and a
  listing; a redirect is not followed).
- Any other answer — 403, a redirect, another status, an empty page, a page that is no
  listing, a failed request — stores nothing as fact. The candidate is recorded apart
  (`SA4_REVISIONS_URL_CANDIDATE`) and shown in the setup message and in *About This Report*.
- A later **Save of Configure Meeting** asks again, while the report has no folder, for the
  candidate of the document folder the report has then. A report that was given another
  meeting gets that meeting's candidate, never the recorded one.
- A stored folder is never asked for and never replaced. The meeting lookup, an update and a
  build ask nothing. Only an https address on a 3GPP host is asked.
- **Two ways a folder becomes `REVISIONS_URL`, on purpose.** Automatically: only a derived
  folder that a request confirmed. By the user: Configure Meeting shows the derived folder as
  "suggested — please review", and a Save with it in the field stores it without a request.
  That is the manual way for when the 3GPP server refuses scripted requests, and it is kept.

Functions (`template/ReportCreator.js`): `confirmAdhocRevisionsFolderWith_()`,
`retryAdhocRevisionsFolderWith_()`, `afterTemplateConfigurationSaved_()`.

### 1.3 No abstract for a withdrawn TDoc

- One rule, `abstractFetchBlockedBy_()`: no abstract is asked for when **either** the Status
  in the report (text, or the value of a status dropdown) **or** the status in the Portal's
  TDoc list says withdrawn.
- It is enforced at the head of `fetchAndAddAbstract_()`, before the token, the "no summary"
  cache and the request, so every path that asks goes through it: the sweep (an update with
  abstracts on, *Update Abstracts*, the abstracts phase of a build), the completion of an
  uploaded TDoc, and the old table builder.
- An abstract that is in the report stays. Nothing is cached for a withdrawn TDoc, so nothing
  stands in the way later.
- A TDoc that is active again gets its abstract from the normal sweep or *Update Abstracts*,
  once neither side says withdrawn. (The status sync does not take a Status off "withdrawn" by
  itself; that rule is unchanged.)
- A withdrawn TDoc is not counted as a candidate or as an attempt.

### 1.4 Revision placement after Build Report from Scratch

A build ordered revisions only inside one agenda item; a revision registered under another
agenda item than its parent stood apart until the first update. `buildSkeletonWithTdocTables()`
now ends its table writing with the pass an update uses (`rearrangeRevisionTables_()`), on the
TDoc list it has already downloaded. A revision stands directly below the document it revises,
chains stay in order, also across agenda items. A failure of that pass is noted in the result
and does not fail the build. The Revisions cell (one revision per line) is not affected.

## 2. Batch 2: the personal Reviewer API token

The Reviewer token of a report is a Script Property of the report's own script project. Every
report created from the template is a new project, so the token had to be entered in each one,
and every editor of a report can read it there.

### 2.1 Architecture (template reports only)

- **Precedence:** the token of the report (`REVIEWER_API_TOKEN`, as before) → the personal
  token → none. Existing reports and deliberate per-report tokens are unaffected. Nothing is
  copied from a report into personal storage unless the user asks.
- **Storage:** one JSON file in the user's own My Drive,
  `{ "schema": "sa4-report-user-settings/1", "reviewerApiToken": "…" }`, named
  `SA4 Report – private settings (do not share).json`.
- **Discovery:** by a Drive custom file property (`sa4ReportUserSettings = 1`) among the files
  the user owns — never by name, and not by anything stored in a report. A fresh copy of the
  template, which has nothing stored, therefore finds it.
- **Who:** an execution reads the Drive of the account it runs as: the person at the keyboard
  for a menu action, and — for automatic updates — the person who switched them on.
- **One resolver**, `resolveReviewerApiToken_()`, is the only place that reads the report's
  token; every reader goes through it. It is asked once per execution and only when the
  Reviewer is asked: with a token in the report, Drive is never queried; an update with
  abstracts off queries nothing.
- **CENTRAL and Legacy** (no template release file): the token of the project and nothing else,
  as before; the token part of their Configure Meeting dialog is unchanged.

### 2.2 What is trusted, and what happens otherwise

A file is used only when it is owned by the user, shared with nobody, not in the trash, of type
`application/json`, at most 2 KB, and exactly the schema with a usable token (1–512 characters,
no line break). Each of these is checked on every read, whatever the query returned.

| Situation | Result |
|---|---|
| No such file | No personal token |
| Exactly one | Its token is used |
| More than one usable file | **None is used**; Configure Meeting says so |
| A candidate that cannot be read | No personal token (it might be a second usable file) |
| A marked file that is shared, or has other content | Not used, and never changed by a lookup |
| Drive cannot be reached | No personal token; the report works without |

- A stored token is never sent to the browser, never logged, and never written into the
  document or its properties. The text of a JSON parse error is not logged either.
- Nothing changes who a file is shared with. A shared settings file is left exactly as it is.
- A write is reported as done only after the file was read back by its id and found private
  and complete.

### 2.3 Configure Meeting

In a template report the dialog says which token is in use (none, personal, the report's, or
the report's with the personal one unused) and, where it applies, why the personal store is
not used. Controls:

- a new token, stored **for all my reports** (default) or **in this report only**;
- remove the token of this report; remove my personal token;
- save the token of this report as my personal token (offered only when the report has one);
  done on the server, without the token passing through the browser.

When a personal token is written and the report's token is to be removed in the same Save, the
order is: write the file, read it back, and only then remove the report's token. If the write
is not confirmed, the report keeps its token.

## 3. Compatibility

- `appsscript.json` is unchanged: the same six OAuth scopes, Drive v3 and Docs v1.
- New Document Properties: `MEETING_START_TIME`, `MEETING_TIME_ZONE`, `MEETING_START_BASIS`,
  and `SA4_REVISIONS_URL_CANDIDATE`. None is read by a background run.
- `getReportConfig_()` no longer looks the token up for every configuration.
  `REVIEWER_API_TOKEN` can still be read from the configuration by name; it is no longer part
  of a copy or a print of it.
- The setup information of a new report has two more fields (`meetingStartTime`,
  `meetingTimeZone`). A report always carries the code of the release it was created with, so
  creator and first run agree.

## 4. Tests

On the release commit: **103 test files, 6,932 passing checks, 0 failures.**

| Suite | Checks | Covers |
|---|---|---|
| `meeting-start-opening` | 52 | §1.1: the Portal values, the sentence, creation, the build |
| `configure-meeting-setup-parity` | 76 | §1.1 and §1.2 through Configure Meeting: the same start as at creation; meeting A → B; the retry |
| `adhoc-revisions-folder-setup` | 58 | §1.2 at Finish Report Setup |
| `withdrawn-no-abstract` | 56 | §1.3, all four paths, text and dropdown statuses |
| `build-revision-placement` | 36 | §1.4 |
| `personal-reviewer-token` | 143 | §2 |
| `user-settings-probe` | 23 | the probe of §5, against a fake Drive |
| `after-10-8-ledger` | 30 | every function changed since T-2026.10.8 is the released one again once the named lines are taken out |

Mutation checks (single changes to the new code, each run against the behavioural suites):
109 applied, 109 caught — 61 for Batch 1 (abstracts, revisions, opening, drafts folder,
Configure Meeting) and 48 for Batch 2 (precedence, owner, sharing, several files, schema, the
order of write / read back / remove, laziness, the dialog).

All of this runs against fakes of the document, the Docs API and Drive.

## 5. Smoke test on real Google services: passed

2026-10-06, on the disposable test report (the same one as for T-2026.10.8), with the candidate
of commit `2e38b0b` in its script project. The smoke was limited to what the fakes cannot
show for Batch 2: real Drive, a second bound script project, and a real installed trigger.

It used the committed probe `template/probe/UserSettingsProbe.gs` and a temporary helper file
with three actions, pushed to that project only. Both call the production storage functions;
the token was a harmless sentinel, and nothing shown or logged contained it (only a status,
counts and whether a SHA-256 matched).

| | Result |
|---|---|
| A. The test report creates the private marked settings file and finds it again | **Pass** |
| B. A different bound script project — a copy of the test report made with `makeCopy()` — finds and reads the same file | **Pass** |
| C. A real one-shot time-driven trigger installed in that copy finds and reads it | **Pass** |
| Cleanup: the trigger, the sentinel file (moved to the trash), the results, the copy (moved to the trash) | **Pass** |

**What this evidence is.** A, B, C and the cleanup are the results of the helper's own checks,
as shown in its dialogs and execution log and reported by the person who ran the three
functions in the script editor. There was no independent inspection of Drive or of the trigger
list from outside.

Afterwards the test report was returned to exactly the seven candidate files: a push that
differed in one file (so that the two extra files were removed), the push of the exact
candidate, and a fresh read-back — seven files, each matching the candidate, no probe or
helper file. The helper file was never committed and was deleted locally.

### Not exercised live

- **Batch 1.** Nothing of §1 was run on real services; it rests on the automated tests.
- **The Configure Meeting controls for the personal token**, and a real abstract request made
  with a personal token. The smoke exercised the storage and discovery functions through the
  helper, not the dialog and not the Reviewer.

### Two things learned

- A copy of a document has a script project with the same functions as the original. A
  function meant for the original, run in the editor of the copy, was refused by its guard —
  correctly, but with a message that did not say which document it was in. The helper was
  changed to say so.
- Running a function of a document-bound script remotely is still not possible with the
  existing tooling; each run was started by hand in the script editor.

## 6. Release: complete

| Step | Result |
|---|---|
| Release commit | **Complete.** One squashed commit on master, `0eb8bedcc04a3552a16bd5af3e6e7463d6624a4c`, with `a960718e5c958f07c2f7a7b72db72e72f5a0d178` (T-2026.10.8 and its closeout) as its only parent and the tree of the accepted candidate `2e38b0bd66c7df87344d0b8c356f5ea8138e2276`. The candidate's commits (`e8660b1`, the ad-hoc automation; `bd25441`, the personal token; `2e38b0b`, the release metadata) are named in its message and were not pushed |
| Automated tests on the release commit | **Pass**: 103 files, 6,932 checks, 0 failures; the release-integrity suites pass |
| Tag and publication | **Complete.** The lightweight tag `template-release/T-2026.10.9` on the release commit; master and the tag pushed together in one atomic push, fast-forward, without force. Every older release tag is where it was, locally and published |
| The release bundle | Built by the release tool from the tagged commit: the seven production files. Its six source files are byte-identical to the candidate of the smoke test; `Release.js` differs only in the commit it names and the time it was built. No probe, helper or test file |
| Deployment to the master template's script project | **Pass** (below) |

### Deployment to the master template

| | |
|---|---|
| Before | The project held T-2026.10.8, `Code.js` 2.20.0: seven files, each matching that release's bundle, no other file |
| The push | One normal `clasp push` from the release bundle. Seven files pushed. No prompt (the manifest is the one of T-2026.10.8). No `--force` |
| Read-back | **Exactly seven files, each matching the release bundle** (line ends normalised; the manifest compared as JSON) |
| `Release.js` as deployed | T-2026.10.9, template, 2.21.0, the release commit, the tag |
| The manifest as deployed | Unchanged: Europe/Berlin, V8, Drive v3 and Docs v1, the same six OAuth scopes |

The deployment replaced the files of the script project and nothing else: the master template
**document** was not modified, no report was built, no trigger was changed.

### What is live

- **The master template: T-2026.10.9.** A report created from it has everything of §1 and §2.
- **The live 6G and MBS reports: not on T-2026.10.9.** They were not read or written.
- **CENTRAL and Legacy** were not deployed and do not use the personal token.

## 7. Follow-ups

Nothing below was changed by this release.

- **Batch 1 awaits its first natural use**: the first new ad-hoc report created from the
  template is the first real run of the opening sentence and of the drafts-folder request.
- **The two Colab pages** (`colab_notebook.html`, `colab_notebook_shared.html`) are part of
  the bundle and are not referenced by the script code. They are saved copies of the Google
  Colab page and contain that page's own browser keys; they are unchanged since T-2026.10.8.
  Whether they are still needed is to be looked at.
  *(Looked at after T-2026.10.11: not needed, and removed from later releases. See
  `PRODUCTION_FILE_SET.md`.)*
- **Attendance**: to be looked at separately, and extended if required.
- **Adoption by a live report** (6G, MBS) is a separate, explicitly authorized step for each.
- **clasp account**: in this environment every `clasp` command must name the account
  explicitly (`--user haudiobe.ts@gmail.com`); the default clasp login is another account.
