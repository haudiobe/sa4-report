# SA4 Report Template: Architecture and One-Click New Report Creation

Design study, 2026-09-30 (overnight). Branch `design/report-template-architecture`, based on
master `d7c23bc` (CENTRAL 2.15.2). Nothing in production was changed: no clasp push, no Google
file created or modified, no trigger, deployment, version or property touched.

**Status, 2026-10-01: TEMPLATE-001 is LIVE VERIFIED / GO** (§19.6). Thomas ran the copy probe
in scratch Google Docs. Programmatic template copying works, so the architecture below is the
recommended implementation path. The next stage is TEMPLATE-002A (§21).

Evidence labels used throughout:

| Label | Meaning |
|---|---|
| **DOCUMENTED** | Stated in current Google developer documentation (URL given). |
| **VERIFIED (repo)** | Read directly in this repository or its sibling checkouts tonight. |
| **PROJECT HISTORY** | Recorded as observed live in this project's changelog/LEGACY_UPGRADES.md. |
| **COMMUNITY** | Stated by experienced community sources, not by Google docs. |
| **INFERRED** | My conclusion from the above; not confirmed. |
| **LIVE VERIFIED** | Observed by Thomas in real Google Docs with the TEMPLATE-001 probe on 2026-10-01 (§19.6). |
| **UNKNOWN** | Not established. |

---

## 1. Executive summary

**Recommendation: one master "SA4 Report Template" Google Doc whose bound script is a tagged
release of today's `Code.js`, plus a small `ReportCreator.js`. New reports are copies of the
template; the creator lives inside the template itself.** No separate creator project, no
add-on test deployment, no central scheduler for new reports.

Why it is the simplest option that works:

- `Code.js` already runs as a bound script with no change. With no execution context, every
  document and state access falls back to `DocumentApp.getActiveDocument()` and Document
  Properties, which is exactly what Legacy does (VERIFIED (repo): `getReportDocument_`,
  `getReportStateStore_`, `Code.js:1881-2042`). The CENTRAL-only parts (registry, scheduler,
  add-on menu, hourly-only intervals) sit alongside it and can be switched off.
- Google documents that copying a container file copies its bound script (DOCUMENTED). So a
  copy of the template works on its own, like the 6G Legacy report does today, but every copy
  starts from a known, tagged release.
- Bound scripts can use 15- and 30-minute triggers again. Add-ons cannot (DOCUMENTED).
- Old reports never change unless someone deliberately pushes to that one report.

**Human effort for a new report:** type the meeting ID, then *Create*, *Open*, and
*Finish Setup & Build*. That is 3 meaningful actions after choosing the meeting, plus Google's
one-time consent screens for that report. A main meeting, or an ad-hoc meeting whose family
is ambiguous, adds one choice (the report family).

**Go/no-go gate: passed.** The one open question was whether a *programmatic* copy
(`DriveApp.makeCopy`) keeps the bound script. The TEMPLATE-001 probe confirmed it live on
2026-10-01: the copy has its own document ID and its own script project, inherits no
properties and no triggers, can use its own Document Properties, can read the description the
creator wrote, and can create its own 30-minute trigger (§19.6).

**Decision (2026-10-01).** One master Google Doc, "SA4 Report Template", holds the bound SA4
report script. The primary new-report workflow is:

1. Open SA4 Report Template.
2. Choose "Create New SA4 Report".
3. Enter the meeting ID.
4. The template creates the new Google Doc with `DriveApp.makeCopy()`.
5. Open the new report.
6. Complete meeting setup / validation.
7. Build the initial report.
8. Optionally enable the normal bound-script continuous-update trigger.

Each report is independent and pinned to the script version the template held when it was
created. New reports need no CENTRAL add-on test deployment. Manual *File → Make a copy* is no
longer the primary workflow; it remains a fallback only. Neither the Advanced Drive service
nor the Apps Script API is needed to copy the bound script.

A local, tested prototype exists (§12, §16). It is not deployed.

---

## 2. Current-state architecture

### 2.1 Three codebases, two live Apps Script projects

| Checkout | Script ID (`.clasp.json`) | What it is |
|---|---|---|
| `sa4-report-legacy/` (master `b95e06e` on 2026-10-01; `e564f70` when this study started) | `1jZOi…oznXA` | **Legacy**: 2.12.0 plus LEGACY-UPGRADE-002…006 and, since 2026-10-01, the integrated reallocation + ad-hoc e-mail release (deployed and live accepted). The script bound to *"6G Media Minutes – SA4-e (AH) on FS_6G_MED"* (86178). |
| `sa4-report/` (master `d7c23bc`) | **`1jZOi…oznXA` (Legacy!)** | **CENTRAL source**: 2.15.2, 14,579 lines, 62 test files. `.clasp.json` is gitignored and points at the *Legacy* project. |
| `sa4-report-central-addon/` (no git) | `13Dpx…zzKc` | Push directory for **CENTRAL**; `rootDir: ..\sa4-report`, so it pushes whatever branch `sa4-report` currently has checked out. |

VERIFIED (repo). Two of these three are deployment hazards (§10.6):

- Running `clasp push` in `sa4-report/` would overwrite the Legacy production script with CENTRAL code.
- CENTRAL's push content depends on which branch happens to be checked out in `sa4-report/`.

### 2.2 How Code.js runs in each mode

| Concern | Bound (Legacy, or CENTRAL code run bound) | CENTRAL add-on |
|---|---|---|
| Document | `DocumentApp.getActiveDocument()`; in a bound time trigger this is the container (PROJECT HISTORY: Legacy's `continuousUpdate` trigger) | Interactive: active doc. Background: `DocumentApp.openById(id)` through the context. |
| State | Document Properties | Adopted documents: Script Properties `SA4_STATE\|<docId>\|<key>`; otherwise Document Properties |
| Registry | none | `SA4_REGISTRY_INDEX` / `SA4_REGISTRY_DOC\|<docId>` |
| Automatic updates | `ScriptApp.newTrigger('continuousUpdate')`, one per report project | One hourly `runAddonSchedulerTrigger` that serves all registered documents |
| Intervals | Legacy offers 15/30/60 via `everyMinutes(n)`, so its "60" is broken (`everyMinutes(60)` is invalid, the 2.15.1 bug) | 2.15.2: hourly only |
| Menu | `⚠️Scripts⚠️` | Same `⚠️Scripts⚠️`, plus `☁️ CENTRAL ADD-ON (hourly)` |
| Concurrency | `LockService.getDocumentLock()` | Scheduler: `ScriptLock` |
| Reviewer token | Script Property `REVIEWER_API_TOKEN` of that project | Script Property of CENTRAL, shared by every document |

### 2.3 The configuration pipeline

The *Configure Meeting Settings* dialog (`configureMeetingSettings`, `Code.js:7023`) already
does almost everything a creator needs. Every step is server-side and reusable:

1. `parseMeetingIdInput_` → `resolveMeetingCoreById_` (Portal GetMeetings: type AH/OR,
   name, dates, `MtgDocURL`, then FTP base and a revisions candidate).
2. `enrichMeetingFromTdocList_` (TdocList.aspx: TDoc families, agenda TDoc candidates).
3. `applyAdhocSourceDiscovery_` → `discoverAdhocMeetingSources_` (ADDON-008A: the Portal
   meeting-ID document list as the TDoc list, and the series `agenda.csv` fallback, validated
   against this meeting's own TDocs).
4. `computeResolvedMeetingPreview_` (merge with existing values; `{value, source}` per field),
   including `inferReportFamily_` (ADDON-007B2: confident / suggested / conflict / unresolved).
5. `evaluateMeetingReadiness_` + `MEETING_READINESS_RULES_` (ADDON-007B3: one rule set for
   the dialog, the save message and the build guard `assertMeetingReadyToBuild_`).
6. `saveConfigurationSettings` → `persistConfigurationSettings_` writes
   `CONFIG_DIALOG_MANAGED_KEYS_` to Document Properties, and mirrors them to central state
   only for registered documents.

The build (`buildSkeletonWithTdocTables`, `Code.js:9217`) checks readiness, downloads and
validates everything, **then clears the whole body** and renames the document with
`generateReportTitle_`. VERIFIED (repo). This matters for the template: nothing in the
template body survives into a built report.

### 2.4 Everything that is specific to CENTRAL

Every dependency that would behave differently in a bound template copy:

| # | Item | Where | In a template copy |
|---|---|---|---|
| C1 | `CONTINUOUS_TRIGGER_INTERVALS_` is hourly only | `Code.js:1566` | Wrong: bound scripts can use 15/30 min. Needs a per-runtime table (§11). |
| C2 | `☁️ CENTRAL ADD-ON (hourly)` submenu | `onOpen`, `Code.js:504` | Confusing; its items manage the central registry. Hide it. |
| C3 | `onInstall(e)` | `Code.js:425` | Never fires; harmless. |
| C4 | Registry + scheduler + `SA4_STATE` backend | `Code.js:1075-2400` | Unused; the registry of the copy's own project is empty, so every path falls back to Document Properties. Harmless, and deliberately left in so Code.js stays one file. |
| C5 | `saveConfigurationSettings` registry check | `Code.js:7637` | Reads the copy's own empty Script Properties, then takes the Document Properties path. Harmless. |
| C6 | `REVIEWER_API_TOKEN` in Script Properties | several | Properties are per project, so each new report starts without a token (§14). |
| C7 | `DEFAULT_MEETING_REPORT_TEMPLATE_DOC_ID_` (`1qP--…`) | `Code.js:7661` | Unrelated "template": the **preamble source** document the build copies the opening/IPR text from. It stays as is. This design calls the new document the *Report Template* to keep the two apart. |
| C8 | Test deployment / "add as test document" | process | Not needed. |
| C9 | Main-meeting save writes SA4#136 defaults when folder/number are blank | `persistConfigurationSettings_` | A creator must derive the folder and number from the FTP base (done in the prototype). This is a latent CENTRAL gap too. |

Nothing else in the call graph assumes add-on mode. The e-mail collector, the exporter
(document links are resolved from the active document's ID), reallocations and the revisions
collector are all document-relative. VERIFIED (repo).

---

## 3. Verified Google Docs / Apps Script copy semantics

| Question | Answer | Evidence |
|---|---|---|
| UI *File > Make a copy* copies the bound script? | **Yes.** "If they make a copy of the container file, they become the owner of the copy and can see and run a copy of the script." | DOCUMENTED ([bound scripts](https://developers.google.com/apps-script/guides/bound)) |
| Source files copied? | Yes: "a copy of the script". | DOCUMENTED |
| Programmatic copy (`DriveApp.File.makeCopy`) copies it? | **Yes.** The copy made by `makeCopy(title, folder)` had the bound script and a working menu. (Drive `files.copy` was not tested and is not used.) | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01); previously COMMUNITY |
| New project / new Script ID for the copy? | **Yes.** Original `REDACTED-PROBE-SCRIPT-ID`, programmatic copy `REDACTED-COPY-SCRIPT-ID`; the document ID differs too. | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01) |
| Document Properties copied? | **No.** The original's marker was missing in the copy, and the copy could write and read its own. | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01); Google: "Properties are never shared between scripts" ([properties](https://developers.google.com/apps-script/guides/properties)) |
| Script Properties copied? | **No.** Marker missing in the copy. | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01) |
| User Properties copied? | **No.** Marker missing in the copy (they are per user per script). | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01) |
| Installable triggers copied? | **No.** The original owned a 30-minute trigger; the copy inherited none. | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01) |
| Authorizations copied? | No. Authorization is per project, so the copy asks for consent on first use. | INFERRED from DOCUMENTED per-project authorization ([authorization](https://developers.google.com/apps-script/guides/services/authorization)); the per-copy prompt was not part of the reported TEMPLATE-001 result |
| Advanced service settings / OAuth scopes preserved? | Yes. They live in `appsscript.json`, a project file, which is copied with the source. | INFERRED. Not part of TEMPLATE-001: copying does not use the Advanced Drive service (§19); the first real template report exercises it (TEMPLATE-006). |
| `onOpen` works in the copy? | Yes, before any authorization: a simple trigger "runs automatically whenever a file is opened by a user who has edit access". | DOCUMENTED; the menu worked in the copy (LIVE VERIFIED). Whether it showed before the permission prompt was not reported. |
| Deployment/version metadata? | Irrelevant. Bound reports are not deployed; the release is recorded in `Release.js` and copied as source. | Design choice |
| Drive description readable by the copy's script? | **Yes.** The text the creator wrote after copying was readable from the copy. | **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01) |
| Is the description itself copied by `makeCopy`? | Not relied upon, and not reported in the TEMPLATE-001 result. The creator writes the description *after* copying, and the payload carries the target document ID. | UNKNOWN (harmless) |
| Can a creator write the copy's Document Properties? | **No.** Properties belong to a script project; the template's project cannot write the copy's store. | DOCUMENTED ("never shared between scripts") |
| Apps Script API alternative (`projects.create` with `parentId`) | Exists, but is **not needed**: `DriveApp.makeCopy` carries the bound script. | DOCUMENTED ([projects.create](https://developers.google.com/apps-script/api/reference/rest/v1/projects/create)); superseded by TEMPLATE-001 |
| Time-driven intervals | `everyMinutes(n)`: "n must be 1, 5, 10, 15 or 30"; `everyHours(n)`. A copied bound script created its own `everyMinutes(30)` trigger. | DOCUMENTED ([ClockTriggerBuilder](https://developers.google.com/apps-script/reference/script/clock-trigger-builder)); creation **LIVE VERIFIED** (TEMPLATE-001, 2026-10-01). The first execution was not waited for (§19.6). |
| Add-on limit | "An add-on can use a time-driven trigger once per hour at most." Bound scripts are not add-ons. | DOCUMENTED ([installable triggers](https://developers.google.com/apps-script/guides/triggers/installable)); PROJECT HISTORY: 2.15.1 failed live with 30 min in CENTRAL |
| Trigger identity | "Installable triggers always run under the account of the person who created them." | DOCUMENTED |
| Quotas (consumer account) | **Triggers total runtime 90 min/day** (across all scripts of the user); 6 min per execution; 20 triggers per user per script; 250 documents created per day. | DOCUMENTED ([quotas](https://developers.google.com/apps-script/guides/services/quotas)) |
| Who can run the bound script? | Editors of the document; viewers can see the code. | DOCUMENTED |

**Design rule that follows:** nothing depends on anything being copied except the script source
and manifest. The template never holds configuration, properties or triggers, and every piece
of per-report state is created by the report's own first run.

---

## 4. Architecture options considered

| | A. CENTRAL add-on (today) | B. Bound template, manual copy | C. Standalone creator + bound template | **D. Template-hosted creator + bound runtime (recommended)** | E2. Creator pushes a script via Apps Script API |
|---|---|---|---|---|---|
| New-report steps | Create doc, add as test document in the test deployment, reopen, Configure, Adopt/Enable, Build | Make a copy, rename, open, Configure (Resolve, Discover, Save), Build | Open creator web app, ID, Create, open, Finish, Build | Open template, ID, **Create**, **Open**, **Finish Setup & Build** | Same as D |
| Authorization | Once for CENTRAL, but the test-deployment dance per doc | Consent once per report | Consent for the creator once, plus per report | Consent once per report (and once for the template) | Creator needs `script.projects`, the Apps Script API user toggle, then consent per report |
| Portability across computers | Poor (observed: add-on visibility differs per session/computer) | Good: menu travels with the doc | Good | Good | Good |
| Version stability | All reports change at every CENTRAL push | Pinned per report | Pinned per report | Pinned per report | Pinned per report |
| Updating an old report | Automatic (a risk as much as a feature) | Manual per report | Manual | Manual, deliberate (§10.5) | Could be scripted |
| Accidental updates | High: one push hits every report | None | None | None | Low |
| 15/30-min triggers | **No** (hourly) | Yes | Yes | Yes | Yes |
| Extra moving parts | Scheduler, registry, central state, test deployment | None | Separate project and web-app deployment | One extra file | Source bundle hosting, API enablement, more scopes |
| Maintenance | One code, one complex runtime | Copies diverge if hand-edited | Two projects | One code, one release step | Most code |
| Existing reports | Untouched | Untouched | Untouched | Untouched | Untouched |

Notes that decided it:

- **A** already has a working codebase, but the problems Thomas reported (test deployment per
  document, add-on visibility, hourly limit, two confusable menus) come from the add-on
  platform itself. They cannot be engineered away while it stays an add-on on a consumer
  account (ADDON-006: private add-ons need Workspace).
- **B** is what Legacy does today, and it is reliable. Its weaknesses are an unknown code
  version per copy and repeating Configure by hand. D keeps B's runtime and removes both.
- **C** adds a second Apps Script project, a web-app deployment and a version lifecycle
  (deployments are exactly the kind of setup to avoid). D gets the same result by putting the
  creator in the template, whose code is the release anyway.
- **E2** avoids relying on copy semantics, but it needs the user-level Apps Script API toggle,
  the `script.projects` scope and a place to host source bundles. TEMPLATE-001 showed it is
  not needed.
- **E3** (considered, rejected): a library that every report references. It keeps one copy
  of the code, but library versions need deployments, and "HEAD" mode would silently update
  every report. That is the CENTRAL risk again.

---

## 5. Recommended architecture

```
Git: sa4-report master ── tag template-release/T-2026.10.0
          │  tools/template-release.js (guards, bundle, manifest)
          ▼
dist/template-release/T-2026.10.0/   Code.js  HyperLink.js  *.html  appsscript.json
          │                          ReportCreator.js  Release.js (generated)
          │  clasp push  (Thomas, from this directory only)
          ▼
"SA4 Report Template" (Google Doc) ── bound script = release T-2026.10.0
          │  ⚠️Scripts⚠️ > 🆕 Create New SA4 Report
          │  DriveApp.makeCopy(title) + Drive description = bootstrap payload
          ▼
"MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e" ── its own bound script (copy of T-2026.10.0)
          │  ⚠️Scripts⚠️ > 🚀 Finish Setup & Build   (first run: consent, config → Document Properties)
          ▼
Normal Legacy-style bound report: Build, Continuous Update trigger (15/30/60), e-mail collector, exporter
```

- **Canonical source:** `sa4-report` master. One `Code.js` serves CENTRAL (for as long as it
  lives) and the template; there is never a template-specific copy of `Code.js`. Legacy is a
  separate lineage that is still receiving accepted fixes (a release went live on 2026-10-01),
  so everything accepted in Legacy has to be present in `sa4-report` master before the first
  template release. That is the TEMPLATE-002A gate (§21).
- **Runtime flavor** is set by the presence of `Release.js` (`SA4_RELEASE_`), which only
  template bundles contain. `Code.js` checks it at call time with
  `typeof SA4_RELEASE_ !== 'undefined'`, so file load order does not matter. CENTRAL and
  Legacy behavior is unchanged.
- **Role** is set by document ID: `templateDocumentRole_(docId, release)`. The template (ID
  equals `release.templateDocumentId`) shows the creator and refuses report operations. Every
  other document is a report. This is independent of how properties behave on copy.

---

## 6. Why this is simpler than today's workflow

| Today (CENTRAL) | Recommended |
|---|---|
| Create an empty doc | Open the bookmarked template |
| Apps Script: add the doc as a test document of the test deployment, *Execute* | none |
| Reload; hope the add-on menu shows on this computer/session | Menu is part of the document |
| Configure Meeting: ID, Resolve, Discover, check family, Save | Type ID; everything is discovered; *Create* |
| Adopt / Enable automatic updates (hourly only) | *Finish Setup & Build*; optionally start a 30-min update in the same dialog |
| Two `⚠️Scripts⚠️`-like menus when a Legacy doc is also open | One menu, no CENTRAL submenu in template reports |
| Every CENTRAL push changes every report | A report changes only if someone pushes to that report |

---

## 7. Exact new-report user journey

### 7.1 Screens

**Template document → ⚠️Scripts⚠️ → 🆕 Create New SA4 Report**

```
Create New SA4 Report                                   Template T-2026.10.0
Meeting ID  [ 86172 ]  (Look up)

86172 · SA4-e (AH) MBS SWG post 137-e · Online · 1 Oct – 5 Nov 2026
  Report family   MBS            ✓ detected (meeting name + TDoc series + FTP path)
  Mailing list    3GPP_TSG_SA_WG4_MBS   (derived from family)
  Document folder …/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/
  TDoc list       Portal document list (10 TDocs)
  Agenda          agenda.csv fallback, 3 of 53 items, sections 2, 3
  Revisions       …/SA4_MBS/Inbox/Drafts   (candidate)

  Title: MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e

                                         [ Create Report ]
```

After *Create*: "✅ Created. [Open report]" (new tab). **In the new report:** a placeholder page
(§8) and the normal `⚠️Scripts⚠️` menu with `🚀 Finish Setup & Build` at the top. Clicking it
shows Google's consent screens for this report once, then the same item is clicked again (the
usual Apps Script behavior after first consent: INFERRED; not part of the reported TEMPLATE-001 result). The
setup dialog then shows:

```
✅ Meeting configuration stored (template T-2026.10.0).
✅ Ready to build.
[ Build report now ]   ☐ Start automatic updates every [30 min ▾]
```

*Build report now* calls `runFullReportBuild` as its own execution (with its own 6-minute budget).

### 7.2 Per meeting type

| Case | What the creator discovers | Questions for Thomas | Verified by |
|---|---|---|---|
| Main SA4 meeting (Portal type `OR`) | Folder + number from `MtgDocURL` (`TSGS4_137_Xian` → 137); TDoc list derived by the runtime; family never inferred | **Report family** (one report per family) | `template-bootstrap.test.js` (main) |
| Normal ad-hoc | Type AH, name, FTP series, Portal document list, agenda TDoc or agenda.csv | none if the family is confident/suggested | tests (86172, 85916) |
| Audio ad-hoc, agenda TDoc | `S4aA` series → Audio; agenda TDoc `S4aA260090` wins over the 0-byte CSV | none | 85916 test |
| Video ad-hoc | `S4aV` → Video | none | same inference code (ADDON-007B2 tests) |
| MBS ad-hoc, agenda.csv | `S4aI` + name → MBS confident; CSV validated against this meeting's TDocs | none | 86172 test |
| RTC ad-hoc | `A4aR` → RTC | none | ADDON-007B2 tests |
| 6G / plenary-style ad-hoc | `S4aP` never decides; `FS_6G_MED` in the name → 6G (suggested) | none; with no 6G token in the name, the family question | 6G test |
| Ambiguous / conflicting family | Nothing is chosen (`conflict` / `unresolved`) | **Report family** | conflict test |
| Missing TDoc list / agenda (created before the agenda exists) | Report **is created**; the items are listed as *pending* | none now; later *Configure Meeting → Discover* | pending test |
| Missing FTP path | `FTP_BASE_REQUIRED` is pending | Advanced field in Configure Meeting later | readiness rules |
| Unusual future meeting (unknown portal type, host) | Refused if the type is unknown; non-3GPP URLs refused | Use *Configure Meeting* manually in the created report | guard tests |

Creation is refused only when the report cannot be identified or named: an unresolved
meeting, an unknown type, a missing family, or a non-3GPP URL. Everything else becomes
*pending* and is shown again by the real build guard.

---

## 8. Template structure

**Document content**: one short, removable placeholder page. The build clears it anyway.

```
SA4 Report (setup pending)
This document was created from the SA4 Report Template.
Use ⚠️Scripts⚠️ > 🚀 Finish Setup & Build. If the menu is missing, reload the page.
```

- No cover page, metadata block, reallocation table or configuration marker in the body.
  `ensureReallocationTable_` and the preamble already create what the minutes need; anything
  else would be clutter or would be wiped by the build.
- **Where things live:**

| Item | Location |
|---|---|
| Script release (ID, commit, tag, template IDs) | `Release.js` in the bound project (copied with the code) |
| "Created from" provenance | Document Properties `SA4_CREATED_FROM_RELEASE`, `SA4_CREATED_AT`, `SA4_SETUP_RELEASE`, `SA4_BOUND_SCRIPT_ID`, `SA4_BOOTSTRAP_STATE`; plus a one-line Drive description written at first run |
| Meeting configuration | Ordinary Document Properties (`CONFIG_DIALOG_MANAGED_KEYS_`), exactly as Legacy and the Configure dialog |
| Schema version | `sa4-report-bootstrap/1` inside the payload |
| Readiness | Computed live (`getBuildReadiness_`), never stored |

- **Menu (TEMPLATE-002)**: in the template, `⚠️Scripts⚠️ > 🆕 Create New SA4 Report`,
  `ℹ️ Template release info` and `🧪 Test All Connections`; report operations are hidden and
  guarded. In a report, the existing menu without `☁️ CENTRAL ADD-ON`, plus
  `🚀 Finish Setup & Build` and `ℹ️ About this report` (release, script ID, created-from).
  `onOpen` can choose between the two because a simple trigger may read the active document's
  ID and Document Properties (DOCUMENTED: simple triggers can access the bound file).
- **Template hygiene**: the template never has configuration, triggers or a token. Every
  mutating entry point calls `assertNotTemplateDocument_` (prototype).

---

## 9. Bootstrap mechanism

Options compared:

| Option | Verdict |
|---|---|
| A. Creator writes the copy's Document Properties | **Impossible.** A different script project cannot reach that store. |
| B. First-run wizard re-discovers from scratch | Works, but it repeats the slow network calls (TdocList.aspx was measured at 90+ s, PROD-016) and asks Thomas again. Kept as the fallback: *Configure Meeting* still works in every report. |
| C. Payload in the document body | Visible, editable by collaborators, and has to be removed from the minutes. Rejected. |
| **D. Payload in the Drive file description** | **Chosen.** Plain Drive metadata, readable and writable by both projects with the existing `drive` scope, invisible in the document, visible under Drive "Details" (useful provenance). Replaced by a one-line provenance note after first run. |
| D'. Drive `properties` (advanced service) | Hidden and key-value, but key+value is limited to 124 bytes, which is too small for URLs. Rejected. |
| E. Derive from title / meeting ID only | Title does not carry the ID; loses Thomas's family choice. Rejected. |

**Payload** (`sa4-report-bootstrap/1`, well under 4,000 characters):

```json
{ "schema": "sa4-report-bootstrap/1", "createdAt": "…", "templateRelease": "T-2026.10.0",
  "templateDocumentId": "…", "targetDocumentId": "<the copy's id>",
  "config": { "meetingId": "86172", "meetingType": "adhoc", "meetingName": "…", "meetingDate": "…",
              "ftpBase": "…", "meetingFolder": "", "meetingNumber": "", "reportType": "MBS",
              "agendaTdoc": "", "agendaCsvUrl": "…", "tdocUrl": "…", "revisionsUrl": "",
              "mailingList": "", "mailingListMode": "derived" },
  "pending": [] }
```

`config` uses the Configure dialog's own field names, so the first run stores it through
**the same `persistConfigurationSettings_()`** a manual *Save* uses. That keeps the existing
protections: skip-if-blank, agenda.csv acceptance only for this series, derived vs. override
mailing list, and no SA4#136 defaults for ad-hoc.

**Creator sequence** (`createReportFromTemplateWith_`): validate the release, check the role is
the template, rebuild the payload server-side from the resolved meeting, validate it, then
`makeCopy(title, sameFolder)`. After that it writes the description with `targetDocumentId`
set. If writing the description fails, the new copy goes to the Drive trash (recoverable),
so Thomas is never left with a half-prepared report.

**First run** (`finishReportSetupWith_`), idempotent:

1. Refuse in the template.
2. `SA4_BOOTSTRAP_STATE = done` means nothing to do.
3. Parse the description. If there is no payload and the document is already configured,
   it is an existing or migrated report: `not-a-new-report`, left untouched.
4. Validate: schema, **`targetDocumentId` equals this document** (so a copy of an un-set-up
   report is refused), whitelisted string keys only, 3GPP https hosts only, known family and
   type, and never a different `MEETING_ID` than one already stored.
5. `persistConfigurationSettings_(config)` (the Reviewer token is untouched: `apiTokenAction: 'keep'`).
6. Record provenance and set `SA4_BOOTSTRAP_STATE = done` last, so an interrupted run is simply repeated.
7. Replace the description with the provenance line.
8. Show readiness
   from the real `getBuildReadiness_()`.

In TEMPLATE-003, `configureMeetingSettings` and `runFullReportBuild` call the same
`ensureReportBootstrapped_()` first, so whichever menu item Thomas clicks first also finishes
the setup.

---

## 10. Script release / version model

### 10.1 Branches and tags

- Development continues on `sa4-report` feature branches, then master (as today).
- **Template production = a Git tag** `template-release/T-YYYY.MM.N` on a master commit.
  There is no long-lived template branch to keep in sync.
- Release IDs `T-2026.10.0` are deliberately distinct from `Code.js` versions (2.15.2): one
  code version can be released to the template several times (for example a manifest fix),
  and the name makes the flavor obvious in every log and property.

### 10.2 Pushing into the template

`node tools/template-release.js --release T-2026.10.0 --write` (prototype, tested) does the following:

1. Refuses unless the tree is clean and HEAD carries the tag.
2. Reads `template/template-target.json` (template Doc ID + bound Script ID, committed after
   TEMPLATE-001; not secret) and **refuses the Legacy and CENTRAL Script IDs** by value.
3. Copies only the whitelisted runtime files (`Code.js`, `HyperLink.js`, two HTML files,
   `appsscript.json`, `ReportCreator.js`). Tests, helpers and fixtures can never be included,
   and any `require(` in a bundled `.js` file is refused (the LEGACY-UPGRADE-002 incident).
   An `addOns` manifest is refused.
4. Generates `Release.js`, a `.clasp.json` for the template script only, a whitelisting
   `.claspignore`, and `RELEASE-MANIFEST.json` (sha256 per file).
5. Writes to `dist/template-release/<id>/` (gitignored) and never overwrites an existing bundle.
6. **Prints** `cd <bundle> && clasp push`. It never runs clasp.

Thomas pushes from the bundle directory, which contains only that `.clasp.json`. The Script
ID never lives in a working checkout's `.clasp.json`.

### 10.3 Verifying before release

- `for f in tests/*.test.js; do node "$f"; done` (the full suite covers the template runtime,
  because `Code.js` is byte-identical; the release test asserts that).
- After pushing: open the template, then *Template release info* (shows `Release.js`).
- The acceptance smoke test is one scratch report created from the template, a build, then
  trashing it. That is TEMPLATE-006, needed only for releases that touch creator/bootstrap code.

### 10.4 Rollback

The template holds exactly one release. To roll back, push the previous bundle (still in
`dist/`, or rebuilt from its tag). Reports already created keep whatever they were created
with, so a bad template release can only affect reports created while it was live. Their
`SA4_CREATED_FROM_RELEASE` says which ones.

### 10.5 Updating one old report deliberately

1. Find the report's Script ID: `ℹ️ About this report` shows `SA4_BOUND_SCRIPT_ID`
   (Apps Script has no API to find a document's bound project, so the report records its own).
2. Build a bundle with a *report* target:
   `--target-script <id> --target-doc <docId>` (TEMPLATE-005 extends the tool). It uses the
   same Legacy/CENTRAL denylist, and the same checks that the Script ID and document belong together.
3. Push from the bundle, open the report, confirm *About this report* shows the new release.

There is no mass-update command, by design.

### 10.6 Guardrails against past mistakes

| Past / possible mistake | Guard |
|---|---|
| Test files pushed, `onOpen` broken (LEGACY-UPGRADE-002) | Bundle whitelist + `require(` refusal |
| `sa4-report/.clasp.json` points at Legacy | Denylist in the tool; recommend renaming that local file to `.clasp.json.legacy-DO-NOT-USE` (Thomas's call; the file is not tracked) |
| CENTRAL pushes whatever branch is checked out | Recommend building CENTRAL pushes with the same bundle tool (a `central` target) |
| Wrong project confused with the right one | Release ID prefix `T-`; `Release.js` shown in the menu; the script ID is printed before the push |
| Template accidentally configured or built | `assertNotTemplateDocument_` on every mutating entry point |
| Template holds a trigger that copies might inherit | Triggers are not copied (LIVE VERIFIED); the guard in `createContinuousTrigger` keeps the template from owning one anyway |

---

## 11. Trigger model

- **Intervals.** The template runtime uses
  `TEMPLATE_CONTINUOUS_TRIGGER_INTERVALS_` = 15 min → `everyMinutes(15)`,
  30 min (default) → `everyMinutes(30)`, 60 min → `everyHours(1)`. This is the 2.15.1 table,
  which was correct for bound scripts. DOCUMENTED as valid for non-add-on scripts; PROJECT
  HISTORY: Legacy's bound 30-minute trigger has been in production use; TEMPLATE-001 verified
  live that a *copied* project can create its own 30-minute trigger. **This is a concrete advantage over CENTRAL.**
  Legacy's own "every hour" option is still broken (`everyMinutes(60)`) and is fixed only in
  new template reports. Legacy stays untouched by policy.
- **Ownership:** each report's trigger belongs to Thomas, in that report's own project, with
  handler `continuousUpdate`. The existing `DocumentLock` guards against overlaps, and
  `DocumentApp.getActiveDocument()` resolves to the report (PROJECT HISTORY: Legacy).
- **Copying:** a copy starts with no triggers (LIVE VERIFIED, TEMPLATE-001). The template can
  never create one (guard).
- **First-run setup:** optional checkbox "Start automatic updates every 30 min" in the setup
  dialog. Nothing starts without that click.
- **Cleanup:** *Manage Auto-Update Trigger → Stop*. Trashing a report stops its trigger from
  doing useful work, but the trigger still exists; stopping it first is the tidy path.
- **Quota warning (consumer account):** all triggers of all scripts share **90 minutes of
  runtime per day**. A 40-second update every 15 minutes is about 64 min/day for *one*
  report. Recommendations: default to 30 min; use 15 min only on meeting days; in TEMPLATE-004,
  auto-stop the trigger two days after the meeting end date (the payload already carries the
  dates). The existing `[PERF]` summary line gives the real per-run cost.

---

## 12. CENTRAL's future role

Decision: **C + B, then A.** CENTRAL stays alive, unchanged, for the reports already adopted
there (for example 86172), until those meetings end. It creates no new reports. Its code keeps
being the canonical `Code.js`, because it *is* the template runtime, so it remains the
place where development and the test suite live. After the last CENTRAL report closes,
disable its scheduler trigger and stop pushing to it (deprecate).

Rejected roles:

- *Scheduler for template reports*: bound triggers are simpler and faster (15/30 min), and
  `openById` cannot reach Document Properties (ADDON-001B), so it would need adoption again.
- *Report creator only*: the add-on UX problems are exactly the creator's front door.

The add-on-specific code (registry, scheduler, central state) stays in `Code.js` for now,
because removing it would mean editing tested code for no user benefit. It can be deleted
after deprecation in one reviewed change.

---

## 13. Existing-report policy

- **Legacy reports (86178 and any other hand copies):** untouched. No push to `1jZOi…`.
- **CENTRAL reports:** untouched; they keep hourly updates until their meetings end.
- **New reports:** created from the template once TEMPLATE-001 through 003 are accepted.
- **No mass migration.** The first run explicitly leaves documents without a payload alone.
- **Migrating one report later, if it is ever worth it** (for example a long-running series
  document). Neither option below needs code:
  1. Preferred: create a fresh report from the template for the next meeting.
  2. In place: create a new report from the template for the same meeting, then copy the
     manual content (minutes text, reallocations) across. Automated in-place migration of a
     CENTRAL document would have to read central `SA4_STATE`, which is another project's
     Script Properties, so it would need an export step in CENTRAL. Not proposed.

---

## 14. Security / authorization considerations

- **Scopes** are unchanged from `appsscript.json`: documents, drive, spreadsheets,
  external_request, scriptapp, container.ui. Consent is shown once per new report (per project).
  An "unverified app" screen is likely, as with Legacy today (INFERRED; not part of the reported TEMPLATE-001 result).
- **Bound code is visible to document editors, and editors can run it.** This is the same as
  Legacy today. Script Properties (the Reviewer token) are visible to editors through the
  project settings. Recommendation: share reports with delegates as *commenter/viewer*, or accept the
  exposure knowingly. Do not put the token in the template (it would not be copied anyway)
  and never in the payload.
- **Reviewer token per report:** since properties are not copied, a new report has no token.
  Abstracts are optional (they are off by default during updates). Offer "Reviewer API token
  (optional)" in the setup dialog (TEMPLATE-003). **Open question 1** covers a single private
  Drive token file instead.
- **Payload is untrusted input:** it is description text that editors can change. Hence strict
  schema/key/host validation and the `targetDocumentId` binding; it can only ever feed the
  same values the Configure dialog accepts.
- **Creator writes only** a new copy in the template's own folder, and its description.

---

## 15. Failure / recovery behavior

| Failure | Behavior | Recovery |
|---|---|---|
| Meeting lookup fails / unknown type | Nothing copied; message | Retry, or create by hand and use Configure Meeting |
| Copy succeeds, description write fails | Copy moved to the trash; message | Retry *Create* |
| First run: no payload in a fresh doc | Refused with instructions | Configure Meeting (the manual path always works) |
| First run: payload for another doc (copy of a copy) | Refused, nothing written | Create again from the template |
| First run interrupted | `done` marker not set | Click again (idempotent) |
| Sources still missing (no agenda yet) | Report configured, not ready; build guard refuses with the same messages | Configure Meeting → Discover later |
| Menu missing in a report | Usually the session; `onOpen` needs no auth | Reload. If still missing, the Apps Script editor shows whether a bad file was pushed (the LEGACY-UPGRADE-002 pattern) |
| Bad template release | Only reports created meanwhile are affected; identifiable by `SA4_CREATED_FROM_RELEASE` | Push the previous bundle to the template; update the affected reports individually (§10.5) |
| Template document deleted | No new reports possible; existing ones unaffected | Restore from Drive trash, or create a new template doc and a new target file (new IDs) |
| Trigger runtime quota exhausted | Updates stop for the day | Lower frequency; auto-stop after the meeting |

---

## 16. Testing strategy

### 16.1 Automated (local, no Google): what exists tonight

`tests/template-bootstrap.test.js` (57 checks), `tests/template-release.test.js` (21 checks) and
`tests/template001-probe.test.js` (46 checks, the probe's own verdict logic):

| Area | Covered by |
|---|---|
| Template copy bootstrap, meeting configuration transfer | bootstrap: creator → description → *fresh sandbox* first run → real `persistConfigurationSettings_` |
| Main meeting | family question, folder/number derivation, runtime TDoc-list URL |
| MBS ad-hoc agenda.csv (86172) | real discovery pipeline on fixtures, zero questions, green readiness |
| Audio agenda TDoc (85916) | agenda TDoc stored, CSV not |
| 6G | family from `FS_6G_MED`, S4aP ignored, Legacy title |
| Ambiguous meeting | conflict → question; choice; unknown family refused |
| Missing required configuration | pending list; real build guard codes after setup |
| Build guard / template guard | `assertNotTemplateDocument_`; readiness via `getBuildReadiness_` |
| First-run behavior | idempotent; template refusal; copy-of-copy; tampered/damaged payload; existing report untouched; different meeting refused |
| Old-report isolation | payload-less configured docs untouched; release tool denies Legacy/CENTRAL IDs; CENTRAL `.claspignore` cannot pick up `template/` |
| Continuous trigger intervals | template table valid for `ClockTriggerBuilder`; CENTRAL table unchanged without `Release.js` |
| Release-version recording | `Release.js` content; provenance properties |
| Code.js unchanged | bundle has byte-identical `Code.js`; no name collisions; same outputs with `ReportCreator.js` loaded |

Reallocation persistence, the e-mail collector and exporter, and relative Docs links are
runtime behavior of the **unchanged** `Code.js`. They are covered by the existing suites
(`addon008a1*`, `addon008a2`, `addon009*`, `email-export-*`), which run against exactly the
code the template bundles. Full suite on 2026-10-01: **65 test files, 0 failures.**

### 16.2 Still to write (TEMPLATE-002/003)

- The flavor-aware `onOpen` menu spec (template vs report, no CENTRAL submenu).
- `createContinuousTrigger` with `Release.js` present: a fake bound `ScriptApp` accepting
  `everyMinutes(15/30)` (mirror of `continuous-trigger-interval.test.js`, which fakes the add-on rules).
- `ensureReportBootstrapped_` wired into Configure/Build.
- Creator dialog client rendering (like `resolve-dialog-client-rendering.test.js`).

### 16.3 What only a real Google experiment can show

The copy semantics in §3 were settled by one experiment, TEMPLATE-001 (LIVE VERIFIED / GO,
§19.6). What remains for a real run is one acceptance of the real template (TEMPLATE-006):
the creator and first-run entry points themselves, the permission prompts per report, and the
first trigger execution in a copy.

---

## 17. Staged implementation plan

| Stage | Goal | Files | Tests | Prod risk | Rollback | Thomas does |
|---|---|---|---|---|---|---|
| **TEMPLATE-001** Copy-semantics proof | **DONE: LIVE VERIFIED / GO, 2026-10-01** (§19.6) | `template/probe/Template001Probe.gs`, `tests/template001-probe.test.js` | probe verdict logic tested locally; result observed live | None: scratch docs only | n/a | Clean up the probe (§19.6) |
| **TEMPLATE-002A** Legacy parity gate | Bring every accepted Legacy behavior into `sa4-report` master, proven by Legacy's own regression suites (§21) | `Code.js`, ported `tests/legacy-*.test.js` | Legacy suites + full CENTRAL suite green | None until pushed; changes CENTRAL behavior only where Legacy's accepted fix differs | Revert commits | Review; decide the listed differences |
| **TEMPLATE-002B** Template runtime hooks | Flavor-aware intervals, menu, template guard in `Code.js` (small, call-time `typeof SA4_RELEASE_` checks) | `Code.js` (onOpen, `CONTINUOUS_TRIGGER_INTERVALS_` accessor, 3 guard calls), `template/ReportCreator.js` | new: bound trigger fake, menu spec; the full suite must stay green (CENTRAL unchanged without `Release.js`) | None until pushed; CENTRAL push is behavior-identical | Revert commit | Review |
| **TEMPLATE-003** Creator + first run UI | Creator dialog, *Finish Setup & Build* dialog, `ensureReportBootstrapped_`, optional token field | `template/ReportCreator.js`, `template/CreatorDialog.html` | client-rendering test, bootstrap tests | None (template-only file) | Revert | Review |
| **TEMPLATE-004** Trigger hygiene | 30-min default, auto-stop after meeting end + 2 days, *About this report* | `ReportCreator.js` (+ one hook in `continuousUpdate`) | trigger/clock tests | None until pushed | Revert | Review |
| **TEMPLATE-005** Release pipeline | Finalize `tools/template-release.js` (report targets, optional CENTRAL target), commit `template-target.json` | `tools/`, `template/` | release tests | None | Revert | Create the template doc; paste its two IDs |
| **TEMPLATE-006** Acceptance: first real report | Tag `template-release/T-2026.10.0`, bundle, push to the template, create one scratch report, then the first real one | none | manual checklist (§20) | Low: only the new template project and new documents | Push the previous bundle, or trash the template | `clasp push` from the bundle; create a report |

Stages 002A to 005 are local and reviewable in isolation; only 001 (done) and 006 touch Google.

---

## 18. Open questions

1. **Reviewer token.** Enter it once per report (simple, visible to editors), or keep one
   private Drive file that each report reads at run time (one entry, not stored in the report)?
   Recommendation: per report for now; decide after the first two reports.
2. **Where reports live.** Same folder as the template (prototype), or a configured
   "SA4 Reports" folder in `Release.js`?
3. **Sharing model.** Are reports shared with delegates as editors? If yes, bound code and
   Script Properties are visible to them (§14).
4. **Main-meeting TDoc list name** for e-meetings (`SA4%23137-e`): derived from the folder,
   covered only by fixtures; confirm on the next main meeting.
5. **Creator in reports too?** The prototype allows creation from the template only, so the
   newest release is always used. It could be offered from any report ("opens the template").
6. **Auto-stop date**: meeting end + 2 days, or a fixed "stop on" date in the setup dialog?
7. **CENTRAL deprecation date**: when the last CENTRAL-adopted meeting (86172, until 5 Nov 2026) ends?

---

## 19. GO / NO-GO criteria and the one experiment (TEMPLATE-001)

**Question:** can a Google Doc with a bound Apps Script be copied, by hand and by script, so
that the copy is an independent document with its own usable bound script?

The probe is `template/probe/Template001Probe.gs`. It is pasted into one scratch Doc. It
needs no clasp, no manifest edit and **no Advanced Drive service**: the production creator
copies with `DriveApp` only (`File.makeCopy(title, folder)` then `File.setDescription`), and
the probe uses that exact call. The Advanced Drive service in `appsscript.json` exists only
for converting a Word agenda (`convertWordBlobToGoogleDoc_`); it has nothing to do with copying.

### 19.1 Checklist for Thomas

1. Create a new, empty scratch Google Doc.
2. **Extensions → Apps Script**: replace everything with `Template001Probe.gs`, save.
3. Reload the Doc. A **Report** menu appears.
4. **Report → Initialize TEMPLATE-001** (allow the permission prompt, then click the item again).
5. **Report → Create Programmatic Copy**, and open the link it shows.
6. Back in the original: **File → Make a copy**.
7. In **each of the two copies**: **Report → Install 30-Minute Test Trigger**. Note whether
   Google asked for permission (it is expected to; click the item again afterwards).
8. Send back the text of the three status boxes (original and both copies). Clicking the
   text selects all of it. **Report → TEMPLATE-001 Status** shows the box again at any time.

There is no waiting step. The decision is made on "trigger created". Later, at any time
after about 35 minutes, *TEMPLATE-001 Status* in a copy also shows whether the trigger ran;
the test trigger deletes itself after its first run. When finished: **Report → Clean Up Test
State** in each of the three documents, then move them to the trash.

If a copy has **no Report menu even after a reload**, that is the result: the copy lost its
script. Report that instead of a status box (Extensions → Apps Script in the copy will show
an empty project).

### 19.2 How the probe tells the cases apart

| Marker | Where | Purpose |
|---|---|---|
| Template marker `TEMPLATE-001 ORIGINAL doc=… script=…` | document body (always copied) | Tells every copy who the original was, so document and script IDs can be compared without opening Apps Script |
| `TEMPLATE001_DOCUMENT_MARKER` | Document Properties of the original | MISSING / PRESENT-COPIED / SET IN THIS DOCUMENT in a copy |
| `TEMPLATE001_SCRIPT_MARKER` | Script Properties of the original | same |
| `TEMPLATE001_USER_MARKER` | User Properties of the original | same |
| `TEMPLATE001_WRITE_TEST` | Document Properties, written by every status check | Shows a copy can create its own Document Properties |
| One `everyMinutes(30)` trigger in the original | created by Initialize | Something a copy could inherit |
| Description `TEMPLATE001-PROGRAMMATIC-COPY target=<id>` | Drive description, written by the creator after copying | Identifies the programmatic copy; shows the bootstrap transport works |

Every stored value carries the document and script ID that wrote it, so a value copied from
the original is never mistaken for one created in the copy. The first check in a document
records which triggers existed before the probe created any there.

### 19.3 Expected status matrix

| Line in the status box | Original | Manual copy | **Programmatic copy** | If different |
|---|---|---|---|---|
| Bound script present, Report menu works | INFO | PASS | **PASS** | No menu, no box: **FAIL** (see 19.5) |
| Document ID differs from the original | INFO | PASS | **PASS** | Same ID: FAIL |
| Script ID differs from the original | INFO | PASS | **PASS** | Same project: FAIL |
| Template marker in body | INFO | INFO | INFO | |
| Document Property marker | SET IN THIS DOCUMENT | EXPECTED: MISSING | EXPECTED: MISSING | PRESENT / COPIED: WARNING |
| Script Property marker | SET IN THIS DOCUMENT | EXPECTED: MISSING | EXPECTED: MISSING | PRESENT / COPIED: WARNING |
| User Property marker | SET IN THIS DOCUMENT | EXPECTED: MISSING | EXPECTED: MISSING | PRESENT / COPIED: WARNING |
| Document Properties usable here | INFO | PASS | **PASS** | Cannot write/read: FAIL |
| Triggers inherited from the original | n/a | PASS: NONE | **PASS: NONE** | Any: FAIL |
| Drive description | INFO | INFO (copied or empty) | PASS: creator's text readable | Not readable / not found: WARNING |
| 30-minute trigger | INFO | PASS: CREATED | **PASS: CREATED** | CREATION FAILED: FAIL |
| Last trigger execution | INFO | EXPECTED: NOT YET | EXPECTED: NOT YET | Later: PASS if it ran and saw this document; WARNING if it saw another or none |
| **RESULT** | ORIGINAL, no verdict | MANUAL COPY OK | **GO** | |

Meaning of the classes: **PASS** and **FAIL** decide the result. **EXPECTED** is the
predicted, harmless outcome. **WARNING** is an outcome the design absorbs, but the bootstrap
has to be adjusted for it:

- Properties copied: the first run must clear inherited keys before storing the configuration.
  The template holds no configuration, so nothing wrong could be inherited either way.
- Description not readable in the copy: the creator cannot hand over the configuration; the
  first run asks for the meeting ID and discovers it there (§9, option B).
- Trigger ran but did not see its own document: automatic updates in copies need a closer look
  before TEMPLATE-004; creation and manual operation are unaffected.

Authorization is observed by Thomas, not by the probe: a permission prompt per copy is
EXPECTED (each copy is a new project). No prompt at all would be a pleasant surprise, not a problem.

### 19.4 The GO / NO-GO rule

**GO** if the **programmatic copy's** status box ends with `RESULT: GO` (with or without
warnings). That line is printed only when all of these hold:

1. the bound script is present and the Report menu works;
2. the document ID differs from the original;
3. the script ID differs from the original (its own project);
4. Document Properties can be written and read in the copy;
5. no installable trigger was inherited;
6. an `everyMinutes(30)` trigger was created in the copy.

**NO-GO** if the programmatic copy has no Report menu, or its box ends with `RESULT: NO-GO`.
`RESULT: INCOMPLETE` only means step 7 has not been done in that copy yet.

The manual copy's result does not decide GO. It decides which fallback applies.

### 19.5 If the programmatic copy does not keep the bound script

Not needed: the programmatic copy kept its script (§19.6). Kept for the record; manual
*File → Make a copy* remains available as a fallback workflow.

| Manual copy result | Preferred fallback | Why |
|---|---|---|
| MANUAL COPY OK | **Manual copy of the template + setup in the copy.** Thomas does *File → Make a copy* on the template; the meeting lookup and "Finish Setup & Build" run in the copy's first run instead of in the template. | It rests on behavior Google documents, needs no new scope, no Apps Script API switch and no place to host source bundles. The runtime, release model, triggers and guards are unchanged. Cost: one extra manual action, and the copy is named after the lookup instead of before. |
| Manual copy also has no script | **Apps Script API** (§4, E2): the creator makes a plain document and attaches a new bound project with `projects.create` + `updateContent`. | Only remaining route to a bound report. It needs the `script.projects` scope, the user-level Apps Script API setting and a hosted source bundle, so it is the last resort. |

### 19.6 Result: LIVE VERIFIED / GO (2026-10-01)

Thomas ran the probe in scratch Google Docs. The programmatic copy's status box ended with
`RESULT: GO`:

| Observation (programmatic copy) | Class |
|---|---|
| Bound script present; Report menu works | PASS |
| Document ID differs from the original | PASS |
| Own Apps Script project: original `REDACTED-TEMPLATE-001-PROBE-SCRIPT-ID`, copy `REDACTED-TEMPLATE-001-COPY-SCRIPT-ID` | PASS |
| Document, Script and User Property markers of the original are missing in the copy | EXPECTED |
| Document Properties usable independently in the copy | PASS |
| No triggers inherited from the original | PASS |
| Drive description written by the creator is readable from the copy | PASS |
| `everyMinutes(30)` installable trigger created in the copy (handler `template001TriggerTick`) | PASS |
| Last trigger execution: not yet | EXPECTED |

Facts established:

- `DriveApp.makeCopy()` preserved the bound Apps Script project.
- The copied document received its own document ID.
- The copied bound script received its own Script ID.
- Installable triggers were not inherited from the source document.
- The copied script can use its own Document Properties.
- The copied script can create its own `everyMinutes(30)` trigger.
- Programmatic template copying is therefore viable for the planned SA4 report architecture.
- Neither the Advanced Drive service nor the Apps Script API is required to copy the bound
  script in the tested environment.
- Manual *File → Make a copy* is no longer required as the primary workflow. It remains a fallback.

Limits of the evidence:

- **The first trigger execution was deliberately not waited for.** It is not required for the
  GO decision: independent trigger creation is the architectural question. Whether a
  time-driven run in a copy sees its own document and Document Properties is the same
  mechanism Legacy's bound trigger uses in production, and TEMPLATE-006 observes it on the
  first real report.
- The manual-copy status box and the per-copy permission prompt were not part of the reported
  result. Neither decides GO.
- The menu item was clicked several times and only one test trigger existed afterwards. That
  is the probe's own behavior (it deletes its previous test trigger before creating one); it
  says nothing about Google.
- One account, one environment, one day.

**Manual cleanup (one step):** in the PROGRAMMATIC COPY, **Report → Clean Up Test State**. It
removes that copy's test trigger and every `TEMPLATE001_` property. The scratch documents can
then be deleted by hand later.

Note for completeness: *Initialize* also created a test trigger in the original scratch
document (so there was something a copy could inherit), and a trigger exists in the manual
copy if step 7 was done there. Each test trigger deletes itself the first time it runs. If
those documents are deleted before that, run the same menu item in them first.

---

## 20. What Thomas does to create a report (one page)

**Once, when the template is first set up (TEMPLATE-005/006):**
1. Create a Google Doc "SA4 Report Template" and paste its Doc ID and bound Script ID into
   `template/template-target.json`.
2. `node tools/template-release.js --release T-2026.10.0 --write`, then `cd` into the printed
   folder and run `clasp push`.
3. Open the template, then *ℹ️ Template release info*, and authorize once. Bookmark it.

**Every new meeting (about 1 minute of attention):**
1. Open the bookmarked **SA4 Report Template**, then **⚠️Scripts⚠️ → 🆕 Create New SA4 Report**.
2. Type the **meeting ID**, then **Look up**. Check the detected family/list/agenda. Answer the
   family question only if asked (main meetings, ambiguous names).
3. **Create Report**, then **Open report**.
4. In the report: **⚠️Scripts⚠️ → 🚀 Finish Setup & Build**. The first time, Google asks for
   permission for this report; allow it and click the item again.
5. **Build report now** (and, if wanted, tick *automatic updates every 30 min*).

If the agenda is not published yet, step 4 says so. Later: **Configure Meeting → Discover →
Save**, then build.

**Never needed:** Apps Script editor, test deployments, "add as test document", the CENTRAL
add-on menu.

**Checklist for the first real report (TEMPLATE-006):** menu present; *About this report* shows
`T-2026.10.0`; build succeeds; one e-mail collector run; export of one TDoc discussion e-mail
(document links absolute); reallocation survives a rebuild; 30-min trigger visible in
*Manage Auto-Update Trigger*; the template still has no configuration and no trigger.

---

## 21. Next stage: turning the prototype into the real SA4 Report Template

### 21.1 Source-of-truth strategy

- **One production `Code.js`: `sa4-report` master.** Template releases bundle it byte-identical
  (`tests/template-release.test.js` asserts this). There is no template fork and no second
  `Code.js` to maintain.
- **Template-only behavior lives outside `Code.js` or behind a runtime check.** `Release.js`
  (generated, template bundles only) switches the flavor; `Code.js` asks
  `typeof SA4_RELEASE_ !== 'undefined'` at call time. Without `Release.js` (CENTRAL, tests)
  nothing changes.
- **Legacy (`sa4-report-legacy`) is not a source for the template.** It stays the script of
  the 86178 report and is not touched. Its accepted fixes are carried forward into
  `sa4-report` master, where they are reviewed and tested like any other change.
- **A report is pinned to its release** (`Release.js` copied with the script,
  `SA4_CREATED_FROM_RELEASE` in its properties). Fixes reach an existing report only by a
  deliberate per-report push (§10.5).

### 21.2 What TEMPLATE-001 allows to be simplified

| Item | Before | Now |
|---|---|---|
| First run checked for inherited triggers and warned | prototype code + one test | **Removed** (triggers are not copied) |
| Fallback creator via the Apps Script API (E2) | design fallback | **Dropped**; no `script.projects` scope, no API setting, no bundle hosting |
| Manual *Make a copy* + setup-in-copy as an alternative primary flow | design fallback | **Fallback only**, no code |
| First-run wipe of inherited properties | planned if properties were copied | **Not needed** (stores start empty) |
| "UNVERIFIED" markers on the live entry points | prototype comments | Replaced by the verified facts |

Kept on purpose, because they do not depend on copy semantics: the `targetDocumentId` check
(someone may copy a report that has not been set up yet), payload validation (the description
is editable text), the template guard, and moving a half-prepared copy to the trash.

### 21.3 Finding: `sa4-report` master does not yet contain everything Legacy has accepted

Legacy received an integrated release on 2026-10-01 (deployed and live accepted on the 86178
report). To size the gap, Legacy's ten regression suites were run, unmodified, against the
current `sa4-report/Code.js` in a scratch folder (read-only for both repositories):

| Legacy suite | Result against `sa4-report` `Code.js` | Reading |
|---|---|---|
| `bugfix-legacy-002` (duplicate TDoc tables) | 4 checks fail | **Real gap.** A second import run, or one TDoc under two agenda groups in one run, duplicates its table. Legacy fixed this (`d1021cf`); master has not. |
| `legacy-adhoc-email-collection`, `diagnose-6g-email-collection`, `legacy-integration-reallocation-email` | 15 + 4 + 2 checks fail; the diagnostic suite then stops on a Legacy-only helper (`extractTdocFromEmailExportSubject_`) | **Real gap** (BUGFIX-LEGACY-003): the saved `EMAIL_START_DATE` is not read by the collector; a Collector Configuration `LIST_NAME` / `RSS_URL_V2` override does not win; the A1 archive does not follow the meeting's list; `…@list.etsi.org` forms are not normalized. Master's ADDON-008A2 covers only part of this. |
| `legacy-0099-reallocation-suppression` | 2 of 82 checks fail | Small difference in one reallocation scenario and its log line; to be read in detail. |
| `legacy-upgrade-004` | 1 check fails (`EMAIL_START_DATE unchanged`), then stops on test mechanics | Same start-date gap. |
| `legacy-email-doc-links` | 1 check fails | **Intended difference**: CENTRAL's introduction text is meeting-neutral (2.15.0). Needs a decision, not a port. |
| `legacy-upgrade-006` | stops: "a list tag is required" | **Intended difference**: CENTRAL derives the subject tag from the mailing list instead of the constant `FS_6G_MED`. |
| `legacy-upgrade-002`, `-003` | stop on test mechanics (Legacy git history, dialog markup) | Not behavior; the tests need adapting, not the code. |

So a template built from today's master would lose accepted Legacy behavior in at least two
places (duplicate-table protection and the ad-hoc e-mail collector). That is why the next
stage is a parity gate and not the runtime hooks.

### 21.4 TEMPLATE-002A: Legacy parity gate (recommended next stage)

Goal: `sa4-report` master contains every accepted Legacy behavior, proven by Legacy's own
tests, so the template never introduces a second, diverging production `Code.js`.

1. Branch from `sa4-report` master (not from this design branch).
2. Copy Legacy's regression suites into `sa4-report/tests/` as `legacy-*.test.js`, adapted only
   in mechanics (loader path, no dependency on Legacy's git history).
3. Port the real gaps into `Code.js`, one reviewed commit each: (a) duplicate-TDoc insertion
   guard; (b) collector list / start-date resolution; (c) the two reallocation-suppression
   differences, after reading them.
4. For each intended difference (introduction text, subject tag, deadline default), record the
   decision in the test: either "template keeps CENTRAL's behavior" or "template keeps Legacy's".
   These are Thomas's calls.
5. Gate: all ported Legacy suites and the full existing suite pass on one `Code.js`.

Files: `Code.js`, new `tests/legacy-*.test.js`. No Google operation. Production risk: none
until something is pushed; CENTRAL would receive the same fixes at its next push, which is a
separate decision. Rollback: revert the commits.

Then, unchanged from §17: **TEMPLATE-002B** (flavor-aware trigger intervals, menu without the
CENTRAL submenu, template guard), **003** (creator and first-run dialogs), **004** (trigger
hygiene), **005** (release pipeline, create the template document), **006** (first real
report, which also observes the first trigger execution in a copy).
