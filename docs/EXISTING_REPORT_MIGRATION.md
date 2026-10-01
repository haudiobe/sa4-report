# Migrating existing reports to the template runtime

2026-10-01. Branch `migration/stage1-tooling` (tooling commit `16a6bb9`), based on master
`d56d628`. Target runtime: template release **T-2026.10.4** (`Code.js` 2.17.4, commit
`962fb7de1f58302b54c51eeafa3c5c57d50f12c3`), unchanged. No new release was created.

| Report | State |
|---|---|
| 6G (meeting 86178) | **COMPLETE / LIVE on T-2026.10.4** (§2) |
| MBS (meeting 86172) | Rehearsal passed (§4). Live migration not started: the live document has no known bound script yet. |

---

## 1. Tooling: adoption mode

`tools/template-release.js --adopt <release> --target-script <id> --label <name> [--write]`
prepares the unchanged payload of an already built release for another bound script project.

- The seven files are copied byte for byte from the release bundle, after a check against the
  bundle's manifest and the release tag. `Release.js` is not regenerated, so an adopted report
  identifies itself as that release.
- Output: `dist/template-adopt/<release>/<label>/` with a deployment-local `.clasp.json` and
  an `ADOPTION.json` record. Neither is uploaded.
- Refused targets: none given; CENTRAL (always); the master template's own script; the Legacy
  bound script unless `--allow-legacy-target` is given.
- The tool never runs clasp and has no force option.
- Tests: `tests/template-adopt.test.js`, 44 checks.

---

## 2. 6G: migrated in place

**Method: script replacement in the existing bound project**, so that the project's property
stores and its token stay where they are.

### 2.1 Rehearsal (on a Drive copy)

- Copy `MIGRATION TEST — 6G — T-2026.10.4`, document `REDACTED-6G-REHEARSAL-DOCUMENT-ID`,
  bound script `REDACTED-6G-REHEARSAL-SCRIPT-ID`.
- The copy's project held the five Legacy files. A normal `clasp push` was skipped because the
  manifest changes (the template manifest adds an explicit scope list); the push was made with
  a one-time, separately approved `--force`. A fresh pull showed seven files byte-identical to
  T-2026.10.4.
- Result, accepted: Finish Report Setup behaved correctly for an adopted report; Configure
  Meeting resolved 86178 (family 6G, agenda S4aP260098, list `3GPP_TSG_SA4_FS_6G_MED`); Update
  Report Now completed; the version-history comparison showed no unexpected change to minutes,
  dispositions, TDoc tables, reallocations or structure.
- A copy starts with empty properties, so the rehearsal validated document and runtime
  compatibility, not state preservation.

### 2.2 Live migration

Before: the live project was five files, byte-identical to the Legacy repository at
`b95e06e`. A document version was named, the Continuous Update trigger stopped, and no
execution was running.

Push: `clasp push --force` (one-time approval, this Script ID only) from
`dist/template-adopt/T-2026.10.4/6g-live` to
`REDACTED-LIVE-6G-SCRIPT-ID`: seven files, 2026-10-01 16:57:30.
A fresh pull showed seven files, each byte-identical to the accepted T-2026.10.4 bundle.

### 2.3 Accepted facts

- The live 6G report runs T-2026.10.4 / `Code.js` 2.17.4 / commit `962fb7d`.
- **Existing Document Properties survived the in-place migration.** Meeting 86178, family 6G
  and agenda S4aP260098 survived.
- Mailing List `3GPP_TSG_SA4_FS_6G_MED`; Discussion E-mail Sender `reporter@example.com`;
  Email Collection Start Date `2026-08-21`.
- The Reviewer API token survived and is configured.
- About This Report shows the correct release, meeting and the live Script ID.
- Update Report Now completed successfully; the version-history comparison showed no
  unexpected changes.
- Automatic Updates are active again at 30 minutes.
- **Rollback source:** the Legacy repository at `b95e06e`.

No further change is planned for 6G.

### 2.4 What this established for later migrations

- Replacing the files of a bound project keeps its Document Properties, Script Properties and
  therefore configuration, collector ledgers and token.
- A Legacy manifest without a scope list always needs the manifest-overwrite confirmation.
- A Legacy trigger has to be stopped before and restarted after, through Automatic Updates…,
  so that the interval is recorded.

---

## 3. MBS: preflight (analysis only)

The MBS report runs under the CENTRAL add-on. Nothing was changed, and no clasp command was
run against any MBS or CENTRAL project.

### 3.1 Why MBS is different

Property stores belong to the script that wrote them. A new bound script in the MBS document
starts empty: it sees neither the add-on's Document Properties for that document nor CENTRAL's
Script Properties. Everything report-specific must be re-entered or rebuilt.

### 3.2 Where MBS state lives today

| State | Location under CENTRAL | In the bound runtime |
|---|---|---|
| Meeting configuration (meeting ID, type, name, date, family, agenda, TDoc list, FTP base, revisions URL) | Add-on Document Properties; copied to CENTRAL Script Properties `SA4_STATE|<doc>|…` each time *Enable Automatic Updates* runs | Re-derived by Configure Meeting's lookup from the meeting ID |
| Mailing-list override | same (`MAILING_LIST`) | Re-entered if it differs from the family list |
| Discussion E-mail Sender | Add-on Document Properties only | Re-entered |
| Email Collection Start Date | **Not configurable in CENTRAL**: the collector uses `2026-08-21` unless a Key/Value table sets `EMAIL_START_DATE` | Must be set to the same date explicitly; the template runtime would otherwise use the meeting start date |
| Abstracts during updates | `FETCH_ABSTRACTS_ON_UPDATE`, same two places | Re-ticked in Automatic Updates… |
| Reviewer API token | CENTRAL Script Property, shared by all CENTRAL reports | Entered again in the report; never copied by tooling |
| E-mail ledgers `DISCUSS_<tdoc>`, revision ledgers `REVIS_<tdoc>`, `REVISION_MAP` | Split: menu actions write the add-on's Document Properties, scheduler runs write the CENTRAL copy | **Empty; rebuilt by the first update** |
| Registration, enabled flag, interval, last run | CENTRAL Script Properties `SA4_REGISTRY_DOC|<doc>` and `SA4_REGISTRY_INDEX` | Replaced by the report's own trigger |
| Minutes, dispositions, reallocation table, deadline-extension table, TDoc tables | The document | Unchanged, no migration |

### 3.3 Can a fresh runtime rebuild the ledgers?

Not proven. What the code does:

- **E-mail.** Each update reads the list's RSS feed (newest 2,000 items, filtered by the start
  date) and the archive index pages of the months touched by the last 14 days. It merges them
  into the ledger and re-renders each e-mail cell from the ledger. With an empty ledger the
  cell is rebuilt from what those sources return now. A message that neither source returns
  any more drops out of its cell. If nothing at all is found for a TDoc, the cell is left as
  it is.
- **Revisions.** The ledger accumulates every draft ever seen; the cell is rendered from it.
  With an empty ledger the cell shows the drafts in the folder now. A draft that was removed
  from the folder drops out.

Only a rehearsal on a copy answers this for MBS.

### 3.4 Disabling CENTRAL for this document only

*CENTRAL ADD-ON › Disable Automatic Updates (this doc)* sets `enabled: false` on this
document's registry entry and nothing else. The scheduler skips disabled entries; other
reports and the scheduler trigger are not affected.

Leaving the entry present but disabled is sufficient and preferable: it is the only option the
menu offers, it keeps CENTRAL's saved state as a fallback, and re-enabling is one click.
Deleting the entry or the saved state has no menu item and would mean running code in CENTRAL.

Two cautions: re-enabling later overwrites CENTRAL's copy with the add-on's Document
Properties; and after migration nobody should run CENTRAL's menu actions in the MBS document,
because CENTRAL would work from its own, older ledgers.

---

## 4. MBS: inventory and rehearsal (accepted)

### 4.1 Live inventory (manual, read-only)

- Document "MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e"; meeting 86172, ad-hoc, family MBS.
- No agenda TDoc; TDoc list URL for meeting 86172; Mailing List `3GPP_TSG_SA_WG4_MBS`; sender
  `reporter@example.com`; revisions / drafts URL saved; e-mail preview snippets enabled;
  Reviewer API token configured in the existing runtime.
- No `LIST_NAME` and no `EMAIL_START_DATE` in the document, so no Key/Value override table.
- **Not registered with the CENTRAL scheduler.** The document is updated by a per-document
  Continuous Update trigger, hourly, with abstracts enabled. This corrects §3.2 and §3.4: there
  is no registry entry to disable, and the MBS state is in one place, the add-on's Document
  Properties for this document. What has to be stopped at migration is that trigger.

### 4.2 Rehearsal

- Copy: document `REDACTED-MBS-REHEARSAL-DOCUMENT-ID`, bound script
  `REDACTED-MBS-REHEARSAL-SCRIPT-ID`.
- The copy's bound project was Google's empty default project: the copy inherited no report
  code, so the original has no bound report script.
- T-2026.10.4 was pushed with a one-time, separately approved `--force` (manifest change) and
  verified by a fresh pull: seven files, byte-identical to the accepted release.
- The first Update Report Now was refused by the ad-hoc readiness guard, before any document
  change: the copy had no agenda source and no TDoc list URL, because only *Resolve* had been
  used. *Discover Agenda / TDocs* supplied both: validated `agenda.csv` fallback (18 of 53
  items, sections 2 and 3) and the TDoc list URL for meeting 86172.
- Configured: Mailing List `3GPP_TSG_SA_WG4_MBS`, sender `reporter@example.com`, Email
  Collection Start Date `2026-08-21`, revisions URL discovered, token entered by hand.
- Update Report Now then completed. The version-history comparison showed no unexpected
  changes, and the existing e-mail discussion and revision information was reconstructed
  sufficiently.

**Decision: no CENTRAL ledger export or import is required. MBS rehearsal passed.**

### 4.3 For the live migration

- A bound script must first exist in the live document; none is known. Creating one is a
  change to the original and is the user's step.
- Configure Meeting needs *Resolve*, then *Discover Agenda / TDocs*, then Save.
- The start date must be set to `2026-08-21` explicitly.
- The add-on keeps its own state for the document, untouched by the migration, so going back
  means restarting the add-on's hourly trigger.
