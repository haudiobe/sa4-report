# Rollout of T-2026.10.28 (Code.js 2.26.12) — plan, not yet approved

T-2026.10.28 = T-2026.10.27 (`b0bc1eb`, in production on all five SA4#138 reports) + the local
listing agent (docs/LOCAL_LISTING_AGENT.md) + Check Connections trying the TDoc list like the
update. Release files: Code.js and Release.js change; HyperLink.js, ReportCreator.js and
appsscript.json are those of T-2026.10.27. No configuration migration: a report without
*Local listing folder* behaves exactly as on T-2026.10.27; *Import Local Inbox Listing…* is unchanged.

Every step below needs approval. Nothing here has been done in production.

## A. Scripts, one report at a time (MBS, Video, 6G Media, RTC, Audio)

For each report:

1. Verify identity: the script is bound to the inventory document (issue #16 mapping), the document
   name is the production report's, the script runs T-2026.10.27 (Code.js `63e3c061…`).
2. Fresh `clasp pull` backup into `dist/backups/sa4-138-deploy-T-2026.10.28/<report>-<UTC>/backup/`;
   it must equal the pre-release baseline (`dist/backups/sa4-138-five-before-T-2026.10.28-*/`).
3. Push the four runtime files of the bundle byte for byte; keep the report's own `appsscript.json`
   when it differs only in formatting (no manifest overwrite, never `--force`).
4. Pull back; compare every runtime file with `RELEASE-MANIFEST.json`; Release.js must say
   T-2026.10.28 / 2.26.12. The document's version rises by one at the script's update time.
5. UI smoke test (about 1 minute): *About This Report* shows T-2026.10.28; *Check Connections* runs and
   shows "Local listing agent: not configured"; *Update TDoc Revisions* runs without error and the
   cells stay as they were.
6. Stop at the first failure; roll that report back (section D).

## B. The agent (once, in the meeting room)

1. A **new** Drive folder for production (not the test folder), e.g. `SA4-138 local listings`, in the
   Drive of the laptop's account; share it as **Viewer** only with the accounts that run the reports.
2. Copy `tools/local-listing-agent/` to `C:\SA4Agent\`; `agent.config.json` with `outDir` = that folder.
3. One manual run: five lines `ok, N links`; the ten files appear in Drive (web) within a minute.
4. `install-scheduled-task.ps1 -ConfigPath C:\SA4Agent\agent.config.json` (every 2 minutes while
   logged on; repeats for 30 days; `-Uninstall` removes it).

## C. Switch each report on, one at a time

1. *Configure Meeting › Local listing folder* = the production folder ID; Save.
2. *Update TDoc Revisions*; *Check Connections*: stage 3 shows a valid scan in Drive, stage 4 an import.
3. *Update TDoc Revisions* again: "already imported" (idempotent), no duplicate lines.
4. Monitor during the meeting: stage 1–3 times stay within a few minutes; "stale" means agent laptop,
   room network or Drive sync — the reports keep what they have.

## D. Rollback (per report, independent of the others)

- Configuration only: clear *Local listing folder* — the report is T-2026.10.27 in behaviour again.
- Script: from that report's `…/backup/` folder: `clasp -u <user> push`, pull back, compare with the
  T-2026.10.27 manifest (Code.js `63e3c061…`). Imported local revisions stay stored (as after a manual
  import); `LOCAL_LISTING_AGENT_STATE` is ignored by T-2026.10.27.
- Agent: `install-scheduled-task.ps1 -Uninstall`; the Drive folder may stay or be deleted.
