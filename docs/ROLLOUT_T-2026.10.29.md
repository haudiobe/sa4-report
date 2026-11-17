# Rollout of T-2026.10.29 (Code.js 2.26.13): plan, not yet approved

T-2026.10.29 is T-2026.10.28 plus two changes:
- **PR #23:** the Revisions-cell cleanup. There is no "Revisions folders" line any more, and
  "No revisions available." appears only where it is known. Check Connections no longer shows the
  manual-import instruction while the agent delivers.
- **The overnight pause** of the local listing agent is reported as **paused**, not **stale**
  (docs/LOCAL_LISTING_AGENT.md, "Scheduled pause vs. genuine failure").

T-2026.10.28 (`feb4c6e`, Code.js `4ed66659…`) is in production on all five SA4#138 reports, deployed
2026-11-17 00:18–00:20 UTC.

Release files: Code.js and Release.js change. HyperLink.js, ReportCreator.js and appsscript.json are
those of T-2026.10.28. There is no configuration migration. Cells that still carry the folders line
are rewritten once, at their next update, with no other change.

The **paused** result needs the meeting time zone of the report (*Configure Meeting › Meeting time
zone*, `America/Edmonton` for SA4#138). A report without one keeps saying **stale** overnight, as on
T-2026.10.28.

Every step below needs approval. Nothing here has been done in production.

## 0. Validation on the disposable MBS copy (before any production step)

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138".
1. **Done 2026-11-17 12:00 UTC.** The candidate bundle (`dist/release-candidate/T-2026.10.29/`, from
   `3927fc3`) was pushed into the copy's bound script. The read-back matches `RELEASE-MANIFEST.json`
   (Code.js `cc32cbb6…`); the copy ran the T-2026.10.28 candidate before. The later commits on the
   release branch change tests, tools and documentation only; the runtime files are byte-identical.
2. In the copy, menu **SA4 Report**:
   - *ℹ️ About This Report* shows T-2026.10.29 / 2.26.13.
   - *⚙️ Configure Meeting…*: **Meeting Time Zone** is `America/Edmonton` (shown with ABT), or
     *Meeting local* resolving to it. **Local listing folder** is the test folder
     `DISPOSABLE TEST - SA4 local listings (T-2026.10.29)`. Save.
   - *📄 Report › Update TDoc Revisions* runs without error. Old "Revisions folders: …" lines are
     gone; revision lines and their links stay; minutes are untouched; no "No revisions available."
     appears in a cell that had none.
   - *🛠 Advanced and Repair › Check Connections*: before 07:45 ABT, the agent line reads
     `Not imported: scheduled overnight pause. … last scan of MBS was at <date> 19:28 ABT, scanning
     resumes at 07:30 ABT. This is not a synchronization failure.` From 07:45 ABT on, with the same
     file, it reads `… is stale: scanned …` (the agent did not resume).
3. Put the copy's *Local listing folder* back to what it was. Roll the copy back if anything differs
   (`dist/integration-test/rc-T-2026.10.29-20261117T120045Z/before/`).

## A. Scripts, one report at a time (MBS, Video, 6G Media, RTC, Audio)

For each report:

1. Verify identity: the script is bound to the inventory document (issue #16 mapping), the document
   name is the production report's, and the script runs T-2026.10.28 (Code.js `4ed66659…`).
2. Take a fresh `clasp -u <user> pull` backup into
   `dist/backups/sa4-138-deploy-T-2026.10.29/<report>-<UTC>/backup/`. It must equal that report's
   T-2026.10.28 read-back (`dist/backups/sa4-138-deploy-T-2026.10.28/<report>-*/readback/`).
3. Push the four runtime files of the bundle byte for byte. Keep the report's own `appsscript.json`
   when it differs only in formatting (no manifest overwrite, never `--force`).
4. Pull back and compare every runtime file with `RELEASE-MANIFEST.json`. Release.js must say
   T-2026.10.29 / 2.26.13.
5. UI smoke test (about 1 minute):
   - *About This Report* shows T-2026.10.29.
   - *Check Connections* runs.
   - *Update TDoc Revisions* runs without error, and cells keep their revisions.
6. Stop at the first failure and roll that report back (section C).

## B. After the scripts

- In each report, check *Configure Meeting › Meeting time zone*: `America/Edmonton` makes the
  overnight pause show as **paused**. Without it, overnight stays **stale**; nothing else changes.
- Agent: PR #25 (window in ABT, UTC−6) was installed on the agent laptop on 2026-11-17. Its task
  runs daily from 07:30 ABT (13:30 UTC) for 12 hours, every 2 minutes; the backup of the previous
  files and task is in `C:\SA4Agent\backup-20261117T122830Z\`. The report's 07:45 ABT boundary
  therefore leaves a quarter of an hour for the first scan to reach Drive.

## C. Rollback (per report, independent of the others)

From that report's `…/backup/` folder: `clasp -u <user> push`, pull back, and compare with the
T-2026.10.28 manifest (Code.js `4ed66659…`). Cells already rewritten without the folders line stay
without it: T-2026.10.28 writes the line again at their next update. `LOCAL_LISTING_AGENT_STATE`
entries with `lastResult: "paused"` are read by T-2026.10.28 as an ordinary message.
