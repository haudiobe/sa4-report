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

Target: "DISPOSABLE TEST COPY (issue #16) - MBS SA4#138", which currently runs the T-2026.10.28 candidate.
1. Push the candidate bundle (`dist/release-candidate/T-2026.10.29/`) into its bound script. Read it
   back and compare every file with `RELEASE-MANIFEST.json`.
2. In the copy:
   - *About This Report* shows T-2026.10.29 / 2.26.13.
   - *Update TDoc Revisions* runs without error.
   - Cells with the old folders line lose it; revision lines and their links stay; minutes are untouched.
   - *Check Connections* shows the local-listing lines.
3. Overnight check (after 20:00 ABT), with the copy's *Local listing folder* set to the test folder and
   its meeting time zone set to `America/Edmonton`: *Check Connections*' first line reads
   `Not imported: scheduled overnight pause. …`. During the day, a stopped agent reads `stale`.
4. Roll the copy back if anything differs (its `before/` folder).

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
- Agent: deploy PR #25 (window in ABT, UTC−6) if approved. The agent installed now scans
  08:30–20:30 ABT, because this laptop's Windows rules for `Mountain Standard Time` predate
  Alberta's change. With that agent, the report says **stale** from 07:45 to 08:30 ABT, which is
  correct: no scan has arrived yet.

## C. Rollback (per report, independent of the others)

From that report's `…/backup/` folder: `clasp -u <user> push`, pull back, and compare with the
T-2026.10.28 manifest (Code.js `4ed66659…`). Cells already rewritten without the folders line stay
without it: T-2026.10.28 writes the line again at their next update. `LOCAL_LISTING_AGENT_STATE`
entries with `lastResult: "paused"` are read by T-2026.10.28 as an ordinary message.
