# Rollout of T-2026.10.32 (Code.js 2.27.0): Video pilot plan, not yet approved

T-2026.10.32 is T-2026.10.31 (`af08ae1`, Code.js 2.26.15, the concurrent-edit hotfix) plus the
Undisposed Documents section and the local document availability (docs/UNDISPOSED_AND_LOCAL_DOCS.md).

- **Release files:** Code.js and Release.js change. ReportCreator.js, HyperLink.js and
  appsscript.json are those of T-2026.10.30.
- **Every step below needs approval.** Nothing here has been done in production.
- **The production reports run T-2026.10.30.** Pushing T-2026.10.32 to a report also brings the
  hotfix of T-2026.10.31 into it. That hotfix changes only what Goes to Plenary does (set-up and the
  end of an update), and no production report but Audio has Goes to Plenary switched on.
- **The Plenary configuration of the reports is not changed.** Nothing is switched on or off, no
  dropdown definition is touched.

## The pilot script

`dist/backups/sa4-138-deploy-T-2026.10.32/pilot-T-2026.10.32.mjs` in the main checkout (not in Git,
like the deploy scripts before it). Every run writes `evidence.json` beside it.

| Command | What it does |
|---|---|
| `--build` | the candidate bundle from the clean worktree at the release commit (no Google access) |
| `--target copy --script <id> --doc <id> [--write]` | a disposable Video copy (guard.js: name, binding, never a production report or the template) |
| `--target video` | dry run: binding, a fresh backup that must equal the T-2026.10.30 read-back, the plan |
| `--target video --write --nobody-editing --named-version "before T-2026.10.32"` | the push, read-back and checks; refused without both confirmations |
| `--target video --rollback [--write --nobody-editing]` | pushes the T-2026.10.30 read-back back, byte for byte |

State on 2026-11-19 (read-only runs): Video runs exactly its T-2026.10.30 read-back. The dry run
and the rollback dry run both pass, and the unsafe invocations are refused before any access.

## Read-only checks only you can make (before § 0 or § 1)

Claude's Google access reads the script and the document's metadata, but not the document's
content, its properties or its executions. In the production Video report:

1. *About This Report*: T-2026.10.30 / 2.26.14.
2. *Check Report Integrity* (read-only): note the lines on Status dropdowns, the Disposition
   Summary and Goes to Plenary.
3. Look at the headings of 9.1: is there a "9.1.3 Document Reallocations", a "9.1.4 Documents"?
4. *Extensions › Apps Script › Executions*: the usual duration of an update.

The answers say which path the section takes in Video (summary, Status dropdowns, or its own Docs
API read) and set the duration to compare with.

## 0. Before the pilot: a disposable Video copy (recommended)

Neither T-2026.10.31 nor T-2026.10.32 has run in a real Google document yet; the tests run on a
fake document.

**By you:**

1. Make a copy of the Video report (*File › Make a copy*) named "DISPOSABLE TEST COPY - Video SA4#138".
2. Send Claude the copy's document ID and script ID.

**By Claude:** `--target copy --script <id> --doc <id>`, first as a dry run, then with `--write`. It
pushes into the copy's script only, and only when the name says "DISPOSABLE TEST COPY", the IDs
match and the script runs T-2026.10.30. Then it reads the script back.

**By you, in the copy:**

3. *About This Report*: T-2026.10.32 / 2.27.0.
4. *Update Report Now*:
   - "9.1.4 Undisposed Documents" appears at the end of 9.1, before 9.2. It is 9.1.5 if the report
     has a "9.1.4 Documents".
   - No other heading has changed.
   - Count the TDocs whose Status is Available, Parked or Reserved and whose Plenary is not Yes:
     the count line shows the same number.
   - Note the duration in *Extensions › Apps Script › Executions* and compare it with an update
     of the production report.
5. *Update Report Now* again: the section is unchanged; the execution log says
   "Undisposed Documents: unchanged".
6. Set one listed TDoc to Agreed and update: its row goes, the count goes down. Set it back.
7. Type a line directly below the section's table and update: the line is kept. Type a line
   between the heading and the table and update: the update says "text typed into it … move it
   below its table", and nothing in the section changes. Remove that line.
8. *Check Report Integrity*: no new problem. The Disposition Summary, if the report has one,
   is unchanged.

The Docs folder needs the agent change in the room (§ 3). Until then the copy's
*Check Connections* has no "Local Docs folder" line, and nothing moves from Reserved to Available.

## 1. Pilot: production Video (separate approval)

**Preconditions, all of them:**

- Step 0 passed, or you decide to skip it.
- **Nobody is editing the Video report:** between sessions, confirmed with the Video rapporteur.
- **A verified backup:**
  - *File › Version history › Name current version*: "before T-2026.10.32".
  - A fresh pull of the Video script must equal its T-2026.10.30 read-back,
    `dist/backups/sa4-138-deploy-T-2026.10.30/video-20261118T180314Z/readback/` (runtime files byte
    for byte, appsscript.json by content). The push refuses otherwise and stops.

**Steps:**

1. Claude runs `--target video` (dry run) and then, with your two confirmations,
   `--target video --write --nobody-editing --named-version "before T-2026.10.32"`. That is the
   push, the read-back and the checks: the plan, Release.js T-2026.10.32 / 2.27.0, and the section's
   code present.
2. You run *Update Report Now* once and check steps 3–5 and 8 of § 0. Do not run 6–7 in production.
3. **It passes when:**
   - the section is in the right place with the right count;
   - nothing else in the report changed;
   - the update completed, and its duration is acceptable to you;
   - the second update writes nothing.
4. **It stops on any mismatch:** rollback, then report.

**Rollback:**

1. `--target video --rollback --write --nobody-editing`: the T-2026.10.30 read-back above, back into
   the Video script, byte for byte; then read back and checked (Release.js T-2026.10.30 / 2.26.14).
2. Remove the section: delete its heading, its two lines and its table (or "None."). T-2026.10.30
   does not know it and would leave it as it is. It holds nothing typed, unless someone typed
   below its table: move that text first.
3. The document properties `LOCAL_DOCS_CONFIRMED` and `LOCAL_DOCS_AGENT_STATE`, if present, are
   unused by T-2026.10.30 and may stay.

Restoring the named version would also undo the minutes typed since; use it only if the report was
not edited after the push.

## 2. The other four reports (separate approval, only after the pilot passes)

MBS, RTC, Audio and 6G Media, one at a time, as § 1:
- each needs a fresh backup equal to its T-2026.10.30 read-back, a time with nobody editing, and a
  read-back after the push;
- in 6G Media the section is 11.0.5 (11.0.4 is the IPR reminder);
- Audio, which has Goes to Plenary switched on, also gets the hotfix's recovery path
  (docs/ROLLOUT_T-2026.10.31.md § A). Its order with this release is decided then.

## 3. The local listing agent (separate approval)

The agent in `C:\SA4Agent` lists the Docs folder only when its configuration has `docsUrl`. It is
changed only after the meeting-server connectivity check is complete. Then:

1. Back up `C:\SA4Agent` (all files) to a dated folder.
2. Copy the new `sa4-local-listing-agent.ps1`.
3. Add `"docsUrl": "http://10.10.10.10/ftp/SA/SA4/Docs/"` to `agent.config.json`.

The scheduled task and its window are not touched. To test: `-Group Docs -IgnoreWindow` on the
meeting-room network. **Rollback:** restore the backed-up files.
