# Local listing agent — automatic local revisions (after T-2026.10.27)

The reports show revisions found on the meeting-room server (`http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/<group>/`)
next to the 3GPP ones (docs/LOCAL_REVISIONS.md). Google cannot reach that server
(docs/LOCAL_REVISIONS_AUTODETECT.md), so a small agent on one laptop in the room reads the
listings and hands them over through Google Drive. *Import Local Inbox Listing…* stays as it is.

```
meeting-room server  --HTTP GET-->  agent (laptop on the room network, every 2 min, 07:30-19:30 Calgary)
                                      | <group>.listing.json, <group>.status.json
                                      v
                     Drive folder, synchronized by Google Drive for desktop
                                      |  read by every update of the report (its own group only)
                                      v
                     report: validate -> existing import (parse, refresh, merge, render)
```

## What is exchanged

Per working group (MBS, Video, FS_6G_MED, RTC, Audio), written atomically (temp file + rename):

- `<group>.listing.json` — `format`, `version`, `group`, `folder`, `fetchedAt` (UTC), `scanId`,
  `agent` (computer name), `linkCount`, `sha256`, `payload`: the listing's links only
  (`<a href="…">name</a>` per line). **No TDoc file is copied**, only names and links (1–6 kB).
- `<group>.status.json` — last attempt, last success, error. Informative only.

## What the report accepts (every update, before the 3GPP folder)

Only a file that is complete (valid JSON, `payload` matches `sha256` and `linkCount`), of the
report's own group, listing that group's folder on the report's meeting-room server, scanned at
most **30 minutes** ago and not more than 5 minutes in the future. Otherwise nothing is imported
and **nothing is erased**; the reason is recorded. A successful local write is no proof that Drive
has uploaded it: the report judges only by what it reads from Drive (`fetchedAt` inside the file),
so a delayed upload shows as "stale", never as success. Several files with the same name (a Drive
conflict copy) — the newest valid one is used. The same scan twice is imported once (`scanId`).

Status: *SA4 Report › 🛠 Advanced and Repair › Check Connections* shows four stages, each from its
own evidence, so that one is never taken for the next:

1. **LAN scan** — from the agent's `<group>.status.json` (`scanOk`, `lastSuccessAt`, last error), as far as
   Drive has synchronized it.
2. **Local write** — from the same file (`writtenAt`). A local write is no proof of an upload.
3. **In Google Drive** — the newest *valid* listing the report finds in Drive, its scan time and its upload
   time. Only this proves synchronization.
4. **Import** — the report's own state: when, files, revisions, new, removed.

Its first line says the result of the last check: imported, unchanged, paused (the scheduled overnight
pause, see below), stale, missing, invalid, unavailable.

Check Connections also tries the TDoc list as the update does (the configured address, then its
`www.3gpp.org/ftp` twin); a 403 from 3GPP is shown as a refusal by the 3GPP server, usually temporary.

## Operating window and overnight staleness

The agent scans **every 2 minutes from 07:30 to 19:30 Calgary local time, every day**. Since 1
November 2026, Calgary has been on **UTC−6 all year (ABT)** (docs/MEETING_TIME_ZONES.md). The window
is therefore a fixed offset: 07:30 ABT is 13:30 UTC in every season. The first scan is at 07:30 and
the last one starts at 19:28. The window is set in `agent.config.json`:

```json
"operatingWindow": { "utcOffset": "-06:00", "label": "ABT", "start": "07:30", "end": "19:30" }
```

Without that entry the agent uses the same values.

> **Do not use the Windows zone `Mountain Standard Time` for Calgary.** Windows installations whose
> time-zone data predates Alberta's change still move it to UTC−7 in winter. With
> `"timeZone": "Mountain Standard Time"`, such a laptop scans from 08:30 to 20:30 ABT. A
> `"timeZone"` without `"utcOffset"` uses that Windows zone and its daylight-saving rules; it is
> meant only for a meeting elsewhere. Check with `-CheckWindow -AtUtc 2026-11-17T13:30`, which must
> say `07:30 ABT (UTC-06:00): inside`.

Two layers keep to the window:

- **Task Scheduler** (`install-scheduled-task.ps1`) starts the agent only during the window. With a
  fixed `utcOffset`, the trigger is fixed in UTC: daily at 07:30 ABT (13:30 UTC), lasting 12 hours.
  The laptop's own time zone and its Windows time-zone rules don't matter, and changing them later
  needs no new installation. (With a Windows `timeZone` instead: 07:30 local time for 12 hours if
  the laptop is on that zone. Otherwise the trigger is fixed in UTC and one hour wider on each side,
  so that daylight saving can't cut the window short.) The installer prints which variant it
  registered.
- **The agent itself** checks the meeting time on each run. Outside the window it writes one log
  line, `paused: … outside the operating window 07:30-19:30; scanning resumes at 07:30 (last listings
  kept)`, and exits with code **0**. Task Scheduler shows that as success, not an error. It does
  **not** touch `<group>.listing.json` or `<group>.status.json`. This also covers a run that Windows
  starts late, for example after the laptop wakes up at night (`StartWhenAvailable`).

**Resuming.** No action is needed: at 07:30 the next scan runs if the laptop is on, the user is
logged in, and the meeting-room network is reachable. If the network is not reachable yet, each
scan records the error and keeps the last good listing, exactly as during the day. A scan follows
automatically once the network is back.

**Expected overnight staleness.** The reports accept a listing for at most 30 minutes. From about
**20:00 Calgary time until the first scan after 07:30**, automatic updates therefore import no local
revisions. Nothing is erased: revisions imported during the day stay in the reports, and no
Revisions cell is newly marked "No revisions available." This is expected and not a fault. The
first update after 07:30 imports again (Drive for desktop needs about a minute to upload).

**Scheduled pause vs. genuine failure (reports after T-2026.10.28).** The report tells the two apart.
It does this by itself, because the agent writes nothing while it pauses:

| The report sees a stale listing … | Check Connections says |
|---|---|
| at night (19:30 to 07:45 meeting time), and the agent's last scan was at 19:00 or later | **paused**: `Not imported: scheduled overnight pause. … its last scan of MBS was at 2026-11-16 19:28 ABT, scanning resumes at 07:30 ABT. This is not a synchronization failure.` |
| during the day (from 07:45), or at night when the last scan was before 19:00 | **stale**, as before: agent stopped, meeting-room network down, or Drive not synchronized. Look into it. |

Times are in the meeting's time zone (*Configure Meeting › Meeting time zone*, e.g. `America/Edmonton`,
which is ABT from 1 November 2026). A report without a meeting time zone guesses nothing and always
says **stale**. 07:45 allows a quarter of an hour for the first scan of the day to reach Drive.
"stale" at 08:00 therefore means the agent did not resume.

Revisions uploaded to the meeting-room server overnight appear after 07:30, or immediately via the
manual fallback:

- *Import Local Inbox Listing…* in the report (unchanged), or
- a single scan on the laptop outside the window:
  `powershell -NoProfile -ExecutionPolicy Bypass -File C:\SA4Agent\sa4-local-listing-agent.ps1 -ConfigPath C:\SA4Agent\agent.config.json -IgnoreWindow`

To check the window without scanning:
`… sa4-local-listing-agent.ps1 -ConfigPath C:\SA4Agent\agent.config.json -CheckWindow` (optionally
`-AtUtc 2026-11-18T14:30`). To check the task: `Get-ScheduledTask 'SA4 local listing agent' |
Get-ScheduledTaskInfo` (`NextRunTime` in the laptop's local time).

## Setup in the meeting room (once, about 10 minutes)

Needs: one Windows laptop that stays on the meeting-room network, logged in, with Google Drive for
desktop signed in to the account that runs the reports (or one they can read).

1. **Drive folder.** In that laptop's Drive (`G:\My Drive` or `G:\Meine Ablage`) create a folder,
   e.g. `SA4-138 local listings`. Share it only with the people whose accounts run the five
   reports, as **Viewer** (the reports only read it). Copy its ID from its URL
   (`https://drive.google.com/drive/folders/<ID>`).
2. **Agent files.** Copy `tools/local-listing-agent/` to e.g. `C:\SA4Agent\`. Copy
   `agent.config.example.json` to `agent.config.json` and set `outDir` to the Drive folder of step 1
   (e.g. `G:\\Meine Ablage\\SA4-138 local listings`). Keep `baseUrl` and the five groups.
3. **Test once.** PowerShell:
   `powershell -NoProfile -ExecutionPolicy Bypass -File C:\SA4Agent\sa4-local-listing-agent.ps1 -ConfigPath C:\SA4Agent\agent.config.json`
   Expected: five lines `ok, N links`. Within a minute the ten files appear in Drive (web).
   Outside 07:30–19:30 Calgary time the agent prints `paused: …` instead; add `-IgnoreWindow` to test.
4. **Schedule.** `powershell -NoProfile -ExecutionPolicy Bypass -File C:\SA4Agent\install-scheduled-task.ps1 -ConfigPath C:\SA4Agent\agent.config.json`
   (every 2 minutes from 07:30 to 19:30 Calgary time, daily, while you are logged on; see
   *Operating window* above; no admin rights; log:
   `%LOCALAPPDATA%\SA4LocalListingAgent\agent.log`). Remove with `-Uninstall`.
5. **Each report.** *⚙️ Configure Meeting… › Local listing folder*: paste the folder ID or URL,
   Save. The next update (automatic or *Update TDoc Revisions*) imports; *Check Connections* shows it.

## The Docs folder (after T-2026.10.31, T-2026.10.32)

With `"docsUrl": "http://10.10.10.10/ftp/SA/SA4/Docs/"` in `agent.config.json`, every run also
lists the Docs folder of the meeting-room server, with the same checks and the same atomic files,
as the group `Docs`: `Docs.listing.json` and `Docs.status.json` beside the groups' files.
`-Group Docs` lists it alone. Without `docsUrl` nothing changes. The reports use it to take a
TDoc whose file is there, named exactly by its number, from Reserved to Available
(docs/UNDISPOSED_AND_LOCAL_DOCS.md § B). A failure of the Docs folder is its own: the groups are
still listed, and its last good listing is kept.

## Failure behaviour

| Situation | Agent | Report |
|---|---|---|
| Meeting-room network or server down | records the error, keeps the last good listing | after 30 min: "stale", nothing imported, nothing erased |
| Outside 07:30–19:30 Calgary time (overnight) | paused: no scan, exit 0, files untouched; resumes at 07:30 | from about 20:00: "paused" (scheduled; "stale" without a meeting time zone), nothing imported, nothing erased |
| Laptop asleep | no new scans (the task does not wake the laptop); after waking, a missed run starts once, then every 2 minutes inside the window (paused outside it) | same until the next successful scan |
| User logged off / agent stopped | no new scans (the task runs only in the logged-on session); after logging on, the next 2-minute run inside the window scans | same |
| Drive for desktop paused or offline | writes locally, not uploaded | "stale" until the upload arrives |
| File half uploaded | — | "incomplete", the previous valid file (if any) is used |
| A group's folder missing on the server | that group fails (HTTP 404), others go on | "not checked" for that group |
| Two scans overlap | the second is skipped (mutex) | — |
| Drive folder not readable | — | "cannot be read", nothing erased |

## Security and privacy

The meeting-room server is only read, from inside the room; it is not changed or exposed. The agent
has no Google credentials of its own (Drive for desktop uploads). Only link names and URLs of the
listing leave the room — the same data the report already stores after a manual import. The reports
keep only links on the configured meeting-room server, of their own group. Share the Drive folder
read-only with the report accounts and nobody else.

## Removing it

Run `install-scheduled-task.ps1 -Uninstall`, clear *Local listing folder* in each report. Imported
revisions stay, as after a manual import.

Running the installer again is safe: it replaces the task in place (same name, same settings). To
be able to go back to a previous task definition, export it first,
`Export-ScheduledTask 'SA4 local listing agent' | Set-Content task-before.xml -Encoding Unicode`,
and keep a copy of the agent files. To restore, put those files back and run
`Register-ScheduledTask -Xml (Get-Content task-before.xml -Raw) -TaskName 'SA4 local listing agent' -Force`.
