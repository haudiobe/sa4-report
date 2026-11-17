# Local listing agent — automatic local revisions (after T-2026.10.27)

The reports show revisions found on the meeting-room server (`http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/<group>/`)
next to the 3GPP ones (docs/LOCAL_REVISIONS.md). Google cannot reach that server
(docs/LOCAL_REVISIONS_AUTODETECT.md), so a small agent on one laptop in the room reads the
listings and hands them over through Google Drive. *Import Local Inbox Listing…* stays as it is.

```
meeting-room server  --HTTP GET-->  agent (laptop on the room network, every 2 min)
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

Its first line says the result of the last check: imported, unchanged, stale, missing, invalid, unavailable.

Check Connections also tries the TDoc list as the update does (the configured address, then its
`www.3gpp.org/ftp` twin); a 403 from 3GPP is shown as a refusal by the 3GPP server, usually temporary.

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
4. **Schedule.** `powershell -NoProfile -ExecutionPolicy Bypass -File C:\SA4Agent\install-scheduled-task.ps1 -ConfigPath C:\SA4Agent\agent.config.json`
   (every 2 minutes while you are logged on; no admin rights; log:
   `%LOCALAPPDATA%\SA4LocalListingAgent\agent.log`). Remove with `-Uninstall`.
5. **Each report.** *⚙️ Configure Meeting… › Local listing folder*: paste the folder ID or URL,
   Save. The next update (automatic or *Update TDoc Revisions*) imports; *Check Connections* shows it.

## Failure behaviour

| Situation | Agent | Report |
|---|---|---|
| Meeting-room network or server down | records the error, keeps the last good listing | after 30 min: "stale", nothing imported, nothing erased |
| Laptop asleep / logged off / agent stopped | no new scans | same |
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
