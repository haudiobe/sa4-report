# Automatic local revision detection — feasibility (research, not released)

Branch `research/local-revision-autodetect`. Production is T-2026.10.27 (`b0bc1eb`) and is not
touched by anything here. Today a person in the meeting room pastes the folder listing of
`http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/<group>/` into *Import Local Inbox Listing…*
(docs/LOCAL_REVISIONS.md). Goal: the same result without copy and paste.

## 1. Integration point (existing code, reused unchanged)

`saveLocalRevisionsListing({ folder, html, text })` already does everything after the listing is
in hand: parse (`parseLocalRevisionsListing_()`, same-server links only), per-folder refresh and
store (`storeLocalRevisions_()`, `LOCALREV_<tdoc>`), merge with the 3GPP revisions
(`mergeLocalRevisions_()`) and render the Revisions cells. Any automatic source only has to supply
`{ folder, html }` -- the page as the server serves it. The update (`updateRevisions_()`) already
shows stored local revisions on every run.

## 2. Can the Apps Script dialog read the listing directly? No.

Measured on 2026-11-16 from a computer on the meeting-room network:

| Fact | Result |
|---|---|
| Server | `Microsoft-IIS/10.0`, own directory page, absolute hrefs (`http://10.10.10.10/...`) |
| HTTPS | none (`https://10.10.10.10/...` does not answer) |
| `GET` with an `Origin` of googleusercontent.com | `200`, **no** `Access-Control-Allow-Origin` |
| `OPTIONS` preflight with `Access-Control-Request-Private-Network` / `-Local-Network` | `200`, **no** `Access-Control-Allow-*` headers |

A dialog of Apps Script is an `https://…googleusercontent.com` page in nested iframes. For its
script to *read* the listing, three things would all have to hold: the request must be allowed
from https to http (mixed content: blocked; Chrome's Local Network Access only exempts it after a
permission prompt that a nested cross-origin iframe may not get), the browser must allow a public
page to reach a private address (Local/Private Network Access: needs the server's preflight
answer, which it does not give), and the response must carry `Access-Control-Allow-Origin`
(CORS: **it does not**). The third alone makes the content unreadable: at best an *opaque*
response (`mode: 'no-cors'`) proves the server was reached, never what it said. Navigating to the
URL in a tab works and says nothing about this.

Without changing the meeting-room server (CORS headers and HTTPS), direct reading is not possible.
`tools/local-revisions-poc/LocalAccessDiagnostic.js` runs every probe in the real dialog
(context, LNA permission, `fetch` with and without `targetAddressSpace`, `no-cors`, an image) and
prints a verdict; the browser-side confirmation is **unverified** until it is run in the room
(section 5), but the server headers above already decide it.

## 3. Alternatives

| | A. Chrome/Edge extension | **B. Local listing agent + Drive (recommended)** |
|---|---|---|
| How it reads | extension service worker with host permission `http://10.10.10.10/*` (extensions are not bound by CORS or mixed content) | PowerShell on one laptop: plain HTTP GET |
| How the report gets it | content script in the *Import* dialog's iframe → `postMessage` → `google.script.run.saveLocalRevisionsListing` | the agent writes `<group>.html` into a folder that Google Drive for desktop syncs; the report reads it with its existing Drive access |
| Automatic? | only while a person has the report and its dialog open, on every reporter's browser | **yes**: also in the time-driven automatic updates, for all five groups from one laptop |
| Install | an unpacked or policy-installed extension per browser; Google's iframe origins change (`n-…-script.googleusercontent.com`) | one script, one Scheduled Task, Drive for desktop on that laptop |
| Credentials | none new | none new (Drive for desktop is already signed in; no OAuth client, no token in the script) |
| Server change | none | none |
| Effort | ~2 days + per-browser rollout, fragile against Google UI changes | ~1.5 days (below) |

**Recommendation: B.** It reuses the import code unchanged, covers all five groups, works when
nobody is watching, and fails safe.

Failure behaviour of B (proven in the PoC tests): meeting-room network or server down → the agent
records the error in `<group>.status.json` and keeps the last good listing; the report imports
only a listing **younger than 30 minutes**, otherwise it says "the newest listing of MBS is from
…" and changes nothing; no listing yet → "not checked"; a file not written by the agent, or a
listing of another server → refused by the existing checks. Revisions already imported stay (they
are only refreshed per folder by a newer listing).

Access control: the Drive folder is shared only with the reporters (editors of the reports); the
agent writes nothing but the listing as served; the report keeps only links on the configured
meeting-room server. The listing contains file names of draft TDocs -- the same data the report
already stores after a paste. The meeting-room server is not exposed: nothing connects to it from
outside, the laptop only reads.

## 4. Proof of concept (this branch)

- `tools/local-revisions-poc/sync-local-listings.ps1` -- the agent (`-Once` or every
  `-IntervalSeconds`; atomic writes; status file per group).
- `tools/local-revisions-poc/LocalListingDriveImport.js` -- `pocImportLocalListingFromDrive()`:
  reads this report's group file (from `DRAFTS_FOLDERS`: MBS, Video, FS_6G_MED, RTC, Audio),
  checks the agent header and its age, calls `saveLocalRevisionsListing()`.
- `tools/local-revisions-poc/LocalAccessDiagnostic.js` -- the browser probe dialog (section 2).
- `tests/local-revisions-autodetect-poc.test.js` (13 checks): the import via the agent file gives
  **exactly** the cells a paste of the same page gives; idempotent; minutes untouched; nothing
  fetched; stale, missing, foreign and other-server files refused; each report takes only its own
  group.

Run on 2026-11-16: the agent read all five group folders of the 3GPP server as a stand-in
(`-BaseUrl https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Inbox/Drafts/`, 17–87 kB
each), and the real MBS folder of the meeting-room server (HTTP 200, 70 kB, 44 files). That real
MBS listing, imported through `pocImportLocalListingFromDrive()` into a sandbox report, gave 44
files / 44 revisions, the same lines the live MBS smoke test showed for S4-261930, and a revision
not seen before (`S4-261964r01.docx`, local only) -- the case the automation is for. A second run:
"no new revisions".

## 5. Manual tests in the meeting room (on a disposable copy only)

1. **Browser probe (optional; confirms section 2).** Push `LocalAccessDiagnostic.js` with the
   release files into a disposable copy. On the meeting-room network, open the copy, in
   *Extensions › Apps Script* run `localAccessDiagnostic()`, click *Run probes*, copy the text.
   Expected: `cors` and `cors+lna` blocked, VERDICT "CANNOT read".
2. **Agent.** On a laptop on the meeting-room network with Google Drive for desktop:
   `.\sync-local-listings.ps1 -OutDir "G:\My Drive\SA4-138 local listings" -Once`
   Expected: five lines "ok, … bytes" (a group whose folder does not exist on the server: FAILED,
   with the HTTP status). The folder appears in Drive within a minute.
3. **Report side.** In the disposable MBS copy: Document Property `LOCAL_LISTING_DRIVE_FOLDER_ID`
   = the Drive folder's ID (set in the editor, PoC only), push `LocalListingDriveImport.js`, run
   `pocImportLocalListingFromDrive()`. Expected: the same message and cells as *Import Local Inbox
   Listing…*; a second run "no new revisions"; with the agent stopped for 30 minutes: "not
   imported … from <time>".

## 6. Proposed next phase (needs approval)

1. Report side, in the release code (~1 day): a Configure Meeting field "Local listing folder
   (Drive)"; `updateRevisions_()` imports a fresh agent listing of the report's group before it
   renders (no lock change: the update holds it); *Check Connections* shows the agent's last
   success; tests as in the PoC, plus the update path.
2. Agent hardening (~0.5 day): Scheduled Task installer (`-IntervalSeconds 120`), log rotation,
   group list from a small config file.
3. Validation (~15 minutes in the room): steps 2–3 above on the disposable MBS copy, then a
   release and the usual one-at-a-time rollout with backups.

Not needed: any change to the meeting-room server, any new Google credential, any external
exposure.
