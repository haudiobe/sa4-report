# Revisions from the meeting-room server (issue #16)

During a physical SA4 meeting, revised TDocs can appear first on the meeting-room server
(`http://10.10.10.10/ftp/SA/SA4/Inbox`, meeting-room LAN only). The 3GPP drafts folder
(`…/TSGS4_<n>_<place>/Inbox/Drafts/<SWG>/`) is synchronized from it later. A report shows
the revisions of both, per TDoc.

## Why the local server is never fetched by the report

| Way | Works? | Why |
|---|---|---|
| `UrlFetchApp` in Apps Script | No | Runs in Google's data centres; `10.10.10.10` is a private (RFC 1918) address there. |
| `fetch()` from a report dialog or sidebar | No | The dialog is an `https://…googleusercontent.com` page: a request to `http://10.10.10.10` is mixed content and is blocked. Chrome's Private Network Access also blocks public→private requests unless the server answers the preflight, and the server sends no CORS headers. None of this is under our control. |
| A bridge (agent, proxy, tunnel) on a laptop in the room | Possible, not used | Has to be installed and kept running; exposes or relays the LAN. Not needed. |
| **The browser of someone in the room, as an ordinary page** | **Yes** | Opening `http://10.10.10.10/…` as a page is allowed. Its listing is copied and pasted into the report. |

So the local source is an **imported listing**: *SA4 Report › 📄 Report › Import Local Inbox Listing…*.

## What a person in the room does (about 20 seconds)

1. In the report: *SA4 Report › 📄 Report › Import Local Inbox Listing…*. The dialog shows when
   a listing was last imported, or "not checked yet".
2. Click the link in the dialog (opens `http://10.10.10.10/ftp/SA/SA4/Inbox`) and go to the
   folder with the group's revisions; copy its address from the address bar.
3. On that page: **Ctrl+A, Ctrl+C**. Back in the dialog: paste the address, click into the box,
   **Ctrl+V**, **Import**.

The dialog says how many files the listing had, how many are revisions of SA4 TDocs, and how
many are new for this report ("no new revisions for this report" when none). The Revisions
cells of the TDocs that got a new local revision are written at once; nothing else.

Repeat whenever useful. The automatic updates keep showing what was imported.

### Unavoidable prerequisites in the meeting room

- A computer **on the meeting-room network** whose browser can open `http://10.10.10.10/…`.
- In that browser, the user is signed in to Google with **edit access** to the report.
- The server shows a **folder listing** (an HTML page of links). Checked: links with
  relative, rooted (`/ftp/…`) and absolute hrefs, double- or single-quoted. The file name is
  taken from the link itself, so a name the listing shows shortened (`S4-2601..>`) is still
  complete, and each file keeps its actual URL.
- If a browser copies **text only** (no links), complete file names are still recognized, but
  no file path is made up: such a revision links to the **listed folder** and is marked
  `[local only, see folder]` (or `[local folder]`). The dialog says so; pasting the page again
  from Chrome or Edge gives file links. A later text-only paste never replaces a file link.
- Unknown until tried in the room: the folder layout under `Inbox` (e.g. whether there is a
  `Drafts/<SWG>` level). It does not matter to the import: any folder of the server can be
  listed, the root included; only files that are revisions of this report's TDocs are kept.

## What the report shows

In the Revisions cell of a TDoc (one line per file, ordered by name):

| Where the file is | Line | Links |
|---|---|---|
| 3GPP only | `S4-260123_r1.docx` | name → 3GPP file (unchanged) |
| both | `S4-260123_r1.docx [local]` | name → 3GPP file, `[local]` → local file |
| meeting-room server only | `S4-260123_r2.docx [local only]` | both → local file |
| from a text-only paste | `… [local only, see folder]` / `… [local folder]` | → the local folder |

Below them, the folder line of PR #15 (`Revisions folders: Local … · Remote …`).

- The same file in both places is one line: identity is the file name, decoded, case ignored.
  Two different names are two files, also if their revision numbers match.
- A local link is always the one the listing gave. It is never built from the 3GPP folder
  layout, and a link to another server than the listed folder's is dropped.
- When the 3GPP folder catches up, the next update links the name to 3GPP and keeps `[local]`.

## What does not change

- The 3GPP folder, its fetch, file matching (`parseDraftAnchor_()`), `REVIS_<tdoc>` and the
  revision tables: a revision seen locally only **never** creates a TDoc table.
- TDocs, titles, Minutes, Dispositions, Statuses, e-mails, abstracts, formatting.
- TDocs without a table in the report are not imported.
- A listing that cannot be read (empty paste, an error page, a folder on another server) or a
  report being updated at the moment changes nothing; the dialog says why.
- If the 3GPP folder is unreachable, the update leaves the cells as they are.

## Stored data

- `LOCALREV_<tdoc>` (Document Property, per TDoc with local revisions):
  `{ key: { text, link, folder, folderOnly? } }`. An adoption prefix (copied with the report state).
- **Refresh:** a listing is the current state of the folder it lists. Entries from that folder
  that it no longer has (renamed, withdrawn) are removed, and their lines leave the cells; entries
  from other folders stay. The same file is stored once. A listing without a single revision
  removes nothing (a wrong or partial paste). The dialog says how many were added and removed.
- **Hand-typed text in a Revisions cell is kept:** a line that is not a revision line,
  "No revisions available." or the folders line stays, below the revisions. A cell whose text
  would not change is not written again.
- `LOCAL_REVISIONS_LAST_IMPORT`: when, from which folder, counts. Shown by the dialog and by
  *Check Connections*.
- No credentials, no tokens. The local server sees only the user's own browser.

## Security and privacy

- Nothing on the LAN is exposed or relayed; nothing is installed.
- The pasted HTML is parsed for links only, never rendered or executed; only `http(s)` links on
  the listed server are kept; `javascript:`, `data:`, protocol-relative and foreign links are
  dropped.
- The private address appears in the report as link targets. It is useless outside the room;
  readers outside see `[local]` markers that do not open.

## Rollout to the five SA4#138 reports (after approval only)

The five live reports are bound scripts of their own (mapped in issue #16; the IDs are not
repeated here). Per the established procedure (docs/EXISTING_REPORT_MIGRATION.md):

1. Merge PR #15 and this PR into `master`; cut a release (`template-release/<id>`, `Code.js`
   version bump, `tools/template-release.js --release <id> --write`).
2. **One smoke check** on a disposable copy of one report (e.g. MBS):
   `tools/integration-test/push-to-disposable-copy.mjs --script <copy> --doc <copy> --write
   --allow-other-base`; in the room, Import Local Inbox Listing once, look at two cells, run
   *Update TDoc Revisions* once. About 3 minutes.
3. Per report, one at a time:
   - back up its script (`clasp pull` into `dist/backups/<label>-before-<id>/`, SHA-256 recorded);
   - `node tools/template-release.js --adopt <id> --target-script <script> --label <label> --write`;
   - `clasp push` from `dist/template-adopt/<id>/<label>/`, then pull back and compare the SHA-256.
   No document is touched by the push. The first automatic update afterwards adds the folder
   line (PR #15) to the Revisions cells: one write per cell, once.
4. Nothing to configure: a main meeting that is not an e-meeting gets
   `http://10.10.10.10/ftp/SA/SA4/Inbox` by default (Configure Meeting › Local Revisions URL;
   `none` switches it off).

## Rollback

- Script: push the backup of step 3 back (`push-to-disposable-copy.mjs --restore <dir>` for a
  copy; `clasp push` of the backup directory for a live script), pull back, compare hashes.
- Data: `LOCALREV_*` and `LOCAL_REVISIONS_LAST_IMPORT` are ignored by older code; they can stay.
- Cells: older code rewrites a Revisions cell from the 3GPP folder whenever it has revisions
  there, which removes the markers and local-only lines. A cell that has **no** 3GPP revision
  keeps its text (older code never overwrites such a cell): its local-only lines and folder
  line stay until edited by hand.
