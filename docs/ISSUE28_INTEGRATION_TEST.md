# Issue #28: integration test on disposable copies

Test tooling only. Nothing here is part of a release bundle.

## Copies

- **main2**: the clean copy of the 6G Media SA4#138 report (2026-11-25 16:18 UTC). It runs T-2026.10.33. Its bound script was registered on 2026-11-26 after a read-only first read: bound to this document, only its HEAD deployment, no versions, every file equal to the production T-2026.10.33 bundle.
- **main**: the first copy of the same report. Retired: updates killed at the 6-minute limit left a TDoc table twice in it. The tooling refuses it, and it is kept as evidence.
- **ahg**: a copy of the Legacy 6G AH report. It runs T-2026.10.19.

Both are private Google Docs named `DISPOSABLE …`. This repository is public, so the tooling knows the copies and the protected reports only by the SHA-256 digest of their IDs.

## Tools

| File | What it does |
|---|---|
| `tools/integration-test/push-issue28-copy.mjs` | Pushes one stage into a copy's bound script. Dry run by default; `--write` pushes. |
| `tools/integration-test/issue28-copy-rules.js` | The push tool's rules: digests, pins, compare-and-swap, stage files. Tested by `tests/issue28-copy-rules.test.js`. |
| `tools/integration-test/Issue28Check.js` | Editor-side functions. The Run button takes no arguments, so use `it28Preflight`, `it28Snap` (numbered snapshots s1, s2, … each compared with the previous one and with s1), `it28Snapshots`, `it28TimedUpdate` and `it28Candidates`. `it28Fingerprint(label)` and `it28Compare(a, b)` need arguments. Pinned by digest (`CHECK_SHA`). Tested by `tests/issue28-integration-check.test.js`. |

```
node tools/integration-test/push-issue28-copy.mjs --target main|ahg --script <Script ID> --stage <stage> [--write]
```

## Stages

| Stage | What is pushed |
|---|---|
| `base` | The copy's own script as first seen, plus `Issue28Check.js`. Report code is unchanged. |
| `t33` | The T-2026.10.33 code (b7299c1) and the production Release.js, plus `Issue28Check.js`. |
| `A`, `B`, `C` | The phase code at its pinned commit (A 03d7aa4, B 8e2d0f5, C 1ea0c06), plus `Issue28Check.js`. Release.js is marked "NOT RELEASED". |
| `original` | The copy's own script exactly as first seen, without `Issue28Check.js`. |

The sequence is:

- **main:** base → A → B → C → t33 → original
- **ahg:** base → t33 → A → B → C → t33 → original

## Safeguards

Each safeguard refuses before anything is written.

1. **Identity.**
   - The script must be the target copy's bound script, and its parent must be the target copy (checked by digest).
   - The parent must be named `DISPOSABLE` and must not be in the trash.
   - Neither may be the master template, one of the five SA4#138 reports, the Legacy 6G AH report, or the Legacy or CENTRAL project. Legacy and CENTRAL are checked by the digests in `tools/template-release.js`.
2. **Pinned original.**
   - On first sight, every file of the copy's script must equal the pinned original digest. The script is then saved to `dist/integration-test/i28-copies/<target>/original/`, which Git ignores.
   - On every later run, the saved original is checked against the pins again.
3. **Compare-and-swap.**
   - The copy's script must equal the last state the tool recorded (`state.json`), or the original before the first push.
   - This is checked when the run starts, and again, together with the identity, immediately before the push.
4. **Pinned commits.** A phase branch whose code differs from its pinned commit is refused. The pushed code always comes from the pinned commit.
5. **Read-back.**
   - After a push, the script is read back and compared file by file. `state.json` is updated only when everything matches.
   - Evidence is written to `dist/integration-test/i28-copies/<target>/<time>-<stage>/evidence.json`.

`it28Preflight()` reads only. It reports problems:

- triggers;
- scheduler registry or central state in the script properties;
- any protected report named in the script or document properties.

It lists foreign IDs for review, and notes a reference to the template.

## Snapshots (fingerprint v2)

Version 1 was replaced on 2026-11-25 after the first real runs. On the main copy, two reads of the same document were reported as different. It hashed objects from the Advanced Docs service exactly as they came, for example `JSON.stringify` of a colour, and the order of their keys is not guaranteed. On the AHG copy, the service's object graph ran out of memory. Version 1 also saw a Status dropdown chip only as the word "dropdown".

How version 2 works:

- **One read.** It makes one Docs REST call (`documents.get`, `includeTabsContent=true`) with a field mask that asks only for what is fingerprinted, and parses the JSON. The response text is released right after parsing. The Docs API cannot read part of a document; the field mask is what keeps the response small.
- **Normalization.** Keys are sorted recursively. Adjacent runs with the same style and link are merged. `startIndex`, `endIndex` and `headingId` are dropped. A link counts by its URL, or by the kind of internal target it points to.
- **Separate digests per location.** Each TDoc row, summary entry and other paragraph or table row gets digests for its text, its links, and its formatting. Formatting is recorded twice: as the sequence of styles, and with the positions where the styles change. Positions count only where the text is unchanged.
- **Chips** (dropdown, person, rich link, date) are text: their words plus a digest of their full content.
- **Comparison.**
  - TDoc tables are compared by TDoc number, with duplicates kept apart, and rows by label (`Discussion #2` for a repeated label).
  - Headings and the Disposition Summary are compared as sequences.
  - Content outside the TDoc tables is compared by its entries, each with an excerpt (600 characters) and the heading it sits under.
  - The order of the TDoc tables is reported separately.
- **Storage.**
  - The detailed snapshot is a gzip file in a private Drive folder named `DISPOSABLE Issue28 snapshots - <document>`.
  - The document properties hold only an index per snapshot: file, SHA-256, size, content SHA-256 and counts.
  - A snapshot whose file is missing, altered or incomplete is refused.
  - The full list of each comparison is stored next to the snapshots as text.
- **Logs.** Each run logs the read, parse, walk and store times. Two snapshots of an unchanged document give the same content SHA-256.

Measured in Node as a proxy, since Apps Script reports no memory figures:

| TDoc tables | Response | Heap after parse | Walk | Snapshot | as gzip |
|---|---|---|---|---|---|
| 1,500 | 7.6 MB | 17 MB | 2.9 s | 3.5 MB | 0.1 MB |
| 4,000 | 20 MB | 45 MB | 10 s | 8.9 MB | 0.4 MB |

The Apps Script platform limits that apply:

- UrlFetch responses up to 50 MB;
- 6 minutes per run;
- property values up to 9 KB, and 500 KB per store;
- no documented memory limit.

An independent review of version 2 led to these fixes before it was installed:

- A moved paragraph is reported as moved, with its old and new place.
- A changed label cell (its text or links) is reported.
- A document with more than one tab is refused instead of only partly fingerprinted.
- A text change in styled text is shown with style markers, such as `a «b»{bold} c`.
- The content SHA-256 leaves out the revision ID and is recomputed when a snapshot is loaded.
- Links to bookmarks keep their target.
- Cell separators are not mistaken for text.
- Chips are shown by their display value.

What the fingerprint does not see, by design. These are not changed by the report code under test.

- Table cell borders and padding, paragraph borders and tab stops, column widths and row heights.
- List glyph definitions (`document.lists`), named style definitions, and the size or content of inline images.
- Headers, footers and footnote content.
- Pending suggested style changes. Suggested insertions and deletions are seen.
- The target of a link to a heading. Heading IDs change when a heading is rebuilt, so only "link to a heading" is kept.

### Documents with tabs (2026-11-25)

The AHG copy has 4 tabs. Its first probe read 18.8 million characters in 16.9 s and parsed them in 0.17 s. The snapshot then refused it, because version 2 covered only a document of one tab. Every tab is now fingerprinted.

How it reads:

- **Structure first.** A small first request reads only the tab tree: tab ID, title, index and parent, for up to three levels. A tab nested deeper is refused, not left out.
- **Then one request per nesting level.** A field mask cannot recurse, and repeating the content mask for every level would exceed the URL limit. Each level request returns the content of that level's tabs only, at most 1,604 characters of URL.
- **Bounded memory.** Each tab is walked and released at once, so memory stays at one level's response at a time.

When the snapshot is refused (nothing stored):

- **The document changed during the read:** a level request returns another revision than the structure request.
- **A tab is missing:** a tab of the structure is not returned with its whole body.
- **An unknown tab appears:** a tab is returned that the structure did not list.

What it compares:

- **Tabs** are identified by tab ID, with their title, nesting path (`2.1` is the first child of the second tab) and parent. The comparison reports tabs added, removed or renamed, and tabs moved to another parent or another place among their siblings.
- **Content** is compared inside each tab, and every change names its tab. The same TDoc in two tabs is two separate tables.
- **Probe output:** `it28Probe()` lists each tab: TDoc tables, tables, paragraphs, whether it has a Disposition Summary, and its first headings.

## Automated validation (the runner)

Running editor functions by hand at every stage was too much work. The validation now runs by itself after a single editor action per copy.

### Why there is one manual step

There is no supported way to run code in a bound Apps Script project without it. These were checked on 2026-11-25:

- **`scripts.run` (Apps Script API).** It needs the script deployed as an API executable, the caller's OAuth client in the same Google Cloud project as the script, and a token with every scope the script uses. The only credential here is the clasp token. Its OAuth client belongs to clasp's own Cloud project, which is Google's, not the user's. Using `scripts.run` would mean creating a project and OAuth client, moving both copies onto it, and granting broader scopes. That is more manual work, not less.
- **Reading the document locally.** The Docs API is not enabled in clasp's project, and the token has no `documents` or full `drive` scope. Only Drive metadata and the script projects can be read.
- **A web app restricted to its owner.** Calling it needs a token with the full `drive` scope, which the clasp token lacks. A public endpoint is out of the question.

### How it works

1. **Once per copy.** Run `it28InstallRunner()` in that copy's editor. It installs a 5-minute time trigger in that copy only, which runs `it28RunnerTick()`.
   - The trigger runs as the user, with the copy's existing scopes. There is no new permission and no web app.
   - It removes itself after 72 hours, or when an uninstall job runs.
   - It refuses to run in any document not named DISPOSABLE, and in the template, the five reports or Legacy.
2. **Every stage after that** is automatic. `tools/integration-test/run-issue28-validation.mjs` does the following:
   - pushes the stage together with a job file (`Issue28Job.js`) through `push-issue28-copy.mjs --job`, with identity checks, the pinned original, compare-and-swap and read-back;
   - waits for the runner, which runs the job's steps from a fixed list: snapshot, compare, timed update, configure the meeting (Resolve and Discover from the 3GPP Portal, then Save, as the dialog does), save sessions, set test settings, the phase B steps (record and close an expected revision, Close Report check), the phase C steps (Meeting Control, rule and carry-forward previews and switches), and cleanup. A failed step that changes something stops the job;
   - reads the result from the descriptions of its Drive files, using Drive metadata only;
   - checks the result against the rules in `tools/integration-test/issue28-jobs.js`.
3. **The rules checked:**
   - Two snapshots of an unchanged document are identical.
   - A second update changes nothing (idempotency).
   - No update and no stage changes a manual row (Discussion, Decision, Comments, Minutes, Notes).
   - No update or stage removes text outside the TDoc tables, removes a TDoc table, or adds or removes a tab.
   - Switching Heading 2 class lines on or off changes only the Disposition Summary.
4. **Cleanup.** The uninstall job, and the runner itself at its 72-hour limit, remove every trigger of the project and every document and script property except the snapshot index; the copies had none before. The result lists what is left, as proof. The orchestrator runs this cleanup after any failure or timeout. It pushes the original script back only after the cleanup is verified, and the push tool refuses `original` unless the copy's last runner record is "uninstalled" or "expired".

### The plan

```
node tools/integration-test/run-issue28-validation.mjs --target main --script <id> --plan --write
node tools/integration-test/run-issue28-validation.mjs --target ahg  --script <id> --plan --write
node tools/integration-test/push-issue28-copy.mjs --target main --script <id> --stage original --write
```

- **main:**
  1. inventory and determinism at T-2026.10.33;
  2. configure the copy's meeting from the Portal: SA4#138 Calgary, portal 60780, main, 6G Media — a completed meeting, the copy only;
  3. the update stage at T-2026.10.33;
  4. A, with the Heading 2 on/off check;
  5. A+B;
  6. A+B+C;
  7. the rollback to T-2026.10.33;
  8. uninstall.
- **ahg:**
  1. inventory and determinism;
  2. configure "3GPPSA4-e (AH) on FS_6G_MED" (portal 86178, ad-hoc, 6G Media) with its two historical sessions from the meeting's agendas S4aP260098 and S4aP260149: 22 Sep 2026, and 26–28 Oct 2026, each 15:00–18:00. They are not taken from the document's tabs;
  3. the T-2026.10.33 stage, then A, A+B and A+B+C;
  4. phase C: Meeting Control, the reservation rule (only if its preview passes), carry-forward on, then both off;
  5. the rollback, then uninstall.
- **Settling first.** A completed report is not a fixed point of its own release. On 2026-11-26 the first T-2026.10.33 update of the clean main copy (production code, unchanged) added a TDoc registered after the meeting (S4-262470). It dropped the label "(revision of S4aP260150)" of S4-262018 and the meeting-room links (the copy has no meeting-room settings, and that agent stays off). It re-rendered 54 e-mail rows and changed the paragraph settings of two Minutes rows; their text was kept. The second update changed nothing. So each plan first runs the `settle` job: update until an update changes nothing, with no typed text of a manual row and no TDoc table lost, then mark that snapshot (`IT28_SNAP_SETTLED`). The stages A, A+B, A+B+C and the rollback compare with it, not with the untouched copy.
- **The 6-minute limit of T-2026.10.33.** A further T-2026.10.33 update of the settled main copy was ended by the 6-minute limit (2026-11-26; its updates of that copy took 316 s, 187 s, then more than 360 s, with no time budget). The main plan therefore takes its reference without another T-2026.10.33 update (`settled-check`: the copy must still be identical to the settled snapshot of run 1, s5). The rollback to T-2026.10.33 (`rollback`) runs an update ended by the limit once more, as the report's own trigger would, and its result says so. The candidate stages A, A+B and A+B+C keep the strict rule: an update ended by the limit fails the stage.
- **Tables by number.** The compare pairs TDoc tables by the TDoc number alone. A label after the number that changes is a change of the row "TDoc", and the table's other rows are still compared.
- **The Meeting Control benchmark** (`--job mc-benchmark`, ahg at A+B+C) does not use the Portal. It saves the settings Configure Meeting resolved for 86178 on 2026-11-25 through the report's own Save, then the two sessions. It sets the Status mark only if the copy's own Status entries are dropdowns, then runs Meeting Control twice and records its time per part. On 2026-11-25 the Portal's TDoc list for 86178 served another meeting's list. Meeting Control does not read that list.
- **main** also runs phase B on A+B (an expected revision whose successor never comes: it blocks Close Report, and is information once closed with a reason) and Meeting Control on A+B+C.

### What still needs a person or another decision

- The one-time `it28InstallRunner()` in each copy.
- Interactive dialogs (Revision Expected, Meeting Control, Review Title Changes) as a person sees them. Their server functions can be run as `call` steps; the dialogs themselves cannot.
