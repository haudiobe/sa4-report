# SA4-ARCH-005/006 — Real SA4 Ad-hoc Meeting Analysis

> This file now contains two appended sections: the original SA4-ARCH-005 evidence (below), and
> the SA4-ARCH-006 follow-up evidence (Plenary/6G deep dive, a second Audio SWG report, and the
> TDoc-identifier-family survey) appended at the end under "SA4-ARCH-006 ADDENDUM". Design
> recommendations built on this evidence (agenda-selection model, profile architecture,
> MeetingContext v2, implementation sequence) were delivered in the SA4-ARCH-006 chat response
> rather than duplicated here, to keep this file as the evidence record.

# SA4-ARCH-005 — Real SA4 Ad-hoc Meeting Analysis

Evidence-based analysis of real, currently-live 3GPP SA4 ad-hoc meeting material, gathered to
inform the design of a future `meeting.type = 'adhoc'` MeetingContext profile. No production
code was modified for this task.

**Every claim below is marked verified (checked against a real, fetched 3GPP document/directory
listing) or inferred (a reasonable conclusion I did not directly confirm). Where I could not
confirm something at all, it is marked unknown.**

All source URLs are under `https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/` unless
otherwise noted. Files were downloaded and parsed directly (XLSX/DOCX are ZIP containers of XML;
I unzipped them and read the raw XML text) rather than relying on any summarization step, so the
quoted text and column names below are exact.

---

## 1. Ad-hoc meetings investigated

### Primary example: SA4 Audio SWG Ad-Hoc (AH) on ULBC-MED

- **Official name** (verified, from the agenda TDoc's own title field): "Meeting agenda for
  SA4-Audio SWG AH" / TDoc-list filename `TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx`
- **Date** (verified, from the agenda document body): physical meeting **28-30 September 2026**;
  TDoc submission deadline **22 September 2026, 23:59 CEST**. This meeting was still upcoming
  (not yet held) at the time of this analysis (today: 2026-09-22).
- **Meeting identifier**: verified — there is **no** `TSGS4_<number>`-style identifier. The
  meeting is identified purely by topic name ("ULBC-MED", a work item) and by its TDoc series
  prefix `S4aA26` (see §3).
- **FTP/archive folder** (verified): `.../3GPP_SA4_AHOC_MTGs/SA4_Audio/`
- **3GPP Portal meeting ID**: **unknown** — portal.3gpp.org/Meetings.aspx requires authenticated
  login; I could not confirm whether this ad-hoc is registered there with its own `MtgId`, and
  general web search found no public page for it. Do not assume one exists.
- **Meeting type/topic** (verified, from the agenda body): "The meeting is dedicated to
  ULBC-MED"; SWG-level ad-hoc, single work item.
- **Physical/electronic/hybrid** (verified, exact quote from the agenda): "a physical SWG meeting
  with the possibility for remote participation (best efforts)".

### Comparison example 1: SA4 Audio SWG Ad-Hoc #1 "post 137-e" (DaCAS)

- **Name** (verified, filename): `TDoc_List_Meeting_SA4-e (AH1) Audio SWG post 137-e_1.xlsx`,
  dated 2026-09-21.
- Electronic ad-hoc, identified **relative to the preceding main meeting** ("post 137-e") rather
  than by a standalone topic name (though its content is also DaCAS-focused).
- All 11 TDocs in this list (`S4aA260051`-`S4aA260061`) sit under a **single flat agenda item
  "1.5"** — narrower in scope than the ULBC-MED example.

### Comparison example 2: SA4 Plenary ad-hoc series (6G)

- **Folder** (verified): `.../3GPP_SA4_AHOC_MTGs/SA4_Plenary/`, TDoc prefix `S4aP` (verified from
  filenames like `S4aP260074.zip`).
- Contains ad-hoc TDoc lists such as `TDoc_List_Meeting_SA4-Ad Hoc group call on FS_6G_MED.xlsx`
  and `TDoc_List_Meeting_SA4-e(AH) post 136.xlsx` — confirms 6G/`FS_6G_MED` ad-hoc activity exists
  under its own series, using the same `FS_6G_MED` work-item name the current `DRAFTS_FOLDERS['6G']`
  constant already uses. **I did not open an agenda or report document from this series** — its
  internal agenda/report structure is not directly verified, only its existence and naming.

Not investigated in this task (out of scope for time): `SA4_RTC`, `SA4_MTSI`, `SA4_EVS`,
`SA4_MBS`, `SA4_SQ`, `SA4_VIDEO` ad-hoc series, `Rel-20_planning`. These folders exist (see
directory listing below) but were not opened.

---

## 2. FTP/archive structure comparison

Directory listing of `.../WG4_CODEC/3GPP_SA4_AHOC_MTGs/` (verified):

```
Agenda/            (2021 — stale/legacy)
Invitation/        (2022 — stale/legacy)
Rel-20_planning/   (2025-06-27)
SA4_Audio/         (2026-07-10)  <- primary example
SA4_EVS/           (2021)
SA4_MBS/           (2021)
SA4_MTSI/          (2022)
SA4_Plenary/       (2026-01-16)  <- 6G-related ad-hoc series
SA4_RTC/           (2022)
SA4_SQ/            (2021)
SA4_VIDEO/         (2021)
```

`SA4_Audio/` itself contains (verified): `Agenda/`, `Docs/`, `Inbox/` (with an `Inbox/Drafts/`
subfolder), and a `Readme_Audio.txt`. This is structurally analogous to a main meeting's
`Docs/` + `Inbox/Drafts/<SWG>/` layout, **except there is no per-meeting-instance subfolder** —
`SA4_Audio/Docs/` is one continuously-reused folder holding TDocs from every Audio-SWG ad-hoc
instance since the group's creation (2022 onward), not a folder scoped to a single dated meeting.

Compared to a normal meeting (`TSGS4_136_Montreal`):

| Question | Finding |
|---|---|
| Does `TSGS4_<number>_<location>` still apply? | **No** (verified). Ad-hoc folders are named by SWG (`SA4_Audio`, `SA4_Plenary`, ...), not by a numbered/located meeting. |
| Is there an `AdHoc`/`AH`/`bis`/topic identifier? | **Yes** (verified) — but it appears in the *TDoc-list filename*, not the folder name: `TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx`, `TDoc_List_Meeting_SA4-e (AH1) Audio SWG post 137-e_1.xlsx`. Note: `TSGS4_135-bis-e` (a "bis" numbered interim meeting, separate from `3GPP_SA4_AHOC_MTGs`) is a genuinely different, main-meeting-style entity — "bis" meetings are **not** the same thing as the `AHOC_MTGs` ad-hoc series and were not analyzed further here. |
| Does `TDoc_List_Meeting_SA4#NNN.xlsx` still exist? | **No** (verified) — see the three real filenames above; none contain a bare `#NNN` in that position, all instead encode the ad-hoc's identity (topic name or "post NNN-e") into the filename. |
| Can current `TDOC_LIST_URL` construction work? | **No, not via auto-detection** (verified). `getReportConfig_()` builds `${FTP_BASE}TDoc_List_Meeting_SA4%23${MEETING_NUMBER}.xlsx` — this pattern matches none of the 7 real ad-hoc TDoc-list filenames found. The existing manual-override field (`TDOC_LIST_URL` document property) *would* work if set explicitly to the real URL — this is an existing escape hatch, not a fix. |
| Can the current `FTP_BASE` model work? | **Partially** (verified/inferred). `FTP_BASE` as "a folder containing TDoc ZIPs" is conceptually fine — `SA4_Audio/Docs/` plays that role — but it is NOT `${FTP_BASE_ROOT}${MEETING_FOLDER}/Docs/` with `MEETING_FOLDER='TSGS4_136_Montreal'`-style naming; it would need to resolve to `3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/` instead, which the current formula cannot construct from a `MEETING_NUMBER` at all (there isn't one). |

---

## 3. TDoc-list comparison

Full, exact column-header row from the real ULBC-MED ad-hoc TDoc list (verified, parsed directly
from `xl/worksheets/sheet1.xml` inside the downloaded `.xlsx`):

```
TDoc | Title | Source | Contact | Contact ID | Type | For | Abstract | Secretary Remarks |
Agenda item sort order | Agenda item | Agenda item description | TDoc sort order within agenda item |
TDoc Status | Reservation date | Uploaded | Is revision of | Revised to | Release | Spec | Version |
Related WIs | CR | CR revision | CR category | TSG CR Pack | UICC | ME | RAN | CN |
Clauses Affected | Reply to | To | Cc | Original LS | Reply in
```

Against `downloadAndGroupTdocs_()`'s assumptions (`Code.js`, header lookups):

| Column code expects | Real ad-hoc list has | Verdict |
|---|---|---|
| `TDoc` | `TDoc` | exact match |
| `Title` | `Title` | exact match |
| `Source` | `Source` | exact match |
| `Contact` | `Contact` | exact match |
| `Agenda item` | `Agenda item` | exact match |
| `Agenda Topic` | **`Agenda item description`** (not `Agenda Topic`) | **mismatch** — `headers.indexOf('Agenda Topic')` returns -1; code already tolerates `agendaTopicCol === -1` gracefully (it's only used for display), so this degrades safely rather than crashing |
| `TDoc Status` | `TDoc Status` | exact match |
| `Type` | `Type` | exact match |
| `For` | `For` | exact match |
| `Revised to` | `Revised to` | exact match |

**Classification for the existing TDoc-list parser: Reusable with configuration.** The column
structure is close enough that `downloadAndGroupTdocs_()`'s header-lookup mechanism would work
almost unchanged against a real ad-hoc TDoc list (one non-critical column name differs and
degrades safely). The actual blocker is not the column model — it's the **agenda-item prefix
filter** applied after parsing (see §4/§7): real ad-hoc agenda items are `1`, `2`, `3`, `4`,
`4.1`-`4.10`, `5`, `6` (verified below), none of which start with any current
`getAgendaPrefixForReportType_()` value (`5.`, `7.`, `8.`, `9.`, `10.`, `11.`, `18.`).

### TDoc identifiers — the `S4-\d{6}` assumption

**Verified, and this is the single most concrete, highest-confidence finding of this task**: ad-hoc
TDoc numbers do **not** match `S4-\d{6}`.

Real TDoc numbers observed (verified, from the parsed XLSX rows): `S4aA260051`, `S4aA260052`, ...,
`S4aA260100`; and from the Plenary/6G series (verified, filenames only): `S4aP260061`, `S4aP260074`.

This is not an inference from the filename pattern alone — it is **officially documented by 3GPP
itself**. `SA4_Audio/Readme_Audio.txt` (verified, fetched directly) states:

> "The prefix designation for adhoc meetings changed to 'S4aA22' with the counter reset to one."

I.e., 3GPP's own SA4 Audio SWG leadership assigned `S4aA` as the deliberate ad-hoc TDoc prefix
(year-coded, so `S4aA26` for 2026), distinct from the `S4-` prefix used for main-meeting TDocs.
The Plenary/6G ad-hoc series independently uses `S4aP`, confirming this is a per-SWG-series
convention, not specific to Audio.

**Concrete consequence for the current code** (verified against `Code.js`):
- `normalizeTdoc_()` (`Code.js:5175-5178`): `/S4-\d{6}/` does not match `S4aA260090` — returns
  `''`.
- `extractTdocId_()`'s default fallback regex (`Code.js:2817-2821`, `'S4-\\d{6}'`) and the
  `TDOC_ID_REGEX` document-property default (`'^S4-\\d{6}$'`) — same mismatch.
- `HyperLink.js`'s hardcoded `/^S4-26\d{4}$/` — same mismatch (this file is not called from
  anywhere in `Code.js`, per SA4-ARCH-001, so its impact is limited, but it shares the same
  wrong assumption).
- Direct, verifiable consequence: `parseAgendaFromZippedTdoc('S4aA260090', ftpBase)` would call
  `normalizeTdoc_('S4aA260090')`, get `''`, and immediately `throw new Error('Invalid agenda TDoc.
  Expected format like S4-260868.')` — **the real agenda TDoc number for this exact meeting
  would be rejected by the existing "Parse Agenda Document" dialog before any network request is
  even made.**

### Cross-numbering-scheme revisions (verified, from real data)

Row 1 of the "post 137-e" DaCAS ad-hoc TDoc list has `TDoc = S4aA260051` and
`Is revision of = S4-261671` — i.e. an ad-hoc TDoc referencing a **main-meeting-numbered** TDoc in
its "Is revision of" column. This is verified data but its exact meaning (ad-hoc doc later
resubmitted at the main meeting under a new number, vs. some other relationship) is **inferred**,
not confirmed from a written explanation.

Same-scheme (ad-hoc-to-ad-hoc) revisions **do** occur: the real chair's report
`S4-26xxxx Report from Audio SWG teleconference on DaCAS (27 July 2026).docx` (verified, opened
and read directly) contains the line:

> "Decision: S4aA260049 is revised to S4aA260050"

This is a same-scheme (`S4aA` → `S4aA`) revision recorded in **narrative report text**, not in the
XLSX "Revised to" column (I did not find a populated `Revised to` cell demonstrating this same
case at the spreadsheet level — that specific combination is **unknown/unverified**). If it does
occur at the spreadsheet level too, `getRevisedTo_()`'s use of `extractTdocId_(text,
cfg.TDOC_ID_REGEX)` with the default `^S4-\d{6}$` pattern would fail to extract an `S4aA`-numbered
target, the same way `normalizeTdoc_` fails above.

---

## 4. Actual agenda structure (verified)

Full agenda tree of the real ULBC-MED ad-hoc, extracted directly from
`S4aA260090 Meeting agenda for SA4 Audio SWG AH.docx` (the actual agenda TDoc for this meeting):

```
1     Opening of the meeting: Monday 28th September 2026 at 09:00 hours local time
2     Approval of the agenda and registration of documents
3     IPR, antitrust and consensus principles reminder
4     ULBC-MED
  4.1   Reports/Liaisons from other groups/meetings
  4.2   Project plan/timeline
  4.3   Performance requirements
  4.4   Design constraints
  4.5   Qualification/Selection rules
  4.6   Deliverables
  4.7   Processing plan
  4.8   Test plan
  4.9   RTP header compression & system integration
  4.10  Other
5     Any Other Business
6     Close of meeting: Wednesday 30th September 2026 at 16:00 hours local time (at the latest)
```

The TDocs in the (different but related) DaCAS ad-hoc TDoc list were seen under agenda items `2`,
`4.4`, `4.5`, `4.9`, `4.10` — consistent with this tree shape (verified cross-check between the
agenda document and the ULBC-MED TDoc list, which had TDocs at exactly those item numbers).

Against the current normal-meeting assumptions:

| Question | Finding |
|---|---|
| Is there an SWG prefix concept (`7.`, `9.`, ...)? | **No** (verified). Numbering starts flat at `1`. |
| Is there one report for the whole ad-hoc? | **Inferred, likely yes** for this SWG — the chair's-report example (§5) is headed "A.I. 1 Audio SWG" as a single document covering the whole session, not split per topic. Not verified for a multi-topic ad-hoc (this example is single-topic). |
| Are there multiple topic reports? | Not observed here (single-topic ad-hoc) — **unknown** for a multi-topic ad-hoc, if any exist. |
| Are agenda numbers flat or hierarchical? | **Both**: flat at the top level (`1`-`6`, no SWG prefix), hierarchical below the topic item (`4.1`-`4.10`) — verified. |
| Would `getAgendaPrefixForReportType_()` work? | **No** (verified) — for `'Audio'` it returns `'7.'`; nothing in the real tree starts with `'7.'`. |
| Would current agenda filtering return useful items or zero items? | **Zero items** (verified by direct comparison — `agendaItem.startsWith('7.')` is false for every real item number `1`-`6`/`4.x`). |

---

## 5. Report/minutes structure (verified)

Real ad-hoc reports **do exist**, but not in the `Docs/` folder — they live as `.docx` files in
`SA4_Audio/Inbox/Drafts/`, consistently named as **drafts of a future main-meeting TDoc**, e.g.:

- `S4-26xxxx Report from Audio SWG teleconference on DaCAS (27 July 2026).docx`
- `Draft S4-xxxxxx Report from Audio SWG AH on FS_ULBC (April 28).docx`

The placeholder `S4-26xxxx`/`S4-xxxxxx` (verified, literal filename text) is deliberate: 3GPP's own
practice is that an ad-hoc/teleconference report becomes an **information TDoc submitted to the
next SA4 main plenary meeting**, receiving its real `S4-NNNNNN` number only at that point. This is
consistent with, and explains, the `S4aA260051` → `S4-261671` cross-scheme reference noted in §3.

I opened one such report in full
(`S4-26xxxx Report from Audio SWG teleconference on DaCAS (27 July 2026).docx`, verified) — a
smaller teleconference, not the ULBC-MED AH itself (that meeting is still upcoming and has no
report yet). Its structure:

```
[document header fields: Source / Title / Document for / Agenda item]
Executive summary  (free-text paragraph)
A.I. 1   Audio SWG
A.I. 1.1   Opening of the session and registration of documents
A.I. 1.2   IPR, antitrust & consensus principles reminder   [full boilerplate text present, near-
                                                               identical in wording to the standard
                                                               3GPP IPR reminder]
A.I. 1.3   Reports/Liaison from other groups/meetings
A.I. 1.4   Release 19 and earlier matters
A.I. 1.5   DaCAS (Diverse audio CApturing system for Smartphone devices)
             [per-TDoc entries: number, title, source, agenda item, Presenter:, Discussion:
              (free-text transcript-style minutes), Decision: (e.g. "agreed" / "noted" /
              "S4aA260049 is revised to S4aA260050")]
```

Comparison against the current generated structure:

| Current generated section | Present in real ad-hoc report? |
|---|---|
| Opening | **Yes**, but combined with Registration into one item (`A.I. 1.1`), not split (6G splits them into `X.0.1`/`X.0.2`) |
| Registration of Documents | **Yes** — folded into `A.I. 1.1`, not a separate heading |
| Document Reallocations | **No** (verified — no such section anywhere in the agenda or the report) |
| IPR / antitrust reminder | **Yes** — `A.I. 1.2`, full boilerplate text present, wording close to what the main-meeting template likely contains |
| Topic agenda sections (TDoc tables) | **Yes** — `A.I. 1.5` in this report, `4.x` in the ULBC-MED agenda; each TDoc gets a real per-document entry |
| AOB | **Yes** in the agenda (item `5`); not yet reached in the report excerpt I read |
| Close of session | **Yes** in the agenda (item `6`) |
| Summary table | Not directly observed as a formal table; the agenda text says "The agenda with Tdoc allocation is agreed (see Annex A)" — **likely present as an annex, not verified** |
| Individual TDoc tables | **Yes, but richer**: real entries have `Presenter:` / `Discussion:` (full narrative) / `Decision:` — conceptually maps onto the current schema's `Minutes` (~Discussion) and `Disposition`/`Status` (~Decision), but the real content is a much longer narrative transcript than the current template's single-line fields anticipate |
| Minutes | **Yes**, as `Discussion:` narrative text |
| Disposition | **Yes**, as `Decision:` text (e.g. "S4aA260048 is agreed") |
| Status | **Implicit** in the `Decision:` text rather than a separate field — the real report doesn't tabulate status the way the XLSX `TDoc Status` column does |
| Revision handling | **Yes** — same-scheme revisions recorded in `Decision:` text (see §3) |

**Two genuinely new sections not present in current code's structure at all**:
`Reports/Liaison from other groups/meetings` (`A.I. 1.3`) and `Release 19 and earlier matters`
(`A.I. 1.4`). Both sit between the IPR reminder and the topic section, every time, in this one
example — whether they're truly fixed boilerplate for all Audio-SWG ad-hoc reports, or specific to
this particular teleconference, is **inferred, not verified** across multiple reports (I only
opened one).

**For the `ReportSkeletonBuilder` design question**: this evidence points toward "a separate
ad-hoc profile with its own skeleton", not "the main-meeting skeleton with configurable sections
toggled off". The front-matter is similar in *spirit* (opening, IPR, topic, AOB, close) but
different enough in *shape* (combined items, flat top-level numbering, two extra sections, no
reallocation concept) that reusing `buildSkeletonWithTdocTables()`'s branching with new
config flags would likely be messier than a dedicated ad-hoc skeleton function sharing only the
generic building blocks (`createSummaryTable_`, the TDoc-table renderer, `orderTdocsByRevision_`).

---

## 6. Mailing list and Drafts/revisions behavior

| Question | Finding |
|---|---|
| Which SA4 mailing list was used? | **Unknown** — not verified. I did not access the ETSI mailing-list archive/RSS feed directly for this ad-hoc series; `Readme_Audio.txt` documents TDoc-numbering policy but not a mailing list name. |
| Does the report-type → mailing-list mapping still make sense? | **Inferred, plausibly yes** — since the ad-hoc series is organized by the same SWG identity (`SA4_Audio`, `SA4_Plenary`, ...) as main meetings, reusing `MAILING_LISTS['Audio'] = '3GPP_TSG_SA_WG4_AUDIO'` is a reasonable guess, but **not confirmed**. |
| Where are draft/revised documents stored? | **Verified**: `SA4_Audio/Inbox/Drafts/` — a **flat** list (111 items, no subfolders), confirmed directly. |
| Do existing `DRAFTS_FOLDERS` conventions apply? | **Partially, and differently** (verified). Main meetings need `DRAFTS_FOLDERS[reportType]` to pick a SWG-named subfolder *within one shared meeting's* `Inbox/Drafts/`. Here, each ad-hoc series (`SA4_Audio`, `SA4_Plenary`, ...) already has its **own** `Inbox/Drafts/` — there is no shared parent to disambiguate, so the lookup table isn't needed at all for ad-hoc; the SWG selection is already implied by which ad-hoc series folder you're in. This is a structural simplification relative to main meetings, not a complication. |
| Does revision-chain processing remain usable? | **Partially** (verified/inferred combination) — `orderTdocsByRevision_()`/`orderRevisionChains_()` themselves are pure string-matching functions and don't care about the `S4-`/`S4aA` distinction (confirmed by reading their implementation — they operate on whatever string is in the TDoc/Revised-to columns). The blocker is upstream: `getRevisedTo_()` extracts that string via a regex (`TDOC_ID_REGEX`) that defaults to `S4-\d{6}` and would not recognize an `S4aA`-prefixed target if one appeared in a "Revised to" cell (unverified whether that specific case occurs at the spreadsheet level; verified that it occurs in narrative report text). |

---

## 7. Current-code compatibility matrix (conceptual walkthrough, verified against real data — no code executed against live data)

| Function | Verdict | Why |
|---|---|---|
| `getReportConfig_()` | **WORK WITH CONFIGURATION** | Runs fine mechanically; its *auto-construction* of `TDOC_LIST_URL` (`TDoc_List_Meeting_SA4%23NNN.xlsx`) does not match any real ad-hoc filename (verified, §2) — a user would have to supply the real URL via the existing `TDOC_LIST_URL` manual-override property. |
| `getAgendaPrefixForReportType_()` | **PRODUCE WRONG STRUCTURE** (functionally equivalent to FAIL downstream) | Returns `'7.'` for Audio; verified that zero real ad-hoc agenda items start with `'7.'`. The function itself doesn't error — it just returns a value that is silently wrong for this input. |
| `downloadAndGroupTdocs_()` | **WORK WITH CONFIGURATION** for column parsing; **effectively FAIL** as currently wired | Verified: header lookups for TDoc/Title/Source/Contact/Agenda item/TDoc Status/Type/For/Revised to all succeed against a real ad-hoc list. But `agendaItem.startsWith(agendaPrefix)` (`agendaPrefix` from the function above) filters out every real row, so `groups` comes back empty. Setting `AGENDA_ITEM_PREFIX` to `''` via the existing property override is a theoretical escape hatch (`startsWith('')` is always true in JS) but is not a real fix — see next row. |
| `parseAgendaForReport_()` | **FAIL** (with default config) or **PRODUCE WRONG STRUCTURE** (with the `AGENDA_ITEM_PREFIX=''` hack) | Same prefix-filter problem as above, applied to agenda-structure parsing. Even with the empty-prefix hack, downstream section-matching in `buildSkeletonWithTdocTables()` builds expected section numbers as `${agendaPrefixNum}.0.1` etc., which becomes a malformed string (`.0.1`) when `agendaPrefixNum` is empty — so the hack does not actually produce a working ad-hoc report, just a differently-broken one. |
| `buildSkeletonWithTdocTables()` | **FAIL** | With realistic configuration (an ad-hoc `AGENDA_TDOC`/`TDOC_LIST_URL`, default `AGENDA_ITEM_PREFIX` for a report type), `parseAgendaForReport_()` returns zero matching items, and the function bails with its own "No agenda items found for prefix: 7." alert (`Code.js`, verified logic path) — **no report is produced at all**, which is arguably the safest of the possible failure modes (loud, not silent). |
| `generateReportTitle_()` | **WORK WITH CONFIGURATION** | Its `/SA4%23(\d+(?:-e)?)/i` regex does not match any real ad-hoc `TDOC_LIST_URL` (verified, §2) and falls back to `SA4#{MEETING_ID}` — produces *a* title (no crash), but a generic one that discards the ad-hoc's real identity ("ULBC-MED", "post 137-e", the actual dates). Not wrong so much as uninformative. |

---

## 8. Minimum MeetingContext additions for `meeting.type = 'adhoc'`

Starting from the current shape (`group`, `meeting: {type, folder, number, portalId}`,
`report: {type, agendaPrefix, structureProfile}`, `sources: {...}`, `options: {...}`), evaluated
strictly against what the real evidence above requires:

| Candidate field | Evidence-based verdict |
|---|---|
| `meeting.number = null` | **Needed.** Verified: no real ad-hoc has a `TSGS4_<number>`-style identifier. |
| `meeting.name` (free-form string) | **Needed**, and should be a plain string, not further decomposed. Verified: the real, official identity of an ad-hoc *is* a free-form name (the TDoc-list filename / agenda-document title *is* the meeting's name, e.g. `"SA4-Audio SWG AH on ULBC-MED"` or `"SA4-e (AH1) Audio SWG post 137-e"`). This single field is the minimum needed to identify the meeting when `meeting.number` is null. |
| `meeting.topic` (separate from `meeting.name`) | **Not evidenced as necessary.** The "topic" (e.g. "ULBC-MED") is already embedded in `meeting.name`, and I found no place in the current code, or in the real documents, where topic is used as a *separate*, structured value independent of the full name. Recommend deferring this field until a concrete consumer needs to isolate it. |
| `meeting.relatedMainMeeting` (e.g. `"137-e"`) | **Optional, not strictly minimal.** Some ad-hocs ("post 137-e") carry this relationship, others (topic-only ad-hocs like ULBC-MED) don't. Worth keeping in mind for SA4-ARCH-006 but not required for a first ad-hoc profile — omit for now rather than add a field only half the real examples populate. |
| `report.agendaPrefix = null` | **Needed.** Verified: there is no meaningful SWG-plenary-style prefix in a real ad-hoc agenda (flat, starts at `1`). Making this explicitly `null` (rather than today's silent `'11.'` fallback for unrecognized types) turns a currently-silent wrong-default into a visible, intentional "this doesn't apply" signal. |
| `report.structureProfile = 'adhoc-flat'` | **Directionally right, but the name undersells the evidence.** The real structure (§5) is flat *only* relative to main-meeting SWG-prefixing — it still has real hierarchy (`4.1`-`4.10`), a fairly standardized front-matter template (combined Opening+Registration, IPR, plus two sections — Reports/Liaison, Release-matters — that don't exist in the current main-meeting skeleton at all), and no Reallocations concept. A name like `adhoc-swg` (paralleling `main-swg`, but capturing that it's a genuinely different template) is closer to the evidence than `adhoc-flat`. **This needs a second real report example (ideally from the Plenary/6G ad-hoc series) before finalizing a name — see §13.** |
| `sources.ftpBase` / `sources.tdocListUrl` / `sources.agendaTdoc` | **Reusable as-is**, but ad-hoc values won't fit the current auto-construction formulas (§2) — these fields stay the same *shape*, only their population needs a different construction rule (or manual entry) for ad-hoc. No new field needed, just a different value-producing path. |
| `sources.mailingList` / `sources.draftsFolder` | **Reusable as-is for mailingList** (unverified but plausible, per §6); **`draftsFolder` likely unnecessary for ad-hoc** since each ad-hoc series already has its own un-shared `Inbox/Drafts/` (§6) — the lookup table's reason for existing (disambiguating SWGs sharing one meeting's Drafts folder) doesn't apply here. |

**Fields I explicitly did NOT find evidence to add**: nothing beyond the above. In particular, I
found no evidence requiring new fields for "physical vs electronic vs hybrid" as a separate
MeetingContext field (it appears only as descriptive agenda text, not consumed anywhere in
`Code.js` today) or for the CR-Pack/UICC/ME/RAN/CN columns visible in the real TDoc list (those
are generic 3GPP CR-tracking columns, unrelated to the ad-hoc/main distinction, and the current
code already ignores them for main meetings too).

---

## 9. `report.type` vs `meeting.type` — independence

**Recommendation: they are independent and should compose, exactly as the task's example
(`meeting.type: 'adhoc'`, `report.type: 'Audio'`) suggests.**

Evidence: every ad-hoc series found is organized by the *same* SWG/report-type taxonomy used at
main meetings — `SA4_Audio`, `SA4_Plenary` (6G), `SA4_RTC`, `SA4_MTSI`, `SA4_EVS`, `SA4_MBS`,
`SA4_SQ`, `SA4_VIDEO` (verified, directory listing, §2) mirror `Audio`/`6G`/`RTC`/.../`Video`
almost one-to-one. The Audio ad-hoc's own report is headed "A.I. 1 **Audio SWG**" (verified, §5) —
the SWG identity is explicitly retained and stated inside the ad-hoc report itself, not discarded.

This means `report.type`'s existing enum (`6G`/`Audio`/`Video`/`MBS`/`RTC`/`Liaison`/`New`) should
be **reused unchanged** for ad-hoc meetings — an Audio ad-hoc is `report.type = 'Audio'` with
`meeting.type = 'adhoc'`, not a fundamentally different report-type concept. What changes between
main and ad-hoc is *how that report type's structure is built* (`report.structureProfile`,
`report.agendaPrefix`), not *whether* the SWG-identity concept applies.

---

## 10. Proposed profile model

**Recommendation: the second shape offered in the task** — two independent axes, not one flat
enum — is the better fit for the evidence:

```text
meeting.type:
    main
    adhoc

report.structureProfile:      (already exists as of SA4-ARCH-003/004; extend, don't replace)
    main-6g
    main-swg
    main-other
    adhoc-swg        <- new, tentative name; needs a second real report example to confirm (§13)
```

Rejected alternative (`main-6g` / `main-swg` / `main-other` / `adhoc` as one flat list): rejected
because it would either lose the `report.type` (SWG identity) axis for ad-hoc meetings entirely
(contradicted by §9's evidence) or require a combinatorial value per SWG-type-and-adhoc combination
(`adhoc-audio`, `adhoc-video`, ...) when the real difference between ad-hoc report types, as far as
verified so far, is primarily in the *front-matter template* (Opening/IPR/Reports-Liaison/Release-
matters), which is plausibly shared across ad-hoc SWGs — but this is itself unverified beyond the
one Audio-SWG example (§13).

---

## 11. Remaining unknowns

Explicitly not inferred, left as open questions for a future task:

1. Whether ad-hoc meetings are registered as `portal.3gpp.org` Meeting entries with their own
   `MtgId` — portal access requires login, not checked.
2. Whether the Plenary/6G ad-hoc series (`SA4_Plenary`, `S4aP` prefix) has the same report/agenda
   structure as the Audio SWG series analyzed in detail here — not opened.
3. Whether same-numbering-scheme (`S4aA`→`S4aA`) revision chains appear in the TDoc-list *XLSX*
   "Revised to" column (verified only in narrative report text, §3).
4. The exact ETSI mailing list used for ad-hoc email/RSS traffic, and whether it matches
   `MAILING_LISTS[reportType]`.
5. Whether "Reports/Liaison from other groups/meetings" and "Release 19 and earlier matters" are
   fixed, every-time boilerplate sections for Audio-SWG ad-hoc reports, or particular to the one
   teleconference report opened here.
6. Whether a "Registration of Documents" summary table (equivalent to `createSummaryTable_()`'s
   output) exists as a real table anywhere in an ad-hoc report, or only as the referenced-but-
   unopened "Annex A" mentioned in the agenda text.
7. Whether other SA4 ad-hoc series (RTC, MTSI, EVS, MBS, SQ, Video) follow the same pattern as
   Audio, or differ — not investigated.
8. Whether a multi-topic ad-hoc (covering more than one work item, unlike both examples analyzed
   here) produces multiple separate reports or one combined report.

---

## 12. Recommendation for SA4-ARCH-006

Do not design/implement the `adhoc` MeetingContext profile or `ReportSkeletonBuilder` variant yet.
Close the highest-value unknowns first, cheaply, using the same read-only method as this task:

1. Open one report from the `SA4_Plenary`/`S4aP` (6G) ad-hoc series (unknown #2) to check whether
   its front-matter template matches the Audio SWG one closely enough to justify a single shared
   `adhoc-swg`/`adhoc-plenary` structure, or whether they diverge enough to need two profiles.
2. If time permits, open a second Audio-SWG ad-hoc report (a different date) to confirm whether
   "Reports/Liaison" and "Release 19 and earlier matters" (unknown #5) are truly fixed boilerplate.
3. Only after those two checks, draft the concrete `MeetingContext` shape for
   `meeting.type: 'adhoc'` (per §8) and a *named but not-yet-implemented* `ReportSkeletonBuilder`
   outline for `adhoc-swg`, still without touching `Code.js`.
4. Continue treating live implementation (an actual `buildAdhocSkeleton_()` or similar) as later
   work, gated on the above.

---
---

# SA4-ARCH-006 ADDENDUM — Plenary/6G deep dive, second Audio report, prefix survey

Continues directly from SA4-ARCH-005 above. Same verified/inferred/unknown labeling. No production
code modified. All URLs relative to
`https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/` unless stated otherwise.

## A.1 — Plenary/6G ad-hoc (`SA4_Plenary`, `S4aP` series) in depth

### Folder structure (verified)
`SA4_Plenary/` contains only `Docs/` and `Inbox/` (no `Agenda/` subfolder, unlike `SA4_Audio/`).
`Inbox/drafts/` (verified, 44 items) contains **no files with 'report', 'minutes', or 'agenda' in
their name** — confirmed by direct listing, and cross-checked a second way: `Docs/` itself
(96 items) also contains **no 'report'/'minutes'-named zip** among its `S4aP26NNNN.zip` series. This
is a double-checked **negative finding**: unlike Audio SWG (30+ named report `.docx` files),
**no equivalent narrative chair's-report artifact was found for the Plenary/6G ad-hoc series.**

### Real agenda (verified — earliest instance, `S4aP260001`, Jan 2026)
Extracted directly from `S4aP260001 - Proposed agenda for SA4 AHG conf. call on FS_6G_MED (Jan.
15th, 2026).docx`:

```
1.    SA WG4
  1.1   Opening of the session and registration of documents
  1.2   IPR, antitrust and consensus principles reminder
          [not embedded text — a HYPERLINK to an external .pptx hosted on a MAIN-MEETING FTP path:
           .../WG4_CODEC/TSGS4_133-e/Inbox/Drafts/Plenary/SA4-IPR-CompetionLaw-ConsensusPinciples.pptx]
  1.3   Reports/Liaisons from other groups/meetings
  1.4   FS_6G_MED (Study on Media aspects for 6G system)
  1.5   Any other business
  1.6   Close of the session
```

### Real TDoc list (verified — most recent instance, `TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx`,
2026-09-21), showing the **same column schema** as the Audio ad-hoc lists (TDoc/Title/Source/
Contact/.../Agenda item/Agenda item description/TDoc Status/.../Is revision of/Revised to/...), but
a **different agenda-item numbering than the January agenda above**:

```
5.0   SA4 WG on FS_6G_MED (Study on Media aspects for 6G System)
5.1   Opening of the session and registration of documents
5.3   Reports/Liaisons from other groups/meetings
5.4   FS_6G_MED - General and working documents
5.5   FS_6G_MED - WT#1: Media Delivery Architecture
5.6   FS_6G_MED - WT#2: 6G Media
  5.6.1   WT#2.1: AI Traffic Characteristics
  5.6.2   WT#2.2: Other AI-Related topics
  5.6.3   WT#2.3: Non-AI-related topics
5.7   FS_6G_MED - WT#3: Media Aspects related to SA2 topics
5.10  FS_6G_MED - Other issues
```

**This is the single most important new finding of SA4-ARCH-006**: the SAME recurring ad-hoc
group's top-level agenda prefix **changed from `1.` (January) to `5.` (September)** over 9 months —
verified by directly comparing the two real documents. The prefix is **not stable even within one
ad-hoc series over time**, which is a materially stronger finding than SA4-ARCH-005's conclusion
that ad-hoc agendas simply "don't use the main-meeting SWG prefix" — they don't have *any* fixed
prefix at all, ad-hoc or otherwise. (Inferred explanation, not confirmed: the prefix likely tracks
whatever slot this study item currently occupies in the *upcoming main meeting's* plenary agenda,
renumbered release-over-release as other agenda items are added/removed ahead of it — but I found
no explicit statement confirming this mechanism.)

Also verified in this September TDoc list: the topic section has matured from a single flat `1.4`
item (January) into a full Work-Task hierarchy (`WT#1`/`WT#2` with `WT#2.1-2.3`/`WT#3`) — the ad-hoc's
own internal structure grows organically as the study progresses, independent of any main-meeting
template. Cross-scheme "Is revision of" references into main-meeting numbers (`S4-261833`,
`S4-261417`, `S4-261521`, `S4-261505`, `S4-261454`, `S4-261588`, `S4-261834`) recur here exactly as
in the Audio series (verified).

### Comparison: can Audio SWG and Plenary/6G share one structural profile?

**No — verified by multiple independent, concrete differences, not just one:**

| Dimension | Audio SWG AH (ULBC-MED) | Plenary/6G AHG |
|---|---|---|
| Top-level numbering | Flat `1`-`6` (verified) | Nested `X.Y`, and `X` itself drifts (`1.`→`5.`, verified) |
| Report/minutes artifact | 30+ named `.docx` reports exist (verified) | **None found**, double-checked (verified) |
| IPR section content | Embedded full boilerplate text (verified) | External hyperlink to a `.pptx` on a main-meeting FTP path (verified) |
| Topic section growth | Fixed template of generic subsection labels (design constraints, qualification rules, ...) (verified for ULBC-MED) | Organic Work-Task hierarchy specific to the study (`WT#1`/`WT#2.x`/`WT#3`) that changed shape between Jan and Sept (verified) |
| TDoc prefix | `S4aA` (verified) | `S4aP` (verified) |

**Recommendation carried into §9 below: these need at least two distinct
`report.structureProfile` values, not one shared `adhoc-*` profile.**

## A.2 — Second Audio SWG ad-hoc report (verified)

Opened `S4-26xxxx Report from Audio SWG teleconference on ATIAS_Ph3-MED (29 June 2026).docx` in
full (a different date and a different work-item topic than the SA4-ARCH-005 DaCAS example).
Structure (verified, exact section list):

```
A.I. 1     Audio SWG
A.I. 1.1   Opening of the session and registration of documents
A.I. 1.2   IPR, antitrust & consensus principles reminder        [full boilerplate, same wording]
A.I. 1.3   Reports/Liaison from other groups/meetings             -- "None."
A.I. 1.4   Release 19 and earlier matters                         -- "None."
A.I. 1.5   DaCAS (Diverse audio CApturing system...)               -- "None"
A.I. 1.6   ATIAS_Ph3-MED (...)                                     -- [the live topic, TDocs discussed]
A.I. 1.7   Others Rel-20 matters including TEI                     -- "None."
A.I. 1.8   Close of the session
Annex A – Meeting agenda
```

### Section-recurrence classification (per the task's requested categories)

| Section | Classification | Evidence |
|---|---|---|
| Opening / Registration (combined) | **consistently present** | Present, same wording, in both reports (verified) |
| IPR reminder | **consistently present** | Present, same near-identical boilerplate text, in both reports (verified) |
| Reports/Liaison | **consistently present** | Present in BOTH reports, in one case explicitly "None." — proving it's a standing template item included even when empty, not conditionally added (verified) |
| Release 19 and earlier matters | **consistently present** | Same as above — present with "None." in the second report (verified) |
| "Others Rel-20 matters including TEI" | **unknown** | Present in the ATIAS report; **not confirmed absent** from the DaCAS report — my SA4-ARCH-005 excerpt of that report stopped before reaching this point in the document, so this is genuinely unconfirmed either way, not "meeting-specific" |
| Topic section(s) | **consistently present, but the SET of topic slots is templated, not per-meeting** | Both reports show MULTIPLE fixed topic slots (DaCAS=1.5, ATIAS_Ph3-MED=1.6) appearing in both reports regardless of whether that topic had live business that day (verified) |
| AOB | **unknown** | Not reached in either captured excerpt |
| Close | **consistently present** | `A.I. 1.8 Close of the session` present with real content in the ATIAS report (verified); DaCAS report excerpt did not reach this far |

### New, important structural finding (verified)

The ATIAS report's Annex A begins: *"Agenda for this telco (keeping only relevant items from the
unique agenda for 3GPP SA4 AH telcos post-136)"* — **direct textual proof that Audio SWG
teleconferences do not each invent a fresh agenda.** There is one shared, versioned **master
agenda template** ("the unique agenda... post-136", i.e., current since main meeting #136), and
each individual telco's agenda is an explicit subset of it. This explains why fixed topic slots
(DaCAS, ATIAS_Ph3-MED, and presumably others) and fixed front-matter sections (Reports/Liaison,
Release 19 matters, Rel-20/TEI matters) recur identically across different telco instances even
when a given slot has nothing to report ("None.") — they are all carried over from the one shared
template, not re-authored per call. This is architecturally significant: it is a much closer
analog to the current code's `AGENDA_SOURCE_DOC_ID`-based template-copying mechanism
(`copySectionContentWithReplacement_`) than SA4-ARCH-005 assumed — an ad-hoc-telco profile could
plausibly reuse the same "copy fixed sections from a template doc" mechanism, just pointed at a
different (ad-hoc, not main-meeting) template document, rather than needing a wholly different
construction strategy.

**Distinguish, going forward**: this "regular multi-topic telco, subset of a shared master
template, nested `1.x` numbering" pattern is evidenced for **Audio SWG teleconferences** and (with
different topic content) **Plenary/6G AHG calls**. It is explicitly a *different* pattern from the
"dedicated, single-topic, flat `1`-`6`, physical AH" pattern seen in the ULBC-MED example
(SA4-ARCH-005). **Both patterns exist within the Audio SWG's own ad-hoc activity** — the
distinction correlates with meeting *format* (periodic multi-topic call vs. dedicated single-topic
gathering), not simply with SWG identity. See §E/§9 in the SA4-ARCH-006 chat response for how this
reshapes the profile recommendation.

## B.3 — TDoc identifier family survey (verified table)

| SWG / area | Prefix | Example (verified) | Source |
|---|---|---|---|
| Audio | `S4aA` | `S4aA260090` | `Readme_Audio.txt` (explicit: *"prefix designation for adhoc meetings changed to 'S4aA22'..."*) + `SA4_Audio/Docs/` listing |
| Plenary / 6G (`FS_6G_MED`) | `S4aP` | `S4aP260098` | `SA4_Plenary/Docs/` listing (92 files, `S4aP260001`-`S4aP260098`) |
| Video | `S4aV` | `S4aV200545` | `SA4_VIDEO/Docs/` listing |
| RTC | `A4aR` | `A4aR260097` | `SA4_RTC/Docs/` listing (most recent files, 2026) — note the **`A4` root, not `S4`** |
| MBS | `S4aI` | `S4aI240064` | `SA4_MBS/Docs/` listing (most recent files, 2024) — letter `I` does not match the SWG name "MBS"; meaning not confirmed |
| MTSI, EVS, SQ | not checked | — | **unknown** — folders exist (verified from the top-level `3GPP_SA4_AHOC_MTGs/` listing in SA4-ARCH-005) but were not opened in this task |

**Two important corrections to the "obvious" pattern** (do not assume a simple
"S4a + first letter of SWG name" rule):
1. RTC uses **`A4aR`**, not `S4aR` — a different root letter (`A4`, not `S4`) entirely.
2. MBS uses **`S4aI`**, not `S4aM` — the letter doesn't match the SWG name at all (reason
   unconfirmed; possibly inherited from a predecessor group name, not verified).

**Direct connection to existing code**: `Code.js`'s `parseEmailSubject_()` already contains the
regex `/(?:S4|A4[a-z]*)-?\d{6}/i` (verified, appears twice) — this ALREADY matches the `A4aR`
(RTC) family (via its `A4[a-z]*` branch) but does **NOT** match `S4aA`/`S4aP`/`S4aV`/`S4aI` (the
`S4` branch requires the literal text "S4" to be followed immediately by an optional hyphen and
then exactly 6 digits — "S4aA260090" has "aA" in between, which doesn't match). This means the
existing code is **already, if accidentally, partially ad-hoc-ID-aware for exactly one family**
(RTC) and unaware of the other four verified families. This is documented in detail in the
SA4-ARCH-006 chat response's TDoc-ID inventory (deliverable #4).
