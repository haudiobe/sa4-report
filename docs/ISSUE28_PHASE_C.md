# Issue #28 phase C: AHG sessions, agenda updates and Meeting Control (Code.js 2.30.0, candidate)

**Status: candidate, not released, not deployed.** Branch `feature/i28-c-ahg-sessions-agenda`, stacked on phase B.
Code: section "ISSUE 28 C" at the end of `Code.js`; one-hook changes in `makeAdhocTdocSessionResolver_`,
`continuousUpdateCore_`, `downloadAndGroupTdocs_`; one menu item. Ledger: `codeWithoutIssue28C` → phase B byte for byte.

## 1. Session assignment by reservation time (opt-in per report)

Setting `ADHOC_SESSION_RULE = reservation-6h` (Meeting Control › Session rule, with confirmation). Without it the
stage B rule (upload time, cut-off = end) is **unchanged**: existing AHG reports do not change by themselves.

**Rule.** A TDoc belongs to the first session whose cut-off is not before its **reservation** time.
`cutoff(session) = instant(last day + end time (or 23:59:59), meeting IANA zone) + 6 hours`, computed on instants.

| Case | Result |
|---|---|
| Reserved exactly at the cut-off | that session (inclusive) |
| One second after the cut-off | the next session |
| Before the first session | the first session (the reason says "before the first session") |
| After the last cut-off | no session, "after the last cut-off" — never the last session by default |
| No reservation time / unreadable | no session, with the reason |
| `05/10/2026 10:05` (day and month ≤ 12, no AM/PM) | read in the **order of its column** (below); only in a column that does not show its order (or contradicts itself): **ambiguous, not read**, no session |
| A date without a time | a session only if the whole day falls into one window |
| A session of several days | cut-off on its last day |
| DST change between end and cut-off | six hours of real time (test: Europe/Amsterdam, 25 Oct 2026) |
| Sessions edited | recomputed; manual assignments stay; the Assign TDoc Sessions dialog shows the automatic result before anything is saved |
| Manual "set" / "add" (ADHOC_TDOC_SESSIONS) | authoritative, as before |

The reservation times are read from the TDoc list's "Reservation date" column like the uploads, and treated as UTC
like "Uploaded" (the model's `clock`).

### The clock of "Reservation date" (investigated 2026-11-23)
**UTC, high confidence** — evidence in [issue28-evidence/reservation-clock/EVIDENCE.md](issue28-evidence/reservation-clock/EVIDENCE.md)
(four SA4 lists, #135 India, #136 Montreal, #137-e, #138 Calgary; 2,081 TDocs in `per_tdoc.csv`; scripts included):

- "Uploaded" is UTC (very high): it equals the GMT `Last-Modified` of the file on www.3gpp.org/ftp (median +3 s, 99 HEADs).
- "Reservation date" on the same clock: read as UTC, no TDoc in any of the four lists is uploaded before its reservation,
  and the shortest reservation-to-upload gap is 0.8–1.5 min in every meeting. Read as meeting-local time (H3), 282/596
  (Calgary) and 272/497 (Montreal) would be uploaded before their reservation. Read as Paris time (H2), no gap below
  61 min (CET) / 121 min (CEST) would exist anywhere, with the busiest gaps at that edge, and the band would move with
  the 25 Oct DST change — it does not (0.8–1.6 min on both sides).
- No public 3GPP source labels the zone (the Portal's TDoc view needs a login; no public web service). Certainty would
  need MCC's statement or one TDoc reserved at a noted time.

**What the code does with it (decision 4: not switched on unverified):**
- *Date order.* "Reservation date" is a **text** cell whose format changed: the SA4#138 list of 2026-11-23 writes
  `13/10/2026 12:34:06` (day/month, 24 h), the earlier lists `7/17/2026 12:33:23 PM`. Read value by value, **361 of
  the 615 SA4#138 reservations were refused as ambiguous**. The order is now told from the whole column
  (`tdocListDateOrder_`: a first field above 12 → day/month; AM/PM or a second field above 12 → month/day; both or
  neither → each value on its own). Test A2: all 615 (#138) and 482 (#137-e) real values read.
- *Clock check on the live list.* The preview of the rule (`previewAdhocSessionRule`) checks the list it reads:
  a TDoc uploaded more than 2 minutes before its reservation as read means the clocks disagree, and the preview
  is refused, naming the TDocs, and nothing is stamped, so the rule cannot be switched on. The confirmation shows
  the check (TDocs with both times, shortest gap).
- *Rollout.* The rule stays off in every report until the user accepts this evidence; it is then switched per report
  by its rapporteur through the preview.

## 2. Carry-forward (its own switch)
**Decision 2026-11-23:** independent of the session rule. `ADHOC_CARRY_FORWARD_MODE=on` (Meeting Control › Carry-forward
› Switch on…, after a preview of what it would record now, within 30 minutes; ad-hoc reports only; under the document
lock). It works with either rule (test D: with the upload rule as well). Off by default: an existing AHG report does not
change. Switching off keeps the recorded carries but no longer shows them.

A TDoc that is not disposed (Status available/parked/reserved or empty) when a later session **starts** is carried
into it: recorded once with the time (`ADHOC_CARRY_FORWARD`, provenance), shown as "A01 → A02, A03" in the Session
column. No second registration row; the TDoc table is not touched. A disposed TDoc is carried no further and keeps what
it has. A "set" override is never carried. Taking a disposed TDoc up again is an explicit "add" assignment.

## 3. Plenary Agenda and Schedule folders (read-only)
Folders derived from FTP_BASE (`…/<meeting>/Docs/` → `…/Inbox/Drafts/Plenary/Agenda/` and `/Schedule/`), or set with
`PLENARY_AGENDA_DIR_URL` / `PLENARY_SCHEDULE_DIR_URL`.

- **Newest allocation:** the highest `revN` of "… Tdoc allocation revN.docx". The listing time is not trusted: on
  2026-11-23 rev3–rev9 of SA4#138 all show 2026/11/23 7:28.
- **Allocations:** the third column of each agenda-item row; 4-digit short numbers mapped to the report's TDocs only;
  dates and remarks in parentheses ("(was 1697 pp)") removed first; chains ("2015->2139") give both.
- **Proposed reallocations:** a TDoc whose single allocation differs from the report's Agenda Item column and from a
  reallocation already recorded. Listed under two items: shown as ambiguous, never proposed. **Accepting** writes the
  Document Reallocations row (the dialog's writer; issue #27's behaviour unchanged); the TDoc tables move only when the
  reallocations are applied.
- **AHG candidates:** the lines of the "Telcos" part (dates, times, zone, group), each with a confidence: high (date,
  time, zone, fits the season); medium (zone that does not fit the date — "CEST" in December —, time not read, year
  taken from the meeting, "pending"); low (no date). Sessions are created **only** from selected candidates, after a
  confirmation, in an ad-hoc report, through Configure Sessions' own writer; their times are then checked by hand
  (agenda zones are not the meeting zone).
- **Schedule:** the files of the Schedule folder, newest listed first, as a low-confidence hint (no version order;
  the SA4#138 grid's day headers were "Monday 20 … Friday 24" for 16–20 November).
- **Robustness:** CacheService, 30 minutes, for listings and the parsed document; each source fails on its own with
  its message; a missing folder or a non-listing page is said, nothing guessed.

## 4. Meeting Control (Report › Meeting Control…)
Opens at once with a progress line and loads its data. One table, searchable, with filter chips:
**new** (reserved), **title** (phase A), **revision** (phase B), **allocation**, **carried**, **unresolved**,
**plenary**. Below: changed allocations (preview + confirm → Document Reallocations), AHG candidates (confirm →
sessions, ad-hoc only), the schedule, and the session rule switch (confirm). Each part reports its own failure.
Bulk actions always show what will be written and ask first.

## Configuration
| Report state key | Meaning |
|---|---|
| `ADHOC_SESSION_RULE` | `upload` (default, stage B) or `reservation-6h` (this phase; needs a preview whose clock check passes) |
| `ADHOC_CARRY_FORWARD_MODE` | `on` carries undisposed TDocs into started sessions (needs a preview); anything else off (default) |
| `PLENARY_AGENDA_DIR_URL`, `PLENARY_SCHEDULE_DIR_URL` | override the derived folders |

## Migration and rollback
- Opt-in: nothing changes in an AHG report until the rule is switched; a main report only gains the Meeting Control
  view. Switch per report after its sessions are configured; the next update rewrites the Session column.
- Rollback of the setting: switch back to "upload" (the next update rewrites the column; carried records stay stored
  but are not shown). Code rollback: older code ignores the new keys; the Session column is rewritten by its next update.

## Tests
`tests/i28c-sessions-agenda.test.js` (A–G) with the real SA4#138 Agenda/Schedule listings and "Tdoc allocation
rev9.docx" (2026-11-23) as fixtures. Test A2: the real "Reservation date" columns of SA4#138 and SA4#137-e (tests/fixtures/tdoc-list-reservation-dates/).

## Open questions
1. ~~Is the Portal's "Reservation date" UTC?~~ Investigated: UTC, high confidence (above); the user's acceptance is
   pending before any report switches the rule on.
2. ~~Carry-forward without the reservation rule?~~ Decided: its own switch (above).
3. The real AHG reports: which sessions/zones to use for the first opt-in (e.g. the Audio post-138 telcos of 4 and 8 Dec)?

## Independent review (2026-11-23) and what changed

| # | Finding | Outcome |
|---|---|---|
| 1 | Candidate dates: "Tuesday 8th December 2026" read as 20 December, high confidence; "26–29 January" as 9 January | **Fixed.** A day is never the first digits of a year; the earliest date of the line is taken; a second date makes a multi-day candidate (`endDate`, medium), and no time is read from the range. Tests on both real lines |
| 2 | Carry-forward shown one update late (the cell read a stale copy) | **Fixed:** the cell reads the resolver's current carry. Test: the same resolver shows "A01 → A02" right after the sync |
| 3 | Years and plenary-level rows produced allocations/proposals | **Fixed:** only allocations within the report's own part of the agenda (its top-level numbers) are compared, so rows 13–22 never move a SWG TDoc; dates are removed before numbers are read |
| 4 | No preview before switching the rule; the dialog's "why" blank under the new rule | **Fixed:** "Preview" reads the TDoc list and lists every TDoc whose session changes (old → new, with the reason) before the confirmation; the switch is refused without a preview of the last 30 minutes, outside an ad-hoc report, and without a "Reservation date" column; it takes the document lock. The Assign TDoc Sessions dialog shows the reservation reason (one hook) |
| 5 | Accepting could overwrite a manual reallocation; the server did not re-check | **Fixed:** a TDoc reallocated by hand to another item is shown as "manual" and never proposed; accepting re-reads sources and report inside the lock and writes only what is still proposed |
| 6 | Carry state stopped silently at the size cap | **Fixed:** compact storage (v 2, about 4× smaller; v 1 still read) and a visible warning in Meeting Control when full |
| 7 | Candidate times stored in the agenda's zone | **Fixed:** converted from the quoted zone (UTC/GMT, CET, CEST) into the meeting zone; an unknown zone with a time is refused |
| 8 | Switching without a "Reservation date" column blanks the column | **Fixed** (refused by the preview) |
| 9 | Repeated zone conversions | **Fixed:** instants memoised |
| 10 | An unknown zone offset was read as UTC | **Fixed:** NaN, so no session (never a guess) |
| 11 | Cache key collisions; a malformed % failed the listing | **Fixed:** SHA-256 key; a malformed name is kept as it is |
| 12 | Carry decided when an update runs, not when a session starts | Accepted and documented: there is no disposition time in the data |
