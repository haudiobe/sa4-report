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
| `05/10/2026 10:05` (day and month ≤ 12, no AM/PM) | **ambiguous: not read**, no session (preview, never a guess); `2/19/2026 3:00:31 PM` and ISO are read |
| A date without a time | a session only if the whole day falls into one window |
| A session of several days | cut-off on its last day |
| DST change between end and cut-off | six hours of real time (test: Europe/Amsterdam, 25 Oct 2026) |
| Sessions edited | recomputed; manual assignments stay; the Assign TDoc Sessions dialog shows the automatic result before anything is saved |
| Manual "set" / "add" (ADHOC_TDOC_SESSIONS) | authoritative, as before |

The reservation times are read from the TDoc list's "Reservation date" column like the uploads, and treated as UTC
like "Uploaded" (the model's `clock`). **To be confirmed on a real list:** that "Reservation date" is UTC.

## 2. Carry-forward (with the reservation rule)
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
| `ADHOC_SESSION_RULE` | `upload` (default, stage B) or `reservation-6h` (this phase; enables carry-forward) |
| `PLENARY_AGENDA_DIR_URL`, `PLENARY_SCHEDULE_DIR_URL` | override the derived folders |

## Migration and rollback
- Opt-in: nothing changes in an AHG report until the rule is switched; a main report only gains the Meeting Control
  view. Switch per report after its sessions are configured; the next update rewrites the Session column.
- Rollback of the setting: switch back to "upload" (the next update rewrites the column; carried records stay stored
  but are not shown). Code rollback: older code ignores the new keys; the Session column is rewritten by its next update.

## Tests
`tests/i28c-sessions-agenda.test.js` (A–G) with the real SA4#138 Agenda/Schedule listings and "Tdoc allocation
rev9.docx" (2026-11-23) as fixtures. Full suite: 135 files, all passing.

## Open questions
1. Is the Portal's "Reservation date" UTC (as "Uploaded")? Decides the clock of the rule.
2. Should carry-forward be available without the reservation rule (its own switch)?
3. The real AHG reports: which sessions/zones to use for the first opt-in (e.g. the Audio post-138 telcos of 4 and 8 Dec)?
