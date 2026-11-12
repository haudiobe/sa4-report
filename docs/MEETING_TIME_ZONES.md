# Meeting time zones (candidate after T-2026.10.23)

Status: **candidate, not released, not deployed.** `Code.js` stays 2.26.7 (T-2026.10.23) until the release. GitHub issue #4.

A meeting can have its local time zone, resolved from the Portal's record, or a time zone selected
explicitly. Generated meeting times then name the zone that is in force on their own date: CET or
CEST, MST or MDT, and so on. No zone name is written into the code.

## What is stored

The existing Document Properties are extended. No competing property is added.

| Property | Values | Meaning |
|---|---|---|
| `MEETING_TIME_ZONE` | an offset: `UTC`, `UTC+2`, `UTC-3:30` | As every earlier release stored it, from the Portal. It is kept and written exactly as stored. |
| | an IANA zone of the allowed list, e.g. `America/Edmonton`, `Europe/Berlin`, `Asia/Seoul`, `UTC` | Written as the abbreviation of the zone's offset on that date (`MST`, `MDT`, `CET`, `CEST`, `KST`), or as the offset where there is no abbreviation (`UTC+8`). |
| `MEETING_TIME_ZONE_MODE` | `local` | **Meeting local**: resolved from the Portal's record and verified (see below). |
| | `explicit` | Selected in Configure Meeting. A later Resolve or Save keeps it. |
| | (absent) | An offset, as before. |
| `MEETING_START_TIME`, `MEETING_START_BASIS` | unchanged | The start time and zone apply only to the meeting ID and date they were stored with. |

The allowed zones are `MEETING_TIME_ZONES_` in the MEETING TIME ZONES section of `Code.js`. A value
outside the list is refused when it is saved and ignored when it is read.

## Meeting local: where it comes from

The zone is never taken from the editor, the account, the browser, or the script's own zone.
Meeting local is resolved from the Portal's record of the meeting:

1. The name after the offset in `StartTimeZone` must be one of `PORTAL_TIME_ZONE_NAMES_`. An
   example is `(GMT-07:00) Mountain Time (US & Canada)`. A partial or unknown name is not
   resolved.
2. The meeting's `Country` picks the zone where the name covers several: `CA` gives
   `America/Edmonton`, `US` gives `America/Denver`.
3. **Verification.** The zone's offset at the Portal's `StartDate`, a wall-clock time, must equal
   the offset the Portal gives. The Portal gives the offset in force on that date (`+02.00` for a
   meeting in summer, `+01:00` in winter), so a mismatch means the Portal and the zone disagree.
   Nothing is stored then. Configure Meeting says why and asks the user to select a zone.

A location alone ("Calgary") is not used: it is free text, and "Online" for e-meetings.

### SA4#138 Calgary, and Alberta's 2026 change

The Portal's record for SA4#138 (meeting 60780) is:

```
StartDate 2026-11-16 09:00, StartTimeZone "(GMT-07:00) Mountain Time (US & Canada)", Location Calgary, Country CA
```

Alberta has kept UTC−6 all year since the spring of 2026. The tz database has it from release 2026c
(July 2026), and Windows has a new zone, "(UTC-06:00) Mountain Time (Alberta)", for it. So Calgary
did **not** return to MST (UTC−7) on 1 November 2026. On 16 November its clocks are at UTC−6, which
the tz database names MDT. The Portal's `-07:00` is the old Windows zone and no longer right for
Calgary.

What the code does:

- **Meeting local for SA4#138 is refused.** The reason given is: the Portal gives UTC−7 for the
  start, but America/Edmonton is UTC−6 then.
- **The report selects `America/Edmonton` explicitly.** Its times are then written as the clocks
  in Calgary show them. With the rules of tzdata 2026c or later, that is `MDT` on 16 November
  2026.
- **The new Windows zone is recognised.** If the Portal changes to "(GMT-06:00) Mountain Time
  (Alberta)", Meeting local resolves and verifies by itself.
- **No abbreviation follows from the season.** Abbreviations come from the offset in force
  (`MEETING_TIME_ZONES_` maps offset to abbreviation). A zone whose rules change is written by what
  its clocks show: Calgary's UTC−6 is never called "MST".

Apps Script uses Google's own copy of the tz database (through `Utilities.formatDate`, pattern `Z`).
The JavaScript runtime's rules (`Intl`) are used only if that is not available. After deployment, a
check in the SA4#138 editor will show which rules Apps Script has.

## Wall-clock times and instants

- A session's start and end, the Portal's `StartDate`, and `MEETING_START_TIME` are **wall-clock
  times** of the meeting zone. They are never converted.
- The "Uploaded" times of the TDoc list are **instants** (UTC). They are converted once, into the
  zone the sessions are in, when a TDoc is assigned to a session. That zone is the meeting's IANA
  zone if one is stored for the meeting, otherwise the script's zone as before. A list whose times
  are already on the sessions' clock (`clock: 'session'`) is not converted.
- Deadlines quoted in e-mails ("23:59 CEST") and the e-mail collection zone (`TIMEZONE`) are not
  meeting times and are unchanged. The only remaining `CEST` in the code reads such deadlines, or
  writes the deadline into a discussion e-mail's subject.

## Where meeting times are written

| Place | What changes |
|---|---|
| Ad-hoc opening sentence, no sessions (Build only) | `at 09:00 MDT` for an IANA zone, by the meeting date. An offset is written as stored, and a missing zone is the placeholder, as before. |
| Session administration lines (Build, Configure Sessions, Edit Opening Details) | Each time names the zone of its own day: `09:00–17:30 MDT`. A session over a DST change reads `from 09:00 CEST on the first day, until 17:00 CET on the last day`. A session through the night of a change reads `00:30 CEST–04:00 CET`. Without an IANA zone the lines are exactly as before. |
| Edit Opening Details, statistics dialog | The same text. |
| Session column (update, Assign TDoc Sessions) | UTC uploads are converted into the meeting zone (see above). |
| Configure Sessions | The hint says which zone the times are meant in. |
| Configure Meeting | A new field: Not set / Meeting local / the allowed zones, with what is stored and what Resolve found. |

A **main meeting** (SA4#138 is one) copies its opening from the report template. The code generates
no meeting timestamp for it, so its stored zone is configuration that is ready for such text; it
changes nothing in the report today. The template document's own text was not inspected or changed
by this work.

## Manual edits

The opening sentence is written by Build Report from Scratch only. The Session administration
section is written by a Build, or by saving Configure Sessions or Edit Opening Details. **No update
writes either**, and this work adds no such call: an edited opening survives every update (see
`tests/meeting-time-zones.test.js`, H).

## New reports and existing reports (migration)

- **New report (creator):** stores the verified meeting-local zone, for any meeting type. If it
  cannot be verified, an ad-hoc meeting stores the Portal's offset as before, and a main meeting
  stores nothing. The opening sentence of a new ad-hoc report of 86172 now reads `15:30 CEST`
  instead of `15:30 UTC+2`.
- **Existing report with an offset** (`UTC+2`): unchanged. It is written as before, and a Save or
  a Resolve keeps it when the Portal's zone says the same offset. Configure Meeting shows it as "the
  fixed offset UTC+2 the Portal gave; it is kept". **To migrate**, select Meeting local (after
  Resolve) or a zone, and Save. A Portal that now gives another offset replaces it, as a Resolve
  always did.
- **Existing report without a zone:** unchanged until a zone is saved. After Resolve, Meeting local
  is preselected when it can be verified.
- **CENTRAL add-on backend:** the meeting zone is not a centrally adopted key. Its sessions keep the
  script's zone for upload times, in interactive and background runs alike.

## Not changed

- The update, Close Report, the agenda, and the TDoc list.
- The e-mail collector, its deadline parsing, and its time zone.
- `computeMeetingTimeZoneLabel_()` and `isValidMeetingTimeZoneLabel_()`.
- Triggers and configuration.
- The TDoc revision feature, which is separate.
