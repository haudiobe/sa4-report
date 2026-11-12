# Meeting time zones (T-2026.10.24, Code.js 2.26.8)

Status: **released as T-2026.10.24 (`Code.js` 2.26.8).** Merged as `23a54e9` (PR #5, GitHub issue #4); release issue #9.
The live check in a copy of SA4#138 and the deployment are recorded in issue #9.

A meeting can have its local time zone, resolved from the Portal's record, or a time zone selected
explicitly. Generated meeting times then name the zone that is in force on their own date: CET or
CEST, MST or MDT, ABT in Alberta from 1 November 2026, and so on. No output hard-codes a zone name; abbreviations come from the offset in force.

## What is stored

The existing Document Properties are extended. No competing property is added.

| Property | Values | Meaning |
|---|---|---|
| `MEETING_TIME_ZONE` | an offset: `UTC`, `UTC+2`, `UTC-3:30` | As every earlier release stored it, from the Portal. It is kept and written exactly as stored. |
| | an IANA zone of the allowed list, e.g. `America/Edmonton`, `Europe/Berlin`, `Asia/Seoul`, `UTC` | Written as the abbreviation of the zone's offset on that date (`MST`, `MDT`, `ABT`, `CET`, `CEST`, `KST`), or as the offset where there is no abbreviation (`UTC+8`). |
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

### SA4#138 Calgary, and Alberta's change on 1 November 2026

The Portal's record for SA4#138 (meeting 60780) is:

```
StartDate 2026-11-16 09:00, StartTimeZone "(GMT-07:00) Mountain Time (US & Canada)", Location Calgary, Country CA
```

From 1 November 2026, Alberta stays at **UTC−6 all year**, and the official name of that time is **ABT**. Its
clocks did not go back to MST (UTC−7) at 02:00 MDT on 1 November 2026 (08:00 UTC). On 16 November 2026, 09:00 in
Calgary is 15:00 UTC. The Portal's `-07:00` is the old Windows zone and no longer right for Calgary.

**Copies of the tz database differ, and Apps Script's copy is not under our control:**

| tz database | Calgary from 1 November 2026 |
|---|---|
| before Alberta's change | UTC−7, `MST` (wrong) |
| 2026c and later | UTC−6, `MDT` (right offset, not the official name) |

So the rule is in the code, not taken from the tz database: `MEETING_TIME_ZONE_RULE_CHANGES_` says that
`America/Edmonton` is at UTC−6 and written `ABT` from 2026-11-01 08:00 UTC on. It is applied before Apps Script's
rules (`Utilities.formatDate`, pattern `Z`) and before the JavaScript runtime's (`Intl`), in every place that needs an
offset: the abbreviation, a wall-clock time to an instant, and the conversion of the TDoc list's UTC upload times.
Before that instant the tz database's rules apply unchanged (MST in winter 2025/26, MDT from 8 March 2026).

What the code does:

- **Times in Calgary from 1 November 2026 are written `ABT`**: `09:00 ABT` on 16 November 2026, with any tz database.
  A session over the night of the change reads `from 09:00 MDT on the first day, until 17:00 ABT on the last day`.
  There is no repeated hour: 01:30 is MDT, 02:30 is ABT, both UTC−6.
- **Meeting local for SA4#138 is refused.** The reason given is: the Portal gives UTC−7 for the start, but
  America/Edmonton is UTC−6 (ABT) then. The report selects `America/Edmonton` explicitly. Nothing is guessed from the
  Portal's out-of-date offset.
- **Windows' new zone is recognised.** If the Portal changes to "(GMT-06:00) Mountain Time (Alberta)" (or `UTC-06:00`),
  Meeting local resolves and verifies by itself.
- **Configure Meeting names a zone by its next twelve months**: `America/Edmonton (ABT)` from November 2026,
  `America/Denver (MST/MDT)`, `Europe/Berlin (CET/CEST)`.
- **Denver is unchanged**: MST from 1 November 2026, MDT in summer. The DST behaviour of Mountain Time is tested with
  it.

The rule table is the only place to change if another zone's rules change before Apps Script's tz database has them.

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
| Ad-hoc opening sentence, no sessions (Build only) | `at 09:00 ABT` for an IANA zone, by the meeting date. An offset is written as stored, and a missing zone is the placeholder, as before. |
| Session administration lines (Build, Configure Sessions, Edit Opening Details) | Each time names the zone of its own day: `09:00–17:30 ABT`. A session over a DST change reads `from 09:00 CEST on the first day, until 17:00 CET on the last day`. A session through the night of a change reads `00:30 CEST–04:00 CET`. Without an IANA zone the lines are exactly as before. |
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

## What a save of Configure Meeting does with the zone

The dialog sends a zone only when the field was **changed** there. A field left as the dialog opened sends nothing,
and the server decides for the meeting being saved:

| Field on Save | Result |
|---|---|
| left as it opened | A zone selected explicitly for **this** meeting stays, and so does a legacy offset. A zone selected for another meeting is not carried over: a main meeting takes its verified meeting-local zone, if Resolve found one. |
| changed to *Meeting local* | The zone verified by this save's Resolve. If this save's Resolve does not verify, the save is refused, even when a meeting-local zone is stored. Without a Resolve, the meeting-local zone stored for this meeting **and date** counts. |
| changed to a zone | Stored as explicit. A start time read on another clock is removed: the opening then shows `<start>` rather than a time that is hours off. A start time stays when the zone has the same offset at that time, or when the zone is the one the Portal's own zone name means (SA4#138: `Mountain Time (US & Canada)` with `CA` is `America/Edmonton`, even though the Portal's offset is out of date). |
| changed to *Not set* | The zone and its mode are removed. A start time stays with its meeting. |

For an adopted document (CENTRAL add-on), every save syncs the zone, its mode and its meeting into the central state,
together with the other Configure Meeting keys.

## Known behaviour (intended)

- **A zone saved later does not rewrite text already in the report.** The opening and the Session administration
  section keep their text until the next Build, or until Configure Sessions or Edit Opening Details is saved. This is
  what keeps manual edits (see above). The Session column follows at the next update.
- **The Session column's clock can move once.** An ad-hoc report that had no IANA zone converts the TDoc list's UTC
  upload times into the script's zone (Europe/Berlin). Once a meeting zone is saved, it converts them into that zone.
  Session times entered under the old hint ("the report's time zone") were meant in the script's zone. For a meeting
  outside it, check the sessions after saving the zone. A TDoc near a session boundary may move at the next update.
  A legacy offset (`UTC+2`) is not an IANA zone and changes nothing here.

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
- **CENTRAL add-on backend:** `MEETING_TIME_ZONE`, `MEETING_TIME_ZONE_MODE` and `MEETING_START_BASIS` are adoption
  keys (`ADDON003_ADOPTION_FIXED_KEYS_`), because every update reads them: they give the zone into which the TDoc
  list's UTC upload times are converted for the Session column. A document adopted after this change converts them in
  the meeting zone, in background runs as in the bound script. A document adopted before it has no zone in the
  central state and keeps the add-on's own zone (Europe/Berlin) until it is adopted again. `MEETING_START_TIME` is not
  an adoption key: only a build reads it, and builds are interactive.

## Not changed

- The update, Close Report, the agenda, and the TDoc list.
- The e-mail collector, its deadline parsing, and its time zone.
- `computeMeetingTimeZoneLabel_()` and `isValidMeetingTimeZoneLabel_()`.
- Triggers and configuration.
- The TDoc revision feature, which is separate.
