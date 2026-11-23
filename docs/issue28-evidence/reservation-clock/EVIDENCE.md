# Which clock are "Reservation date" and "Uploaded" in? (3GPP TDoc list xlsx)

> In the repository: this file, `per_tdoc.csv`, `analysis_output.md` (written to `data/` by the scripts) and `scripts/`.
> The raw downloads (`downloads/`) and the per-meeting extracts (`data/*.csv`) are not committed: they are the public
> files at the URLs below, and `scripts/extract.py` makes the extracts again.

Investigation date: 2026-11-23 (all retrieval times below are UTC, taken from the local clock and
cross-checked against the server `Date:` headers, which agree to the second).

## Verdict

| Column | Clock | Confidence |
|---|---|---|
| **Uploaded** | **UTC** (equals the HTTP `Last-Modified` time of the file on www.3gpp.org/ftp, which is GMT by definition; median difference +3 s) | Very high |
| **Reservation date** | **UTC** (the same clock as Uploaded). Europe/Paris (H2) and the meeting's local time (H3) are rejected | High. It is inferred from the data; no 3GPP source labels the zone |

The two values are not stored the same way. Uploaded is a real Excel date-time cell with milliseconds.
Reservation date is a **text** cell, and its format differs between files (see caveat 1).

## Sources

| What | URL | Retrieved (UTC) |
|---|---|---|
| WG4 folder index (to check the folder name) | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/ | 2026-11-23T16:41:23Z |
| SA4#138 Docs listing (HTML, per-file date and time) | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/ | 16:41:33Z |
| SA4#138 TDoc list (Last-Modified Mon, 23 Nov 2026 15:16:53 GMT) | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx | 16:41:39Z |
| HEAD of 99 SA4#138 zips (every 6th row of the list), `Last-Modified` | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/S4-26xxxx.zip | 16:42:13Z – 16:42:56Z |
| SA4#137-e listing + TDoc list (e-meeting, Aug 2026, CEST; Last-Modified 22 Nov 2026 11:31:30 GMT) | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/Docs/ and .../Docs/TDoc_List_Meeting_SA4%23137-e.xlsx | 16:43:04Z / 16:43:05Z |
| SA4#136 Montreal listing + TDoc list (Apr–May 2026, CEST; meeting at UTC-4; LM 30 Oct 2026 14:21:08 GMT) | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/ and .../Docs/TDoc_List_Meeting_SA4%23136.xlsx | 16:43:05Z / 16:43:06Z |
| SA4#135 India listing + TDoc list (Jan–Feb 2026, CET; meeting at UTC+5:30; LM 30 Oct 2026 14:29:10 GMT) | https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_135_India/Docs/ and .../Docs/TDoc_List_Meeting_SA4%23135.xlsx | 16:43:07Z / 16:43:07Z |
| Portal TDoc view (looking for a zone label) | https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=S4-261857 (also `contributionId=`) | 16:44:31Z: redirects to login.aspx. Not used further |
| S4 meeting calendar (meeting ID 60780 = SA4#138) | https://www.3gpp.org/dynareport?code=Meetings-S4.htm | 16:44:40Z |
| Public portal TDoc list page | https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=60780 | 16:44:46Z: public, but shows no reservation or upload time and no zone |
| Portal web-service roots | https://portal.3gpp.org/webservices/Rest/ and https://portal.3gpp.org/webservices/ | 16:45:02Z: HTTP 500. No public endpoint found |

About 110 requests in total, sent one at a time; HEAD was used for the zips. No login was used and no
personal data was sent.

Together the four meetings cover both Paris seasons (CEST: #136, #137-e and the start of #138; CET:
#135 and most of #138) and three meeting time zones (UTC-6, UTC-4, UTC+5:30), plus one e-meeting. SA4#138
itself spans the 25 Oct 2026 EU DST change: reservations run from 13 Oct to 20 Nov.

## Method

1. I parsed `TDoc`, `TDoc Status`, `Reservation date` and `Uploaded` from each xlsx, and the per-file
   date and time from the FTP HTML listing (`scripts/extract.py`).
2. (a) I compared Uploaded with the HTTP `Last-Modified` (99 HEADs) and with the listing time for every file.
3. (b) For each hypothesis I converted Reservation date to UTC and computed `Uploaded − Reservation`.
   A TDoc cannot be uploaded before it is reserved, and Uploaded is the latest upload, so it is later
   still. Under the true clock there should be **no inversions** (negative gaps), and the smallest gaps
   should be about a minute: the fastest a person can reserve a number, put it in the document and
   upload it.
   - Inversions test whether the reservation clock is *behind* UTC (H3 for the Americas).
   - For a clock *ahead* of UTC (H2, India), the test is the shape of the gap distribution. Under the
     wrong hypothesis, every gap shorter than the clock offset disappears, and the peak sits exactly at
     the edge of the empty band.
4. (d) I split the gaps by the Paris DST state at reservation time, both within SA4#138 and across
   meetings. An H2 clock has an offset of 2 h in CEST and 1 h in CET.
5. (c) I counted reservations and uploads by hour of day during the SA4#138 meeting week, 16–20 Nov 2026.
6. (e) I looked for explicit zone labels in the xlsx parts and on the public 3GPP portal pages.

Analysis: `scripts/analyse.py`, with its output in `data/analysis_output.md`. The per-TDoc numbers are
in `per_tdoc.csv` (2,081 rows).

## Results

### (a) Uploaded vs file time on the FTP server

- 99 HEADs on SA4#138: `Uploaded − Last-Modified` has a median of **+3 s**, and 69/99 are within
  0–60 s. None is near ±1 h, ±2 h or −6 h. The rest are negative: the file was written again after the
  recorded upload (re-uploads and bulk copies), so the gap is days or tens of minutes, never a whole-hour step.
- The listing time equals `Last-Modified` truncated to the minute for **99/99** sampled files, so the
  listing is also UTC and can stand in for HEAD.
- Uploaded vs listing for every file, grouped by nearest whole hour:

| meeting | n | uploaded − listing in [0,60) s | nearest-hour buckets |
|---|---|---|---|
| SA4#138 | 596 | 355 | 0 h: 521; −1 h: 62; scattered re-uploads |
| SA4#137-e | 470 | 405 | 0 h: 470 |
| SA4#136 | 497 | 415 | 0 h: 495 |
| SA4#135 | 469 | 247 | 0 h: 466 |

The 62 "−1 h" cases in #138 are not a clock offset. All 62 have the same Uploaded value
(2026-11-19 18:49:09), and their files were written 44–45 min later (19:33–19:34). That is one bulk
copy, and the offset is not a whole hour. The modal offset is 0 h in all four meetings, in both seasons,
so **Uploaded = UTC**.

### (b) Upload − Reservation (minutes) under each hypothesis

| meeting | hypothesis | n | inversions | min gap | 1st pct | 0–10 min | 10–60 min | 60–130 min |
|---|---|---|---|---|---|---|---|---|
| SA4#138 (Calgary) | **H1 UTC** | 596 | **0** | **0.8** | 1.5 | 91 | 62 | 39 |
| | H2 Paris | 596 | 0 | 60.8 | 61.5 | **0** | **0** | 159 |
| | H3 Calgary (UTC-6) | 596 | **282** | −359.2 | −358.5 | 0 | 8 | 7 |
| SA4#136 (Montreal) | **H1 UTC** | 497 | **0** | **1.1** | 2.0 | 69 | 115 | 57 |
| | H2 Paris | 497 | 0 | 121.1 | 122.0 | 0 | 0 | 69 |
| | H3 Montreal (EDT, UTC-4) | 497 | **272** | −238.9 | −237.9 | 2 | 12 | 7 |
| SA4#135 (India) | **H1 UTC** | 469 | **0** | **0.8** | 2.0 | 48 | 90 | 39 |
| | H2 Paris | 469 | 0 | 60.8 | 62.0 | 0 | 0 | 145 |
| | H3 India (IST, UTC+5:30) | 469 | 0 | 330.8 | 332.0 | 0 | 0 | 0 |
| SA4#137-e | **H1 UTC** | 471 | **0** | **1.5** | 2.1 | 66 | 100 | 43 |
| | H2 Paris | 471 | 0 | 121.5 | 122.1 | 0 | 0 | 66 |

Gaps when Reservation is read as UTC, per minute from 0 to 9 min (sharp onset at about 1 min, mode at 2–3 min):

| meeting | 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 |
|---|---|---|---|---|---|---|---|---|---|---|
| SA4#138 | 3 | 8 | 18 | 18 | 11 | 4 | 8 | 8 | 6 | 7 |
| SA4#137-e | 0 | 4 | 15 | 7 | 7 | 7 | 8 | 8 | 6 | 4 |
| SA4#136 | 0 | 4 | 10 | 15 | 8 | 11 | 8 | 3 | 4 | 6 |
| SA4#135 | 1 | 4 | 7 | 6 | 6 | 6 | 3 | 9 | 1 | 5 |

Gaps when Reservation is read as UTC, all meetings together, in 10-minute bins: 0–10: **273**,
10–20: 153, 20–30: 56, 30–40: 65, 40–50: 50, 50–60: 44, 60–70: 25, 70–80: 41, 80–90: 36, …
The counts fall off steadily from zero.

Examples: S4-262209 was reserved 2026-11-10 20:01:14 and uploaded 20:02:00. S4-260289 was reserved
2026-02-08 10:18:18 and uploaded 10:19:05. S4-260923 was reserved 2026-04-30 08:28:02 and uploaded
08:29:08. All examples are in `data/small_gaps.csv`.

What the results mean for each hypothesis:

- **H3 for the Americas** (Calgary UTC-6, Montreal UTC-4) is rejected outright: about 47–55 % of TDocs
  would have been uploaded hours before they were reserved.
- **Reservation clock behind UTC by any amount**: also bounded tightly. With no inversions and a
  minimum gap of 0.8 min, the reservation clock cannot be more than about 1 minute behind the Uploaded clock.
- **H2 (Paris)**: if it were true, the real gaps would be the H1 gaps plus 60 min (CET) or 120 min (CEST).
  In all four meetings, then, nobody would ever have uploaded within 1 h (CET meetings) or 2 h (CEST
  meetings) of reserving. The largest group of TDocs, about 13 %, would sit exactly at the edge of that
  empty band, and the band's width would change with the European season. The H1 reading gives what
  people actually do in a single, season-independent shape: reserve, then upload a minute or a few
  minutes later.
- **H3 for India** (UTC+5:30): rejected by the same argument, with an empty band of 5.5 h.

### (d) DST behaviour

Upload − Reservation when Reservation is read as UTC, split by the Paris offset at reservation time:

| meeting / period | n | min gap | n < 10 min | n < 60 min | n < 120 min |
|---|---|---|---|---|---|
| SA4#138, CEST (before 25 Oct) | 28 | 1.6 | 18 | 19 | 19 |
| SA4#138, CET (after 25 Oct) | 568 | 0.8 | 73 | 134 | 169 |
| SA4#136 (CEST) | 497 | 1.1 | 69 | 184 | 235 |
| SA4#137-e (CEST) | 471 | 1.5 | 66 | 166 | 204 |
| SA4#135 (CET) | 469 | 0.8 | 48 | 138 | 172 |

The smallest gap stays at about 1 min on both sides of the 25 Oct 2026 change, inside SA4#138 and
across meetings in both seasons. A Paris clock would make the smallest gap jump from about 120 to about
60 min at the change. It does not.

### (c) Hour of day during the SA4#138 meeting week (16–20 Nov 2026)

Reservations, counted by the hour as written: 14 h: 21, 15 h: 18, 16 h: 15, 17 h: 27, 18 h: 25,
19 h: 10, 20 h: 4, 21 h: 11, 22 h: 12, 23 h: 10, 00 h: 11, 01 h: 7, 03 h: 23. Only 0–6 per hour
between 04 and 13 h. Read as UTC, the busy hours are **08:00–21:00 Calgary** (UTC-6), which is the
meeting day plus the evening. Uploads peak at 14, 18 and 22 h as written, which is 08, 12 and 16 h
Calgary.

Read as Calgary local time (H3), the same data would mean heavy activity at 00–03 h at night and almost
none from 08 to 13 h during the sessions, which is implausible. This test cannot separate H1 from H2,
because they differ by only 1 h in November. It agrees with H1 and argues against H3. The full table is
in `data/analysis_output.md`.

### (e) Explicit zone labels

- None in the xlsx. The only time-zone text is an agenda title, "Opening of the meeting: Monday 24th
  August 2026, at 9:00 CEST" (#137-e).
- The public portal TDoc list has no times. The TDoc view page needs a login, and the web-service roots
  return HTTP 500.
- **I found no public source that labels the reservation zone.**

## Counter-evidence and caveats

1. **The format of Reservation date varies.**
   - SA4#138, regenerated 2026-11-23: `13/10/2026 12:34:06` (dd/mm/yyyy, 24 h).
   - SA4#135, #136 and #137-e, generated 2026-10-30 and 2026-11-22: `7/17/2026 12:33:23 PM`
     (m/d/yyyy, 12 h with AM/PM).

   The value is text, not an Excel date. Any parser has to handle both formats. Day and month cannot be
   told apart reliably when both are 12 or less unless the AM/PM marker is used as the switch, as in
   `scripts/extract.py`. This is not a zone issue, but it is a correctness risk for the ad-hoc feature.
2. Reservation is checked *relative to* Uploaded, and Uploaded relative to HTTP GMT. The conclusion
   assumes both are stamped by the same back-end clock. The roughly 1-minute minimum gap, the same in
   every meeting and season, supports that.
3. The "one-sided" problem: for a clock ahead of UTC, the test cannot produce inversions, so rejecting
   H2 rests on the shape of the gap distribution (an empty band whose width follows Paris DST, in four
   independent datasets), not on impossible values. I consider that argument strong but statistical,
   not a proof.
4. Some TDocs are re-uploaded or bulk-copied after their recorded upload, for example 62 files in #138
   written 44 min after their recorded upload. That is why minima and modes were used, not every row.
5. Only SA4 main meetings were tested, all generated by the same system. Ad-hoc and telco TDoc lists come
   from the same portal export and are expected to behave the same, but none was checked.
6. Small samples: only 28 SA4#138 reservations fall before 25 Oct. The DST argument rests mainly on the
   comparison across meetings.

## What would make it certain

- A TDoc page shown while logged in (portal.3gpp.org `CreateTdoc.aspx?mode=view&contributionUid=…`) or
  3GU showing a reservation time with an explicit zone, or a statement from 3GPP MCC / the Portal team.
- A controlled test: reserve a TDoc at a noted wall-clock time (with the zone written down) and read its
  Reservation date in the next generated TDoc list.

## Files

- `per_tdoc.csv`: per-TDoc numbers for all four meetings: meeting, TDoc, status, raw Reservation and
  Uploaded, FTP listing time (UTC to the minute), HTTP Last-Modified (HEAD sample), and both gaps.
- `data/*.csv`: per-meeting extracts. `data/small_gaps.csv` holds all gaps under 3 h.
  `data/analysis_output.md` is the raw analysis output.
- `scripts/extract.py`, `scripts/analyse.py`, `scripts/combine.py`.
- `downloads/`: raw xlsx files, listings, HEAD results (`head_138.tsv`) and the portal probes.
