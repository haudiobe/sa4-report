# No update without a saved meeting

**Status: IMPLEMENTED, NOT RELEASED, NOT DEPLOYED. On branch `fix/report-structure-order`, the
commit after the report-structure fix (`docs/REPORT_STRUCTURE.md`). `Code.js` is still
2.22.0.**

## What happened

The first real-document smoke of the structure fix ran on a Drive copy of the MBS report. A
Drive copy starts with empty Document Properties (verified in TEMPLATE-001), and the copy's
meeting had not been saved. Update Report Now still ran: it downloaded the TDoc list of
SA4#136, grouped it by agenda prefix 11., inserted 60 TDocs that have nothing to do with the
report, and read the FS_6G_MED mailing list.

## Why

`getReportConfig_()` fills every missing meeting value with a built-in default:

| Missing | Becomes |
|---|---|
| `MEETING_FOLDER`, `MEETING_NUMBER` | `TSGS4_136_Montreal`, `136` |
| `TDOC_LIST_URL` | `…/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx` |
| `REPORT_SUFFIX` | `6G`: mailing list `3GPP_TSG_SA4_FS_6G_MED`, agenda prefix `11.` |

A missing `MEETING_TYPE` counts as `main` (`normalizeMeetingType_()`), and the readiness rules
never block a main meeting (ADDON-007B3: "legacy fallbacks keep working"). The ad-hoc rules —
TDoc list, agenda source, document folder, family — apply only once `MEETING_TYPE` is `adhoc`,
which only a Save writes. Resolve writes nothing.

## The rule now

In the **template runtime** (a report with `Release.js`), an update or a build needs a saved
meeting. One of:

- `MEETING_TYPE` = `adhoc` — and then the ad-hoc readiness rules, as before;
- `MEETING_TYPE` = `main`;
- no `MEETING_TYPE`, but a saved report family **and** a saved main-meeting location
  (`MEETING_FOLDER`, `TDOC_LIST_URL` or `FTP_BASE`): a main meeting saved by hand, without
  Resolve.

Anything else — an empty copy, Resolve without Save, a meeting number and name alone, a
family alone, an unknown meeting type — is refused **before anything is downloaded and before
the document is touched**:

> The report was not updated: This report has not been configured for a meeting yet: no saved
> meeting configuration was found. Open Configure Meeting, Resolve the meeting, Discover
> Agenda / TDocs where it is an ad-hoc meeting, and Save. Nothing was changed.

The check is in `assertMeetingReadyToBuild_()`, which every update path and the build already
pass first: Update Report Now, the automatic update (`continuousUpdate()`, the trigger) and
Build Report from Scratch. A refusal of the ad-hoc readiness rules now says "The report was
not updated" too, instead of "may have been updated in part".

What does not change:

- **The defaults themselves.** `getReportConfig_()` is untouched; the guard decides whether
  to run, not what the defaults are.
- **The CENTRAL add-on and Legacy reports** (no `Release.js`): unchanged. They may rely on the
  defaults, and they are not copied as reports.
- **Configure Meeting** opens and saves as before; after its Save the report updates.

## Live reports

The live 6G and MBS reports update every hour from their own ad-hoc TDoc lists (MBS:
`meetingId=86172`). That needs a saved family and a saved TDoc list or document folder — the
built-in defaults would give the SA4#136 list — so they pass the guard whatever their saved
meeting type; and MBS was configured with Resolve and Save (`docs/EXISTING_REPORT_MIGRATION.md`
§5.4), which stores `adhoc`. Not read from the live documents. A template report created by
the report creator gets its meeting type from Resolve. A main template report saved with a
meeting type, or with family and folder, passes as well.

Not covered: Update E-mail Discussions, Update TDoc Revisions and Update Abstracts in an
unconfigured report still read the defaults. None of them adds a TDoc or a section.

## Tests

`tests/configuration-guard.test.js`.
