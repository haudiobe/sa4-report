# The 3GPP Portal TDoc list of meeting 86178, twice

Both files come from the same public address, `https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178`, on 2026-11-25. The meeting is "3GPPSA4-e (AH) on FS_6G_MED".

| File | Downloaded | Content |
|---|---|---|
| `portal-1355utc-118rows.xlsx` | about 13:55 UTC | The meeting's list: 117 TDocs, S4aP260067–S4aP260182, including the meeting's agenda S4aP260149. |
| `portal-1515utc-38rows-other-meeting.xlsx` | about 15:15 UTC | Another meeting's list: 37 TDocs, S4aP260008–S4aP260043, without S4aP260149. The Portal kept serving it with cache-busters and header variants, and its meeting record still said 116 documents. |

The report refused the second list with "it has no row for S4aP260149, the agenda TDoc of this meeting -- it may be the list of another meeting". That refusal is correct. `tests/issue28-portal-identity.test.js` runs the report's own identity check on both files.
