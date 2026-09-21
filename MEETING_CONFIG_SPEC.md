# Meeting Configuration Specification

## Configuration Inputs

### User Provides:
1. **Meeting Name** (e.g., "TSGS4_136_Montreal")
2. **Meeting ID** (e.g., "60777") - from portal.3gpp.org
3. **Report Type** (6G, Audio, Video, MBS, RTC, Liaison, New)
4. **Reviewer API Token** (optional)

### System Automatically Constructs:

#### 1. Mailing List Name
Based on report type:
- Audio → `3GPP_TSG_SA_WG4_AUDIO`
- Video → `3GPP_TSG_SA_WG4_VIDEO`
- MBS → `3GPP_TSG_SA_WG4_MBS`
- RTC → `3GPP_TSG_SA_WG4_RTC`
- 6G/Liaison/New → `3GPP_TSG_SA_WG4`

#### 2. TDOC List URL
Pattern: `https://ftp.3gpp.org/tsg_sa/WG4_CODEC/{MEETING_NAME}/Docs/TDoc_List_Meeting_SA4#{MEETING_NUMBER}.xlsx`

Examples:
- `https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4#136.xlsx`
- `https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_137_Sophia_Antipolis-e/Docs/TDoc_List_Meeting_SA4#137-e.xlsx`

#### 3. Revisions Folder URL
Pattern: `https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/{MEETING_NAME}/Inbox/Drafts/{DRAFTS_FOLDER}`

Drafts folder by report type:
- Audio → `Audio` (SWG name)
- Video → `Video` (SWG name)
- MBS → `MBS` (SWG name)
- RTC → `RTC` (SWG name)
- 6G → `FS_6G_MED`
- Liaison → `Plenary`
- New → `Plenary`

Examples:
- `https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED`
- `https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/Audio`

#### 4. Email Archive Date Range
From portal.3gpp.org meeting page:
- Extract registration deadline date
- Search emails from 7 days before registration deadline until meeting end

---

## URL Construction Functions

### extractMeetingNumber(meetingName)
```javascript
// "TSGS4_136_Montreal" → "136"
// "TSGS4_137_Sophia_Antipolis-e" → "137-e"
```

### constructTdocListUrl(meetingName, meetingNumber)
```javascript
const baseUrl = 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC';
return `${baseUrl}/${meetingName}/Docs/TDoc_List_Meeting_SA4#${meetingNumber}.xlsx`;
```

### constructRevisionsUrl(meetingName, reportType)
```javascript
const baseUrl = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC';
const folder = DRAFTS_FOLDERS[reportType];
return `${baseUrl}/${meetingName}/Inbox/Drafts/${folder}`;
```

### getMailingListName(reportType)
```javascript
return MAILING_LISTS[reportType];
```

---

## Configuration Dialog Flow

### Step 1: Basic Information
```
Meeting Name: [TSGS4_136_Montreal]
Meeting Number: [136] (extracted from name, editable)
Report Type: [6G ▼]
```

### Step 2: Auto-Generated URLs (Preview)
```
✅ TDOC List URL:
https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4#136.xlsx

✅ Revisions Folder:
https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED

✅ Mailing List:
3GPP_TSG_SA_WG4

[Test Connections] [Save] [Cancel]
```

### Step 3: Optional Settings
```
Reviewer API Token: [crv1_...] (optional)
Email Archive Days: [7] (days before registration)
```

---

## Validation Rules

1. **Meeting Name Format:**
   - Must match pattern: `TSGS4_\d+_.*`
   - Examples: `TSGS4_136_Montreal`, `TSGS4_137_Sophia_Antipolis-e`

2. **Meeting Number:**
   - Extracted from meeting name
   - Can include suffix (e.g., "137-e" for electronic meeting)

3. **URL Validation:**
   - Test TDOC list URL (HTTP 200)
   - Test revisions folder (HTTP 200)
   - Test mailing list RSS feed (HTTP 200)

---

## Storage

### Document Properties:
- `MEETING_NAME`: "TSGS4_136_Montreal"
- `MEETING_NUMBER`: "136"
- `REPORT_SUFFIX`: "6G"
- `TDOC_LIST_URL`: (auto-constructed)
- `REVISIONS_URL`: (auto-constructed)
- `MAILING_LIST`: (auto-constructed)

### Script Properties:
- `REVIEWER_API_TOKEN`: "crv1_..."

---

## Example Configurations

### Example 1: SA4#136 Montreal (6G Report)
```
Meeting Name: TSGS4_136_Montreal
Meeting Number: 136
Report Type: 6G

Generated:
- TDOC List: .../TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4#136.xlsx
- Revisions: .../TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED
- Mailing List: 3GPP_TSG_SA_WG4
```

### Example 2: SA4#136 Montreal (Audio Report)
```
Meeting Name: TSGS4_136_Montreal
Meeting Number: 136
Report Type: Audio

Generated:
- TDOC List: .../TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4#136.xlsx
- Revisions: .../TSGS4_136_Montreal/Inbox/Drafts/Audio
- Mailing List: 3GPP_TSG_SA_WG4_AUDIO
```

### Example 3: SA4#137-e Electronic Meeting
```
Meeting Name: TSGS4_137_Sophia_Antipolis-e
Meeting Number: 137-e
Report Type: Video

Generated:
- TDOC List: .../TSGS4_137_Sophia_Antipolis-e/Docs/TDoc_List_Meeting_SA4#137-e.xlsx
- Revisions: .../TSGS4_137_Sophia_Antipolis-e/Inbox/Drafts/Video
- Mailing List: 3GPP_TSG_SA_WG4_VIDEO
```

---

## Implementation Checklist

- [ ] Add URL construction helper functions
- [ ] Update configuration dialog with meeting name input
- [ ] Auto-extract meeting number from name
- [ ] Show preview of generated URLs
- [ ] Add "Test Connections" button in dialog
- [ ] Update getCollectorConfig_() to use dynamic mailing list
- [ ] Update RSS URL construction with correct list name
- [ ] Add validation for meeting name format
- [ ] Store all configuration in document properties
- [ ] Update all functions to use new configuration

---

## Migration from Old System

Old system required manual entry of:
- TDOC_LIST_URL
- REVISIONS_URL
- LIST_NAME (hardcoded)

New system only requires:
- Meeting Name
- Report Type

Everything else is auto-constructed!