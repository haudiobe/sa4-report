# Integration Plan: Colab Notebook + Apps Script Report Generation

## Overview

This document outlines the integration strategy between your Google Colab notebook (data processing) and the Google Apps Script (report generation).

**Colab Notebook:** https://colab.research.google.com/drive/1WqWdhDVoo2SQRS99WJJa55KdOTgjxxOT

**Apps Script Project:** Script ID `1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev`

---

## Current Architecture

### Component 1: Google Colab Notebook
**Purpose:** Data processing and analysis for 3GPP SA4 reports

**Typical Workflow:**
1. Processes meeting data
2. Analyzes contributions
3. Generates statistics
4. Produces structured data outputs

### Component 2: Google Apps Script
**Purpose:** Report generation and document management

**Current Functions:**
- Copies data from Excel templates to Google Docs
- Collects email discussions from listserv
- Tracks document revisions
- Formats and styles reports
- Manages TDOC tables

---

## Integration Strategies

### Option 1: Direct Data Export (Recommended)

**Flow:**
```
Colab Notebook → Google Sheets → Apps Script → Google Doc Report
```

**Implementation:**

#### Step 1: Colab Exports to Google Sheets
```python
# In your Colab notebook
from google.colab import auth
from google.auth import default
import gspread

# Authenticate
auth.authenticate_user()
creds, _ = default()
gc = gspread.authorize(creds)

# Create or open spreadsheet
spreadsheet = gc.create('3GPP_Report_Data')
# Or open existing: spreadsheet = gc.open_by_key('SPREADSHEET_ID')

# Write data
worksheet = spreadsheet.worksheet('Sheet1')
worksheet.update('A1', your_data_array)

# Share with your account
spreadsheet.share('your.email@example.com', perm_type='user', role='writer')
```

#### Step 2: Apps Script Reads from Sheets
```javascript
// In Code.js - Add new function
function importColabData() {
  const COLAB_SHEET_ID = 'YOUR_SPREADSHEET_ID';
  const ss = SpreadsheetApp.openById(COLAB_SHEET_ID);
  const sheet = ss.getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();
  
  // Process and insert into document
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Create table from data
  const table = body.appendTable(data);
  
  // Apply formatting
  styleTable(table);
  
  return data.length + ' rows imported';
}
```

**Advantages:**
- Simple and reliable
- Google Sheets acts as intermediate storage
- Easy to debug and verify data
- Can manually review before importing

---

### Option 2: Direct API Integration

**Flow:**
```
Colab Notebook → Google Drive API → Apps Script
```

**Implementation:**

#### Step 1: Colab Writes JSON to Drive
```python
# In your Colab notebook
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Prepare data
report_data = {
    'tdocs': [...],
    'statistics': {...},
    'metadata': {...}
}

# Save to Drive
drive_service = build('drive', 'v3', credentials=creds)

file_metadata = {
    'name': 'report_data.json',
    'mimeType': 'application/json'
}

# Write JSON
with open('/tmp/report_data.json', 'w') as f:
    json.dump(report_data, f)

media = MediaFileUpload('/tmp/report_data.json', mimetype='application/json')
file = drive_service.files().create(
    body=file_metadata,
    media_body=media,
    fields='id'
).execute()

print(f'File ID: {file.get("id")}')
```

#### Step 2: Apps Script Reads JSON
```javascript
// In Code.js - Add new function
function importColabJSON() {
  const FILE_ID = 'YOUR_JSON_FILE_ID';
  const file = DriveApp.getFileById(FILE_ID);
  const content = file.getBlob().getDataAsString();
  const data = JSON.parse(content);
  
  // Process data
  processTDocs(data.tdocs);
  updateStatistics(data.statistics);
  
  return 'Data imported successfully';
}

function processTDocs(tdocs) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  tdocs.forEach(tdoc => {
    // Create table for each TDOC
    const table = body.appendTable();
    table.appendTableRow().appendTableCell('TDoc').appendTableCell(tdoc.number);
    table.appendTableRow().appendTableCell('Title').appendTableCell(tdoc.title);
    table.appendTableRow().appendTableCell('Source').appendTableCell(tdoc.source);
    // ... more fields
  });
}
```

**Advantages:**
- More flexible data structure
- Can include complex nested data
- Single source of truth

---

### Option 3: Hybrid Approach (Best for Complex Workflows)

**Flow:**
```
Colab → Sheets (raw data) + JSON (metadata) → Apps Script → Doc
```

**Use Cases:**
- Raw data in Sheets for easy viewing/editing
- Metadata and configuration in JSON
- Apps Script combines both sources

---

## Recommended Implementation Plan

### Phase 1: Setup (Week 1)

1. **Identify Data Requirements**
   - [ ] List all data types from Colab
   - [ ] Define data schema
   - [ ] Determine update frequency

2. **Create Shared Resources**
   - [ ] Create Google Sheet for data exchange
   - [ ] Set up folder structure in Drive
   - [ ] Configure permissions

3. **Update Colab Notebook**
   - [ ] Add Google Sheets export function
   - [ ] Add data validation
   - [ ] Add error handling

### Phase 2: Apps Script Integration (Week 2)

1. **Create Import Functions**
   - [ ] Add `importColabData()` function
   - [ ] Add data validation
   - [ ] Add error handling

2. **Update Existing Functions**
   - [ ] Modify `copyIndividualToReport()` to use Colab data
   - [ ] Update table creation logic
   - [ ] Preserve existing functionality

3. **Add Menu Items**
   ```javascript
   function onOpen() {
     DocumentApp.getUi()
       .createMenu('⚠️Scripts⚠️')
       .addItem('Update (ALL – recommended)', 'updateAll')
       .addItem('Import Colab Data', 'importColabData')  // NEW
       .addItem('Update (Copy tables only)', 'copyIndividualToReport')
       .addSeparator()
       .addItem('Format only', 'removeRowHeightAndSpacing')
       .addToUi();
   }
   ```

### Phase 3: Testing (Week 3)

1. **Unit Testing**
   - [ ] Test Colab export
   - [ ] Test Apps Script import
   - [ ] Test data transformation

2. **Integration Testing**
   - [ ] Test end-to-end workflow
   - [ ] Test error scenarios
   - [ ] Test with real data

3. **Performance Testing**
   - [ ] Measure execution time
   - [ ] Check quota usage
   - [ ] Optimize if needed

### Phase 4: Deployment (Week 4)

1. **Documentation**
   - [ ] Update README with new workflow
   - [ ] Create user guide
   - [ ] Document troubleshooting steps

2. **Rollout**
   - [ ] Deploy to production
   - [ ] Train users
   - [ ] Monitor for issues

---

## Data Schema Design

### Example: TDOC Data Structure

```javascript
// JSON format from Colab
{
  "tdocs": [
    {
      "number": "S4-260879",
      "title": "2D Video Codecs in the 6G Media Study",
      "source": "Orange, Qualcomm, Dolby",
      "agenda_item": "11.3.3",
      "status": "Available",
      "type": "Discussion",
      "revised_to": null,
      "revisions": [],
      "email_discussion": []
    }
  ],
  "metadata": {
    "meeting": "SA4#136",
    "location": "Montreal",
    "date": "2026-05-11",
    "generated": "2026-07-23T17:00:00Z"
  }
}
```

### Google Sheets Format

| TDoc | Title | Source | Agenda Item | Status | Type |
|------|-------|--------|-------------|--------|------|
| S4-260879 | 2D Video Codecs... | Orange, Qualcomm | 11.3.3 | Available | Discussion |

---

## Configuration Management

### Colab Configuration
```python
# config.py in Colab
CONFIG = {
    'output_sheet_id': 'YOUR_SHEET_ID',
    'output_folder_id': 'YOUR_FOLDER_ID',
    'apps_script_url': 'YOUR_SCRIPT_URL',
    'meeting_number': 136,
    'meeting_location': 'Montreal'
}
```

### Apps Script Configuration
```javascript
// In Code.js
const COLAB_CONFIG = {
  SHEET_ID: 'YOUR_SHEET_ID',
  JSON_FILE_ID: 'YOUR_JSON_FILE_ID',
  AUTO_IMPORT: false,  // Set to true for automatic import
  IMPORT_ON_OPEN: false
};
```

---

## Error Handling

### Colab Error Handling
```python
def export_to_sheets(data, sheet_id):
    try:
        gc = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(sheet_id)
        worksheet = spreadsheet.worksheet('Sheet1')
        worksheet.clear()
        worksheet.update('A1', data)
        print(f"✓ Exported {len(data)} rows to Google Sheets")
        return True
    except Exception as e:
        print(f"✗ Export failed: {str(e)}")
        return False
```

### Apps Script Error Handling
```javascript
function importColabData() {
  try {
    const data = fetchColabData();
    if (!data || data.length === 0) {
      throw new Error('No data received from Colab');
    }
    
    processData(data);
    Logger.log('✓ Import successful');
    return 'Success';
    
  } catch (error) {
    Logger.log('✗ Import failed: ' + error.message);
    DocumentApp.getUi().alert('Import failed: ' + error.message);
    return 'Failed';
  }
}
```

---

## Automation Options

### Option A: Manual Trigger
- User runs Colab notebook
- User clicks "Import Colab Data" in Apps Script menu
- Simple and controlled

### Option B: Time-Based Trigger
```javascript
// Set up in Apps Script
function createTimeTrigger() {
  ScriptApp.newTrigger('importColabData')
    .timeBased()
    .everyHours(1)
    .create();
}
```

### Option C: Event-Based Trigger
- Colab writes to specific location
- Apps Script watches for changes
- Automatically imports new data

---

## Next Steps

1. **Review this plan** and decide on integration strategy
2. **Share Colab notebook** code so I can see the data structure
3. **Identify specific data** that needs to flow from Colab to Apps Script
4. **Choose implementation option** (1, 2, or 3)
5. **Start with Phase 1** setup

---

## Questions to Answer

1. What data does your Colab notebook produce?
2. How often does the data need to be updated?
3. Should the integration be automatic or manual?
4. Are there any data transformations needed?
5. What's the current data format in Colab?

---

## Resources

- [Google Colab + Sheets Integration](https://colab.research.google.com/notebooks/io.ipynb)
- [Apps Script + Sheets](https://developers.google.com/apps-script/reference/spreadsheet)
- [Apps Script + Drive](https://developers.google.com/apps-script/reference/drive)
- [Google Drive API](https://developers.google.com/drive/api/v3/about-sdk)

---

## Contact

Once you provide more details about your Colab notebook's output, I can create specific implementation code for the integration.