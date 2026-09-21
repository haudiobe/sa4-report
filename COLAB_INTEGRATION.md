# Colab Notebook Integration Guide

## Overview

Your Colab notebook downloads TDOC data from 3GPP and creates structured Google Docs/Sheets templates. This guide shows how to integrate it with your Apps Script for a complete workflow.

---

## 📊 What Your Colab Notebook Does

### Input
- Downloads Excel file from: `https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx`
- Processes TDOC data (number, title, source, agenda item, status, etc.)

### Processing
1. **Filters by report type** (6G, MBS, Video, Audio, RTC, Liaison, New)
2. **Tracks agenda items** - maintains an agenda tracking sheet for revisions
3. **Sorts documents** by agenda item and TDOC number
4. **Preserves hyperlinks** from the original Excel

### Output (per report type)
1. **Google Doc** (`Template-Report-{type}.docx`) - Simple table format
2. **Google Sheet** (`Template-Report-{type}.xlsx`) - Single sheet with all TDOCs
3. **Google Sheet Individual** (`Template-Report-{type}-Individual.xlsx`) - One sheet per TDOC with detailed structure
4. **Agenda Tracking Sheet** (`Template-Report-AgendaItems.xlsx`) - Tracks agenda item changes

---

## 🔄 Integration Architecture

### Current Workflow
```
3GPP Website → Colab Downloads → Creates Templates → Manual Copy to Apps Script
```

### Improved Workflow
```
3GPP Website → Colab Downloads → Creates Templates → Apps Script Auto-Imports → Final Report
```

---

## 🎯 Integration Strategy

### Option 1: Apps Script Reads from Colab Output (Recommended)

**Flow:**
```
Colab Notebook → Google Sheets (Individual) → Apps Script → Google Doc Report
```

**Why this works:**
- Your Colab already creates `Template-Report-{type}-Individual.xlsx`
- Each sheet has the exact structure your Apps Script expects
- Apps Script already has code to read from Excel/Sheets

**Implementation:**

#### Step 1: Modify Apps Script to Use Colab Output

Add this function to your `Code.js`:

```javascript
/**
 * Import data from Colab-generated Individual sheets
 * This replaces the need for manual Excel templates
 */
function importFromColabSheets() {
  const cfg = getReportConfig_();
  const suffix = cfg.REPORT_SUFFIX || '6G';
  
  // The file created by your Colab notebook
  const colabFileName = `Template-Report-${suffix}-Individual.xlsx`;
  
  // Find the file in Drive
  const files = DriveApp.getFilesByName(colabFileName);
  if (!files.hasNext()) {
    Logger.log('Colab output file not found: ' + colabFileName);
    DocumentApp.getUi().alert('Please run the Colab notebook first to generate: ' + colabFileName);
    return;
  }
  
  const file = files.next();
  const spreadsheet = SpreadsheetApp.open(file);
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Process each sheet (each TDOC)
  const sheets = spreadsheet.getSheets();
  Logger.log(`Found ${sheets.length} TDOCs in Colab output`);
  
  sheets.forEach(function(sheet) {
    copyTableToDoc_(sheet, body);
  });
  
  Logger.log('Import from Colab sheets complete');
}

/**
 * Update menu to include Colab import
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('⚠️Scripts⚠️')
    .addItem('Update (ALL – recommended)', 'updateAll')
    .addItem('Import from Colab Output', 'importFromColabSheets')  // NEW
    .addItem('Update (Copy tables only)', 'copyIndividualToReport')
    .addItem('Update (Email + Revisions only)', 'collectorUpdate_')
    .addSeparator()
    .addItem('Format only', 'removeRowHeightAndSpacing')
    .addItem('Fix Links (portal→FTP)', 'rewritePortalLinksInDoc_')
    .addItem('Clear A1 empty cache', 'clearA1EmptyCache_')
    .addToUi();
}
```

#### Step 2: Update Your Workflow

**New Process:**
1. Run Colab notebook (creates templates in Google Drive)
2. Open your report Google Doc
3. Click **⚠️Scripts⚠️ → Import from Colab Output**
4. Click **⚠️Scripts⚠️ → Update (Email + Revisions only)**
5. Click **⚠️Scripts⚠️ → Format only**

---

### Option 2: Direct API Integration

**For advanced automation:**

#### Colab: Add Export Function

Add this to the end of your Colab notebook:

```python
def export_to_apps_script_format(report_df, target_folder_id, report_type):
    """
    Export data in a format optimized for Apps Script import
    Creates a JSON file with all TDOC data
    """
    import json
    from googleapiclient.http import MediaFileUpload
    
    # Prepare data structure
    export_data = {
        'report_type': report_type,
        'generated': pd.Timestamp.now().isoformat(),
        'tdocs': []
    }
    
    for _, row in report_df.iterrows():
        export_data['tdocs'].append({
            'number': row['TDoc'],
            'title': row['Title'],
            'source': row['Source'],
            'agenda_item': row['Agenda item'],
            'status': row.get('TDoc Status', ''),
            'contact': row.get('Contact', ''),
            'hyperlink': row.get('Hyperlink', '')
        })
    
    # Save as JSON
    json_filename = f'colab-export-{report_type}.json'
    json_path = f'/tmp/{json_filename}'
    
    with open(json_path, 'w') as f:
        json.dump(export_data, f, indent=2)
    
    # Upload to Drive
    file_metadata = {
        'name': json_filename,
        'parents': [target_folder_id]
    }
    
    media = MediaFileUpload(json_path, mimetype='application/json')
    
    # Delete existing file if present
    results = drive_service.files().list(
        q=f"name='{json_filename}' and '{target_folder_id}' in parents",
        spaces='drive',
        fields='files(id)'
    ).execute()
    
    for item in results.get('files', []):
        drive_service.files().delete(fileId=item['id']).execute()
    
    # Upload new file
    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    
    print(f'Exported {len(export_data["tdocs"])} TDOCs to {json_filename}')
    return file.get('id')

# Call after creating templates
if target_folder_id_for_templates:
    export_to_apps_script_format(
        report_df, 
        target_folder_id_for_templates, 
        '6G'
    )
```

#### Apps Script: Add JSON Import

```javascript
function importFromColabJSON() {
  const cfg = getReportConfig_();
  const suffix = cfg.REPORT_SUFFIX || '6G';
  const jsonFileName = `colab-export-${suffix}.json`;
  
  // Find JSON file
  const files = DriveApp.getFilesByName(jsonFileName);
  if (!files.hasNext()) {
    Logger.log('Colab JSON not found: ' + jsonFileName);
    return;
  }
  
  const file = files.next();
  const content = file.getBlob().getDataAsString();
  const data = JSON.parse(content);
  
  Logger.log(`Importing ${data.tdocs.length} TDOCs from Colab`);
  
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Process each TDOC
  data.tdocs.forEach(tdoc => {
    createTDocTable(body, tdoc);
  });
  
  Logger.log('Import complete');
}

function createTDocTable(body, tdoc) {
  // Create table structure matching your format
  const table = body.appendTable();
  
  table.appendTableRow().appendTableCell('TDoc').appendTableCell(tdoc.number);
  table.appendTableRow().appendTableCell('Title').appendTableCell(tdoc.title);
  table.appendTableRow().appendTableCell('Source').appendTableCell(tdoc.source);
  table.appendTableRow().appendTableCell('Contact').appendTableCell(tdoc.contact);
  table.appendTableRow().appendTableCell('Agenda Item').appendTableCell(tdoc.agenda_item);
  table.appendTableRow().appendTableCell('E-mail Discussion').appendTableCell('');
  table.appendTableRow().appendTableCell('Revisions').appendTableCell('');
  table.appendTableRow().appendTableCell('Minutes').appendTableCell('');
  table.appendTableRow().appendTableCell('Disposition').appendTableCell('');
  table.appendTableRow().appendTableCell('Status').appendTableCell(tdoc.status);
  
  // Apply hyperlink to TDoc if present
  if (tdoc.hyperlink) {
    const cell = table.getRow(0).getCell(1);
    const text = cell.editAsText();
    text.setLinkUrl(0, tdoc.number.length - 1, tdoc.hyperlink);
  }
  
  return table;
}
```

---

## 🔧 Configuration

### Colab Notebook Settings

Update these variables in your Colab notebook:

```python
# Base folder in Google Drive
base_url = '/content/drive/MyDrive/3GPP'

# Meeting folder name
meeting_folder = 'SA4#136'

# Excel source URL
excel_url = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx'

# Report types to generate
report_types = ['6G', 'MBS', 'Video', 'Audio', 'RTC']
```

### Apps Script Settings

Update in your Google Doc's configuration table:

| Key | Value |
|-----|-------|
| REPORT_SUFFIX | 6G |
| COLAB_FOLDER_ID | (Your Google Drive folder ID) |

---

## 📋 Complete Workflow

### One-Time Setup

1. **Run Colab notebook** to create initial templates
2. **Create your report Google Doc**
3. **Add the integration code** to your Apps Script
4. **Configure folder paths** in both Colab and Apps Script

### Regular Use

**Every Meeting:**

1. **Update Colab settings:**
   ```python
   meeting_folder = 'SA4#137'  # New meeting
   excel_url = 'https://..../SA4%23137.xlsx'  # New URL
   ```

2. **Run Colab notebook:**
   - Downloads latest TDOC list
   - Creates/updates templates
   - Tracks agenda item changes

3. **In Google Doc, run:**
   - `Import from Colab Output` (gets TDOC tables)
   - `Update (Email + Revisions only)` (adds discussions)
   - `Format only` (applies styling)

4. **Review and finalize** the report

---

## 🎯 Benefits of Integration

### Before Integration
- ❌ Manual copying from Excel to Doc
- ❌ Lose hyperlinks
- ❌ Manual agenda item tracking
- ❌ Tedious updates for revisions

### After Integration
- ✅ Automatic data import
- ✅ Hyperlinks preserved
- ✅ Agenda items tracked automatically
- ✅ One-click updates
- ✅ Consistent formatting

---

## 🐛 Troubleshooting

### "Colab output file not found"
**Solution:** Ensure the Colab notebook completed successfully and the files are in the correct Google Drive folder.

### "Permission denied"
**Solution:** Make sure both Colab and Apps Script have access to the same Google Drive folder.

### "Data mismatch"
**Solution:** Check that the REPORT_SUFFIX in Apps Script matches the report type generated by Colab.

### "Hyperlinks not working"
**Solution:** Verify that the Excel file from 3GPP contains hyperlinks in the TDoc column.

---

## 📝 Next Steps

1. **Test Option 1** (recommended) - Add `importFromColabSheets()` to your Apps Script
2. **Run Colab notebook** to generate test data
3. **Test the import** in your Google Doc
4. **Verify** email collection and formatting still work
5. **Document** your specific workflow

---

## 💡 Future Enhancements

### Potential Improvements

1. **Automatic Scheduling**
   - Use Google Cloud Functions to run Colab on schedule
   - Trigger Apps Script automatically after Colab completes

2. **Change Detection**
   - Compare current vs. previous TDOC lists
   - Highlight new/changed documents

3. **Multi-Meeting Support**
   - Process multiple meetings in one run
   - Generate comparison reports

4. **Email Notifications**
   - Send email when Colab completes
   - Alert when new revisions are detected

---

## 📚 Related Documentation

- `README.md` - Project overview
- `WORKFLOW.md` - Development workflow
- `CODE_ANALYSIS.md` - Code structure
- `BUGS_AND_ISSUES.md` - Known issues

---

## 🤝 Support

If you need help with the integration:
1. Check the troubleshooting section above
2. Review the code comments in both Colab and Apps Script
3. Test with a small dataset first
4. Verify folder permissions in Google Drive

Your Colab notebook is well-structured and the integration should be straightforward! 🚀