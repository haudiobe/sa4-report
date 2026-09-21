# Document Reallocation System

## Overview

This system allows flexible reallocation of documents between agenda items during the meeting, with persistent configuration stored in a table at the beginning of the report.

---

## Use Cases

### 1. Cross-SWG Reallocation
- **Example**: Liaison document S4-260123 initially under 5.3 → moved to MBS agenda 8.3
- **Reason**: Document scope changed or was initially misclassified

### 2. Within-SWG Reallocation
- **Example**: MBS document S4-260456 initially under 8.7 → moved to 8.3
- **Reason**: Better fit with discussion topic

### 3. Persistent Configuration
- Reallocations stored in configuration table
- Survives report updates and rebuilds
- Can be edited manually or via UI

---

## Configuration Table Structure

### Table: "Document Reallocations"

Located at the beginning of the report (after Configuration tables).

| TDoc | Original Agenda | New Agenda | Reason (Optional) |
|------|----------------|------------|-------------------|
| S4-260123 | 5.3 | 8.3 | Scope changed to MBS |
| S4-260456 | 8.7 | 8.3 | Better topic fit |
| S4-260789 | 11.2 | 11.5 | Moved during discussion |

### Fields:
- **TDoc**: Document number (e.g., S4-260123)
- **Original Agenda**: Where it was initially placed (e.g., 5.3)
- **New Agenda**: Where it should be moved (e.g., 8.3)
- **Reason**: Optional explanation for the move

---

## Implementation Strategy

### Phase 1: Configuration Table Management

```javascript
/**
 * Create or ensure reallocation table exists
 */
function ensureReallocationTable_() {
  const body = DocumentApp.getActiveDocument().getBody();
  
  // Check if table already exists
  for (const t of body.getTables()) {
    if (isReallocationTable_(t)) return;
  }
  
  // Create new table after config tables
  const insertIndex = findInsertIndexAfterConfig_(body);
  
  body.insertParagraph(insertIndex, 'Document Reallocations')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.insertTable(insertIndex + 1, [
    ['TDoc', 'Original Agenda', 'New Agenda', 'Reason'],
    // Empty rows for user to fill
  ]);
}

function isReallocationTable_(table) {
  try {
    const row0 = table.getRow(0);
    return row0.getCell(0).getText() === 'TDoc' &&
           row0.getCell(1).getText() === 'Original Agenda' &&
           row0.getCell(2).getText() === 'New Agenda';
  } catch (e) {
    return false;
  }
}

/**
 * Read reallocation configuration
 * Returns: { 'S4-260123': { original: '5.3', new: '8.3', reason: '...' }, ... }
 */
function getReallocationMap_() {
  const body = DocumentApp.getActiveDocument().getBody();
  const map = {};
  
  for (const table of body.getTables()) {
    if (!isReallocationTable_(table)) continue;
    
    for (let r = 1; r < table.getNumRows(); r++) {
      const tdoc = table.getCell(r, 0).getText().trim();
      const original = table.getCell(r, 1).getText().trim();
      const newAgenda = table.getCell(r, 2).getText().trim();
      const reason = table.getCell(r, 3).getText().trim();
      
      if (tdoc && newAgenda) {
        map[tdoc] = {
          original: original,
          new: newAgenda,
          reason: reason
        };
      }
    }
  }
  
  return map;
}
```

### Phase 2: Apply Reallocations During Report Build

```javascript
/**
 * Modified processWebDownloadedSheet_ to apply reallocations
 */
function processWebDownloadedSheet_(sheet) {
  const cfg = getReportConfig_();
  const suffix = cfg.REPORT_SUFFIX || '6G';
  const agendaPrefix = getAgendaPrefixForReportType_(suffix);
  
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  const data = sheet.getDataRange().getValues();
  const richTextValues = sheet.getDataRange().getRichTextValues();
  
  // Get reallocation map
  const reallocations = getReallocationMap_();
  
  // Find column indices
  const headers = data[0];
  const tdocCol = headers.indexOf('TDoc');
  const agendaCol = headers.indexOf('Agenda item');
  
  // Group TDOCs by agenda item (with reallocations applied)
  const agendaGroups = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tdoc = String(row[tdocCol] || '').trim();
    let agendaItem = String(row[agendaCol] || '').trim();
    
    // Apply reallocation if exists
    if (reallocations[tdoc]) {
      agendaItem = reallocations[tdoc].new;
      Logger.log(`Reallocating ${tdoc}: ${reallocations[tdoc].original} → ${agendaItem}`);
    }
    
    // Filter by report type (after reallocation)
    if (!agendaItem.startsWith(agendaPrefix)) {
      continue;
    }
    
    // ... rest of grouping logic
  }
  
  // ... rest of processing
}
```

### Phase 3: UI for Managing Reallocations

```javascript
/**
 * Menu function to add reallocation
 */
function addDocumentReallocation() {
  const ui = DocumentApp.getUi();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
    </style>
    
    <h2>Add Document Reallocation</h2>
    
    <label>TDoc Number:</label>
    <input type="text" id="tdoc" placeholder="S4-260123">
    
    <label>Original Agenda:</label>
    <input type="text" id="original" placeholder="5.3">
    
    <label>New Agenda:</label>
    <input type="text" id="newAgenda" placeholder="8.3">
    
    <label>Reason (Optional):</label>
    <input type="text" id="reason" placeholder="Scope changed to MBS">
    
    <button onclick="saveReallocation()">Add Reallocation</button>
    <button onclick="google.script.host.close()" style="background: #666;">Cancel</button>
    
    <script>
      function saveReallocation() {
        const data = {
          tdoc: document.getElementById('tdoc').value,
          original: document.getElementById('original').value,
          newAgenda: document.getElementById('newAgenda').value,
          reason: document.getElementById('reason').value
        };
        google.script.run
          .withSuccessHandler(() => {
            alert('Reallocation added! Rebuild report to apply changes.');
            google.script.host.close();
          })
          .addReallocationToTable(data);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(450);
  
  ui.showModalDialog(html, 'Add Document Reallocation');
}

function addReallocationToTable(data) {
  ensureReallocationTable_();
  
  const body = DocumentApp.getActiveDocument().getBody();
  
  for (const table of body.getTables()) {
    if (!isReallocationTable_(table)) continue;
    
    // Add new row
    const row = table.appendTableRow();
    row.appendTableCell(data.tdoc);
    row.appendTableCell(data.original);
    row.appendTableCell(data.newAgenda);
    row.appendTableCell(data.reason || '');
    
    Logger.log(`Added reallocation: ${data.tdoc} → ${data.newAgenda}`);
    return;
  }
}
```

---

## Agenda Structure from Meeting Agenda Document

### Concept

Instead of hardcoding agenda items, download and parse the official meeting agenda document (e.g., S4-260868 for SA4#136).

### Benefits:
1. **Automatic Structure**: All clauses and subclauses created automatically
2. **Proper Headings**: Official agenda topic descriptions used
3. **Complete Coverage**: No missing agenda items
4. **Consistent**: Matches official meeting structure

### Implementation

```javascript
/**
 * Download and parse meeting agenda document
 */
function downloadMeetingAgenda(agendaTdoc) {
  // Example: S4-260868 for SA4#136
  const cfg = getReportConfig_();
  const ftpBase = cfg.FTP_BASE || DEFAULT_S4_FTP_BASE;
  const url = `${ftpBase}${agendaTdoc}.zip`;
  
  Logger.log('Downloading agenda document: ' + url);
  
  try {
    // Download ZIP
    const response = UrlFetchApp.fetch(url);
    const zipBlob = response.getBlob();
    
    // Extract Word document from ZIP
    const wordBlob = extractWordFromZip_(zipBlob);
    
    // Parse agenda structure
    const agendaStructure = parseAgendaDocument_(wordBlob);
    
    return agendaStructure;
    
  } catch (e) {
    Logger.log('Error downloading agenda: ' + e.message);
    throw e;
  }
}

/**
 * Parse agenda document to extract structure
 * Returns: [
 *   { number: '5.3', title: 'Liaison statements', level: 2 },
 *   { number: '7.1', title: 'Audio codec enhancements', level: 2 },
 *   { number: '7.1.1', title: 'EVS enhancements', level: 3 },
 *   ...
 * ]
 */
function parseAgendaDocument_(wordBlob) {
  // Convert Word to Google Doc temporarily
  const tempFile = DriveApp.createFile(wordBlob);
  const docId = tempFile.getId();
  
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const structure = [];
    
    // Parse paragraphs looking for agenda items
    for (let i = 0; i < body.getNumChildren(); i++) {
      const element = body.getChild(i);
      
      if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const para = element.asParagraph();
        const text = para.getText().trim();
        
        // Match patterns like "5.3 Liaison statements"
        const match = text.match(/^(\d+(?:\.\d+)*)\s+(.+)$/);
        
        if (match) {
          const number = match[1];
          const title = match[2];
          const level = (number.match(/\./g) || []).length + 1;
          
          structure.push({
            number: number,
            title: title,
            level: level,
            heading: para.getHeading()
          });
        }
      }
    }
    
    return structure;
    
  } finally {
    // Clean up temp file
    DriveApp.getFileById(docId).setTrashed(true);
  }
}

/**
 * Create document structure based on agenda
 */
function createDocumentStructureFromAgenda(agendaStructure, reportType) {
  const body = DocumentApp.getActiveDocument().getBody();
  const agendaPrefix = getAgendaPrefixForReportType_(reportType);
  
  // Filter agenda items for this report type
  const relevantItems = agendaStructure.filter(item => 
    item.number.startsWith(agendaPrefix)
  );
  
  // Create headings for each agenda item
  relevantItems.forEach(item => {
    const headingText = `${item.number} ${item.title}`;
    const heading = body.appendParagraph(headingText);
    
    // Set heading level based on depth
    if (item.level === 1) {
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (item.level === 2) {
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else {
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    }
    
    Logger.log(`Created heading: ${headingText}`);
  });
}
```

---

## Complete Workflow

### Step 1: Configure Meeting
```
1. Meeting Name: TSGS4_136_Montreal
2. Report Type: 6G
3. Agenda TDoc: S4-260868
```

### Step 2: Build Initial Structure
```
1. Download agenda document (S4-260868)
2. Parse agenda structure
3. Create all headings for report type
4. Download TDOC list
5. Place documents under correct headings
```

### Step 3: During Meeting - Reallocate Documents
```
1. Use menu: "Add Document Reallocation"
2. Enter: S4-260123, from 5.3 to 8.3
3. Rebuild report → document moves automatically
```

### Step 4: Update Report
```
1. Reallocations persist in table
2. New documents placed correctly
3. Moved documents stay in new location
```

---

## Menu Structure Update

```javascript
function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('⚠️Scripts⚠️');
  
  // ... existing menus ...
  
  // DOCUMENT MANAGEMENT submenu
  const docMenu = ui.createMenu('📋 DOCUMENT MANAGEMENT');
  docMenu.addItem('➕ Add Document Reallocation', 'addDocumentReallocation');
  docMenu.addItem('📊 View Reallocations', 'viewReallocations');
  docMenu.addItem('🗑️ Clear Reallocations', 'clearReallocations');
  
  menu.addSubMenu(docMenu);
  menu.addToUi();
}
```

---

## Benefits

### 1. Flexibility
- Move documents any time during meeting
- Changes persist across updates
- Easy to track what was moved and why

### 2. Accuracy
- Official agenda structure used
- All subclauses included
- Proper heading hierarchy

### 3. Transparency
- Reallocation table shows all moves
- Reason field documents why
- Easy to review and audit

### 4. Automation
- Agenda structure auto-generated
- Documents auto-placed
- Reallocations auto-applied

---

## Implementation Priority

1. **Phase 1**: Reallocation table and basic functionality
2. **Phase 2**: UI for adding reallocations
3. **Phase 3**: Agenda document parsing
4. **Phase 4**: Auto-structure creation

---

## Example Configuration

### Meeting Configuration
```
Meeting Name: TSGS4_136_Montreal
Meeting Number: 136
Report Type: 6G
Agenda TDoc: S4-260868
```

### Document Reallocations Table
```
| TDoc       | Original | New  | Reason                    |
|------------|----------|------|---------------------------|
| S4-260123  | 5.3      | 8.3  | Scope changed to MBS      |
| S4-260456  | 8.7      | 8.3  | Better topic fit          |
| S4-260789  | 11.2     | 11.5 | Moved during discussion   |
```

### Result
- S4-260123 appears under heading "8.3 [Topic]" instead of "5.3 [Topic]"
- Original placement preserved in table for reference
- Reason documented for future reference