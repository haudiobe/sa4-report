# Incremental Update Logic Analysis

## Current Update Strategy

### 1. **Table Identification & Matching**
```javascript
// Line 657-668: Find existing table by TDOC number
for (var i = 0; i < tables.length; i++) {
  var t = tables[i];
  if (t.getNumRows() < 1) continue;
  var row0 = t.getRow(0);
  if (row0.getNumCells() < 2) continue;
  if (row0.getCell(0).getText().trim() !== 'TDoc') continue;
  if (row0.getCell(1).getText().trim() === b1Value) { 
    existingTable = t; 
    break; 
  }
}
```

**✅ SAFE:** Tables are matched by TDOC number, not position

---

### 2. **Existing Table Update (Lines 670-676)**
```javascript
if (existingTable) {
  mergeMissingRowsFromSheet_(sheet, existingTable, ftpBase);
  ensureAgendaItemRow_(sheet, existingTable);
  updateExistingTable(sheet, existingTable, ftpBase);
  return;
}
```

**What happens:**
1. **mergeMissingRowsFromSheet_**: Adds NEW rows from Excel that don't exist in doc
2. **ensureAgendaItemRow_**: Updates agenda item if changed
3. **updateExistingTable**: Updates status with strict rules

---

### 3. **Protected Fields (Lines 1305-1309)**
```javascript
// never override these
if (kl === 'tdoc status' || kl === 'status') continue;
if (kl.startsWith('revised')) continue;
if (kl.indexOf('e-mail discussion') !== -1) continue;
if (kl.indexOf('revisions') !== -1) continue;
```

**✅ SAFE:** These fields are NEVER overwritten during merge:
- Status (protected by strict rules)
- Revised to
- E-mail discussion (accumulated, never cleared)
- Revisions (accumulated, never cleared)

---

### 4. **Status Update Rules (Lines 742-763)**
```javascript
var docStatus = normalizeStatus_(d.value);
var sheetStatusRaw = String(s.value || '').trim().toLowerCase();
var sheetIsRevised = sheetStatusRaw.indexOf('revised') !== -1;

// Allowed updates:
// 1) doc is reserved OR available
// 2) sheet is revised
var allowUpdate = (docStatus === 'reserved' || docStatus === 'available' || sheetIsRevised);
if (!allowUpdate) return;
```

**✅ SAFE:** Status can ONLY be updated if:
- Current status is "Reserved" or "Available" (not decided yet)
- OR new status contains "revised" (always allow revision updates)

**❌ PROTECTED:** Once status is "Agreed", "Noted", "Endorsed", "Withdrawn" → LOCKED

---

### 5. **Email Discussion Accumulation (Lines 823-880)**
```javascript
const storeKey = 'DISCUSS_' + tdoc;
const store = loadJsonObject_(props.getProperty(storeKey));

msgs.forEach(m => {
  const id = (m.messageId || m.link || '').trim();
  if (!store[id]) { store[id] = m; added++; }
});

props.setProperty(storeKey, JSON.stringify(store));
```

**✅ SAFE:** 
- Emails stored in Document Properties (persistent)
- New emails ADDED to existing collection
- Never cleared or overwritten
- Sorted by date before display

---

### 6. **Revisions Accumulation (Lines 1021-1063)**
```javascript
const storeKey = 'REVIS_' + tdoc;
const store = loadJsonObject_(props.getProperty(storeKey));

anchors.forEach(a => {
  if (!store[id]) { store[id] = { text, link: abs }; added++; }
});

props.setProperty(storeKey, JSON.stringify(store));
```

**✅ SAFE:**
- Revisions stored in Document Properties (persistent)
- New revisions ADDED to existing collection
- Never cleared or overwritten

---

## Summary: Incremental Update Safety

### ✅ **SAFE Operations (Non-Destructive)**
1. **New TDOCs**: Added to document
2. **Email discussions**: Accumulated over time
3. **Revisions**: Accumulated over time
4. **Missing rows**: Added from Excel if not in doc
5. **Agenda item**: Updated if changed in Excel
6. **Status updates**: Only if "Reserved"/"Available" or contains "revised"

### ❌ **PROTECTED Fields (Never Overwritten)**
1. **Decided status**: "Agreed", "Noted", "Endorsed", "Withdrawn"
2. **Minutes**: Never touched by automation
3. **Disposition**: Never touched by automation
4. **Email discussion**: Only additions, never deletions
5. **Revisions**: Only additions, never deletions

### ⚠️ **UPDATED Fields (Conditional)**
1. **Status**: Only if Reserved/Available or revised
2. **Agenda Item**: Updated if changed in Excel
3. **Title/Source/Contact**: Updated from Excel (safe, factual data)

---

## Recommendation

**The current logic is SAFE for incremental updates** with these caveats:

1. ✅ Run updates multiple times during meeting
2. ✅ Email/revisions accumulate automatically
3. ✅ Decided statuses are protected
4. ⚠️ Minutes/Disposition must be entered in doc (not Excel)
5. ⚠️ If agenda item changes in Excel, it updates in doc

**Best Practice:**
- Update frequently during meeting
- Enter Minutes/Disposition directly in Google Doc
- Status decisions are protected once made
- Email/revisions auto-collect in background