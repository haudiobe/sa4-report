# Implementation Plan for 6 Requested Features

## 1. Fix Heading Format (11.1 + Description)

**Current:** `Architecture and Framework`
**Requested:** `11.1 General and working documents`

**Implementation:**
- Change line 437 from `const headingText = group.topic || agendaItem;`
- To: `const headingText = agendaItem + (group.topic ? ' ' + group.topic : '');`

---

## 2. Add "Revised to" Logic

**Current:** Status shows "Revised" but doesn't show what it's revised to
**Requested:** Show "Revised to S4-260879r01"

**Implementation:**
- Extract "Revised to" information from Excel "TDoc Status" column
- Parse patterns like "Revised to S4-XXXXXX" or "Revised by S4-XXXXXX"
- Add new row "Revised to" with the target document number
- Make it clickable link to the revision

**Code location:** After status update (line 760)

---

## 3. Implement Dropdown for Status

**Current:** Plain text status field
**Requested:** Dropdown menu (like example doc)

**Challenge:** Google Docs API doesn't support dropdowns in tables
**Workaround:** 
- Add data validation to the cell (not perfect but closest)
- OR: Add comment with instructions
- OR: Use colored buttons/chips (visual only)

**Best Solution:** Add a note in documentation that dropdowns must be added manually via Google Docs UI

---

## 4. Agenda Item Override Configuration

**Current:** Agenda items come from Excel only
**Requested:** Configuration to override/move documents between agenda items

**Implementation:**
- Add configuration table: "Agenda Item Overrides"
- Format: `TDoc | Override Agenda Item`
- Example: `S4-260879 | 11.2`
- Apply overrides during document processing

---

## 5. Fix Summary Table Column Widths

**Current:** All columns equal width
**Requested:** Optimize column widths (TDoc column smaller)

**Implementation:**
- Set column widths after creating summary table
- TDoc: 80pt (just fits S4-XXXXXX)
- Title: 250pt (longest)
- Source: 100pt
- Agenda Item: 60pt

**Code location:** Line 521 (after createSummaryTable_)

---

## 6. Verify Incremental Update Logic

**Status:** ✅ VERIFIED SAFE (see INCREMENTAL_UPDATE_LOGIC.md)

**Summary:**
- Tables matched by TDOC number (not position)
- Email/revisions accumulated (never cleared)
- Decided statuses protected
- Minutes/Disposition never touched
- Safe for multiple updates during meeting

---

## Implementation Priority

1. **HIGH**: #1 Heading format (simple, 1 line change)
2. **HIGH**: #5 Summary table widths (simple, improves readability)
3. **MEDIUM**: #2 Revised to logic (moderate complexity)
4. **MEDIUM**: #4 Agenda override config (moderate complexity)
5. **LOW**: #3 Dropdown status (API limitation, document workaround)
6. **DONE**: #6 Incremental update verification (documented)

---

## Testing Plan

After implementation:
1. Test with sample Excel file
2. Verify heading format: "11.1 Description"
3. Verify "Revised to" appears when status is revised
4. Test agenda item overrides
5. Check summary table column widths
6. Run multiple updates to verify no data loss