# Bugs and Issues Tracker

## Overview
This document tracks known bugs, issues, and improvement opportunities in the Google Apps Script project.

**Last Updated:** 2026-07-23

## Status Legend
- 🔴 **Critical** - Blocking functionality, needs immediate attention
- 🟡 **High** - Important issue, should be fixed soon
- 🟢 **Medium** - Should be addressed but not urgent
- 🔵 **Low** - Nice to have, can be deferred
- ✅ **Fixed** - Issue has been resolved
- 🔄 **In Progress** - Currently being worked on

---

## Known Issues

### Issue Template
Use this template when documenting new issues:

```markdown
### [Priority] Issue Title

**Status:** 🔴/🟡/🟢/🔵/✅/🔄

**Description:**
Brief description of the issue

**Steps to Reproduce:**
1. Step 1
2. Step 2
3. Step 3

**Expected Behavior:**
What should happen

**Actual Behavior:**
What actually happens

**Impact:**
Who/what is affected

**Possible Solution:**
Ideas for fixing (if any)

**Related Files:**
- File1.gs
- File2.gs

**Notes:**
Additional context
```

---

## Critical Issues Found

### 🔴 Issue #1: Duplicate Function Definition

**Status:** 🔴 Critical

**Description:**
The function `removeRowHeightAndSpacing()` is defined twice in Code.js (lines 1116-1139 and 1217-1286). The second definition overwrites the first, causing the simpler version to never execute.

**Steps to Reproduce:**
1. Open Code.js
2. Search for "function removeRowHeightAndSpacing"
3. Observe two definitions

**Expected Behavior:**
Only one definition should exist

**Actual Behavior:**
Two definitions exist, second one overwrites the first

**Impact:**
- Code confusion and maintenance issues
- First implementation is dead code
- Potential for unexpected behavior

**Possible Solution:**
Remove lines 1116-1139 (the first, simpler definition)

**Related Files:**
- Code.js (lines 1116-1139, 1217-1286)

**Notes:**
The second definition is more complete and includes `setTwoColumnTDocTableWidths_()` call

---

### 🟡 Issue #2: Duplicate Code Block

**Status:** 🟡 High

**Description:**
In `copyTableToDoc_()` function, lines 401-417 are an exact duplicate of lines 381-399, including the rich text hyperlink application and column width setting.

**Steps to Reproduce:**
1. Open Code.js
2. Go to function `copyTableToDoc_()` (line 318)
3. Compare lines 381-399 with lines 401-417

**Expected Behavior:**
Code should only execute once

**Actual Behavior:**
Same code block runs twice for every new table

**Impact:**
- Performance degradation (unnecessary processing)
- Maintenance confusion
- Potential for bugs if one block is updated but not the other

**Possible Solution:**
Remove lines 401-417 (the duplicate block)

**Related Files:**
- Code.js (lines 381-417)

---

### 🟡 Issue #3: Disabled Revised Text Update

**Status:** 🟡 High

**Description:**
Critical revised text update logic is disabled with `if (0)` condition on line 428, preventing revision information from being displayed in row 10.

**Steps to Reproduce:**
1. Open Code.js
2. Go to line 428 in `updateExistingTable()` function
3. Observe `if (0)` condition

**Expected Behavior:**
Revised text should be updated when present

**Actual Behavior:**
Code never executes due to `if (0)` condition

**Impact:**
- Users don't see revision information
- Feature is completely disabled
- Data loss for revised documents

**Possible Solution:**
Either:
1. Remove the disabled code if no longer needed
2. Re-enable with: `if (revisedText && existingTable.getNumRows() >= 10 && existingTable.getRow(9).getNumCells() >= 2)`

**Related Files:**
- Code.js (lines 428-443)

---

## Medium Priority Issues

### 🟢 Issue #4: Hardcoded Meeting Number

**Status:** 🟢 Medium

**Description:**
HyperLink.js has hardcoded pattern for meeting 136 (S4-26xxxx), which will fail for future meetings.

**Steps to Reproduce:**
1. Open HyperLink.js
2. Check line 5: `const pattern = /^S4-26\d{4}$/;`

**Expected Behavior:**
Pattern should work for all meetings

**Actual Behavior:**
Only works for S4-26xxxx documents

**Impact:**
- Will fail for meetings 137+ (S4-27xxxx, S4-28xxxx, etc.)
- Requires code update for each meeting

**Possible Solution:**
Change to: `const pattern = /^S4-\d{6}$/;`

**Related Files:**
- HyperLink.js (line 5)

---

### 🟢 Issue #5: Inconsistent Configuration Systems

**Status:** 🟢 Medium

**Description:**
Three separate configuration functions exist with different approaches:
- `getReportConfig_()` uses PropertiesService
- `getCollectorConfig_()` reads from document table
- `getConfig_()` reads from document table

**Impact:**
- Confusion about which config to use
- Potential for configuration conflicts
- Maintenance complexity

**Possible Solution:**
Consolidate into a single configuration system

**Related Files:**
- Code.js (lines 52, 63, 1399)

---

### 🟢 Issue #6: Missing Error Handling in safeFetch_()

**Status:** 🟢 Medium

**Description:**
The `safeFetch_()` function doesn't handle rate limiting (429), implement retry logic, or handle timeouts.

**Impact:**
- May fail on rate limits
- No automatic recovery from transient failures
- Could hit quota limits

**Possible Solution:**
Add retry logic with exponential backoff

**Related Files:**
- Code.js (lines 789-803)

---

### 🟢 Issue #7: Potential Race Condition

**Status:** 🟢 Medium

**Description:**
Multiple tables are processed in forEach without proper synchronization. If script times out mid-execution, some tables may be updated while others aren't.

**Impact:**
- Inconsistent state if execution is interrupted
- No transaction-like behavior
- Difficult to resume after failure

**Possible Solution:**
Add progress tracking and resume capability

**Related Files:**
- Code.js (lines 533-592)

---

## Low Priority Issues

### 🔵 Issue #8: Inconsistent String Comparison

**Status:** 🔵 Low

**Description:**
Mix of case-sensitive and case-insensitive string comparisons throughout the code.

**Examples:**
- Line 109: `getText() === 'Key'` (case-sensitive)
- Line 129: `.toLowerCase() === 'true'` (case-insensitive)
- Line 339: `getText().trim() !== 'TDoc'` (case-sensitive)

**Impact:**
- Potential for bugs if user input varies in case
- Inconsistent behavior

**Possible Solution:**
Standardize on case-insensitive comparisons for user-facing strings

**Related Files:**
- Code.js (throughout)

---

### 🔵 Issue #9: Magic Numbers

**Status:** 🔵 Low

**Description:**
Hardcoded numbers without explanation throughout the code.

**Examples:**
- Line 69: `LIMIT=2000` in RSS URL
- Line 70: `LIMIT=200` in RSS URL
- Line 73: `'14'` days back
- Line 74: `'24'` hours cache TTL

**Impact:**
- Difficult to understand why specific values were chosen
- Hard to adjust without understanding context

**Possible Solution:**
Move to configuration with comments explaining the values

**Related Files:**
- Code.js (various locations)

---

### 🔵 Issue #10: Incomplete JSDoc Comments

**Status:** 🔵 Low

**Description:**
Most functions lack documentation, parameter descriptions, and return type information.

**Impact:**
- Difficult for new developers to understand
- No IDE autocomplete help
- Hard to maintain

**Possible Solution:**
Add JSDoc comments to all functions

**Related Files:**
- Code.js (throughout)
- HyperLink.js (throughout)

---

## Code Quality Issues

### Performance Issues
- [x] Identified slow operations (duplicate code execution)
- [x] Found unnecessary API calls (none identified)
- [x] Identified inefficient loops (regex compilation in loops)
- [x] Reviewed quota usage (needs rate limiting)

### Security Issues
- [x] No exposed credentials found
- [x] Authorization uses Google's built-in system
- [ ] Input validation needed for URLs
- [ ] SSRF protection needed

### Functionality Issues
- [x] Identified disabled features (revised text update)
- [x] Found duplicate code execution
- [ ] Need to test all main functions
- [ ] Need to test error scenarios

### Documentation Issues
- [x] Missing function documentation identified
- [x] Some unclear variable names found
- [x] Lack of usage examples
- [x] Missing inline comments for complex logic

---

## Testing Checklist

### Manual Testing
- [ ] Test with valid inputs
- [ ] Test with invalid inputs
- [ ] Test with edge cases (empty, null, undefined)
- [ ] Test with large datasets
- [ ] Test error handling
- [ ] Test with different user permissions

### Integration Testing
- [ ] Test Google Docs integration
- [ ] Test Google Sheets integration
- [ ] Test external API calls (if any)
- [ ] Test triggers and time-based functions

### Performance Testing
- [ ] Measure execution time
- [ ] Check quota usage
- [ ] Test with concurrent users (if applicable)
- [ ] Monitor memory usage

---

## Improvement Opportunities

### Code Improvements
- [ ] Add TypeScript definitions for better IDE support
- [ ] Implement proper logging system
- [ ] Add configuration management
- [ ] Create reusable utility functions

### Feature Enhancements
- [ ] Add data validation
- [ ] Implement batch processing
- [ ] Add progress indicators
- [ ] Create user-friendly error messages

### Documentation Improvements
- [ ] Add JSDoc comments
- [ ] Create usage examples
- [ ] Document API endpoints (if any)
- [ ] Add troubleshooting guide

---

## Fixed Issues

### Example Fixed Issue
**Status:** ✅ Fixed on 2026-07-23

**Description:**
Example of a fixed issue

**Solution:**
How it was fixed

**Files Changed:**
- Code.gs

---

## Notes

### Common Patterns to Watch For

1. **Quota Limits:**
   - Apps Script has daily quotas
   - Watch for operations that might hit limits
   - Implement retry logic for quota errors

2. **Execution Time Limits:**
   - Scripts have 6-minute execution limit
   - Break long operations into smaller chunks
   - Use time-based triggers for long processes

3. **Data Size Limits:**
   - Be aware of size limits for documents
   - Implement pagination for large datasets
   - Consider using external storage for large files

4. **Error Handling:**
   - Always wrap API calls in try-catch
   - Log errors for debugging
   - Provide meaningful error messages to users

### Testing Tips

1. **Use Logger.log():**
   ```javascript
   Logger.log('Debug info: ' + variable);
   ```

2. **Check Execution Logs:**
   ```bash
   clasp logs
   ```

3. **Test in Stages:**
   - Test individual functions first
   - Then test integration
   - Finally test end-to-end workflows

4. **Create Test Data:**
   - Use sample data for testing
   - Don't test with production data
   - Create edge case test scenarios

---

## Action Items

### Immediate Actions (After Cloning)
1. [ ] Clone the script: `clasp clone SCRIPT_ID`
2. [ ] Review all `.gs` files
3. [ ] Document all functions
4. [ ] Identify critical bugs
5. [ ] Create test cases

### Short-term Actions (This Week)
1. [ ] Fix critical bugs
2. [ ] Add error handling
3. [ ] Improve documentation
4. [ ] Create backup procedures

### Long-term Actions (This Month)
1. [ ] Refactor code for maintainability
2. [ ] Add comprehensive tests
3. [ ] Optimize performance
4. [ ] Create deployment procedures

---

## Resources

- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Apps Script Quotas](https://developers.google.com/apps-script/guides/services/quotas)
- [Apps Script Troubleshooting](https://developers.google.com/apps-script/guides/support/troubleshooting)

---

## Contact

For questions or to report new issues:
- Review this document
- Check the Apps Script documentation
- Search Stack Overflow
- Consult with team members