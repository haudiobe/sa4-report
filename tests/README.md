# Testing Infrastructure

## Overview
This directory contains test files and testing documentation for the Google Apps Script project.

## Current automated tests (SA4-ARCH-002 baseline)

This section describes what's actually implemented and runnable today, as opposed to the
generic guidance further down. It was written as part of SA4-ARCH-002, which established a
regression baseline before introducing `MeetingContext` (SA4-ARCH-003+).

### How to run everything

All tests are plain Node scripts (no test framework/runner dependency). Each prints its own
`ok`/`FAIL` lines and exits non-zero if anything failed.

```bash
node tests/revision-order.test.js     # revision-chain ordering / Disposition text
node tests/no-duplicates.test.js      # structural integrity + architecture invariants
node tests/pure-logic.test.js         # pure/near-pure function regression tests
node tests/report-structure.test.js   # report-type -> structure characterization
```

Run them individually, or all four in sequence; there is no shared state between files, and no
network access, Google account, or Apps Script project is required.

### What each file covers

- **`revision-order.test.js`** — `orderTdocsByRevision_`, `orderRevisionChains_`, `getRevisedTo_`,
  `tdocNumberOf_`, `setDispositionRevisedTo_`. Pre-existing (added alongside Code.js 2.5.0).
- **`no-duplicates.test.js`** — restores the file referenced by the Code.js changelog but missing
  from the repo (see SA4-ARCH-001). Checks: no top-level `function name(...)` is declared twice in
  `Code.js`; every function referenced by an `onOpen()` menu item exists; the canonical entry points
  `runFullReportBuild`, `buildSkeletonWithTdocTables`, `continuousUpdate` exist; the configuration
  functions `getReportConfig_`, `getAgendaPrefixForReportType_`, `generateReportTitle_` exist;
  `orderTdocsByRevision_` exists. These are the invariants meant to catch accidental damage while
  `Code.js` is later split into modules.
- **`pure-logic.test.js`** — `normalizeStatus_`, `getAgendaPrefixForReportType_`, the report-type →
  mailing-list/drafts-folder mapping (observed through `getReportConfig_`, since the underlying
  `MAILING_LISTS`/`DRAFTS_FOLDERS` consts aren't reachable from outside the VM sandbox),
  `generateReportTitle_`, `stripReplyPrefixes_`, `parseEmailSubject_`, `extractDeadlineFromTitle_`,
  `parseExtendedDeadline_`, `normalizeTdoc_`, `extractTdocId_`, `computeShortNumber_`. Expected
  values were captured by running the current, unmodified `Code.js` and recording its real output —
  they describe CURRENT behavior, not what the "correct" behavior should be, and may need updating
  if that behavior is deliberately changed later.
- **`report-structure.test.js`** — diffs live `getReportConfig_()` / `getAgendaPrefixForReportType_()`
  / `generateReportTitle_()` output, per report type, against the golden snapshot in
  `tests/fixtures/main-meeting-profile.expected.json`.
- **`tests/fixtures/main-meeting-profile.expected.json`** — a regression snapshot (not a new
  configuration format) of the main-meeting/report-type knowledge currently embedded in `Code.js`:
  the 7 report types, their agenda prefixes/mailing lists/drafts folders, the 6G-vs-SWG-vs-other
  structural branch, default configuration values, and title-generation examples.

### What is NOT covered

- **No fake `DocumentApp`.** Nothing in this baseline builds, mutates, or reads a Google Doc — real
  or simulated. `buildSkeletonWithTdocTables()`, `continuousUpdate()`, `runFullReportBuild()`, and
  every "ensure/apply/copy to doc" helper are exercised only for *existence* (via
  `no-duplicates.test.js`), never for behavior.
- **The `is6G` / `isSWGReport` branch inside `buildSkeletonWithTdocTables()`** (which report types
  get the `11.0.x` subsections vs. the `X.1`/`X.2` subsections) is documented in the fixture as
  read from source, not exercised as running code — it isn't extracted into a separate, callable
  function, and extracting it purely for testability is out of scope for this baseline.
- **Live network dependencies are intentionally excluded.** The RSS/A1 e-mail collector
  (`checkRSSFeed_`, `collectA1_`, `collectHybridListservMessages_`), the TDoc-list/agenda downloads
  (`downloadAndGroupTdocs_`, `downloadMeetingAgenda_`), and the Reviewer API calls are not tested —
  they need live HTTP responses (or fixture captures of them) and, in several cases, wall-clock time
  (late-response greying, deadline countdowns), which this baseline does not attempt to mock.
- **`calculateTimeRemaining_`/`formatDeadline_`** are not tested for exact output because they read
  `new Date()` ("now") directly; testing them meaningfully would require mocking the clock, which
  this baseline treats as future work rather than doing cheaply here.
- **Google Docs rendering is not tested end-to-end.** There is no fixture "before" report and no
  fixture "after" report, and no way to compare rendered documents byte-for-byte or even
  structurally yet. Building that (real or fake `DocumentApp`, fixture TDoc-list/agenda inputs, a
  normalized structural snapshot of the output) is future work, not part of this baseline.
- The legacy/parallel report-building paths (`processWebDownloadedSheet_`/`copyTableToDoc_`,
  `createReportStructure_`/`autoCreateReportStructure`) are only checked for existence, same as the
  canonical paths — their behavior is untested.

### Why this baseline exists

SA4-ARCH-003 (and later tasks) will introduce a `getMeetingContext_()` compatibility wrapper around
the existing configuration system, then migrate call sites to use it. That work touches
`getReportConfig_()` and the report-type mapping logic directly. This baseline exists so that work
can be checked against `tests/pure-logic.test.js` and `tests/report-structure.test.js` after each
step: if either file starts failing, the wrapper is not yet behavior-preserving. It is deliberately
narrow ("cheap, deterministic protection first") rather than a comprehensive fake-Docs test harness —
see SA4-ARCH-002's report for the full rationale and its limitations.

## Testing Approach

Since Google Apps Script runs in Google's cloud environment, true local testing requires mocking Google services. This document outlines different testing strategies.

## Testing Strategies

### 1. Manual Testing (Recommended for Apps Script)

**Process:**
1. Make changes locally
2. Push to Google: `clasp push`
3. Open in browser: `clasp open`
4. Run functions manually in the Apps Script editor
5. Check logs: `clasp logs`

**Advantages:**
- No mocking required
- Tests real Google services
- Easy to debug

**Disadvantages:**
- Manual process
- Time-consuming
- Not automated

### 2. Unit Testing with Mocks (Advanced)

For local unit testing, you'll need to mock Google Apps Script services.

**Setup:**
```bash
npm install --save-dev jest
npm install --save-dev @types/google-apps-script
```

**Example Test File:** `tests/example.test.js`
```javascript
// Mock Google Apps Script services
global.Logger = {
  log: jest.fn()
};

global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(),
  openById: jest.fn()
};

// Import your functions (you'll need to export them)
// const { myFunction } = require('../Code.gs');

describe('Example Tests', () => {
  test('should log a message', () => {
    // Your test here
    expect(true).toBe(true);
  });
});
```

### 3. Integration Testing

Test the script with real Google services but in a test environment.

**Approach:**
1. Create a test Google Doc/Sheet
2. Create a test version of your script
3. Run tests against the test environment
4. Verify results

## Test Cases Template

### Test Case Template
```markdown
### Test Case: [Function Name]

**Purpose:** What this test verifies

**Prerequisites:**
- Required setup
- Test data needed

**Steps:**
1. Step 1
2. Step 2
3. Step 3

**Expected Result:**
What should happen

**Actual Result:**
What actually happened

**Status:** ✅ Pass / ❌ Fail

**Notes:**
Additional observations
```

## Test Scenarios

### Scenario 1: Basic Functionality
- [ ] Test with valid inputs
- [ ] Test with typical use cases
- [ ] Verify expected outputs

### Scenario 2: Edge Cases
- [ ] Test with empty inputs
- [ ] Test with null/undefined values
- [ ] Test with very large inputs
- [ ] Test with special characters

### Scenario 3: Error Handling
- [ ] Test with invalid inputs
- [ ] Test error messages
- [ ] Test recovery from errors
- [ ] Test timeout scenarios

### Scenario 4: Performance
- [ ] Measure execution time
- [ ] Test with large datasets
- [ ] Check quota usage
- [ ] Monitor memory usage

### Scenario 5: Integration
- [ ] Test Google Docs integration
- [ ] Test Google Sheets integration
- [ ] Test external API calls
- [ ] Test triggers

## Test Data

Create test data files in this directory:

```
tests/
├── README.md           # This file
├── test-data/
│   ├── sample-input.json
│   ├── expected-output.json
│   └── edge-cases.json
├── manual-tests/
│   └── test-checklist.md
└── automated-tests/
    └── example.test.js
```

## Running Tests

### Manual Tests
```bash
# 1. Push your code
clasp push

# 2. Open in browser
clasp open

# 3. Run functions in the Apps Script editor

# 4. Check logs
clasp logs
```

### Automated Tests (if set up)
```bash
npm test
```

## Test Results Log

### Test Run: [Date]

**Tester:** [Name]

**Environment:**
- Script Version: [version]
- Test Data: [description]

**Results:**
- Total Tests: X
- Passed: X
- Failed: X
- Skipped: X

**Issues Found:**
1. Issue 1
2. Issue 2

**Notes:**
Additional observations

## Best Practices

### 1. Test Early and Often
- Test after each significant change
- Don't accumulate untested changes

### 2. Document Test Results
- Keep a log of test runs
- Document any issues found
- Track fixes and retests

### 3. Use Realistic Test Data
- Test with data similar to production
- Include edge cases
- Test with various data sizes

### 4. Test Error Scenarios
- Don't just test the happy path
- Test what happens when things go wrong
- Verify error messages are helpful

### 5. Performance Testing
- Monitor execution time
- Check quota usage
- Test with realistic data volumes

## Common Testing Patterns

### Pattern 1: Logger Testing
```javascript
function testFunction() {
  Logger.log('Starting test...');
  
  try {
    // Your code here
    var result = someFunction();
    Logger.log('Result: ' + JSON.stringify(result));
    
    // Verify result
    if (result === expectedValue) {
      Logger.log('✅ Test passed');
    } else {
      Logger.log('❌ Test failed');
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
  }
}
```

### Pattern 2: Assertion Helper
```javascript
function assertEqual(actual, expected, message) {
  if (actual === expected) {
    Logger.log('✅ PASS: ' + message);
    return true;
  } else {
    Logger.log('❌ FAIL: ' + message);
    Logger.log('  Expected: ' + expected);
    Logger.log('  Actual: ' + actual);
    return false;
  }
}

function testMyFunction() {
  var result = myFunction(5);
  assertEqual(result, 10, 'myFunction should double the input');
}
```

### Pattern 3: Test Suite
```javascript
function runAllTests() {
  Logger.log('=== Starting Test Suite ===');
  
  var tests = [
    testFunction1,
    testFunction2,
    testFunction3
  ];
  
  var passed = 0;
  var failed = 0;
  
  tests.forEach(function(test) {
    try {
      test();
      passed++;
    } catch (error) {
      Logger.log('❌ Test failed: ' + error.message);
      failed++;
    }
  });
  
  Logger.log('=== Test Results ===');
  Logger.log('Passed: ' + passed);
  Logger.log('Failed: ' + failed);
}
```

## Debugging Tips

### 1. Use Logger.log() Extensively
```javascript
Logger.log('Variable value: ' + myVariable);
Logger.log('Object: ' + JSON.stringify(myObject));
```

### 2. Check Execution Logs
```bash
clasp logs
clasp logs --watch  # Real-time logs
```

### 3. Use Try-Catch Blocks
```javascript
try {
  // Your code
} catch (error) {
  Logger.log('Error: ' + error.message);
  Logger.log('Stack: ' + error.stack);
}
```

### 4. Test in Isolation
- Test individual functions first
- Then test integration
- Finally test end-to-end

## Resources

- [Apps Script Testing Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Apps Script Debugging](https://developers.google.com/apps-script/guides/support/troubleshooting)
- [Jest Documentation](https://jestjs.io/) (for local testing)

## Next Steps

1. [ ] Review existing code
2. [ ] Identify critical functions to test
3. [ ] Create test cases
4. [ ] Run manual tests
5. [ ] Document results
6. [ ] Fix any issues found
7. [ ] Retest after fixes