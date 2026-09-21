/**
 * SA4-ARCH-004 — Title-consumer migration tests.
 *
 * setDocumentTitleFromTemplate_() is the first real production consumer
 * migrated to MeetingContext (Code.js ~5809-5845). This suite complements,
 * rather than replaces, the existing generateReportTitle_() tests in
 * tests/pure-logic.test.js: those test the pure title-formatting function in
 * isolation, this file tests the migrated CONSUMER around it --
 * i.e. that the production call chain
 * (runFullReportBuild -> buildSkeletonWithTdocTables -> setDocumentTitleFromTemplate_)
 * now gets its title inputs from getMeetingContext_() and still produces the
 * exact same document title as the pre-migration code would have.
 *
 * generateReportTitle_() itself is intentionally untouched (same
 * {REPORT_SUFFIX, TDOC_LIST_URL, MEETING_ID} contract as before), so "same
 * behavior" is proven here by comparing:
 *
 *   OLD call shape: generateReportTitle_(getReportConfig_())
 *     (this is exactly what pre-SA4-ARCH-004 setDocumentTitleFromTemplate_
 *      did -- it received the full getReportConfig_() result as `cfg` and
 *      passed it straight through)
 *
 *   NEW call path: setDocumentTitleFromTemplate_(templateDocId)
 *     (reads getMeetingContext_() internally, builds a small
 *      {REPORT_SUFFIX, TDOC_LIST_URL, MEETING_ID} shim from
 *      context.report.type / context.sources.tdocListUrl /
 *      context.meeting.portalId, and calls the SAME unchanged
 *      generateReportTitle_())
 *
 * for the same underlying configuration, using a minimal, single-purpose
 * DocumentApp.getActiveDocument() double that only supports what
 * setDocumentTitleFromTemplate_() actually calls (setName, and
 * getBody().insertParagraph(...).setHeading(...)). This is not a general
 * fake DocumentApp -- it exists only to observe the one side effect this
 * function has (setting the document title) and nothing else.
 *
 * Run: node tests/title-consumer.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log(`  ok   ${name}`);
  } else {
    failures++;
    console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`);
  }
}

/**
 * Minimal double for DocumentApp.getActiveDocument(), sized exactly to what
 * setDocumentTitleFromTemplate_() calls: doc.setName(...) and
 * doc.getBody().insertParagraph(0, ...).setHeading(...).
 */
function makeFakeActiveDocument() {
  const calls = { setName: null, insertParagraphText: null, insertParagraphIndex: null, heading: null };
  const doc = {
    setName: (name) => { calls.setName = name; },
    getBody: () => ({
      insertParagraph: (index, text) => {
        calls.insertParagraphIndex = index;
        calls.insertParagraphText = text;
        return { setHeading: (h) => { calls.heading = h; } };
      }
    })
  };
  return { doc, calls };
}

// ------------------------------------- per report type: old shape vs. new path

console.log('migrated setDocumentTitleFromTemplate_() vs. legacy generateReportTitle_(getReportConfig_()) call shape');

const REPORT_TYPES = ['6G', 'Audio', 'Video', 'MBS', 'RTC', 'Liaison', 'New'];

REPORT_TYPES.forEach(type => {
  const { sandbox } = loadCode({
    documentProperties: {
      REPORT_SUFFIX: type,
      MEETING_FOLDER: 'TSGS4_140_Example',
      MEETING_NUMBER: '140',
      MEETING_ID: '61111'
    }
  });

  // OLD call shape: pre-SA4-ARCH-004, setDocumentTitleFromTemplate_(id, cfg)
  // called generateReportTitle_(cfg) with the full getReportConfig_() result.
  const legacyCfg = sandbox.getReportConfig_();
  const legacyTitle = sandbox.generateReportTitle_(legacyCfg);

  // NEW path: the actual migrated production function.
  const { doc, calls } = makeFakeActiveDocument();
  sandbox.DocumentApp.getActiveDocument = () => doc;
  sandbox.setDocumentTitleFromTemplate_('someTemplateDocId');

  check(`${type}: migrated consumer's document title matches the legacy call shape's title`,
    calls.setName, legacyTitle);
  check(`${type}: migrated consumer inserts the title paragraph at index 0`,
    calls.insertParagraphIndex, 0);
  check(`${type}: inserted paragraph text matches the set document name`,
    calls.insertParagraphText, calls.setName);
  check(`${type}: inserted title paragraph uses the TITLE heading style`,
    calls.heading, sandbox.DocumentApp.ParagraphHeading.TITLE);
});

// -------------------------------------------------------- edge/fallback cases

console.log('edge/fallback cases (mirroring the existing generateReportTitle_ tests in pure-logic.test.js)');

{
  // TDOC_LIST_URL present but does not match the SA4%23NNN pattern -> falls
  // back to "SA4#{MEETING_ID}" (same fallback exercised in pure-logic.test.js).
  const { sandbox } = loadCode({
    documentProperties: {
      REPORT_SUFFIX: 'Video',
      MEETING_ID: '60777',
      TDOC_LIST_URL: 'https://example.com/no-pattern-here.xlsx'
    }
  });
  const legacyTitle = sandbox.generateReportTitle_(sandbox.getReportConfig_());
  const { doc, calls } = makeFakeActiveDocument();
  sandbox.DocumentApp.getActiveDocument = () => doc;
  sandbox.setDocumentTitleFromTemplate_('someTemplateDocId');

  check('TDOC_LIST_URL without SA4%23 pattern -> migrated consumer falls back to SA4#{MEETING_ID}, matching legacy',
    calls.setName, legacyTitle);
  check('  -> concretely, the title is "Video SWG Minutes SA4#60777"',
    calls.setName, 'Video SWG Minutes SA4#60777');
}

{
  // Unrecognized report type -> used verbatim as the topic name (same
  // fallback exercised in pure-logic.test.js).
  const { sandbox } = loadCode({
    documentProperties: { REPORT_SUFFIX: 'SomeUnknownType', MEETING_ID: '99999' }
  });
  const legacyTitle = sandbox.generateReportTitle_(sandbox.getReportConfig_());
  const { doc, calls } = makeFakeActiveDocument();
  sandbox.DocumentApp.getActiveDocument = () => doc;
  sandbox.setDocumentTitleFromTemplate_('someTemplateDocId');

  check('unrecognized report type -> migrated consumer matches legacy behavior',
    calls.setName, legacyTitle);
}

// ---------------------------------- prove MeetingContext is actually consumed

console.log('source-structure assertion: setDocumentTitleFromTemplate_ calls getMeetingContext_(), not getReportConfig_()');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const startMatch = source.match(/^function setDocumentTitleFromTemplate_\(/m);
  if (!startMatch) {
    failures++;
    console.log('  FAIL could not locate "function setDocumentTitleFromTemplate_(" in Code.js');
  } else {
    const startIndex = startMatch.index;
    const nextFnMatch = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnMatch.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnMatch.exec(source);
    const endIndex = next ? next.index : source.length;
    const body = source.slice(startIndex, endIndex);

    check('function body calls getMeetingContext_()',
      /\bgetMeetingContext_\s*\(\s*\)/.test(body), true);
    check('function body does NOT call getReportConfig_() directly',
      /\bgetReportConfig_\s*\(\s*\)/.test(body), false);
    check('function signature no longer takes a `cfg` parameter',
      /^function setDocumentTitleFromTemplate_\(\s*sourceDocId\s*\)/.test(body), true);
  }
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll title-consumer migration tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
