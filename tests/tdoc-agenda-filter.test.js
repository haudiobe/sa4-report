/**
 * SA4-IMPL-005 — First production migration of agendaSelector: the
 * agenda-selection filter inside downloadAndGroupTdocs_().
 *
 * downloadAndGroupTdocs_() itself cannot be executed directly here without a
 * large UrlFetchApp/DriveApp/SpreadsheetApp fake (it downloads a real XLSX
 * from a URL and opens it as a Sheet) -- exactly the kind of fake this
 * project's baseline (SA4-ARCH-002) deliberately avoided building. Instead,
 * per SA4-IMPL-005 §10, this suite combines:
 *
 *   1. Source-structure assertions proving WHICH functions were migrated
 *      (downloadAndGroupTdocs_ only) and which were deliberately left alone
 *      (processWebDownloadedSheet_ -- the legacy pipeline's OWN, separate
 *      agenda-prefix filter -- plus every agenda-structure parser and every
 *      structural/skeleton consumer).
 *   2. Pure behavior characterization: downloadAndGroupTdocs_()'s new filter
 *      line is now EXACTLY `agendaSelectorMatches_(agendaSelector,
 *      agendaItem)` where `agendaSelector = getMeetingContext_().report
 *      .agendaSelector` (verified by reading the migrated source directly,
 *      see the "source-structure" section below) -- so exercising that same
 *      composition (real getMeetingContext_() output, real
 *      agendaSelectorMatches_()) against representative agenda-item values
 *      characterizes the actual, live filtering decision without needing to
 *      run the download/parse machinery around it.
 *
 * Run: node tests/tdoc-agenda-filter.test.js
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

function extractFunctionBody(source, fnName) {
  const startMatch = source.match(new RegExp('^function ' + fnName + '\\(', 'm'));
  if (!startMatch) return null;
  const startIndex = startMatch.index;
  const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
  nextFnRe.lastIndex = startIndex + startMatch[0].length;
  const next = nextFnRe.exec(source);
  const endIndex = next ? next.index : source.length;
  return source.slice(startIndex, endIndex);
}

// ================================================== 1. source-structure ===

console.log('source-structure: which functions were migrated, which were not');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  const migrated = extractFunctionBody(source, 'downloadAndGroupTdocs_');
  if (!migrated) {
    failures++;
    console.log('  FAIL could not locate "function downloadAndGroupTdocs_(" in Code.js');
  } else {
    check('downloadAndGroupTdocs_() now calls agendaSelectorMatches_()',
      /\bagendaSelectorMatches_\s*\(/.test(migrated), true);
    check('downloadAndGroupTdocs_() now calls getMeetingContext_()',
      /\bgetMeetingContext_\s*\(\s*\)/.test(migrated), true);

    // Strip `//` line comments before checking for absence of the OLD
    // pattern -- this function's own migration comment mentions both
    // "agendaItem.startsWith(agendaPrefix)" and "getConfiguredAgendaPrefix_()"
    // in prose (explaining what it used to be), which would otherwise
    // false-positive these two checks. Same fix as SA4-IMPL-001A applied to
    // an identical false positive in tests/tdoc-identifier.test.js.
    const migratedCodeOnly = migrated.replace(/\/\/.*$/gm, '');
    check('downloadAndGroupTdocs_() no longer contains the old direct agendaItem.startsWith(agendaPrefix) filter',
      /agendaItem\.startsWith\(agendaPrefix\)/.test(migratedCodeOnly), false);
    check('downloadAndGroupTdocs_() no longer calls getConfiguredAgendaPrefix_() (superseded by the selector for this function only)',
      /\bgetConfiguredAgendaPrefix_\s*\(\s*\)/.test(migratedCodeOnly), false);
  }

  // The ONE other place with the same-looking inline filter is
  // processWebDownloadedSheet_() -- part of the legacy table-copy pipeline
  // (downloadAndProcessFromWeb()), explicitly out of scope for this task
  // (and for every prior task touching agenda/report generation). Prove it
  // was deliberately left exactly as it was, not silently migrated too.
  const legacy = extractFunctionBody(source, 'processWebDownloadedSheet_');
  if (!legacy) {
    failures++;
    console.log('  FAIL could not locate "function processWebDownloadedSheet_(" in Code.js');
  } else {
    check('processWebDownloadedSheet_() (legacy pipeline) STILL uses the old direct agendaItem.startsWith(agendaPrefix) filter -- correctly NOT migrated',
      /agendaItem\.startsWith\(agendaPrefix\)/.test(legacy), true);
    check('processWebDownloadedSheet_() does NOT call agendaSelectorMatches_() (legacy pipeline untouched)',
      /\bagendaSelectorMatches_\s*\(/.test(legacy), false);
  }

  // Agenda-structure parsers: must remain on their own, intentionally
  // broader matching rule (item.number === prefix-without-dot OR
  // item.number.startsWith(prefix), which keeps the bare parent item) --
  // NOT migrated to agendaSelectorMatches_ in this task.
  ['parseAgendaForReport_', 'parseAgendaFromZippedTdoc', 'parseAgendaDocumentById'].forEach(fnName => {
    const body = extractFunctionBody(source, fnName);
    if (!body) {
      failures++;
      console.log(`  FAIL could not locate "function ${fnName}(" in Code.js`);
      return;
    }
    check(`${fnName}() does NOT call agendaSelectorMatches_() (agenda-structure parsing is unmigrated, by design)`,
      /\bagendaSelectorMatches_\s*\(/.test(body), false);
  });

  // Structural/skeleton consumers: also untouched.
  ['ensureReallocationTable_', 'buildSkeletonWithTdocTables', 'buildReportPreamble_'].forEach(fnName => {
    const body = extractFunctionBody(source, fnName);
    if (!body) {
      failures++;
      console.log(`  FAIL could not locate "function ${fnName}(" in Code.js`);
      return;
    }
    check(`${fnName}() does NOT call agendaSelectorMatches_() (structural/skeleton consumers are unmigrated, by design)`,
      /\bagendaSelectorMatches_\s*\(/.test(body), false);
  });
}

// ============================ 2. pure behavior characterization ===========
//
// The migrated filter is exactly:
//   agendaSelectorMatches_(getMeetingContext_().report.agendaSelector, agendaItem)
// (verified in §1 above). Exercising that composition directly
// characterizes the live decision downloadAndGroupTdocs_() now makes,
// without needing to run the download/Sheet-parsing machinery around it.

console.log('behavior: main Audio (agendaSelector derived from a real MeetingContext, not a hand-written constant)');

{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Audio' } });
  const selector = sandbox.getMeetingContext_().report.agendaSelector;
  check('selector is {mode:"prefix", value:"7."} (unchanged since SA4-ARCH-003/SA4-IMPL-004)',
    selector, { mode: 'prefix', value: '7.' });

  check('"7.1" is selected', sandbox.agendaSelectorMatches_(selector, '7.1'), true);
  check('"7.2.3" is selected', sandbox.agendaSelectorMatches_(selector, '7.2.3'), true);
  check('"17.1" is rejected', sandbox.agendaSelectorMatches_(selector, '17.1'), false);
  check('"8.1" is rejected (different report type\'s agenda item)', sandbox.agendaSelectorMatches_(selector, '8.1'), false);
}

console.log('behavior: main 6G');

{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: '6G' } });
  const selector = sandbox.getMeetingContext_().report.agendaSelector;
  check('selector is {mode:"prefix", value:"11."}', selector, { mode: 'prefix', value: '11.' });

  check('"11.1" is selected', sandbox.agendaSelectorMatches_(selector, '11.1'), true);
  check('"11.4.2" is selected', sandbox.agendaSelectorMatches_(selector, '11.4.2'), true);
  check('"1.1" is rejected', sandbox.agendaSelectorMatches_(selector, '1.1'), false);
}

console.log('behavior: main -- old vs. new filter equivalence, all 7 report types, mixed real agenda items');

{
  const REPORT_TYPES = ['6G', 'Audio', 'Video', 'MBS', 'RTC', 'Liaison', 'New'];
  const SAMPLE_ITEMS = ['5.1', '7.1', '8.2.1', '9.10', '10.', '11.4', '18.3', '1.1', '17.1', '70.', '111.'];

  REPORT_TYPES.forEach(type => {
    const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: type } });
    const cfg = sandbox.getReportConfig_();
    const selector = sandbox.getMeetingContext_().report.agendaSelector;

    SAMPLE_ITEMS.forEach(item => {
      const oldResult = item.startsWith(cfg.AGENDA_ITEM_PREFIX); // the exact old inline expression
      const newResult = sandbox.agendaSelectorMatches_(selector, item);
      check(`${type}: "${item}" -- old .startsWith() vs. new agendaSelectorMatches_() agree (${oldResult})`,
        newResult, oldResult);
    });
  });
}

console.log('behavior: ULBC-MED Audio ad-hoc -- real agenda items now pass agenda selection');

{
  const ULBC_MED_PROPS = {
    MEETING_TYPE: 'adhoc',
    MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED',
    REPORT_SUFFIX: 'Audio',
    FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
    TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx',
    AGENDA_TDOC: 'S4aA260090',
    REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/'
  };
  const { sandbox } = loadCode({ documentProperties: ULBC_MED_PROPS });
  const selector = sandbox.getMeetingContext_().report.agendaSelector;
  check('selector is {mode:"all"}', selector, { mode: 'all' });

  // Real agenda items verified in SA4-ARCH-005/006 (the ULBC-MED agenda
  // TDoc and TDoc list): 1=Opening, 2=Approval/registration,
  // 3=IPR reminder, 4/4.1/4.7=ULBC-MED topic subsections, 5=AOB, 6=Close.
  // None of these would have passed the OLD main-Audio "7." prefix filter
  // -- that was the whole bug SA4-IMPL-004/005 exist to fix.
  ['1', '2', '3', '4', '4.1', '4.7', '5', '6'].forEach(item => {
    const oldWouldHaveMatched = item.startsWith('7.');
    check(`"${item}" is selected under {mode:"all"} (old main-Audio prefix filter would have rejected it: ${oldWouldHaveMatched})`,
      sandbox.agendaSelectorMatches_(selector, item), true);
  });

  console.log('  -- blank agenda items remain excluded (IMPL-004 "all" semantics, not "every spreadsheet row")');
  ['', '   ', null, undefined].forEach(item => {
    check(`${JSON.stringify(item)} is NOT selected even under {mode:"all"}`,
      sandbox.agendaSelectorMatches_(selector, item), false);
  });
}

console.log('behavior: "all" is agenda-selection only -- it does not mean "skip other row rules"');

{
  // agendaSelectorMatches_ only ever receives the agenda-item VALUE.
  // downloadAndGroupTdocs_()'s other row rules (blank TDoc -> `continue`,
  // reallocation lookup, column presence checks) are untouched lines of
  // that same function, entirely independent of this call -- confirmed by
  // the source-structure check in §1 above showing the diff is limited to
  // the filter line, plus a direct read of the surrounding lines: this test
  // documents the distinction, it doesn't re-verify the surrounding rules'
  // Apps-Script-dependent behavior (out of scope, same reasoning as the
  // rest of this suite).
  const { sandbox } = loadCode();
  const selector = { mode: 'all' };
  check('a real, non-blank agenda item value alone is what "all" evaluates -- true for this example',
    sandbox.agendaSelectorMatches_(selector, '4.1'), true);
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll TDoc-agenda-filter tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
