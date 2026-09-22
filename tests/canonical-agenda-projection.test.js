/**
 * SA4-IMPL-007 — Migrate canonical parseAgendaForReport_() agenda filtering
 * (ZIP/AGENDA_TDOC path only) from the old inline "parent-or-prefix" filter
 * to projectAgendaItems_(agendaStructure, context.report.agendaSelector).
 *
 * parseAgendaForReport_() cannot be executed directly here without a large
 * UrlFetchApp/Utilities.unzip fake (it downloads a real ZIP over HTTP and
 * unzips a .docx from it) -- exactly the kind of fake this project's
 * baseline (SA4-ARCH-002) deliberately avoided building. Instead, per the
 * pattern established by tests/tdoc-agenda-filter.test.js (SA4-IMPL-005),
 * this suite combines:
 *
 *   1. Source-structure assertions proving WHICH function was migrated
 *      (parseAgendaForReport_ only) and which were deliberately left alone
 *      (the Google-Doc fallback branch inside the SAME function, every
 *      standalone/dialog parser, buildSkeletonWithTdocTables(), etc).
 *   2. Pure behavior characterization: the migrated ZIP-path line is now
 *      EXACTLY `projectAgendaItems_(agendaStructure, context.report
 *      .agendaSelector)` where `context = getMeetingContext_()` -- so
 *      exercising that same composition (real getMeetingContext_() output,
 *      real projectAgendaItems_()) against representative agenda
 *      structures characterizes the actual, live filtering decision
 *      without needing to run the download/unzip/parse machinery around
 *      it.
 *   3. Observable-equivalence checks: SA4-ARCH-007 established that
 *      buildSkeletonWithTdocTables() immediately re-filters out the bare
 *      parent item before doing anything else with the array, so this
 *      suite compares the OLD and NEW pipelines' output AFTER that same
 *      parent-removal step, not the raw parseAgendaForReport_() return
 *      value.
 *
 * Run: node tests/canonical-agenda-projection.test.js
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

function item(number, extra) {
  return Object.assign({ number: number, title: 'Title for ' + number, level: 1, heading: 'H2', text: '' }, extra || {});
}
function numbers(items) {
  return items.map(i => i.number);
}

// OLD ZIP-path filter, verbatim, for reference/equivalence comparison only.
function oldZipPathFilter(agendaStructure, prefix) {
  return agendaStructure.filter(item =>
    item.number === prefix.replace(/\.$/, '') ||
    item.number.startsWith(prefix)
  );
}

// The same parent-removal step buildSkeletonWithTdocTables() performs
// immediately after calling parseAgendaForReport_(), verbatim.
function skeletonParentRemoval(agendaItems, agendaPrefix) {
  return agendaItems.filter(item => item.number !== agendaPrefix.replace(/\.$/, ''));
}

// ================================================== 1. source-structure ===

console.log('source-structure: parseAgendaForReport_() migration, and what was deliberately left alone');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  const migrated = extractFunctionBody(source, 'parseAgendaForReport_');
  if (!migrated) {
    failures++;
    console.log('  FAIL could not locate "function parseAgendaForReport_(" in Code.js');
  } else {
    check('parseAgendaForReport_() now calls projectAgendaItems_()',
      /\bprojectAgendaItems_\s*\(/.test(migrated), true);
    check('parseAgendaForReport_() now calls getMeetingContext_()',
      /\bgetMeetingContext_\s*\(\s*\)/.test(migrated), true);
    check('parseAgendaForReport_() no longer contains the old inline parent-or-prefix filter (item.number === prefix...)',
      /item\.number\s*===\s*prefix/.test(migrated), false);
    check('the Google-Doc fallback branch still calls parseAgendaStructureWithText_() with a prefix (unmigrated, by design)',
      /parseAgendaStructureWithText_\s*\(\s*agendaDoc\.getBody\(\)\s*,\s*prefix\s*\)/.test(migrated), true);
  }

  const skeleton = extractFunctionBody(source, 'buildSkeletonWithTdocTables');
  if (!skeleton) {
    failures++;
    console.log('  FAIL could not locate "function buildSkeletonWithTdocTables(" in Code.js');
  } else {
    check('buildSkeletonWithTdocTables() still contains its own parent-removal filter (item.number !== agendaPrefix...)',
      /item\.number\s*!==\s*agendaPrefix\.replace/.test(skeleton), true);
    check('buildSkeletonWithTdocTables() does not itself call projectAgendaItems_() (migration is inside parseAgendaForReport_ only)',
      /\bprojectAgendaItems_\s*\(/.test(skeleton), false);
  }

  ['parseAgendaFromZippedTdoc', 'parseAgendaDocumentById', 'parseAgendaStructureWithText_', 'parseAgendaFromHeadings_', 'parseAgendaFromTables_'].forEach(fnName => {
    const body = extractFunctionBody(source, fnName);
    if (!body) {
      failures++;
      console.log(`  FAIL could not locate "function ${fnName}(" in Code.js`);
      return;
    }
    check(`${fnName}() does not call projectAgendaItems_() (standalone/dialog parsers untouched)`,
      /\bprojectAgendaItems_\s*\(/.test(body), false);
  });

  // Characterize (freeze), not fix: the known parseAgendaFromHeadings_() bug
  // is unchanged by this task. This suite does not attempt to reproduce or
  // assert the bug itself (that's SA4-ARCH-007's job) -- only that this
  // function's source was not touched by the IMPL-007 migration, which is
  // covered by the "does not call projectAgendaItems_()" check above.
  check('parseAgendaFromHeadings_() exists (still present, still unmigrated, still unfixed)',
    typeof extractFunctionBody(source, 'parseAgendaFromHeadings_'), 'string');
}

// ============================================ 2. main Audio equivalence ===

console.log('canonical ZIP path -- main Audio ("7.") old vs. new, AFTER skeleton parent-removal');

{
  const numbersList = ['1', '5', '6', '7', '7.1', '7.1.1', '7.2', '8', '17'];
  const source = numbersList.map(n => item(n));

  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Audio' } });
  const ctx = sandbox.getMeetingContext_();
  check('main Audio agendaSelector is {mode:"prefix", value:"7."}', ctx.report.agendaSelector, { mode: 'prefix', value: '7.' });

  const oldRaw = oldZipPathFilter(source, '7.');
  const newRaw = sandbox.projectAgendaItems_(source, ctx.report.agendaSelector);

  check('raw (pre-skeleton-filter) output DOES differ: old retains bare "7", new does not',
    numbers(oldRaw).indexOf('7') !== -1 && numbers(newRaw).indexOf('7') === -1, true);

  const oldEffective = skeletonParentRemoval(oldRaw, '7.');
  const newEffective = skeletonParentRemoval(newRaw, '7.');

  check('effective (post-skeleton-parent-removal) output is IDENTICAL old vs. new',
    numbers(newEffective), numbers(oldEffective));
  check('effective output is exactly the expected sub-item list',
    numbers(newEffective), ['7.1', '7.1.1', '7.2']);
}

// =============================================== 3. main 6G equivalence ===

console.log('canonical ZIP path -- main 6G ("11.") old vs. new, AFTER skeleton parent-removal');

{
  const numbersList = ['1', '10', '11', '11.1', '11.1.1', '11.2', '12', '111'];
  const source = numbersList.map(n => item(n));

  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: '6G' } });
  const ctx = sandbox.getMeetingContext_();
  check('main 6G agendaSelector is {mode:"prefix", value:"11."}', ctx.report.agendaSelector, { mode: 'prefix', value: '11.' });

  const oldRaw = oldZipPathFilter(source, '11.');
  const newRaw = sandbox.projectAgendaItems_(source, ctx.report.agendaSelector);

  const oldEffective = skeletonParentRemoval(oldRaw, '11.');
  const newEffective = skeletonParentRemoval(newRaw, '11.');

  check('effective (post-skeleton-parent-removal) output is IDENTICAL old vs. new',
    numbers(newEffective), numbers(oldEffective));
  check('effective output is exactly the expected sub-item list ("111" correctly excluded)',
    numbers(newEffective), ['11.1', '11.1.1', '11.2']);
}

// ============================================= 4. ULBC-MED canonical all ==

console.log('canonical ZIP path -- ULBC-MED ({mode:"all"}) -- main functional gain of IMPL-007');

{
  const numbersList = ['1', '2', '3', '4', '4.1', '4.2', '4.3', '4.4', '4.5', '4.6', '4.7', '4.8', '4.9', '4.10', '5', '6'];
  const source = numbersList.map(n => item(n));

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
  const ctx = sandbox.getMeetingContext_();
  check('ULBC-MED agendaSelector is {mode:"all"}', ctx.report.agendaSelector, { mode: 'all' });

  const result = sandbox.projectAgendaItems_(source, ctx.report.agendaSelector);

  check('every real agenda item survives canonical ZIP-path projection, in source order',
    numbers(result), numbersList);
  check('nothing is dropped (old Audio "7." assumption would have retained NONE of these)',
    result.length, source.length);
}

// ==================================================== 5. itemList result ==

console.log('canonical projection boundary -- itemList {value:["1.5","1.6"]} retains exactly those two items');

{
  const source = ['1', '1.1', '1.5', '1.5.1', '1.5.2', '1.6', '1.6.1', '1.7', '2'].map(n => item(n));

  const { sandbox } = loadCode();
  const selector = sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: ['1.5', '1.6'] });
  const result = sandbox.projectAgendaItems_(source, selector);

  check('retains exactly "1.5" and "1.6", no ancestors, no descendants',
    numbers(result), ['1.5', '1.6']);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
