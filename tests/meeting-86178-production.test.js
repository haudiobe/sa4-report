/**
 * SA4-PROD-001 / SA4-PROD-002 / SA4-PROD-003 — Meeting 86178 (SA4-e (AH) on
 * FS_6G_MED, 2026-09-22) production-readiness regression coverage.
 *
 * Real evidence for this meeting was gathered directly from the public
 * 3GPP FTP tree on 2026-09-22 (not inferred from convention):
 *
 *   - TDoc list:  https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/
 *                 SA4_Plenary/Docs/TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx
 *   - Agenda TDoc: S4aP260098 ("Proposed agenda for SA4 AH conf. call on
 *                 FS_6G_MED (22nd September 2026)"), same Docs folder.
 *
 * tests/fixtures/meeting-86178-tdoc-list.json is a trimmed, real snapshot of
 * that TDoc list's rows (TDoc/Title/Source/Type/For/AgendaItem/
 * AgendaItemDesc/Status columns only), captured so this suite runs
 * deterministically without live network access. The real agenda structure
 * below (numbers 5, 5.1-5.11) was read directly out of S4aP260098's own
 * Word document body.
 *
 * SA4-PROD-002 addendum: a real Google Docs production run of SA4-PROD-001's
 * fix exposed further main-meeting assumptions in
 * buildSkeletonWithTdocTables()/ensureReallocationTable_()/
 * generateReportTitle_()/testAllConnections() that the is6G guard alone
 * didn't reach.
 *
 * SA4-PROD-003 correction: PROD-002's premise that an ad-hoc meeting never
 * wants the SWG X.1.1/X.1.2/X.1.3 subsection structure was WRONG (clarified
 * production requirement) -- meeting 86178 explicitly wants that structure
 * under its real X.1 item, PLUS a new {agendaPrefixNum}.1.4 "Documents"
 * bucket for TDocs whose original agenda assignment predates the normal
 * X.3+ sections (so they are never silently dropped). PROD-002's ad-hoc
 * early-return guards in the openingSection branch and
 * ensureReallocationTable_() are reverted here; the new registration
 * boundary/isBeforeRegistrationBoundary_()/appendTdocDetailTable_()
 * mechanism is characterized below (sections 7-9).
 *
 * This suite does NOT execute buildSkeletonWithTdocTables() itself -- doing
 * so requires a fake DocumentApp, which SA4-ARCH-002 explicitly scoped out
 * (see tests/report-structure.test.js's header comment) and this task does
 * not reopen. Instead, per that established precedent, this suite combines:
 *
 *   1. Source-structure assertions proving the SA4-PROD-001/002/003
 *      skeleton fixes exist exactly as intended, and that main-meeting-
 *      affecting code paths were not touched otherwise.
 *   2. Real-data pipeline characterization: MeetingContext -> agenda
 *      projection -> the (untouched) skeleton parent-removal filter ->
 *      TDoc-list grouping -> registration-boundary extraction, exercised
 *      against the real captured meeting 86178 data using the real,
 *      unmodified production functions (getMeetingContext_,
 *      projectAgendaItems_, agendaSelectorMatches_, parseSA4DocumentId_
 *      family, generateReportTitle_, isBeforeRegistrationBoundary_).
 *
 * Run: node tests/meeting-86178-production.test.js
 */

const fs = require('fs');
const path = require('path');
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

function item(number, title) {
  return { number, title, level: 1, heading: 'H2', text: '' };
}

// Strips // line comments and /* */ block comments so a source-structure
// regex checks only executable code, not an explanatory comment that
// legitimately names something (e.g. "S4aP260089", "getCollectorConfig_()")
// for context without actually using it.
function stripComments(src) {
  return src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
}

// =========================================================== 1. source ===

console.log('source-structure: buildSkeletonWithTdocTables() SA4-PROD-001 fixes');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const body = extractFunctionBody(source, 'buildSkeletonWithTdocTables');
  if (!body) {
    failures++;
    console.log('  FAIL could not locate "function buildSkeletonWithTdocTables(" in Code.js');
  } else {
    check('buildSkeletonWithTdocTables() now calls getMeetingContext_()',
      /\bgetMeetingContext_\s*\(\s*\)/.test(body), true);
    check('is6G now also requires meeting.type !== "adhoc" (not reportType alone)',
      /is6G\s*=\s*\(reportType === '6G'\)\s*&&\s*context\.meeting\.type\s*!==\s*'adhoc'/.test(body), true);
    check('AOB auto-insert guard now also requires !hasCloseOfSession',
      /if\s*\(!hasAOB\s*&&\s*!hasCloseOfSession\)/.test(body), true);
    // Everything else in this function (parent-removal filter, TDoc-table
    // schema, isSWGReport IPR gating, close-of-session handling) is
    // untouched -- only these two conditions changed.
    check('the pre-existing parent-removal filter is still present, unchanged',
      /item\.number\s*!==\s*agendaPrefix\.replace/.test(body), true);
    // SA4-PROD-003: PROD-002's "&& context.meeting.type !== 'adhoc'" guard
    // on the openingSection branch is reverted -- meeting 86178 explicitly
    // wants that branch's X.1.1/X.1.2 template-copy behavior.
    check('the openingSection branch is unconditional again (no meeting.type guard -- PROD-002\'s guard was reverted per the clarified PROD-003 requirement)',
      /if\s*\(item\.number === openingSection\)\s*\{/.test(stripComments(body)), true);
    check('no code-level special-case comparison against "S4aP260089" exists in buildSkeletonWithTdocTables() (comments MAY reference it for context; only an actual comparison would be a remapping)',
      /['"]S4aP260089['"]\s*[=!]==/.test(stripComments(body)), false);
    // SA4-PROD-003: the new registration-boundary extraction, gated to the
    // SWG-style branch only (!is6G), and the new X.1.4 "Documents" section
    // inside the openingSection branch.
    check('buildSkeletonWithTdocTables() now builds registrationDocs via isBeforeRegistrationBoundary_(), gated on !is6G',
      /if\s*\(!is6G\)\s*\{[\s\S]*?isBeforeRegistrationBoundary_\(/.test(body), true);
    check('buildSkeletonWithTdocTables() now creates a "{agendaPrefixNum}.1.4" Documents section when registrationDocs is non-empty',
      /documentsSection\s*=\s*`\$\{agendaPrefixNum\}\.1\.4`/.test(body), true);
    check('the new X.1.4 section renders each doc via appendTdocDetailTable_() with its OWN preserved original agenda-item label (not the loop item.number)',
      /appendTdocDetailTable_\(body,\s*tdocData,\s*tdocData\.row\[tdocData\.agendaCol\]\)/.test(body), true);
  }

  const reallocBody = extractFunctionBody(source, 'ensureReallocationTable_');
  if (!reallocBody) {
    failures++;
    console.log('  FAIL could not locate "function ensureReallocationTable_(" in Code.js');
  } else {
    // SA4-PROD-003: PROD-002's unconditional ad-hoc early return is
    // reverted -- meeting 86178 explicitly wants a real X.1.3 Document
    // Reallocations section, exactly like a main SWG meeting.
    check('ensureReallocationTable_() no longer has an ad-hoc early return (reverted to its pre-PROD-002, meeting-type-agnostic form)',
      /getMeetingContext_\(\)\.meeting\.type === 'adhoc'/.test(stripComments(reallocBody)), false);
    check('no code-level special-case comparison against "S4aP260089" exists in ensureReallocationTable_() (its own comment names the finding for context; that is not a remapping)',
      /['"]S4aP260089['"]\s*[=!]==/.test(stripComments(reallocBody)), false);
  }

  const boundaryBody = extractFunctionBody(source, 'isBeforeRegistrationBoundary_');
  if (!boundaryBody) {
    failures++;
    console.log('  FAIL could not locate "function isBeforeRegistrationBoundary_(" in Code.js');
  } else {
    check('isBeforeRegistrationBoundary_() does not hardcode the literal "5" (it is parameterized by agendaPrefixNum)',
      /['"]5['"]/.test(stripComments(boundaryBody)), false);
    check('isBeforeRegistrationBoundary_() does not reference any meeting ID or TDoc ID literal',
      /86178|S4aP260089|S4aP260098/.test(stripComments(boundaryBody)), false);
  }

  const appendBody = extractFunctionBody(source, 'appendTdocDetailTable_');
  if (!appendBody) {
    failures++;
    console.log('  FAIL could not locate "function appendTdocDetailTable_(" in Code.js');
  }

  check('no code-level special-case comparison against "S4aP260089" exists ANYWHERE in Code.js (the anomaly remains a manual verification item, not a code-level special case)',
    /['"]S4aP260089['"]\s*[=!]==/.test(stripComments(source)), false);

  const titleBody = extractFunctionBody(source, 'generateReportTitle_');
  if (!titleBody) {
    failures++;
    console.log('  FAIL could not locate "function generateReportTitle_(" in Code.js');
  } else {
    check('generateReportTitle_() now supports an optional cfg.meetingLabel, used before the SA4#-number formula',
      /if\s*\(cfg\.meetingLabel\)/.test(titleBody), true);
  }

  const connBody = extractFunctionBody(source, 'testAllConnections');
  if (!connBody) {
    failures++;
    console.log('  FAIL could not locate "function testAllConnections(" in Code.js');
  } else {
    check('testAllConnections() Revisions Folder test now reads getMeetingContext_().sources.revisionsUrl, not the main-meeting-only cfg.REVISIONS_URL formula',
      /getMeetingContext_\(\)\.sources\.revisionsUrl/.test(connBody), true);
    check('testAllConnections() no longer calls the old getCollectorConfig_().REVISIONS_URL pattern as executable code (a comment MAY still name it for context)',
      /getCollectorConfig_\(\)\.REVISIONS_URL/.test(stripComments(connBody)), false);
  }
}

// ============================================ 2. TDoc identifier family ===

console.log('TDoc identifier family: real meeting 86178 IDs (S4aP / plenary-adhoc)');

{
  const { sandbox } = loadCode();
  const realIds = ['S4aP260067', 'S4aP260074', 'S4aP260090', 'S4aP260098'];
  realIds.forEach(id => {
    const parsed = sandbox.parseSA4DocumentId_(id);
    check(`parseSA4DocumentId_("${id}") recognizes it as plenary-adhoc, unchanged casing`,
      { isValid: parsed.isValid, familyKey: parsed.familyKey, raw: parsed.raw },
      { isValid: true, familyKey: 'plenary-adhoc', raw: id });
    check(`normalizeTdoc_("${id}") round-trips exactly`, sandbox.normalizeTdoc_(id), id);
  });
}

// ================================================= 3. real MeetingContext =

console.log('getMeetingContext_() -- meeting 86178 with the required Document Properties');

const MEETING_86178_PROPS = {
  MEETING_TYPE: 'adhoc',
  MEETING_NAME: 'SA4-e (AH) on FS_6G_MED',
  REPORT_SUFFIX: '6G',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/',
  TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx',
  AGENDA_TDOC: 'S4aP260098',
  AGENDA_ITEM_PREFIX: '5.'
};

{
  const { sandbox } = loadCode({ documentProperties: MEETING_86178_PROPS });
  const ctx = sandbox.getMeetingContext_();

  check('meeting.type is "adhoc"', ctx.meeting.type, 'adhoc');
  check('meeting.name matches the real 3GPP meeting name', ctx.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('report.type is "6G"', ctx.report.type, '6G');
  check('report.agendaSelector is {mode:"all"} -- correct here because the real agenda TDoc contains ONLY item 5 and its sub-items (verified by direct inspection of S4aP260098), so "all" and prefix "5." are equivalent for this meeting',
    ctx.report.agendaSelector, { mode: 'all' });
  check('sources.ftpBase is the real SA4_Plenary/Docs folder', ctx.sources.ftpBase, MEETING_86178_PROPS.FTP_BASE);
  check('sources.tdocListUrl is the real TDoc list filename', ctx.sources.tdocListUrl, MEETING_86178_PROPS.TDOC_LIST_URL);
  check('sources.agendaTdoc is the real agenda TDoc "S4aP260098"', ctx.sources.agendaTdoc, 'S4aP260098');

  // getConfiguredAgendaPrefix_() -- NOT context.report.agendaPrefix (which
  // stays null for every ad-hoc meeting, unchanged) -- is what
  // buildSkeletonWithTdocTables()'s parent-removal filter and
  // agendaPrefixNum actually read. This is a Document Properties
  // configuration requirement, not a code path: the existing
  // AGENDA_ITEM_PREFIX override (already used by every main meeting) is
  // reused as-is, no new property introduced.
  check('getConfiguredAgendaPrefix_() reflects the AGENDA_ITEM_PREFIX override ("5."), not the 6G default ("11.")',
    sandbox.getConfiguredAgendaPrefix_(), '5.');
}

// ============================================== 4. real agenda projection =

console.log('real meeting 86178 agenda -- projection + skeleton parent-removal, using the ACTUAL agenda from S4aP260098');

// Read directly out of S4aP260098's own Word document body (Agenda: item 5,
// with sub-items 5.1-5.11 -- no items outside "5" exist in that document at
// all, confirmed by direct inspection, not assumed from a main-meeting
// convention).
const REAL_86178_AGENDA = [
  item('5', 'SA4 WG on FS_6G_MED (Study on Media aspects for 6G System)'),
  item('5.1', 'Opening of the session and registration of documents'),
  item('5.2', 'IPR, antitrust and consensus principles reminder'),
  item('5.3', 'Reports/Liaisons from other groups/meetings'),
  item('5.4', 'FS_6G_MED - General and working documents'),
  item('5.5', 'FS_6G_MED - WT#1: Media Delivery Architecture'),
  item('5.6', 'FS_6G_MED - WT#2: 6G Media'),
  item('5.6.1', 'FS_6G_MED - WT#2.1: AI Traffic Characteristics'),
  item('5.6.2', 'FS_6G_MED - WT#2.2: Other AI-Related topics'),
  item('5.6.3', 'FS_6G_MED - WT#2.3: Non-AI-related topics'),
  item('5.7', 'FS_6G_MED - WT#3: Media Aspects related to SA2 topics'),
  item('5.8', 'FS_6G_MED - WT#4: Media for ubiquitous access'),
  item('5.9', 'FS_6G_MED - WT#5: Trusted and private media communication'),
  item('5.10', 'FS_6G_MED - Other issues'),
  item('5.11', 'Close of the session')
];

{
  const { sandbox } = loadCode({ documentProperties: MEETING_86178_PROPS });
  const ctx = sandbox.getMeetingContext_();

  const projected = sandbox.projectAgendaItems_(REAL_86178_AGENDA, ctx.report.agendaSelector);
  check('projectAgendaItems_ under {mode:"all"} retains every real agenda item, including the bare "5" parent',
    projected.map(i => i.number),
    ['5', '5.1', '5.2', '5.3', '5.4', '5.5', '5.6', '5.6.1', '5.6.2', '5.6.3', '5.7', '5.8', '5.9', '5.10', '5.11']);

  // The EXACT parent-removal filter buildSkeletonWithTdocTables() applies,
  // verbatim (unchanged by this task).
  const agendaPrefix = sandbox.getConfiguredAgendaPrefix_();
  const effective = projected.filter(i => i.number !== agendaPrefix.replace(/\.$/, ''));

  check('effective agenda (after the existing parent-removal filter) drops the bare "5" and keeps every real sub-item, in order',
    effective.map(i => i.number),
    ['5.1', '5.2', '5.3', '5.4', '5.5', '5.6', '5.6.1', '5.6.2', '5.6.3', '5.7', '5.8', '5.9', '5.10', '5.11']);

  check('the effective agenda ends at "Close of the session" with no explicit AOB item (real evidence for the SA4-PROD-001 AOB/Close guard)',
    /close of/i.test(effective[effective.length - 1].title) && !effective.some(i => /any other business|\baob\b/i.test(i.title)),
    true);
}

// =============================================== 5. real TDoc grouping ====

console.log('real meeting 86178 TDoc list -- agenda-selector grouping against captured live data');

{
  const { sandbox } = loadCode({ documentProperties: MEETING_86178_PROPS });
  const ctx = sandbox.getMeetingContext_();
  const rows = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86178-tdoc-list.json'), 'utf8'));

  const groups = {};
  rows.forEach(r => {
    const agendaItem = String(r.AgendaItem || '').trim();
    if (!sandbox.agendaSelectorMatches_(ctx.report.agendaSelector, agendaItem)) return;
    (groups[agendaItem] = groups[agendaItem] || []).push(r.TDoc);
  });

  check('every real TDoc row (non-blank agenda item) is retained under {mode:"all"}',
    Object.values(groups).reduce((n, g) => n + g.length, 0),
    rows.filter(r => String(r.AgendaItem || '').trim() !== '').length);

  check('agenda item "5.6.1" (AI Traffic Characteristics) groups its real 13 TDocs',
    (groups['5.6.1'] || []).length, 13);
  check('agenda item "5.4" (General and working documents) groups its real 6 TDocs',
    (groups['5.4'] || []).length, 6);
  check('the agenda TDoc itself (S4aP260098) is grouped under "5.1" (Opening)',
    (groups['5.1'] || []), ['S4aP260098']);

  // KNOWN DATA GAP (not a code defect): the live TDoc list files
  // S4aP260089 under agenda item "5.0", which does not match any heading in
  // the real agenda (S4aP260098 has no "5.0" section; the agenda's own
  // Word document places 260089's content under section 5.4, and its own
  // "Note from the chair: 0089 and 0094 agenda item were changed" confirms
  // this was a deliberate late edit to the live spreadsheet). Because no
  // agenda item in REAL_86178_AGENDA is literally "5.0", this TDoc will NOT
  // appear under any section of the auto-generated skeleton -- this is
  // flagged in the deliverable as a manual-verification item for Thomas,
  // not something this task's production fix attempts to auto-correct
  // (guessing a chair's intended reassignment in code would be exactly the
  // kind of speculative, meeting-specific special-casing this task's
  // "smallest safe fix" principle rules out).
  check('agenda item "5.0" exists in the live data (orphaned; not a heading in REAL_86178_AGENDA -- see comment above)',
    (groups['5.0'] || []), ['S4aP260089']);
  check('"5.0" is confirmed absent from the real projected/effective agenda (the orphan is real, not a fixture typo)',
    REAL_86178_AGENDA.some(i => i.number === '5.0'), false);
}

// ==================================== 6. main-meeting protection (6G) =====

console.log('main-meeting protection: is6G is UNCHANGED for a real main 6G meeting (meeting.type defaults to "main")');

{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: '6G' } });
  const ctx = sandbox.getMeetingContext_();
  check('a main meeting (no MEETING_TYPE override) has meeting.type "main"', ctx.meeting.type, 'main');
  // is6G's new expression is (reportType === '6G') && context.meeting.type
  // !== 'adhoc'. For every existing main-meeting test/fixture,
  // meeting.type is 'main' (never 'adhoc'), so this reduces to exactly the
  // OLD (reportType === '6G') value -- i.e. still true for REPORT_SUFFIX
  // '6G', unchanged.
  const oldIs6G = (ctx.report.type === '6G');
  const newIs6G = (ctx.report.type === '6G') && ctx.meeting.type !== 'adhoc';
  check('newIs6G === oldIs6G for a main 6G meeting (no behavior change)', newIs6G, oldIs6G);
  check('newIs6G is true for a main 6G meeting', newIs6G, true);
}

// ================================== 7. isBeforeRegistrationBoundary_() ====

console.log('isBeforeRegistrationBoundary_() -- pure function, generic/prefix-driven, real meeting 86178 boundary values');

{
  const { sandbox } = loadCode();
  const fn = sandbox.isBeforeRegistrationBoundary_;

  // Prefix "5" (meeting 86178): the exact examples from the clarified
  // production requirement.
  check('"5" is before the "5.3" boundary (bare parent)', fn('5', '5'), true);
  check('"5.0" is before the "5.3" boundary', fn('5.0', '5'), true);
  check('"5.1" is before the "5.3" boundary', fn('5.1', '5'), true);
  check('"5.2" is before the "5.3" boundary', fn('5.2', '5'), true);
  check('"5.3" is NOT before the boundary (normal agenda sections begin here)', fn('5.3', '5'), false);
  check('"5.4" is NOT before the boundary', fn('5.4', '5'), false);
  check('"5.6.1" is NOT before the boundary (nested sub-item, still >= 5.3)', fn('5.6.1', '5'), false);

  // Generic across a DIFFERENT prefix, proving no literal "5" is baked in.
  check('for prefix "11" (a main 6G-plenary AGENDA_ITEM_PREFIX), "11.1" is before the "11.3" boundary', fn('11.1', '11'), true);
  check('for prefix "11", "11.3" is NOT before the boundary', fn('11.3', '11'), false);

  // A value belonging to a DIFFERENT top-level item must never match.
  check('"6.1" is never before the "5.3" boundary (different top-level item)', fn('6.1', '5'), false);
  check('blank/garbage input does not throw and returns false', fn('', '5'), false);
}

// ======================= 8. registration-boundary extraction (real data) ==

console.log('meeting 86178 -- registration-boundary extraction against the real captured TDoc list (5.1.4 Documents bucket)');

{
  const { sandbox } = loadCode({ documentProperties: MEETING_86178_PROPS });
  const ctx = sandbox.getMeetingContext_();
  const rows = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86178-tdoc-list.json'), 'utf8'));
  const agendaPrefixNum = sandbox.getConfiguredAgendaPrefix_().replace(/\.$/, '');
  check('agendaPrefixNum derived from Document Properties is "5" (not hardcoded)', agendaPrefixNum, '5');

  // The SAME grouping downloadAndGroupTdocs_() performs (agenda-selector
  // filtered), followed by the SAME extraction buildSkeletonWithTdocTables()
  // now performs (isBeforeRegistrationBoundary_(), gated on !is6G -- is6G is
  // false for this ad-hoc 6G-type meeting per SA4-PROD-001).
  const groups = {};
  rows.forEach(r => {
    const agendaItem = String(r.AgendaItem || '').trim();
    if (!sandbox.agendaSelectorMatches_(ctx.report.agendaSelector, agendaItem)) return;
    (groups[agendaItem] = groups[agendaItem] || []).push(r);
  });

  const registrationDocs = [];
  Object.keys(groups).forEach(key => {
    if (sandbox.isBeforeRegistrationBoundary_(key, agendaPrefixNum)) {
      registrationDocs.push(...groups[key]);
      delete groups[key];
    }
  });

  check('registrationDocs collects exactly S4aP260098 (5.1) and S4aP260089 (5.0), preserving each one\'s ORIGINAL agenda item',
    registrationDocs.map(d => ({ tdoc: d.TDoc, originalAgendaItem: d.AgendaItem })).sort((a, b) => a.tdoc.localeCompare(b.tdoc)),
    [{ tdoc: 'S4aP260089', originalAgendaItem: '5.0' }, { tdoc: 'S4aP260098', originalAgendaItem: '5.1' }]);

  check('S4aP260098 no longer sits in the raw "5.1" group after extraction (it would otherwise render directly under the 5.1 heading)',
    groups['5.1'], undefined);
  check('S4aP260089\'s orphaned "5.0" group no longer exists after extraction (it now has a home: 5.1.4, not nowhere)',
    groups['5.0'], undefined);
  check('agenda item "5.3" (Reports/Liaisons) is untouched by the extraction -- still renders under its own real section',
    (groups['5.3'] || []).map(d => d.TDoc), ['S4aP260090']);
  check('agenda item "5.6.1" is untouched by the extraction -- still 13 real TDocs under its own real section',
    (groups['5.6.1'] || []).length, 13);
}

// ============================ 9. real 5.1/5.2 remain single, real headings

console.log('meeting 86178 -- exactly one real 5.1 and 5.2, agenda continues through 5.11, no synthetic AOB after Close');

{
  const numbers = REAL_86178_AGENDA.map(i => i.number);
  check('"5.1" (Opening) appears exactly once in the real agenda', numbers.filter(n => n === '5.1').length, 1);
  check('"5.2" (IPR) appears exactly once in the real agenda', numbers.filter(n => n === '5.2').length, 1);
  check('the real agenda continues through "5.11" (Close)', numbers[numbers.length - 1], '5.11');
  // "5.1.1"/"5.1.2"/"5.1.3"/"5.1.4" are SYNTHESIZED (by the reverted
  // openingSection branch / ensureReallocationTable_() / the new X.1.4
  // step) -- they are legitimately NOT present in the raw parsed agenda
  // itself; this just documents that fact so it is not mistaken for a gap.
  ['5.1.1', '5.1.2', '5.1.3', '5.1.4'].forEach(synthetic => {
    check(`"${synthetic}" is not itself a parsed agenda item (it is synthesized under the real "5.1" item)`,
      numbers.indexOf(synthetic), -1);
  });
}

// ==================================== 10. ad-hoc title, no meeting number ==

console.log('ad-hoc title: uses context.meeting.name, never invents a meeting number, never "SA4#null"');

{
  const { sandbox } = loadCode({ documentProperties: MEETING_86178_PROPS });
  const ctx = sandbox.getMeetingContext_();
  check('meeting 86178: context.meeting.portalId is null (no real portal number -- confirms the old formula had nothing valid to fall back on)',
    ctx.meeting.portalId, null);

  const title = sandbox.generateReportTitle_({
    REPORT_SUFFIX: ctx.report.type,
    TDOC_LIST_URL: ctx.sources.tdocListUrl,
    MEETING_ID: ctx.meeting.portalId,
    meetingLabel: ctx.meeting.type === 'adhoc' ? ctx.meeting.name : undefined
  });
  check('ad-hoc title contains the real meeting name', title.indexOf('SA4-e (AH) on FS_6G_MED') !== -1, true);
  check('ad-hoc title never contains "SA4#null"', title.indexOf('SA4#null'), -1);
  check('ad-hoc title is exactly "6G Media Minutes – SA4-e (AH) on FS_6G_MED"',
    title, '6G Media Minutes – SA4-e (AH) on FS_6G_MED');
}

console.log('main-meeting title: generateReportTitle_() output is byte-identical to before (meetingLabel never set for main meetings)');

{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: '6G', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/x/TDoc_List_Meeting_SA4%23137-e.xlsx', MEETING_ID: '60777' } });
  const ctx = sandbox.getMeetingContext_();
  check('a main meeting has meeting.type "main" (meetingLabel stays undefined)', ctx.meeting.type, 'main');

  const titleWithoutLabel = sandbox.generateReportTitle_({
    REPORT_SUFFIX: ctx.report.type,
    TDOC_LIST_URL: ctx.sources.tdocListUrl,
    MEETING_ID: ctx.meeting.portalId
  });
  const titleThroughWrapperContract = sandbox.generateReportTitle_({
    REPORT_SUFFIX: ctx.report.type,
    TDOC_LIST_URL: ctx.sources.tdocListUrl,
    MEETING_ID: ctx.meeting.portalId,
    meetingLabel: ctx.meeting.type === 'adhoc' ? ctx.meeting.name : undefined
  });
  check('main-meeting title unchanged: "6G Media Minutes SA4#137-e"', titleWithoutLabel, '6G Media Minutes SA4#137-e');
  check('setDocumentTitleFromTemplate_()\'s conditional meetingLabel (undefined for main) produces the identical title',
    titleThroughWrapperContract, titleWithoutLabel);
}

// ============================== 11. revisions "Not configured", not 403 ====

console.log('meeting 86178 -- revisions source is genuinely absent, characterized as "Not configured", no fabricated URL/403');

{
  const { sandbox } = loadCode({ documentProperties: MEETING_86178_PROPS });
  const ctx = sandbox.getMeetingContext_();
  check('meeting 86178 has no revisionsUrl (no REVISIONS_URL override was configured, and none should be invented)',
    ctx.sources.revisionsUrl, undefined);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
