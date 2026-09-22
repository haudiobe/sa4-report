/**
 * PROD-017 — Support ad-hoc SA4 draft documents as revisions.
 *
 * Characterization (see this task's report for the full write-up):
 * updateRevisions_(cfg) is the actual "revision workflow" consumer -- it
 * fetches a drafts/revisions folder's directory listing, parses its <a>
 * anchors (parseAnchors_()), matches each anchor against every TDoc table
 * already in the report body, and writes matches into that table's
 * "Revisions" cell (plus inserts a minimal new table via
 * insertRevisedDocTablesAfter_() for any matched file whose TDoc number
 * doesn't already have a table of its own). This is DISTINCT from the
 * separate "Revised to" column mechanism (getRevisedTo_()/tdocNumberOf_()/
 * orderTdocsByRevision_()) used during report BUILD to order same-meeting
 * TDoc chains -- that mechanism was already migrated to the central SA4
 * TDoc identifier registry (SA4-IMPL-001) and is untouched here.
 *
 * Before this task, updateRevisions_() had two production-relevant defects
 * for ad-hoc meetings:
 *   1. It read cfg.REVISIONS_URL, which getCollectorConfig_()/
 *      getReportConfig_() ALWAYS compute via the main-meeting-only
 *      `${INBOX_BASE}Drafts/${DRAFTS_FOLDER}` formula -- fabricating a URL
 *      for meeting 86178 that was never its real ad-hoc drafts location
 *      (testAllConnections()'s own Test 4 already carried an identical fix,
 *      see its PROD-002 comment -- this task applies the same fix here).
 *   2. It matched a report's own TDoc against a candidate filename with
 *      `new RegExp(escapeRegExp_(tdoc)).test(text)` -- a literal substring
 *      test with no digit-boundary guard, which could accidentally match an
 *      unrelated, longer TDoc number sharing the same leading digits, and
 *      which (via cfg.TDOC_ID_REGEX's '^S4-\d{6}$' default) could never
 *      recognize an ad-hoc identifier family (S4aA/S4aP/S4aV/S4aI/A4aR) at
 *      all on the "this table's own TDoc" side either.
 *
 * This suite proves both fixes, using a fake DocumentApp Body/Table/Cell
 * (same minimal-surface pattern as tests/meeting-86178-production.test.js's
 * ensureReallocationTable_() coverage) so updateRevisions_() itself -- the
 * real, unmodified production function -- can be executed end-to-end, not
 * just characterized via source-structure regex.
 *
 * Run: node tests/revision-drafts-folder.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');

const DRAFTS_FIXTURE = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86178-drafts-folder.json'), 'utf8'));

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

// ---- minimal fake DocumentApp surface, just enough for updateRevisions_()/
// insertRevisedDocTablesAfter_() to run against a hand-built document state.

function makeFakeTextEditor(state) {
  return {
    getText: () => state.text,
    appendText: (s) => { state.text += s; },
    setLinkUrl: (start, end, url) => { state.links.push({ start, end, url }); },
    setForegroundColor: () => {},
    setBold: () => {}
  };
}

function makeFakeCell(initialText) {
  const state = { text: String(initialText === undefined || initialText === null ? '' : initialText), links: [] };
  return {
    getText: () => state.text,
    setText: (t) => { state.text = String(t || ''); state.links = []; },
    editAsText: () => makeFakeTextEditor(state),
    _state: state
  };
}

function makeFakeRow(cellTexts) {
  const cells = cellTexts.map(t => makeFakeCell(t));
  return {
    getNumCells: () => cells.length,
    getCell: (i) => cells[i],
    appendTableCell: (t) => { const c = makeFakeCell(t); cells.push(c); return c; },
    _cells: cells
  };
}

function makeFakeTable(rows) {
  const _rows = rows.map(r => makeFakeRow(r));
  return {
    getType: () => 'TABLE',
    getNumRows: () => _rows.length,
    getRow: (i) => _rows[i],
    getCell: (r, c) => _rows[r].getCell(c),
    appendTableRow: () => { const row = makeFakeRow([]); _rows.push(row); return row; },
    removeRow: (i) => { _rows.splice(i, 1); },
    _rows
  };
}

function makeFakeParagraph(text) {
  return { getType: () => 'PARAGRAPH', getText: () => text };
}

function makeFakeBody(children) {
  const _children = children.slice();
  return {
    _children,
    getNumChildren: () => _children.length,
    getChild: (i) => _children[i],
    getChildIndex: (child) => _children.indexOf(child),
    getTables: () => _children.filter(c => c.getType() === 'TABLE'),
    insertTable: (idx, data) => { const t = makeFakeTable(data); _children.splice(idx, 0, t); return t; }
  };
}

// A single, standard-schema TDoc table, matching the exact row order
// insertRevisedDocTablesAfter_() itself creates (and what a real report
// table has) -- row 6 is "Revisions", matching
// findOrFallbackCell_(t, [...], 6, 1)'s coordinate fallback.
function makeTdocTable(tdoc, agendaItem) {
  return makeFakeTable([
    ['TDoc', tdoc],
    ['Title', 'Some title'],
    ['Source', 'Some source'],
    ['Contact', ''],
    ['Agenda Item', agendaItem || '5.1'],
    ['E-mail Discussion', ''],
    ['Revisions', ''],
    ['Minutes', ''],
    ['Disposition', ''],
    ['Status', 'Available']
  ]);
}

function stubUrlFetch(sandbox, urlToHtml) {
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      const html = urlToHtml[url];
      if (html === undefined) throw new Error('Unexpected fetch URL in test: ' + url);
      return { getResponseCode: () => 200, getContentText: () => html };
    }
  };
}

// ============================== 1. parseDraftAnchor_() -- canonical parsing

console.log('parseDraftAnchor_() -- canonical TDoc parsing via the central SA4 identifier registry (6dc526b)');

{
  const { sandbox } = loadCode();
  const fn = sandbox.parseDraftAnchor_;

  const qcom = fn({ href: 'S4aP260071_QCOM.docx', text: 'S4aP260071_QCOM.docx' }, 'https://x/drafts/');
  check('S4aP260071_QCOM.docx -> tdocId S4aP260071 (exact task example)', qcom.tdocId, 'S4aP260071');
  check('S4aP260071_QCOM.docx -> fileName preserved with its "_QCOM" suffix, not discarded', qcom.fileName, 'S4aP260071_QCOM.docx');
  check('S4aP260071_QCOM.docx -> url resolved absolute', qcom.url, 'https://x/drafts/S4aP260071_QCOM.docx');

  // Representative supported families.
  const cases = [
    ['S4-261834_r1.docx', 'S4-261834'],
    ['S4aA260090_contribution.docx', 'S4aA260090'],
    ['S4aP260071_QCOM.docx', 'S4aP260071'],
    ['S4aV260012_draft.docx', 'S4aV260012'],
    ['S4aI260005_v2.docx', 'S4aI260005'],
    ['A4aR260033_proposal.docx', 'A4aR260033']
  ];
  cases.forEach(([fileName, expectedId]) => {
    const draft = fn({ href: fileName, text: fileName }, 'https://x/drafts/');
    check(`${fileName} -> tdocId ${expectedId}`, draft && draft.tdocId, expectedId);
    check(`${fileName} -> fileName round-trips exactly`, draft && draft.fileName, fileName);
  });

  // Unsupported/unrecognized family -- never guessed.
  const unsupported = fn({ href: 'MTSI-260001_draft.docx', text: 'MTSI-260001_draft.docx' }, 'https://x/drafts/');
  check('an unsupported/unregistered family (e.g. "MTSI-") is NOT guessed -- returns null', unsupported, null);

  // A working document with no TDoc identifier at all -- real negative case.
  const noId = fn({ href: '3GPP_TR_26870-0.5.1-6G-MED%20rm.docx', text: '3GPP_TR_26870-0.5.1-6G-MED rm.docx' }, 'https://x/drafts/');
  check('a filename with no recognizable SA4 identifier -> null (real, observed negative case)', noId, null);

  check('empty anchor text -> null, does not throw', fn({ href: 'x', text: '' }, 'https://x/'), null);
  check('missing anchor -> null, does not throw', fn(null, 'https://x/'), null);
}

// ==================== 2. exact matching -- no accidental neighbor match =====

console.log('canonical matching -- exact identity, never a substring/prefix accident');

{
  const { sandbox } = loadCode();
  const fn = sandbox.parseDraftAnchor_;

  const target = 'S4aP260071';
  const match = fn({ href: 'S4aP260071_QCOM.docx', text: 'S4aP260071_QCOM.docx' }, 'https://x/');
  check('S4aP260071_QCOM.docx matches S4aP260071 by canonical identity', match.tdocId === target, true);

  // The exact regression the old `escapeRegExp_(tdoc)` substring test was
  // vulnerable to: "S4aP260071" IS a literal leading substring of
  // "S4aP2600711_other_contributor.docx" (an unrelated 7-digit run, not a
  // real/valid TDoc number), so the OLD `rx.test(text)` check would have
  // WRONGLY matched this file to S4aP260071. The central registry's
  // `(?!\d)` digit-boundary guard means no family matches here at all --
  // it is correctly unrecognized, never mis-attributed to S4aP260071.
  const neighborDigitRun = fn({ href: 'S4aP2600711_other_contributor.docx', text: 'S4aP2600711_other_contributor.docx' }, 'https://x/');
  check('S4aP2600711_other_contributor.docx (old code\'s false-positive case) is correctly UNRECOGNIZED, not matched to S4aP260071',
    neighborDigitRun, null);

  // A genuinely different, adjacent TDoc number -- must never be conflated.
  const adjacentTdoc = fn({ href: 'S4aP260072_someone.docx', text: 'S4aP260072_someone.docx' }, 'https://x/');
  check('S4aP260072_someone.docx parses to the distinct canonical id S4aP260072', adjacentTdoc.tdocId, 'S4aP260072');
  check('S4aP260071 does NOT match S4aP260072\'s draft', adjacentTdoc.tdocId !== target, true);

  // A different family/series sharing the same numeric suffix must not match.
  const differentFamily = fn({ href: 'S4aA260071_unrelated.docx', text: 'S4aA260071_unrelated.docx' }, 'https://x/');
  check('S4aA260071 (Audio family) is a different canonical id from S4aP260071 (Plenary family)',
    differentFamily.tdocId !== target, true);
}

// ================ 3. real 86178 drafts-folder fixture (trimmed, evidence-based)

console.log('real meeting 86178 drafts folder -- trimmed fixture (S4aP260071_BBC.docx, S4aP260071_QCOM.docx confirmed real)');

{
  const { sandbox } = loadCode();
  const anchors = sandbox.parseAnchors_(DRAFTS_FIXTURE.html);
  check('the fixture listing parses to the expected anchor count (4 real files + "../")', anchors.length, 5);

  const drafts = anchors.map(a => sandbox.parseDraftAnchor_(a, 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/'))
    .filter(Boolean);
  check('exactly 3 anchors parse to a recognized SA4 identifier ("../" and the working-document filename do not)', drafts.length, 3);

  const s071 = drafts.filter(d => d.tdocId === 'S4aP260071');
  check('S4aP260071 has TWO real, distinct draft files (BBC and QCOM) -- confirmed via live investigation, not synthetic',
    s071.map(d => d.fileName).sort(), ['S4aP260071_BBC.docx', 'S4aP260071_QCOM.docx']);

  const s096 = drafts.find(d => d.tdocId === 'S4aP260096');
  check('S4aP260096_stabilization_huawei.docx resolves to S4aP260096', !!s096, true);
}

// =================== 4. ad-hoc source selection (MeetingContext) ===========

console.log('updateRevisions_() -- ad-hoc meeting uses context.sources.revisionsUrl, never the main-meeting DRAFTS_FOLDERS formula');

const AD_HOC_86178_URL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/';

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_86178_URL } });
  const ctx = sandbox.getMeetingContext_();
  const cfg = sandbox.getReportConfig_();

  check('the always-computed cfg.REVISIONS_URL (main-meeting formula) is NOT the real ad-hoc URL -- confirms the two sources genuinely differ for this meeting',
    cfg.REVISIONS_URL !== AD_HOC_86178_URL, true);
  check('getMeetingContext_().sources.revisionsUrl IS the configured ad-hoc REVISIONS_URL', ctx.sources.revisionsUrl, AD_HOC_86178_URL);

  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };

  const fetchedUrls = [];
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      fetchedUrls.push(url);
      return { getResponseCode: () => 200, getContentText: () => DRAFTS_FIXTURE.html };
    }
  };

  const collectorCfg = sandbox.getCollectorConfig_();
  sandbox.updateRevisions_(collectorCfg);

  check('updateRevisions_() fetched the CONFIGURED ad-hoc drafts URL, not the fabricated main-meeting formula',
    fetchedUrls.length === 1 && fetchedUrls[0] === AD_HOC_86178_URL, true);
  check('the Revisions cell for S4aP260071 now contains both real draft filenames',
    parentTable.getCell(6, 1).getText().split('\n').filter(Boolean).sort(),
    ['S4aP260071_BBC.docx', 'S4aP260071_QCOM.docx']);
}

// ============================ 5. main-meeting regression ====================

console.log('updateRevisions_() -- main-meeting source/matching behavior is unchanged');

{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: '6G' } });
  const ctx = sandbox.getMeetingContext_();
  const cfg = sandbox.getReportConfig_();
  check('a main meeting: getMeetingContext_().sources.revisionsUrl is IDENTICAL to cfg.REVISIONS_URL (the same main-meeting formula, unchanged)',
    ctx.sources.revisionsUrl, cfg.REVISIONS_URL);

  const parentTable = makeTdocTable('S4-260123', '11.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };

  const mainHtml = '<html><body><a href="S4-260123_r1.docx">S4-260123_r1.docx</a></body></html>';
  const fetchedUrls = [];
  sandbox.UrlFetchApp = {
    fetch: (url) => { fetchedUrls.push(url); return { getResponseCode: () => 200, getContentText: () => mainHtml }; }
  };

  const collectorCfg = sandbox.getCollectorConfig_();
  sandbox.updateRevisions_(collectorCfg);

  check('main meeting: updateRevisions_() fetched exactly cfg.REVISIONS_URL', fetchedUrls[0], cfg.REVISIONS_URL.replace(/\/$/, '') + '/');
  check('main meeting: the matching draft file (S4-260123_r1.docx) is still recognized and linked',
    parentTable.getCell(6, 1).getText().trim(), 'S4-260123_r1.docx');

  // The exact substring-overlap safety property, now verified for the real
  // main-meeting family too -- a fresh sandbox/props so this TDoc's own
  // REVIS_ store starts empty (isolated from the assertion just above).
  const { sandbox: overlapSandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: '6G' } });
  const overlapTable = makeTdocTable('S4-260200', '11.1');
  const overlapBody = makeFakeBody([overlapTable]);
  overlapSandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => overlapBody }) };
  const overlapHtml = '<html><body><a href="S4-2602001_unrelated.docx">S4-2602001_unrelated.docx</a></body></html>';
  overlapSandbox.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => overlapHtml }) };
  overlapSandbox.updateRevisions_(overlapSandbox.getCollectorConfig_());
  check('S4-260200 does NOT accidentally match the unrelated, longer S4-2602001\'s draft (digit-boundary safe, unlike the old substring test)',
    overlapTable.getCell(6, 1).getText().trim(), 'No revisions available.');
}

// ============================ 6. missing revisions URL ======================

console.log('updateRevisions_() -- graceful no-op when no revisions URL is configured/derivable');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G' } }); // no REVISIONS_URL set
  const ctx = sandbox.getMeetingContext_();
  check('ad-hoc meeting with no REVISIONS_URL configured: sources.revisionsUrl is genuinely absent', ctx.sources.revisionsUrl, undefined);

  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  let fetchAttempted = false;
  sandbox.UrlFetchApp = { fetch: () => { fetchAttempted = true; return { getResponseCode: () => 200, getContentText: () => '' }; } };

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());
  check('no fetch is attempted when no revisions URL is configured', fetchAttempted, false);
  check('the Revisions cell is left untouched (no-op, not an error, not "No revisions available.")',
    parentTable.getCell(6, 1).getText(), '');
}

// ============================ 7. multiple candidates =========================

console.log('multiple draft files for one TDoc -- deterministic, all surfaced, no invented "latest"');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_86178_URL } });
  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  stubUrlFetch(sandbox, { [AD_HOC_86178_URL.replace(/\/$/, '') + '/']: DRAFTS_FIXTURE.html });

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());

  const lines = parentTable.getCell(6, 1).getText().split('\n').filter(Boolean);
  check('BOTH real draft files for S4aP260071 are surfaced (no silent "latest" pick)', lines.length, 2);
  check('the two lines are deterministically ordered (alphabetical by filename, existing, unchanged sort rule)',
    lines, ['S4aP260071_BBC.docx', 'S4aP260071_QCOM.docx']);

  // Running it again (idempotency / stable store) does not duplicate entries.
  sandbox.updateRevisions_(sandbox.getCollectorConfig_());
  const linesAfterRerun = parentTable.getCell(6, 1).getText().split('\n').filter(Boolean);
  check('re-running does not duplicate entries (existing REVIS_ store dedupe, unchanged)', linesAfterRerun.length, 2);
}

// ============================ 8. report integration =========================

console.log('report integration -- a registered S4aP260071 receives the discovered draft/revision link through the existing workflow');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_86178_URL } });
  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  stubUrlFetch(sandbox, { [AD_HOC_86178_URL.replace(/\/$/, '') + '/']: DRAFTS_FIXTURE.html });

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());

  check('the Revisions cell is no longer empty', parentTable.getCell(6, 1).getText().trim() !== '', true);
  const cellState = parentTable.getCell(6, 1)._state;
  check('a real hyperlink was attached for at least one draft', cellState.links.length >= 1, true);
  check('the attached link points to the real absolute drafts-folder URL',
    cellState.links.every(l => l.url.indexOf(AD_HOC_86178_URL) === 0), true);

  // Because each draft file's own canonical id (S4aP260071) is the SAME as
  // its parent table's own TDoc -- these are alternate drafts of the
  // SAME, already-registered TDoc, not a genuinely different/newer TDoc
  // number -- insertRevisedDocTablesAfter_() must NOT create a redundant
  // duplicate table. This is the correct, pre-existing behavior (the
  // "already exists" check in insertRevisedDocTablesAfter_() finds the
  // parent table itself), not a new special case introduced here.
  check('no duplicate/synthetic child table was created for S4aP260071 (it already has a table -- itself)',
    body.getTables().length, 1);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
