/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-upgrade-004.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-upgrade-004.test.js
 */

/**
 * LEGACY-UPGRADE-004 -- generic ad-hoc administrative-section anchoring,
 * ported from the central add-on's ADDON-007A to this legacy bound
 * script's own buildSkeletonWithTdocTables()/ensureReallocationTable_().
 *
 * Central's pre-acceptance analysis (meeting 85916, an ad-hoc series whose
 * real agenda is numbered 1..4.x) showed that buildSkeletonWithTdocTables()
 * anchored the Opening / Registration-of-Documents / Document-Reallocations /
 * IPR sections on `${REPORT_SUFFIX prefix}.1` / `.1.2` / `.2` (e.g. "7.1",
 * "7.2" for Audio). For an ad-hoc meeting whose real numbering is independent
 * of the report family those anchors never matched, so the administrative
 * sections were silently skipped. Legacy had the exact same defect (it
 * predates ADDON-007A), reproduced and fixed identically here.
 *
 * Rule under test: report family (Audio/Video/6G...) and meeting structure
 * (main/adhoc) are separate. MAIN meetings keep the configured-prefix
 * anchors exactly. ADHOC meetings derive the anchors from the REAL parsed
 * agenda and never consult the report-family prefix.
 *
 * These tests run the REAL buildSkeletonWithTdocTables()/
 * ensureReallocationTable_() against a fake document body. Only the network/
 * template boundaries (agenda download, TDoc-list download, template doc)
 * are stubbed. No meeting number, series name or agenda numbering is
 * hard-coded in production code; the agendas below are synthetic. Known real
 * meetings (86178 = FS_6G_MED, 85916 = Audio SWG/ULBC-MED) are never
 * referenced here -- this suite is fully generic/table-driven, matching the
 * LEGACY-UPGRADE-004 task's requirement.
 *
 * This harness is isolated -- excluded from the Apps Script push payload by
 * .claspignore.
 *
 * Run: node tests/legacy-upgrade-004.test.js
 */

const { loadCode } = require('./helpers/legacy-load-code.js');

// TEMPLATE-002A harness adaptation: Legacy's historical commits exist only in
// the sa4-report-legacy repository. Blocks that load one (to show Legacy's own
// before/after, or that a Legacy stage touched only certain functions) are
// skipped here; they say nothing about the current Code.js.
const LEGACY_GIT_HISTORY = false;
function skipLegacyHistory(what) { console.log('  skip ' + what + ' [needs Legacy git history]'); }

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

// ------------------------------------------------------------ fake document

function makeFakeDocumentBody(sandbox) {
  const PARAGRAPH = sandbox.DocumentApp.ElementType.PARAGRAPH;
  const TABLE = sandbox.DocumentApp.ElementType.TABLE;
  const NORMAL = sandbox.DocumentApp.ParagraphHeading.NORMAL;
  const children = [];

  function makeParagraph(text, heading) {
    const p = {
      _text: text, _heading: heading || NORMAL,
      getType: () => PARAGRAPH, asParagraph: () => p,
      getText: () => p._text,
      getHeading: () => p._heading,
      setHeading: (h) => { p._heading = h; return p; },
      setGlyphType: () => p
    };
    return p;
  }
  function makeTable(rowsData) {
    const rows = [];
    function makeRow(cellTexts) {
      const cells = cellTexts.map(t => ({ _t: String(t), getText: () => String(t), editAsText: () => ({ setLinkUrl() {} }) }));
      return {
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { const c = { _t: String(t), getText: () => String(t), editAsText: () => ({ setLinkUrl() {} }) }; cells.push(c); return c; }
      };
    }
    (rowsData || []).forEach(r => rows.push(makeRow(r)));
    const t = {
      getType: () => TABLE, asTable: () => t,
      getNumRows: () => rows.length, getRow: (i) => rows[i],
      getCell: (r, c) => rows[r].getCell(c),
      appendTableRow: () => { const r = makeRow([]); rows.push(r); return r; },
      removeRow: (i) => { rows.splice(i, 1); }
    };
    return t;
  }

  const body = {
    _children: children,
    clear: () => { children.length = 0; return body; },
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
    getChildIndex: (c) => children.indexOf(c),
    getTables: () => children.filter(c => c.getType() === TABLE),
    getParagraphs: () => children.filter(c => c.getType() === PARAGRAPH),
    appendParagraph: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendListItem: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendTable: () => { const t = makeTable([]); children.push(t); return t; },
    insertParagraph: (idx, text) => { const p = makeParagraph(text); children.splice(idx, 0, p); return p; },
    insertTable: (idx, data) => { const t = makeTable(data); children.splice(idx, 0, t); return t; }
  };
  body._headingTexts = () => children.filter(c => c.getType() === PARAGRAPH && c.getHeading() !== NORMAL).map(c => c.getText());
  return body;
}

function tdoc(id, agendaItem) {
  return {
    row: [id, 'Title of ' + id, 'Source Co', agendaItem],
    richTextRow: null,
    tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 3,
    agendaTopicCol: -1, statusCol: -1, typeCol: -1, forCol: -1, revisedToCol: -1
  };
}

function item(number, title) {
  return { number, title, level: number.split('.').length, heading: 'NORMAL', text: '' };
}

function runBuild(props, agendaItems, tdocsByItem) {
  const { sandbox, docProps } = loadCode({ documentProperties: props });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  sandbox.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(sandbox) });
  sandbox.DocumentApp.getUi = () => ({ alert() {}, ButtonSet: { OK: 'OK' } });
  sandbox.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  sandbox.setDocumentTitleFromTemplate_ = () => {};
  sandbox.parseAgendaForReport_ = () => agendaItems;
  sandbox.findHeading_ = () => null; // shared main-meeting template: not under test
  sandbox.styleStatusCell_ = () => {};
  sandbox.removeRowHeightAndSpacing = () => {};
  const groups = {};
  Object.keys(tdocsByItem || {}).forEach(k => { groups[k] = { tdocs: tdocsByItem[k].map(id => tdoc(id, k)) }; });
  sandbox.downloadAndGroupTdocs_ = () => groups;
  sandbox.buildSkeletonWithTdocTables();
  return { sandbox, body, docProps };
}

const ADHOC = (suffix) => ({ MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic AH', REPORT_SUFFIX: suffix, AGENDA_TDOC: 'S4aA000001', FTP_BASE: 'https://example.invalid/Docs/', TDOC_LIST_URL: 'https://example.invalid/list.xlsx' });

const SYNTHETIC_ADHOC_AGENDA = [
  item('1', 'Opening of the session'),
  item('1.1', "Chair's notes"),
  item('2', 'Approval of the agenda'),
  item('3', 'IPR and antitrust reminder'),
  item('4', 'Technical contributions'),
  item('4.1', 'Requirements'),
  item('5', 'Any other business'),
  item('6', 'Closing of the session')
];

// ============================================ 1. ad-hoc Audio, non-7.x numbering

console.log('ad-hoc + Audio with real 1..6 numbering: administrative sections anchor on the REAL agenda');

{
  const { body } = runBuild(ADHOC('Audio'), SYNTHETIC_ADHOC_AGENDA, {
    '1': ['S4aA000010'],       // filed under the opening item itself
    '1.1': ['S4aA000011'],     // filed under a REAL child of the opening item
    '4.1': ['S4aA000012']
  });
  const h = body._headingTexts();

  ['1 Opening of the session', "1.1 Chair's notes", '2 Approval of the agenda', '3 IPR and antitrust reminder',
   '4 Technical contributions', '4.1 Requirements', '5 Any other business', '6 Closing of the session'
  ].forEach(real => check(`real agenda heading preserved: "${real}"`, h.indexOf(real) !== -1, true));

  check('Opening content subsection anchored under the real opening item, after its real child (1.1 -> 1.2)',
    h.indexOf('1.2 Opening of the session') > h.indexOf("1.1 Chair's notes"), true);
  check('Registration of Documents summary heading is present, anchored at 1.3',
    h.indexOf('1.3 Registration of Documents') !== -1, true);
  check('"Documents" bucket (TDocs filed under the opening item) is present at 1.5 (1.4 reserved for Document Reallocations)',
    h.indexOf('1.5 Documents') !== -1, true);
  check('IPR children are anchored under the REAL IPR item (3.1..3.4)',
    ['3.1 Introduction', '3.2 Call for IPRs', '3.3 Statement regarding competition law', '3.4 Consensus principles reminder']
      .map(t => h.indexOf(t) !== -1), [true, true, true, true]);
  check('document order matches numeric order: 1, 1.1, 1.2, 1.3, 1.5, 2, 3, 3.1',
    ['1 Opening of the session', "1.1 Chair's notes", '1.2 Opening of the session', '1.3 Registration of Documents', '1.5 Documents', '2 Approval of the agenda', '3 IPR and antitrust reminder', '3.1 Introduction']
      .map(t => h.indexOf(t)).every((v, i, a) => v !== -1 && (i === 0 || v > a[i - 1])), true);
  check('no heading imposes the report-family prefix (nothing starts with "7")',
    h.filter(t => /^7(\.|\s)/.test(t)), []);

  const tables = body._children.filter(c => c.getType() === 'TABLE');
  const tdocCells = tables.map(t => t.getRow(0).getNumCells() > 1 ? t.getRow(0).getCell(1).getText() : t.getRow(0).getCell(0).getText());
  const idxOf = (pred) => body._children.findIndex(pred);
  check('TDoc grouping still follows the REAL item numbers (S4aA000012 rendered under 4.1, after its heading)',
    idxOf(c => c.getType() === 'TABLE' && c.getRow(0).getCell(1).getText() === 'S4aA000012') >
      idxOf(c => c.getType() === 'PARAGRAPH' && c.getText() === '4.1 Requirements'), true);
  check('a TDoc filed under a REAL child of the opening item (1.1) stays under it -- not captured into Documents',
    body._children.findIndex(c => c.getType() === 'TABLE' && c.getRow(0).getCell(1).getText() === 'S4aA000011') >
      body._children.findIndex(c => c.getType() === 'PARAGRAPH' && c.getText() === "1.1 Chair's notes"), true);
  check('the TDoc filed under the opening item itself is rendered (not silently dropped)',
    tdocCells.indexOf('S4aA000010') !== -1, true);
}

console.log('ad-hoc administrative block, no real opening children: same shape as the main layout');

{
  const agenda = [item('1', 'Opening of the session'), item('2', 'IPR and antitrust reminder'), item('3', 'Topics'), item('4', 'Closing of the session')];
  const { body } = runBuild(ADHOC('Audio'), agenda, {});
  const h = body._headingTexts();
  check('subsections are X.1 / X.2 under the real opening item',
    ['1.1 Opening of the session', '1.2 Registration of Documents'].map(t => h.indexOf(t) !== -1), [true, true]);
  check('IPR children under the real IPR item (2.1..2.4)',
    ['2.1 Introduction', '2.2 Call for IPRs', '2.3 Statement regarding competition law', '2.4 Consensus principles reminder'].map(t => h.indexOf(t) !== -1), [true, true, true, true]);
}

// ============================================ 2. different family + numbering

console.log('ad-hoc with a different report family and numbering: anchors independent of report family');

{
  const agenda = [
    item('3', 'Opening of the meeting'), item('4', 'Agenda'), item('5', 'IPR policy reminder'),
    item('6', 'Contributions'), item('7', 'Any other business'), item('8', 'Close of the session')
  ];
  const video = runBuild(ADHOC('Video'), agenda, { '3': ['S4aV000001'] });
  const mbs = runBuild(ADHOC('MBS'), agenda, { '3': ['S4aV000001'] });
  const hv = video.body._headingTexts();
  check('anchors come from the real agenda: opening subsections under 3, IPR children under 5',
    ['3.1 Opening of the session', '3.2 Registration of Documents', '5.1 Introduction', '5.4 Consensus principles reminder'].map(t => hv.indexOf(t) !== -1),
    [true, true, true, true]);
  check('Video prefix ("9") is never used', hv.filter(t => /^9(\.|\s)/.test(t)), []);
  check('the generated heading sequence is IDENTICAL for two different report families (Video vs MBS)',
    hv, mbs.body._headingTexts());
}

console.log('ad-hoc: a real item numbered like the report family digit is NOT dropped by the main-meeting parent filter');

{
  const agenda = [item('1', 'Opening of the session'), item('2', 'IPR and antitrust reminder'), item('7', 'Real item seven'), item('8', 'Closing of the session')];
  const audio = runBuild(ADHOC('Audio'), agenda, {}); // Audio family digit is 7
  check('real item "7" is kept for an ad-hoc Audio report', audio.body._headingTexts().indexOf('7 Real item seven') !== -1, true);
  const main = runBuild({ REPORT_SUFFIX: 'Audio', AGENDA_TDOC: 'S4-260868' }, [item('7', 'Audio SWG'), item('7.1', 'Opening of the session'), item('7.2', 'IPR'), item('7.3', 'Topic'), item('7.4', 'Close of the session')], {});
  check('main meeting: the parent "7 Audio SWG" is still removed (unchanged)', main.body._headingTexts().indexOf('7 Audio SWG'), -1);
}

console.log('ad-hoc with no recognizable opening/IPR titles: documented generic fallback (first real top-level item; synthetic IPR child)');

{
  const agenda = [item('1', 'Welcome'), item('2', 'Topics'), item('3', 'Wrap-up')];
  const { body, sandbox } = runBuild(ADHOC('Audio'), agenda, {});
  const h = body._headingTexts();
  check('opening block hangs off the first real top-level item', h.indexOf('1.2 Registration of Documents') !== -1, true);
  check('IPR content is not lost: synthetic "1.5 IPR and antitrust reminder" with 1.5.1..1.5.4 children',
    ['1.5 IPR and antitrust reminder', '1.5.1 Introduction', '1.5.4 Consensus principles reminder'].map(t => h.indexOf(t) !== -1), [true, true, true]);
  const anchors = sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'adhoc' } }, agenda, '7.');
  check('anchor source is reported as the fallback', anchors.source, 'adhoc-fallback');
  check('the report-family prefix "7" never appears in any fallback anchor',
    JSON.stringify(anchors).indexOf('"7') === -1 && JSON.stringify(anchors).indexOf('7.') === -1, true);
}

// ================================================ 3. helper unit checks

console.log('getAdministrativeAgendaAnchors_() -- unit');

{
  const { sandbox } = loadCode();
  const main = sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'main' } }, [], '7.');
  check('main meeting: exact configured-prefix anchors (unchanged strings)', {
    o: main.openingSection, os: main.openingSubSection, r: main.registrationSection, re: main.reallocationSection,
    d: main.documentsSection, a: main.afterDocumentsSection, i: main.iprSection, src: main.source
  }, { o: '7.1', os: '7.1.1', r: '7.1.2', re: '7.1.3', d: '7.1.4', a: '7.1.5', i: '7.2', src: 'main-prefix' });

  const adhoc = sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'adhoc' } }, SYNTHETIC_ADHOC_AGENDA, '7.');
  check('ad-hoc: the prefix argument is ignored entirely', adhoc, sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'adhoc' } }, SYNTHETIC_ADHOC_AGENDA, '99.'));
  check('ad-hoc: opening/IPR resolved semantically from the real titles',
    [adhoc.openingSection, adhoc.iprSection, adhoc.source], ['1', '3', 'adhoc-semantic']);
  check('ad-hoc: synthetic children numbered after the real child (k=1)',
    [adhoc.openingSubSection, adhoc.registrationSection, adhoc.reallocationSection, adhoc.documentsSection], ['1.2', '1.3', '1.4', '1.5']);
  check('ad-hoc: the admin block is emitted after the last real descendant of the opening item', adhoc.openingEmitAfter, '1.1');
  check('ad-hoc: empty agenda yields no anchors (nothing to anchor on, never a main prefix)',
    sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'adhoc' } }, [], '7.').openingSection, null);

  check('isBeforeAdminBoundary_: main uses the unchanged prefix rule',
    [sandbox.isBeforeAdminBoundary_('7.2', main), sandbox.isBeforeAdminBoundary_('7.3', main), sandbox.isBeforeAdminBoundary_('1', main)], [true, false, false]);
  check('isBeforeAdminBoundary_: ad-hoc captures the opening/IPR anchor items only',
    [sandbox.isBeforeAdminBoundary_('1', adhoc), sandbox.isBeforeAdminBoundary_('3', adhoc), sandbox.isBeforeAdminBoundary_('1.1', adhoc), sandbox.isBeforeAdminBoundary_('4.1', adhoc)],
    [true, true, false, false]);
}

// ===================================== 4. Document Reallocations placement

console.log('ensureReallocationTable_() -- ad-hoc: Document Reallocations placed relative to the real opening/admin block');

{
  const built = runBuild(ADHOC('Audio'), SYNTHETIC_ADHOC_AGENDA, { '1': ['S4aA000010'] });
  const { sandbox, body } = built;
  built.docProps.setProperty('PARSED_AGENDA', JSON.stringify(SYNTHETIC_ADHOC_AGENDA));
  sandbox.ensureReallocationTable_();
  const h = body._headingTexts();
  check('"1.4 Document Reallocations" was created', h.indexOf('1.4 Document Reallocations') !== -1, true);
  check('it sits after the registration block (1.3) and before the Documents bucket (1.5)',
    h.indexOf('1.4 Document Reallocations') > h.indexOf('1.3 Registration of Documents') &&
    h.indexOf('1.4 Document Reallocations') < h.indexOf('1.5 Documents'), true);
  check('no report-family-prefixed reallocation section exists', h.filter(t => /^7\.1\.3/.test(t)), []);
  check('exactly one reallocation table exists',
    body._children.filter(c => c.getType() === 'TABLE' && sandbox.isReallocationTable_(c)).length, 1);
}

console.log('ensureReallocationTable_() -- ad-hoc with nothing to anchor on creates nothing (never a main-prefix section)');

{
  const { sandbox } = loadCode({ documentProperties: ADHOC('Audio') });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });
  sandbox.ensureReallocationTable_();
  check('no heading and no table were inserted', body._children.length, 0);
}

// ============================================= 5. main-meeting regression

console.log('MAIN meeting (Audio): administrative anchors and headings unchanged from the configured prefix');

{
  const agenda = [
    item('7', 'Audio SWG'), item('7.1', 'Opening of the session'), item('7.2', 'IPR and antitrust reminder'),
    item('7.3', 'Contributions'), item('7.4', 'Any other business'), item('7.5', 'Close of the session')
  ];
  const { body } = runBuild({ REPORT_SUFFIX: 'Audio', AGENDA_TDOC: 'S4-260868' }, agenda, { '7.3': ['S4-260001'], '7.2': ['S4-260002'] });
  check('main headings: exact sequence, all keyed to the configured "7" prefix', body._headingTexts(), [
    '7.1 Opening of the session',
    '7.1.2 Registration of Documents',
    '7.1.4 Documents',
    '7.2 IPR and antitrust reminder',
    '7.2.1 Introduction',
    '7.2.2 Call for IPRs',
    '7.2.3 Statement regarding competition law',
    '7.2.4 Consensus principles reminder',
    '7.3 Contributions',
    '7.4 Any other business',
    '7.5 Close of the session'
  ]);
}

console.log('MAIN meeting: ensureReallocationTable_() still targets the configured-prefix section');

{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(sandbox);
  body.appendParagraph('7.1 Opening').setHeading(sandbox.DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('7.1.2 Registration of Documents').setHeading(sandbox.DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('7.2 IPR').setHeading(sandbox.DocumentApp.ParagraphHeading.HEADING2);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });
  sandbox.ensureReallocationTable_();
  const h = body._headingTexts();
  check('main: "7.1.3 Document Reallocations" inserted before the 7.2 boundary', h, ['7.1 Opening', '7.1.2 Registration of Documents', '7.1.3 Document Reallocations', '7.2 IPR']);
}

// =================================== 6. invalid/non-real item numbers (task item 9)

console.log('isRealItemNumber_() -- invalid/non-real item numbers are rejected');

{
  const { sandbox } = loadCode();
  const adhoc = sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'adhoc' } }, SYNTHETIC_ADHOC_AGENDA, '7.');
  check('a real agenda number is recognized', sandbox.isRealItemNumber_(adhoc, '4.1'), true);
  check('a number that was never in the parsed agenda is rejected', sandbox.isRealItemNumber_(adhoc, '4.2'), false);
  check('a synthetic IPR-child-shaped number is rejected as real (it is not a real agenda item)', sandbox.isRealItemNumber_(adhoc, '3.9'), false);
  check('garbage input is rejected, not thrown on', [sandbox.isRealItemNumber_(adhoc, ''), sandbox.isRealItemNumber_(adhoc, null)], [false, false]);
  // A TDoc filed directly under the IPR item's number that is NOT a real
  // child is captured as admin overhead (synthetic-slot cleanup); one that
  // IS a real child of the IPR item is left alone.
  const iprHasRealChild = sandbox.getAdministrativeAgendaAnchors_({ meeting: { type: 'adhoc' } },
    [item('1', 'Opening of the session'), item('2', 'IPR and antitrust reminder'), item('2.1', 'Real IPR sub-point'), item('3', 'Topics')], '7.');
  check('a real child of the IPR item is NOT treated as before the admin boundary',
    sandbox.isBeforeAdminBoundary_('2.1', iprHasRealChild), false);
  check('a non-real number directly under the IPR item IS treated as before the admin boundary (synthetic slot)',
    sandbox.isBeforeAdminBoundary_('2.9', iprHasRealChild), true);
}

// =================================== 7. regression: earlier stages preserved

console.log('regression: BUGFIX-LEGACY-001 / LEGACY-UPGRADE-002 / LEGACY-UPGRADE-003 preserved');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_FOLDER: 'F', MEETING_NUMBER: '1' } });
  // LEGACY-0099 (as ADDON-008A1b): the build's first document access is now
  // reading the saved Document Reallocations (before anything is cleared),
  // so reaching the document means reaching either that read or clear().
  const reachDocument = () => { throw new Error('REACHED_BODY_CLEAR'); };
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ clear: reachDocument, getTables: reachDocument }) });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('BUGFIX-LEGACY-001: main meetings remain unblocked', thrown && thrown.message, 'REACHED_BODY_CLEAR');
}
{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('BUGFIX-LEGACY-001: incomplete ad-hoc build still fails before touching the agenda/TDoc logic', thrown && /Cannot build report yet/.test(thrown.message), true);
}
{
  const REAL_TOKEN = 'crv1_TESTONLY_not_a_real_token_0123456789';
  const { sandbox, scriptProps } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_TOKEN } });
  sandbox.saveConfigurationSettings({ apiTokenAction: 'keep', apiToken: '', showPreview: true });
  check('LEGACY-UPGRADE-002: token keep unchanged', scriptProps._store.REVIEWER_API_TOKEN, REAL_TOKEN);
  const { sandbox: s2, docProps } = loadCode();
  s2.saveConfigurationSettings({ meetingType: 'adhoc', meetingFolder: '', meetingNumber: '', meetingId: '', reportType: '', showPreview: true });
  check('LEGACY-UPGRADE-002: ad-hoc stale-default protection unchanged', [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID, docProps._store.REPORT_SUFFIX], [undefined, undefined, undefined, undefined]);
  // INTENTIONAL DIFFERENCE (CENTRAL ADDON-007B1, commit 0510461): the
  // Meeting Configuration dialog has no Email Collection Start Date field
  // and Save does not write EMAIL_START_DATE. Legacy keeps the field. The
  // collector honours a stored EMAIL_START_DATE either way (BUGFIX-LEGACY-003,
  // tests/legacy-parity-adhoc-email-collection.test.js).
  const { sandbox: s3, docProps: dp3 } = loadCode({ documentProperties: { EMAIL_START_DATE: '2026-09-01' } });
  s3.saveConfigurationSettings({ emailStartDate: '2026-09-28', showPreview: true });
  check('[template] Save neither writes nor erases EMAIL_START_DATE', dp3._store.EMAIL_START_DATE, '2026-09-01');
}
{
  const { sandbox } = loadCode();
  check('LEGACY-UPGRADE-003: family inference still present and functioning', JSON.parse(JSON.stringify(sandbox.inferReportFamily_({ meetingType: 'adhoc', meetingName: 'SA4 Audio SWG' }))).family, 'Audio');
  const { sandbox: s2, docProps } = loadCode();
  s2.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '1', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('LEGACY-UPGRADE-003: derived mailing list still never persisted as an override', docProps._store.MAILING_LIST, undefined);
}

// =================================== 8. no meeting-specific runtime constants

console.log('structural: no meeting-specific constants; exactly the expected functions changed');

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('LEGACY-UPGRADE-004 stage scope (3 checks)'); else {
  const fs = require('fs');
  const { execFileSync } = require('child_process');
  const path = require('path');
  const { CODE_JS_PATH } = require('./helpers/legacy-load-code.js');
  const PRE_004_COMMIT = 'b97b46c';
  // LEGACY-UPGRADE-004's own commit (not the live/current file) -- keeps
  // "exactly these new functions" a fact about the 004 stage itself, valid
  // even after LEGACY-UPGRADE-006+ add their own functions on top (same
  // pattern LEGACY-UPGRADE-002/003's own structural checks use).
  const POST_004_COMMIT = '83de6cb';
  const REPO_ROOT = path.join(__dirname, '..');
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const pre004Source = execFileSync('git', ['show', `${PRE_004_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  const post004Source = execFileSync('git', ['show', `${POST_004_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  const names = (s) => new Set((s.match(/^function ([A-Za-z0-9_]+)\(/gm) || []).map((m) => m.replace(/^function |\($/g, '')));
  const before = names(pre004Source);
  const after = names(src);
  const removed = [...before].filter((n) => !after.has(n));
  check('no function present as of LEGACY-UPGRADE-003 has since been removed', removed, []);
  const addedByThisStage = [...names(post004Source)].filter((n) => !before.has(n));
  check('LEGACY-UPGRADE-004 itself added exactly its own new functions', addedByThisStage.sort(),
    ['deriveAdminAnchorsFromBuiltDocument_', 'getAdministrativeAgendaAnchors_', 'isBeforeAdminBoundary_', 'isRealItemNumber_'].sort());

  const anchorBlockStart = src.indexOf('function getAdministrativeAgendaAnchors_(');
  const anchorBlockEnd = src.indexOf('function appendTdocDetailTable_(') > anchorBlockStart ? src.indexOf('function appendTdocDetailTable_(') : src.length;
  const anchorBlock = src.slice(anchorBlockStart, anchorBlockEnd) + src.slice(src.indexOf('function deriveAdminAnchorsFromBuiltDocument_('), src.indexOf('function ensureReallocationTable_('));
  check('no 85916/86178/FS_6G_MED/ULBC constant in the anchoring code added by this stage',
    /85916|86178|FS_6G_MED|ULBC/i.test(anchorBlock), false);
}

// ================================================================ summary

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
