/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-0099-reallocation-suppression.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-0099-reallocation-suppression.test.js
 */

/**
 * LEGACY-0099 -- a Document Reallocation to Removed / withdrawn / N/A keeps
 * a TDoc out of the report across Continuous Update AND Full Build.
 *
 * Diagnosis (commit a1e2312, the earlier version of this file): S4aP260099
 * kept reappearing in the 6G (86178) report because
 *   - downloadAndGroupTdocs_() rewrote a "Removed" TDoc's agenda item to the
 *     literal word "Removed" instead of dropping it; the ad-hoc {mode:'all'}
 *     selector accepted that, so every Continuous Update re-inserted it
 *     (under a stray "Removed" heading, then moved back below S4aP260074 by
 *     rearrangeRevisionTables_());
 *   - buildSkeletonWithTdocTables() cleared the document before reading the
 *     reallocation table, so a Full Build lost every reallocation;
 *   - Apply Document Reallocations' move path did
 *     body.removeChild(table).asTable() (Body.removeChild() returns the
 *     body: "BODY_SECTION can't be cast to TABLE.").
 * Fixed by porting the central add-on's ADDON-008A1/008A1b semantics.
 *
 * Source rows are the real upstream rows of the 86178 TDoc list
 * (TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx, fetched 2026-09-30):
 *   S4aP260074  Draft TR 26.870 v0.6.0  "revised", Revised to S4aP260099
 *   S4aP260099  Draft TR 26.870 v0.6.2  "reserved", Is revision of S4aP260074
 * plus S4aP260088 as an unrelated control. The agenda is 86178's real one
 * (S4aP260098). The REAL downloadAndGroupTdocs_(), continuousUpdateCore_(),
 * buildSkeletonWithTdocTables() and applyDocumentReallocations() run; only
 * network / Drive / Sheets / template access is faked.
 *
 * Run: node tests/legacy-0099-reallocation-suppression.test.js
 */

const { loadCode } = require('./helpers/legacy-load-code.js');

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
function attempt(fn) {
  try { return { value: fn(), error: null }; } catch (e) { return { value: null, error: e.message }; }
}

// ------------------------------------------------------------ source data

const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Is revision of', 'Revised to'];
const ROW_0074 = ['S4aP260074', 'Draft TR 26.870 FS_6G_MED v0.6.0', 'VODAFONE Group Plc', 'Elmira Ramazanirend', 'draft TR', 'Agreement', '5.4', 'FS_6G_MED - General and working documents', 'revised', 'S4-261521', 'S4aP260099'];
const ROW_0099 = ['S4aP260099', 'Draft TR 26.870 FS_6G_MED v0.6.2', 'VODAFONE Group Plc', 'Elmira Ramazanirend', 'draft TR', 'Agreement', '5.4', 'FS_6G_MED - General and working documents', 'reserved', 'S4aP260074', ''];
const ROW_0088 = ['S4aP260088', 'Proposed Planning for AHG and SA4#138', 'Qualcomm Germany', '', 'discussion', '-', '5.4', 'FS_6G_MED - General and working documents', 'available', '', ''];
const SHEET_86178 = [HEADER, ROW_0074, ROW_0099, ROW_0088];

const PROPS_86178 = {
  MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) on FS_6G_MED', REPORT_SUFFIX: '6G',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/',
  TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx',
  AGENDA_TDOC: 'S4aP260098', AGENDA_ITEM_PREFIX: '5.'
};

const item = (number, title) => ({ number, title, level: number.split('.').length, heading: 'NORMAL', text: '' });
const AGENDA_86178 = [
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

// ------------------------------------------------------------ fake document

// Behaves like Apps Script where it matters here: Body.removeChild() returns
// the BODY (whose asTable() throws the real cast error), tables know their
// parent, and copy() is a detached deep copy. Formatting setters are no-ops.
function makeFakeDocumentBody(sandbox) {
  const PARAGRAPH = sandbox.DocumentApp.ElementType.PARAGRAPH;
  const TABLE = sandbox.DocumentApp.ElementType.TABLE;
  const NORMAL = sandbox.DocumentApp.ParagraphHeading.NORMAL;
  const children = [];
  const noop = function () { return this; };
  const styling = { setBold: noop, setItalic: noop, setUnderline: noop, setForegroundColor: noop, setBackgroundColor: noop,
    setFontSize: noop, setFontFamily: noop, setAttributes: noop, setLinkUrl: noop, setSpacingBefore: noop, setSpacingAfter: noop,
    setLineSpacing: noop, setIndentStart: noop, setIndentFirstLine: noop, setAlignment: noop, setGlyphType: noop,
    setPaddingTop: noop, setPaddingBottom: noop, setPaddingLeft: noop, setPaddingRight: noop, setWidth: noop,
    setColumnWidth: noop, setBorderWidth: noop, setBorderColor: noop, setMinimumHeight: noop, setVerticalAlignment: noop };

  function makeCell(text) {
    const c = Object.assign({}, styling, {
      _t: String(text),
      getText: () => c._t,
      setText: (t) => { c._t = String(t); return c; },
      editAsText: () => Object.assign({}, styling, { getText: () => c._t, getLinkUrl: () => null }),
      getLinkUrl: () => null,
      getNumChildren: () => 0
    });
    return c;
  }
  function makeParagraph(text, heading) {
    const p = Object.assign({}, styling, {
      _text: String(text), _heading: heading || NORMAL,
      getType: () => PARAGRAPH, asParagraph: () => p,
      getText: () => p._text,
      setText: (t) => { p._text = String(t); return p; },
      getHeading: () => p._heading,
      setHeading: (h) => { p._heading = h; return p; },
      editAsText: () => Object.assign({}, styling, { getText: () => p._text }),
      getParent: () => body
    });
    return p;
  }
  function makeTable(rowsData) {
    const rows = [];
    function makeRow(cellTexts) {
      const cells = cellTexts.map(makeCell);
      const r = Object.assign({}, styling, {
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { const c = makeCell(t === undefined ? '' : t); cells.push(c); return c; }
      });
      return r;
    }
    (rowsData || []).forEach((r) => rows.push(makeRow(r)));
    const t = Object.assign({}, styling, {
      getType: () => TABLE, asTable: () => t,
      getNumRows: () => rows.length, getRow: (i) => rows[i],
      getCell: (r, c) => rows[r].getCell(c),
      appendTableRow: () => { const r = makeRow([]); rows.push(r); return r; },
      removeRow: (i) => { rows.splice(i, 1); },
      copy: () => makeTable(rows.map((r) => { const out = []; for (let i = 0; i < r.getNumCells(); i++) out.push(r.getCell(i).getText()); return out; })),
      getParent: () => (children.indexOf(t) >= 0 ? body : null),
      removeFromParent: () => { const i = children.indexOf(t); if (i >= 0) children.splice(i, 1); return t; }
    });
    return t;
  }

  const body = {
    _children: children,
    getType: () => 'BODY_SECTION',
    asTable: () => { throw new Error("BODY_SECTION can't be cast to TABLE."); },
    clear: () => { children.length = 0; return body; },
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
    getChildIndex: (c) => children.indexOf(c),
    getTables: () => children.filter((c) => c.getType() === TABLE),
    getParagraphs: () => children.filter((c) => c.getType() === PARAGRAPH),
    appendParagraph: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendListItem: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendTable: (data) => { const t = makeTable(data || []); children.push(t); return t; },
    insertParagraph: (idx, text) => { const p = makeParagraph(text); children.splice(idx, 0, p); return p; },
    insertTable: (idx, data) => {
      const t = (data && data.getType && data.getType() === TABLE) ? data : makeTable(data || []);
      children.splice(idx, 0, t); return t;
    },
    // Real Apps Script: returns the Body, not the removed element.
    removeChild: (c) => { const i = children.indexOf(c); if (i >= 0) children.splice(i, 1); return body; },
    _heading: (text, h) => { const p = makeParagraph(text, h); children.push(p); return p; },
    _table: (data) => { const t = makeTable(data); children.push(t); return t; }
  };
  return body;
}

function tdocTable(id, agenda) {
  return [['TDoc', id], ['Title', ''], ['Source', ''], ['Contact', ''], ['Agenda Item', agenda], ['Type/For', ''],
    ['E-mail Discussion', ''], ['Revisions', ''], ['Minutes', ''], ['Disposition', ''], ['Status', '']];
}

// ------------------------------------------------------------ harness

function makeSandbox(props, sheet) {
  const loaded = loadCode({ documentProperties: props });
  const sandbox = loaded.sandbox;
  const body = makeFakeDocumentBody(sandbox);
  const alerts = [];
  const logs = [];
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  sandbox.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(sandbox) });
  sandbox.DocumentApp.getUi = () => ({ alert: (t, m) => { alerts.push([t, m]); return 'YES'; }, Button: { YES: 'YES' }, ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' } });
  sandbox.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  sandbox.UrlFetchApp.fetch = () => ({ getResponseCode: () => 200, getBlob: () => ({ setName() { return this; } }) });
  sandbox.DriveApp.createFile = () => ({ setTrashed() {} });
  sandbox.SpreadsheetApp.open = () => ({ getSheets: () => [{ getDataRange: () => ({
    getValues: () => sheet.map((r) => r.slice()),
    getRichTextValues: () => sheet.map((r) => r.map(() => null))
  }) }] });
  sandbox.downloadMeetingAgenda_ = () => AGENDA_86178.map((i) => Object.assign({}, i));
  sandbox.setDocumentTitleFromTemplate_ = () => {};
  sandbox.findHeading_ = () => null;
  sandbox.styleStatusCell_ = () => {};
  sandbox.removeRowHeightAndSpacing = () => {};
  sandbox.getFetchAbstractsSetting_ = () => false;
  sandbox.collectorUpdate_ = () => {};
  sandbox.Logger = { log: (m) => logs.push(String(m)) };
  return { sandbox, body, alerts, logs, docProps: loaded.docProps };
}

// A built-looking 86178 report: 5.1.3 Document Reallocations (empty) and the
// three TDoc tables under 5.4, 0099 directly below its parent 0074.
function setupReport() {
  const env = makeSandbox(PROPS_86178, SHEET_86178);
  const { sandbox, body } = env;
  const H2 = sandbox.DocumentApp.ParagraphHeading.HEADING2;
  const H3 = sandbox.DocumentApp.ParagraphHeading.HEADING3;
  body._heading('5.1 Opening of the session and registration of documents', H2);
  body._heading('5.1.2 Registration of Documents', H3);
  body._heading('5.1.3 Document Reallocations', H3);
  body._table([['TDoc', 'Original Agenda', 'New Agenda', 'Reason']]);
  body._heading('5.4 FS_6G_MED - General and working documents', H2);
  body._table(tdocTable('S4aP260074', '5.4'));
  body._table(tdocTable('S4aP260099', '5.4'));
  body._table(tdocTable('S4aP260088', '5.4'));
  body._heading('5.5 FS_6G_MED - WT#1: Media Delivery Architecture', H2);
  body._heading('5.10 FS_6G_MED - Other issues', H2);
  body._heading('5.11 Close of the session', H2);
  return env;
}

const isTable = (c) => c.getType() === 'TABLE';
const cellText = (t, r, k) => { try { return t.getRow(r).getCell(k).getText(); } catch (e) { return null; } };
function tablesFor(env, id) {
  return env.body.getTables().filter((t) => env.sandbox.isTDocTable_(t) && env.sandbox.safeCellText_(t, 0, 1).trim() === id);
}
function rowValue(t, label) { for (let r = 0; r < t.getNumRows(); r++) if (cellText(t, r, 0) === label) return cellText(t, r, 1); return null; }
function sectionOf(env, el) {
  const NORMAL = env.sandbox.DocumentApp.ParagraphHeading.NORMAL;
  for (let i = env.body.getChildIndex(el) - 1; i >= 0; i--) {
    const c = env.body.getChild(i);
    if (c.getType() === 'PARAGRAPH' && c.getHeading() !== NORMAL) return c.getText().split(' ')[0];
  }
  return null;
}
const ids = (env) => env.body.getTables().filter((t) => env.sandbox.isTDocTable_(t) && /^S4/.test(cellText(t, 0, 1))).map((t) => cellText(t, 0, 1));
function strayRemovalHeadings(env) {
  const NORMAL = env.sandbox.DocumentApp.ParagraphHeading.NORMAL;
  return env.body._children.filter((c) => c.getType() === 'PARAGRAPH' && c.getHeading() !== NORMAL &&
    /^(removed|withdrawn|n\/a)\b/i.test(c.getText().trim())).map((c) => c.getText());
}
function reallocationRows(env) {
  const t = env.body.getTables().find((x) => env.sandbox.isReallocationTable_(x));
  return t ? Array.from({ length: t.getNumRows() - 1 }, (_, i) => [0, 1, 2, 3].map((k) => cellText(t, i + 1, k))) : null;
}
function addReallocation(env, tdoc, original, target, reason) {
  env.sandbox.saveReallocation(tdoc, original, target, reason || 'test');
}
function removeReallocation(env, tdoc) {
  const t = env.body.getTables().find((x) => env.sandbox.isReallocationTable_(x));
  for (let r = 1; r < t.getNumRows(); r++) if (cellText(t, r, 0) === tdoc) { t.removeRow(r); return; }
}
function snapshot(env) {
  return JSON.stringify(env.body._children.map((c) => isTable(c)
    ? ['T', Array.from({ length: c.getNumRows() }, (_, r) => Array.from({ length: c.getRow(r).getNumCells() }, (_, k) => cellText(c, r, k)))]
    : ['P', c.getText(), c.getHeading()]));
}
function runCU(env) {
  env.logs.length = 0;
  const r = attempt(() => env.sandbox.continuousUpdateCore_());
  return { error: r.error, internalErrors: env.logs.filter((l) => /^ERROR:/.test(l)) };
}
const cuOk = (r) => r.error === null && r.internalErrors.length === 0;

// ============================================================ A. baseline

console.log('A. baseline source (0074 revised to 0099, 0099 reserved), no reallocation: both imported normally');
{
  const env = makeSandbox(PROPS_86178, SHEET_86178);
  const build = attempt(() => env.sandbox.buildSkeletonWithTdocTables());
  check('Full Build ok', build.error, null);
  check('both documents imported (with the control), each exactly once', ids(env).filter((i) => /00(74|99|88)$/.test(i)).sort(), ['S4aP260074', 'S4aP260088', 'S4aP260099']);
  check('0099 sits directly below 0074 under 5.4', [sectionOf(env, tablesFor(env, 'S4aP260099')[0]),
    env.body.getChildIndex(tablesFor(env, 'S4aP260099')[0]) === env.body.getChildIndex(tablesFor(env, 'S4aP260074')[0]) + 1], ['5.4', true]);
  check('0074 Disposition records the revision', rowValue(tablesFor(env, 'S4aP260074')[0], 'Disposition'), 'Revised to S4aP260099');
  check('no reallocation table is created when there are none', reallocationRows(env), null);

  const cuEnv = setupReport();
  check('Continuous Update on a built report: ok and nothing inserted', [cuOk(runCU(cuEnv)), cuEnv.logs.some((l) => /^Inserted /.test(l))], [true, false]);
}

// ============================================================ B. manual deletion only

console.log('\nB. manual deletion only (no reallocation entry): Continuous Update re-imports 0099 -- intentional');
{
  const env = setupReport();
  env.body.removeChild(tablesFor(env, 'S4aP260099')[0]);
  check('Continuous Update ok', cuOk(runCU(env)), true);
  check('0099 is back once, directly below 0074 under 5.4',
    [tablesFor(env, 'S4aP260099').length, sectionOf(env, tablesFor(env, 'S4aP260099')[0])], [1, '5.4']);
}

// ============================================================ C/D/E. removal targets in Continuous Update

[['C', 'Removed'], ['D', 'Withdrawn'], ['E', 'N/A']].forEach(([label, target]) => {
  console.log(`\n${label}. S4aP260099 -> "${target}", Apply, then Continuous Update`);
  const env = setupReport();
  addReallocation(env, 'S4aP260099', '5.4', target, 'excluded from this report');
  const applied = attempt(() => env.sandbox.applyDocumentReallocations());
  check('Apply ok and removes the 0099 table', [applied.error, tablesFor(env, 'S4aP260099').length], [null, 0]);
  check('Continuous Update ok', cuOk(runCU(env)), true);
  check('0099 is NOT re-imported', tablesFor(env, 'S4aP260099').length, 0);
  check('nothing was inserted', env.logs.some((l) => /^Inserted /.test(l)), false);
  check(`H: no "${target}" heading anywhere`, strayRemovalHeadings(env), []);
  check('the reallocation entry is untouched', reallocationRows(env), [['S4aP260099', '5.4', target, 'excluded from this report']]);
  check('0074 and the control are untouched in place', [tablesFor(env, 'S4aP260074').length, tablesFor(env, 'S4aP260088').length,
    sectionOf(env, tablesFor(env, 'S4aP260074')[0])], [1, 1, '5.4']);
  check('a second Continuous Update still does not import it', [cuOk(runCU(env)), tablesFor(env, 'S4aP260099').length], [true, 0]);
});

// ============================================================ F. accepted normalization

console.log('\nF. case / surrounding-space variants accepted by the CENTRAL normalization (trim + lower-case)');
{
  const variants = ['removed', 'REMOVED', ' Removed ', 'withdrawn', 'WITHDRAWN', ' Withdrawn', 'n/a', 'N/a', 'N/A '];
  variants.forEach((v) => {
    const env = setupReport();
    addReallocation(env, 'S4aP260099', '5.4', v);
    env.body.removeChild(tablesFor(env, 'S4aP260099')[0]);
    check(`"${v}": Continuous Update keeps 0099 out, no stray heading`, [cuOk(runCU(env)), tablesFor(env, 'S4aP260099').length, strayRemovalHeadings(env)], [true, 0, []]);
  });
  const { sandbox } = makeSandbox(PROPS_86178, SHEET_86178);
  check('interpretReallocationTarget_ matches CENTRAL exactly',
    ['2.7', ' 2.7 ', 'Withdrawn', 'n/a', 'REMOVED', ' N/A ', 'later', '', 'N / A', 'remove'].map((v) => sandbox.interpretReallocationTarget_(v).kind),
    ['move', 'move', 'remove', 'remove', 'remove', 'remove', 'invalid', 'invalid', 'invalid', 'invalid']);
}

// ============================================================ G. lifting the suppression

console.log('\nG. deleting the reallocation entry lifts the suppression: the next update imports 0099 again');
{
  const env = setupReport();
  addReallocation(env, 'S4aP260099', '5.4', 'Removed');
  env.sandbox.applyDocumentReallocations();
  runCU(env);
  check('suppressed while the entry exists', tablesFor(env, 'S4aP260099').length, 0);
  removeReallocation(env, 'S4aP260099');
  check('Continuous Update ok', cuOk(runCU(env)), true);
  check('0099 is imported again, once, directly below 0074 under 5.4',
    [tablesFor(env, 'S4aP260099').length, sectionOf(env, tablesFor(env, 'S4aP260099')[0]),
      env.body.getChildIndex(tablesFor(env, 'S4aP260099')[0]) === env.body.getChildIndex(tablesFor(env, 'S4aP260074')[0]) + 1],
    [1, '5.4', true]);

  const changed = setupReport();
  addReallocation(changed, 'S4aP260099', '5.4', 'Removed');
  changed.sandbox.applyDocumentReallocations();
  runCU(changed);
  addReallocation(changed, 'S4aP260099', '5.4', '5.10'); // the same row, changed to a destination
  check('changing the entry to a destination imports 0099 there instead',
    [cuOk(runCU(changed)), tablesFor(changed, 'S4aP260099').map((t) => rowValue(t, 'Agenda Item'))], [true, ['5.10']]);
}

// ============================================================ I. Full Build

console.log('\nI. Full Build with S4aP260099 -> Removed');
{
  const env = makeSandbox(PROPS_86178, SHEET_86178);
  check('initial build ok', attempt(() => env.sandbox.buildSkeletonWithTdocTables()).error, null);
  addReallocation(env, 'S4aP260099', '5.4', 'Removed', 'excluded from this report');
  check('the reallocation table is created under the real 5.1 admin block', sectionOf(env, env.body.getTables().find((t) => env.sandbox.isReallocationTable_(t))), '5.1.3');
  const rebuilt = attempt(() => env.sandbox.buildSkeletonWithTdocTables());
  check('rebuild ok', rebuilt.error, null);
  check('0099 absent after the rebuild', tablesFor(env, 'S4aP260099').length, 0);
  check('0074 and the control present once each, under 5.4', ['S4aP260074', 'S4aP260088'].map((id) => [tablesFor(env, id).length, sectionOf(env, tablesFor(env, id)[0])]), [[1, '5.4'], [1, '5.4']]);
  check('0099 is not listed in Registration of Documents either',
    env.body.getTables().filter((t) => cellText(t, 0, 1) === 'Title' && cellText(t, 0, 3) === 'Agenda Item')
      .some((t) => Array.from({ length: t.getNumRows() }, (_, r) => cellText(t, r, 0)).indexOf('S4aP260099') !== -1), false);
  check('no "Removed" heading', strayRemovalHeadings(env), []);
  check('the Document Reallocations table survives with the exact entry', reallocationRows(env), [['S4aP260099', '5.4', 'Removed', 'excluded from this report']]);
  check('... still under 5.1.3', sectionOf(env, env.body.getTables().find((t) => env.sandbox.isReallocationTable_(t))), '5.1.3');
  const once = snapshot(env);
  check('a second Full Build gives the identical report', [attempt(() => env.sandbox.buildSkeletonWithTdocTables()).error, snapshot(env) === once], [null, true]);
  check('Apply afterwards: no error, nothing changes', [attempt(() => env.sandbox.applyDocumentReallocations()).error, snapshot(env) === once], [null, true]);
  check('Continuous Update afterwards: ok, 0099 still absent', [cuOk(runCU(env)), tablesFor(env, 'S4aP260099').length], [true, 0]);
}

console.log('\nI2. Full Build refuses an unusable saved destination BEFORE clearing (the report is unchanged)');
{
  const env = setupReport();
  addReallocation(env, 'S4aP260088', '5.4', '5.99');
  const before = snapshot(env);
  const r = attempt(() => env.sandbox.buildSkeletonWithTdocTables());
  check('refused, naming TDoc / source / destination / reason',
    /report was not changed[\s\S]*S4aP260088 \(5\.4 → 5\.99\): this meeting's agenda has no item 5\.99\./.test(r.error || ''), true);
  check('document untouched', snapshot(env), before);
  const word = setupReport();
  addReallocation(word, 'S4aP260099', '5.4', 'N / A');
  const before2 = snapshot(word);
  const r2 = attempt(() => word.sandbox.buildSkeletonWithTdocTables());
  check('a value outside removed/withdrawn/n/a and not a number is refused, document untouched',
    [/"N \/ A" is not an agenda item number or removed\/withdrawn\/n\/a\./.test(r2.error || ''), snapshot(word) === before2], [true, true]);
}

// ============================================================ J. revision rearrangement

console.log('\nJ. revision re-arrangement cannot resurrect a suppressed 0099');
{
  const env = setupReport();
  env.docProps.setProperty('REVISION_MAP', JSON.stringify({ S4aP260074: 'S4aP260099' })); // as stored by earlier live runs
  addReallocation(env, 'S4aP260099', '5.4', 'Removed');
  env.sandbox.applyDocumentReallocations();
  check('Continuous Update (with the stored revision map) ok, 0099 absent', [cuOk(runCU(env)), tablesFor(env, 'S4aP260099').length], [true, 0]);
  const stats = env.sandbox.rearrangeRevisionTables_();
  check('the menu re-arrangement worker: 0099 is a missing child, nothing moved or created',
    [stats.missingChild, stats.moved, tablesFor(env, 'S4aP260099').length], [1, 0, 0]);
  check('0074 stays valid: in place under 5.4, Disposition keeps the source fact "Revised to" 0099 (CENTRAL behavior)',
    [sectionOf(env, tablesFor(env, 'S4aP260074')[0]), /^Revised to S4aP260099$/i.test(rowValue(tablesFor(env, 'S4aP260074')[0], 'Disposition'))], ['5.4', true]);
  const revisions = attempt(() => env.sandbox.insertRevisedDocTablesAfter_(env.body, tablesFor(env, 'S4aP260074')[0], [{ text: 'S4aP260074_rev1.docx', link: 'x' }], {}));
  check('revision-draft insertion for 0074 does not create 0099', [revisions.error, tablesFor(env, 'S4aP260099').length], [null, 0]);
}

// ============================================================ K. main meetings

console.log('\nK. main-meeting reallocation behavior unchanged');
{
  const MAIN_SHEET = [HEADER,
    ['S4-260101', 'Main A', 'X', '', 'pCR', 'Agreement', '7.3', '', 'available', '', ''],
    ['S4-260102', 'Main B', 'X', '', 'pCR', 'Agreement', '7.4', '', 'available', '', ''],
    ['S4-260103', 'Main C', 'X', '', 'pCR', 'Agreement', '7.4', '', 'available', '', ''],
    ['S4-260104', 'Main D', 'X', '', 'pCR', 'Agreement', '7.5', '', 'available', '', '']];
  const env = makeSandbox({ MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137', MEETING_NUMBER: '137', REPORT_SUFFIX: 'Audio', AGENDA_ITEM_PREFIX: '7.',
    TDOC_LIST_URL: 'https://example.invalid/list.xlsx' }, MAIN_SHEET);
  const snap = {
    'S4-260101': { original: '7.3', new: '7.5', reason: '' },
    'S4-260102': { original: '7.4', new: 'removed', reason: '' },
    'S4-260103': { original: '7.4', new: '9.3', reason: '' }
  };
  // TEMPLATE-002A harness adaptation: in sa4-report the reallocation
  // snapshot is the third parameter (the second is the execution context,
  // ADDON-004); Legacy's signature is (cfg, snapshot).
  const groups = env.sandbox.downloadAndGroupTdocs_(env.sandbox.getReportConfig_(), undefined, snap);
  check('move within the SWG, removal, move to another SWG, untouched',
    Object.keys(groups).sort().map((k) => [k, groups[k].tdocs.map((d) => d.row[0])]),
    [['7.5', ['S4-260101', 'S4-260104']]]);
  check('main: another SWG\'s item is not a build problem', env.sandbox.findReallocationBuildProblems_({ 'S4-260001': { original: '7.3', new: '9.3' } }, null, null), []);
  check('main: a non-number is still refused by a build', env.sandbox.findReallocationBuildProblems_({ 'S4-260001': { original: '7.3', new: 'later' } }, null, null).length, 1);
}

// ============================================================ L. normal destination reallocation

console.log('\nL. a normal destination reallocation (S4aP260088 5.4 -> 5.10) still works everywhere');
{
  const env = setupReport();
  addReallocation(env, 'S4aP260088', '5.4', '5.10');
  const applied = attempt(() => env.sandbox.applyDocumentReallocations());
  check('Apply ok: moved under 5.10, Agenda Item 5.10, exactly one table',
    [applied.error, tablesFor(env, 'S4aP260088').map((t) => [sectionOf(env, t), rowValue(t, 'Agenda Item')])], [null, [['5.10', '5.10']]]);
  check('Continuous Update leaves it there', [cuOk(runCU(env)), tablesFor(env, 'S4aP260088').map((t) => sectionOf(env, t))], [true, ['5.10']]);

  const missing = setupReport();
  addReallocation(missing, 'S4aP260088', '5.4', '5.10');
  missing.body.removeChild(tablesFor(missing, 'S4aP260088')[0]);
  check('Continuous Update imports a missing reallocated TDoc at its destination',
    [cuOk(runCU(missing)), tablesFor(missing, 'S4aP260088').map((t) => [sectionOf(missing, t), rowValue(t, 'Agenda Item')])], [true, [['5.10', '5.10']]]);

  const built = makeSandbox(PROPS_86178, SHEET_86178);
  built.sandbox.buildSkeletonWithTdocTables();
  addReallocation(built, 'S4aP260088', '5.4', '5.10');
  check('Full Build places it under 5.10 and keeps the entry',
    [attempt(() => built.sandbox.buildSkeletonWithTdocTables()).error, tablesFor(built, 'S4aP260088').map((t) => [sectionOf(built, t), rowValue(t, 'Agenda Item')]), reallocationRows(built)],
    [null, [['5.10', '5.10']], [['S4aP260088', '5.4', '5.10', 'test']]]);
}

// ============================================================ M/N. Apply move path

console.log('\nM. Apply move path: no BODY_SECTION -> TABLE cast, no duplicate; plan validated before any change');
{
  const env = setupReport();
  check('the fake body reproduces the real API (removeChild returns the body, which is not a table)',
    attempt(() => env.body.removeChild(env.body.getChild(0)).asTable()).error, "BODY_SECTION can't be cast to TABLE.");
  const env2 = setupReport();
  addReallocation(env2, 'S4aP260088', '5.4', '5.5');
  const r = attempt(() => env2.sandbox.applyDocumentReallocations());
  check('move succeeds without the cast error', r.error, null);
  check('exactly one 0088 table, under 5.5', tablesFor(env2, 'S4aP260088').map((t) => sectionOf(env2, t)), ['5.5']);

  const bad = setupReport();
  addReallocation(bad, 'S4aP260088', '5.4', '5.10');
  addReallocation(bad, 'S4aP260074', '5.4', '5.8'); // the report has no 5.8 heading
  const before = snapshot(bad);
  const r2 = attempt(() => bad.sandbox.applyDocumentReallocations());
  check('one unresolvable entry aborts the whole run with a clear message',
    /report was not changed[\s\S]*S4aP260074 \(5\.4 → 5\.8\): the report has no agenda item 5\.8 heading\./.test(r2.error || ''), true);
  check('... and nothing at all was changed (not even the valid 0088 move)', snapshot(bad), before);
  check('no heading is appended for a missing destination', bad.body._children.some((c) => c.getType() === 'PARAGRAPH' && /^5\.8\b/.test(c.getText())), false);
}

console.log('\nN. Apply twice is idempotent');
{
  const env = setupReport();
  addReallocation(env, 'S4aP260088', '5.4', '5.10');
  addReallocation(env, 'S4aP260099', '5.4', 'Removed');
  env.sandbox.applyDocumentReallocations();
  const once = snapshot(env);
  check('second Apply: no error, document identical', [attempt(() => env.sandbox.applyDocumentReallocations()).error, snapshot(env) === once], [null, true]);
  const last = env.alerts[env.alerts.length - 1][1];
  check('second Apply reports nothing moved/removed and 0099 as not in the report',
    [/Moved tables: 0\b/.test(last), /Removed tables: 0\b/.test(last), /Not in this report \(skipped\): S4aP260099/.test(last)], [true, true, true]);
}

// ============================================================ legacy web-sheet import

console.log('\nlegacy "Build Initial Report" import path applies the same interpretation');
{
  const env = setupReport();
  env.body.removeChild(tablesFor(env, 'S4aP260099')[0]);
  addReallocation(env, 'S4aP260099', '5.4', 'Removed');
  const sheet = { getDataRange: () => ({ getValues: () => SHEET_86178.map((r) => r.slice()), getRichTextValues: () => SHEET_86178.map((r) => r.map(() => null)) }) };
  check('processWebDownloadedSheet_ skips the removed TDoc', [attempt(() => env.sandbox.processWebDownloadedSheet_(sheet)).error, tablesFor(env, 'S4aP260099').length], [null, 0]);
  check('... and logs why', env.logs.some((l) => /^Skipping S4aP260099: reallocated to "Removed"/.test(l)), true);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
