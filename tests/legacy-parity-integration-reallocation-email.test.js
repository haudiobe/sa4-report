/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-integration-reallocation-email.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-integration-reallocation-email.test.js
 */

/**
 * LEGACY integration -- Document Reallocation suppression (LEGACY-0099) and
 * ad-hoc discussion e-mail collection (BUGFIX-LEGACY-003) in ONE runtime.
 *
 * Each fix has its own suite, but each stubs the other half of
 * continuousUpdate(): legacy-0099-reallocation-suppression stubs
 * collectorUpdate_(), legacy-adhoc-email-collection stubs
 * downloadAndGroupTdocs_(). Here both run for real, in the same
 * continuousUpdate() call, against the same document:
 *   A. S4aP260099 reallocated to "Removed" stays out of the report;
 *   B. S4aP260091 stays a normal, eligible TDoc table;
 *   C. a reply to the Legacy exporter's subject naming S4aP260091 is
 *      associated with S4aP260091 only;
 *   D. e-mail collection neither recreates nor stores anything for the
 *      suppressed S4aP260099 -- even for a reply that names it;
 *   E. continuousUpdate() runs the corrected collector while its TDoc
 *      grouping honours the reallocation.
 *
 * Real production code: continuousUpdate(), downloadAndGroupTdocs_(),
 * rearrangeRevisionTables_(), collectorUpdate_(), getCollectorConfig_(),
 * collectHybridListservMessages_(), checkRSSFeed_(). Only the platform
 * (network, Drive, Sheets, XmlService, document body) is faked.
 *
 * Run: node tests/legacy-integration-reallocation-email.test.js
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

// ------------------------------------------------------------ source data

const LIST_6G = '3GPP_TSG_SA4_FS_6G_MED';
const NOW = '2026-10-01T09:00:00Z';
const MSG_DATE = 'Thu, 01 Oct 2026 08:00:00 GMT';

const PROPS_86178 = {
  MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) on FS_6G_MED', MEETING_ID: '86178', REPORT_SUFFIX: '6G',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/',
  TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx',
  AGENDA_TDOC: 'S4aP260098', AGENDA_ITEM_PREFIX: '5.',
  MAILING_LIST: LIST_6G, EMAIL_START_DATE: '2026-08-21'
};

const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Is revision of', 'Revised to'];
const TITLE_0091 = '[FS_6G_MED] Clarification of AI-native Traffic Characteristics in Clause 6.3.3';
const SHEET = [HEADER,
  ['S4aP260074', 'Draft TR 26.870 FS_6G_MED v0.6.0', 'VODAFONE Group Plc', 'Elmira Ramazanirend', 'draft TR', 'Agreement', '5.4', 'FS_6G_MED - General and working documents', 'revised', 'S4-261521', 'S4aP260099'],
  ['S4aP260099', 'Draft TR 26.870 FS_6G_MED v0.6.2', 'VODAFONE Group Plc', 'Elmira Ramazanirend', 'draft TR', 'Agreement', '5.4', 'FS_6G_MED - General and working documents', 'reserved', 'S4aP260074', ''],
  ['S4aP260091', TITLE_0091, 'Qualcomm Germany', '', 'pCR', 'Agreement', '5.6.1', 'FS_6G_MED - WT#2.1: AI Traffic Characteristics', 'available', '', '']];

// The accepted Legacy exporter subject (asserted against the real exporter
// in tests/legacy-adhoc-email-collection.test.js).
const subjectFor = (id, agenda, title) => `[FS_6G_MED,${agenda},26-10-15-1500CEST][${id}] Discussion: ${title}`;
const REPLY_0091 = { title: 'Re: ' + subjectFor('S4aP260091', '5.6.1', TITLE_0091), author: 'Jane Delegate' };
const REPLY_0099 = { title: 'Re: ' + subjectFor('S4aP260099', '5.4', 'Draft TR 26.870 FS_6G_MED v0.6.2'), author: 'John Delegate' };

// ------------------------------------------------------------ fake document

// Body.removeChild() returns the BODY and insertTable() only takes a
// detached table, as in Apps Script. Formatting setters are no-ops.
function makeFakeDocumentBody(sandbox) {
  const PARAGRAPH = sandbox.DocumentApp.ElementType.PARAGRAPH;
  const TABLE = sandbox.DocumentApp.ElementType.TABLE;
  const NORMAL = sandbox.DocumentApp.ParagraphHeading.NORMAL;
  const children = [];
  const noop = function () { return this; };
  const styling = {};
  ['setBold', 'setItalic', 'setUnderline', 'setForegroundColor', 'setBackgroundColor', 'setFontSize', 'setFontFamily', 'setAttributes',
    'setLinkUrl', 'setSpacingBefore', 'setSpacingAfter', 'setLineSpacing', 'setIndentStart', 'setIndentFirstLine', 'setAlignment',
    'setGlyphType', 'setPaddingTop', 'setPaddingBottom', 'setPaddingLeft', 'setPaddingRight', 'setWidth', 'setColumnWidth',
    'setBorderWidth', 'setBorderColor', 'setMinimumHeight', 'setVerticalAlignment'].forEach((m) => { styling[m] = noop; });

  function makeCell(text) {
    const c = Object.assign({}, styling, {
      _t: String(text),
      getText: () => c._t,
      setText: (t) => { c._t = String(t); return c; },
      getLinkUrl: () => null,
      getNumChildren: () => 0
    });
    const te = Object.assign({}, styling, { getText: () => c._t, getLinkUrl: () => null, appendText: (v) => { c._t += String(v); return te; } });
    c.editAsText = () => te;
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
      return Object.assign({}, styling, {
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { const c = makeCell(t === undefined ? '' : t); cells.push(c); return c; }
      });
    }
    (rowsData || []).forEach((r) => rows.push(makeRow(r)));
    const t = Object.assign({}, styling, {
      getType: () => TABLE, asTable: () => t,
      getNumRows: () => rows.length, getRow: (i) => rows[i],
      getCell: (r, c) => rows[r].getCell(c),
      appendTableRow: () => { const r = makeRow([]); rows.push(r); return r; },
      insertTableRow: (i) => { const r = makeRow([]); rows.splice(i, 0, r); return r; },
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
      let t;
      if (data && data.getType && data.getType() === TABLE) {
        if (children.indexOf(data) !== -1) throw new Error('Element must be detached.');
        t = data;
      } else {
        t = makeTable(data || []);
      }
      children.splice(idx, 0, t); return t;
    },
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

// ------------------------------------------------------------ fake platform

const fakeXmlService = {
  getNamespace: () => ({}),
  parse: (xml) => {
    const items = (xml.match(/<item>[\s\S]*?<\/item>/g) || []).map((it) => ({
      getChildText: (tag) => { const m = it.match(new RegExp('<' + tag + '>([\\s\\S]*?)</' + tag + '>')); return m ? m[1] : null; }
    }));
    const channel = { getChildren: () => items };
    return { getRootElement: () => ({ getChild: () => channel, getChildren: () => items }) };
  }
};
const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const rssXml = (msgs) => '<rss><channel>' + msgs.map((m, i) =>
  `<item><title>${esc(m.title)}</title><author>${m.author}</author><pubDate>${MSG_DATE}</pubDate><link>https://list.etsi.org/rss/m${i}</link><guid>rss-${i}</guid><description></description></item>`).join('') + '</channel></rss>';

/**
 * A built-looking 86178 report: the Document Reallocations table, 0074 under
 * 5.4, 0091 under 5.6.1, and no S4aP260099 table. The TDoc list (with 0099
 * "reserved") and the 6G list's RSS feed (`messages`) are served by one fake
 * UrlFetchApp; every other list is empty.
 */
function setupReport(messages) {
  const loaded = loadCode({ documentProperties: PROPS_86178 });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const logs = [];
  const fetched = [];
  const RealDate = Date;
  s.Date = class extends RealDate {
    constructor(...a) { if (a.length === 0) super(NOW); else super(...a); }
    static now() { return new RealDate(NOW).getTime(); }
  };
  s.Logger = { log: (m) => logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.XmlService = fakeXmlService;
  s.Utilities.Charset = { UTF_8: 'UTF-8' };
  s.UrlFetchApp = {
    fetch: (url) => {
      fetched.push(url);
      if (url.indexOf('list.etsi.org') === -1) return { getResponseCode: () => 200, getBlob: () => ({ setName() { return this; } }) };
      const hit = new RegExp('[?&]L=' + LIST_6G + '(&|$)').test(url);
      const text = url.indexOf('?RSS&') !== -1 ? rssXml(hit ? messages : []) : '<html></html>';
      return { getResponseCode: () => 200, getContentText: () => text };
    }
  };
  s.DriveApp.createFile = () => ({ setTrashed() {} });
  s.SpreadsheetApp.open = () => ({ getSheets: () => [{ getDataRange: () => ({
    getValues: () => SHEET.map((r) => r.slice()),
    getRichTextValues: () => SHEET.map((r) => r.map(() => null))
  }) }] });
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.getFetchAbstractsSetting_ = () => false;
  // The revisions half of the collector needs the FTP drafts folder; it is
  // not part of either fix. collectorUpdate_() and checkRSSFeed_() are real.
  s.updateRevisions_ = () => {};

  const H2 = s.DocumentApp.ParagraphHeading.HEADING2;
  const H3 = s.DocumentApp.ParagraphHeading.HEADING3;
  body._heading('5.1 Opening of the session and registration of documents', H2);
  body._heading('5.1.3 Document Reallocations', H3);
  body._table([['TDoc', 'Original Agenda', 'New Agenda', 'Reason']]);
  body._heading('5.4 FS_6G_MED - General and working documents', H2);
  body._table(tdocTable('S4aP260074', '5.4'));
  body._heading('5.6 FS_6G_MED - WT#2: 6G Media', H2);
  body._heading('5.6.1 FS_6G_MED - WT#2.1: AI Traffic Characteristics', H3);
  body._table(tdocTable('S4aP260091', '5.6.1'));
  body._heading('5.10 FS_6G_MED - Other issues', H2);
  return { s, body, logs, fetched, docProps: loaded.docProps };
}

const cellText = (t, r, k) => t.getRow(r).getCell(k).getText();
const tablesFor = (env, id) => env.body.getTables().filter((t) => env.s.isTDocTable_(t) && cellText(t, 0, 1).trim() === id);
function rowValue(t, label) { for (let r = 0; r < t.getNumRows(); r++) if (cellText(t, r, 0) === label) return cellText(t, r, 1); return null; }
const stored = (env, id) => { const v = env.docProps._store['DISCUSS_' + id]; return v === undefined ? null : Object.keys(JSON.parse(v)).length; };
function sectionOf(env, el) {
  const NORMAL = env.s.DocumentApp.ParagraphHeading.NORMAL;
  for (let i = env.body.getChildIndex(el) - 1; i >= 0; i--) {
    const c = env.body.getChild(i);
    if (c.getType() === 'PARAGRAPH' && c.getHeading() !== NORMAL) return c.getText().split(' ')[0];
  }
  return null;
}
const headings = (env) => env.body._children.filter((c) => c.getType() === 'PARAGRAPH' && c.getHeading() !== env.s.DocumentApp.ParagraphHeading.NORMAL).map((c) => c.getText().split(' ')[0]);
function reallocationRows(env) {
  const t = env.body.getTables().find((x) => env.s.isReallocationTable_(x));
  return Array.from({ length: t.getNumRows() - 1 }, (_, i) => [0, 1, 2, 3].map((k) => cellText(t, i + 1, k)));
}
function runContinuousUpdate(env) {
  env.logs.length = 0;
  env.fetched.length = 0;
  let error = null;
  try { env.s.continuousUpdate(); } catch (e) { error = e.message; }
  return { error, problems: env.logs.filter((l) => /^ERROR|\[WARN /.test(l)) };
}
const HEADINGS = ['5.1', '5.1.3', '5.4', '5.6', '5.6.1', '5.10'];
const LINE_0091 = 'Jane Delegate on 2026-10-01 08:00\n';

// ================================================= suppressed + collected

console.log('S4aP260099 -> "Removed"; replies to S4aP260091 and S4aP260099 arrive on the 6G list');
{
  const env = setupReport([REPLY_0091, REPLY_0099]);
  env.s.saveReallocation('S4aP260099', '5.4', 'Removed', 'excluded from this report');

  const run = runContinuousUpdate(env);
  check('E continuousUpdate() completes: no error, no warning (real grouping + real collector)', run, { error: null, problems: [] });
  check('E the TDoc list and the 6G list (RSS + A1) were both fetched',
    [env.fetched.some((u) => u === PROPS_86178.TDOC_LIST_URL), env.fetched.some((u) => u.indexOf('?RSS&L=' + LIST_6G + '&') !== -1),
      env.fetched.some((u) => u.indexOf('?A1=') !== -1 && u.indexOf('&L=' + LIST_6G) !== -1)], [true, true, true]);
  check('E no list other than the 6G list was read',
    Array.from(new Set(env.fetched.filter((u) => u.indexOf('list.etsi.org') !== -1).map((u) => (u.match(/[?&]L=([^&]+)/) || [])[1]))), [LIST_6G]);

  check('A S4aP260099 is not imported; nothing was inserted', [tablesFor(env, 'S4aP260099').length, env.logs.some((l) => /^Inserted /.test(l))], [0, false]);
  check('A no "Removed" heading, headings unchanged', headings(env), HEADINGS);
  check('A the reallocation entry is untouched', reallocationRows(env), [['S4aP260099', '5.4', 'Removed', 'excluded from this report']]);

  check('B S4aP260091 is an eligible TDoc table, once, under 5.6.1',
    [env.s.parseExactSA4DocumentId_('S4aP260091').isValid, tablesFor(env, 'S4aP260091').length, sectionOf(env, tablesFor(env, 'S4aP260091')[0])], [true, 1, '5.6.1']);
  check('B S4aP260074 is still in place under 5.4', [tablesFor(env, 'S4aP260074').length, sectionOf(env, tablesFor(env, 'S4aP260074')[0])], [1, '5.4']);

  check('C the S4aP260091 reply is stored for S4aP260091 only', [stored(env, 'S4aP260091'), stored(env, 'S4aP260074')], [1, 0]);
  check('C ... and shown in its E-mail Discussion cell only',
    [rowValue(tablesFor(env, 'S4aP260091')[0], 'E-mail Discussion'), rowValue(tablesFor(env, 'S4aP260074')[0], 'E-mail Discussion').indexOf('Delegate')], [LINE_0091, -1]);

  check('D the reply naming S4aP260099 recreated nothing and stored nothing', [tablesFor(env, 'S4aP260099').length, stored(env, 'S4aP260099')], [0, null]);
  check('D ... and reached no other table', ['S4aP260074', 'S4aP260091'].map((id) => rowValue(tablesFor(env, id)[0], 'E-mail Discussion').indexOf('John Delegate')), [-1, -1]);

  const again = runContinuousUpdate(env);
  check('a second continuousUpdate(): still suppressed, still one message, one line',
    [again, tablesFor(env, 'S4aP260099').length, stored(env, 'S4aP260091'), rowValue(tablesFor(env, 'S4aP260091')[0], 'E-mail Discussion'), headings(env)],
    [{ error: null, problems: [] }, 0, 1, LINE_0091, HEADINGS]);
}

// ======================================================== control (no entry)

console.log('\ncontrol: the same runtime without the reallocation entry imports S4aP260099 and collects its reply');
{
  const env = setupReport([REPLY_0091, REPLY_0099]);
  const run = runContinuousUpdate(env);
  check('continuousUpdate() completes', run, { error: null, problems: [] });
  check('S4aP260099 is imported once, directly below S4aP260074 under 5.4',
    [tablesFor(env, 'S4aP260099').length, sectionOf(env, tablesFor(env, 'S4aP260099')[0]),
      env.body.getChildIndex(tablesFor(env, 'S4aP260099')[0]) === env.body.getChildIndex(tablesFor(env, 'S4aP260074')[0]) + 1], [1, '5.4', true]);
  check('each reply is associated with exactly the document it names',
    ['S4aP260074', 'S4aP260091', 'S4aP260099'].map((id) => stored(env, id)), [0, 1, 1]);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
