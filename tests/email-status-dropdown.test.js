/**
 * T-2026.10.10 (Code.js 2.21.1): the Status of a TDoc in a discussion e-mail,
 * when that Status is a native dropdown.
 *
 * The regression: the e-mail copies the TDoc table through DocumentApp, and
 * DocumentApp has no text for a dropdown -- so the Status cell of the e-mail
 * was empty in every report with status dropdowns (since T-2026.10.7).
 *
 * The fix is in docTableToHtml_() only: for the value cell of the Status
 * row, when it holds a dropdown and no text, it writes the value
 * readTdocStatus_() reads. This file runs the REAL build of a report, the
 * REAL renderer, both REAL e-mail builders and the REAL Generate RPC on the
 * fake document and the fake Docs API (tests/helpers/template-report.js).
 * All names and numbers are synthetic.
 */

const fs = require('fs');
const path = require('path');
const { templateReport } = require('./helpers/template-report.js');
const { installExportStubs, blobText } = require('./helpers/email-export-stubs.js');

let failures = 0;
function check(name, actual, expected) {
  const pass = JSON.stringify(actual) === JSON.stringify(expected);
  if (!pass) failures++;
  console.log(`  ${pass ? 'ok  ' : 'FAIL'} ${name}`);
  if (!pass) {
    console.log(`         expected ${JSON.stringify(expected)}`);
    console.log(`         actual   ${JSON.stringify(actual)}`);
  }
}

const CODE = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8').replace(/\r\n/g, '\n');
const URL = 'https://docs.google.com/document/d/SYNTHETIC/edit';
const STATUS_NAMES = ['available', 'noted', 'agreed', 'revised', 'parked', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'other', 'replied', 'Plenary', 'postponed'];
const TDOCS = [['S4aA269001', 'agreed'], ['S4aA269002', 'withdrawn'], ['S4aA269003', 'noted'], ['S4aA269004', 'available'], ['S4aA269005', 'not treated'], ['S4aA269006', 'revised']];
const P = (inner) => '<p style="margin:0 0 4px 0;">' + inner + '</p>';
const EMPTY = P('&nbsp;');

/** A report with these TDocs, built by the real build. `release: false`: not a template report, so its statuses are text. */
function report(options) {
  const r = templateReport(Object.assign({ tdocs: TDOCS.map(([id, status]) => ({ id: id, agenda: '4.3', status: status, uploaded: true })),
    props: { MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', DISCUSSION_EMAIL_SENDER: 'reporter@example.com',
      REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts' } }, options || {}));
  installExportStubs(r.s);
  // What must not happen when an e-mail is made: a save of the document, a look at the TDoc list.
  r.saves = 0;
  r.sheetOpens = 0;
  const doc = r.s.DocumentApp.getActiveDocument();
  const realSave = doc.saveAndClose;
  doc.saveAndClose = () => { r.saves++; return realSave(); };
  const realOpen = r.s.SpreadsheetApp.open;
  r.s.SpreadsheetApp.open = (...args) => { r.sheetOpens++; return realOpen(...args); };
  r.build();
  r.files = [];
  r.s.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { r.files.push(b); return { getUrl: () => 'https://drive.example/' + b.getName() }; } }) }) };
  r.html = (id) => r.exec(() => r.s.docTableToHtml_(r.table(id), URL));
  r.generate = (ids) => r.exec(() => r.s.generateTdocDiscussionEmails(ids.map((id) => ({ tableIndex: r.body.getTables().indexOf(r.table(id)), deadline: { date: '2026-10-15', time: '15:00', tz: 'CEST' } })), null));
  return r;
}
/** The rows of an e-mail table: [[label, the HTML of its value cell]]. */
function rowsOf(html) {
  const out = [];
  const re = /<tr><td style="[^"]*">(.*?)<\/td><td style="[^"]*">(.*?)<\/td><\/tr>/g;
  let m;
  while ((m = re.exec(html))) out.push([m[1].replace(/<[^>]+>/g, ''), m[2]]);
  return out;
}
const statusOf = (html) => rowsOf(html).filter((row) => row[0] === 'Status').map((row) => row[1]);
function qpDecode(text) {
  const unwrapped = text.replace(/=\r\n/g, '');
  const bytes = [];
  for (let i = 0; i < unwrapped.length; i++) {
    if (unwrapped[i] === '=' && /^[0-9A-F]{2}$/.test(unwrapped.substr(i + 1, 2))) { bytes.push(parseInt(unwrapped.substr(i + 1, 2), 16)); i += 2; } else bytes.push(unwrapped.charCodeAt(i));
  }
  return Buffer.from(bytes).toString('utf8');
}
const bodyOf = (eml) => qpDecode(eml.split('\r\n\r\n').slice(1).join('\r\n\r\n'));
/** Everything a report holds: every cell of every table with its dropdown, the dropdown definitions, the stored properties. */
function snapshot(r) {
  const tables = r.body.getTables().map((t) => { const rows = []; for (let i = 0; i < t.getNumRows(); i++) { const row = []; for (let c = 0; c < t.getRow(i).getNumCells(); c++) { const cell = t.getRow(i).getCell(c); row.push([cell._t, cell._dropdown, cell._links]); } rows.push(row); } return rows; });
  return JSON.stringify({ tables: tables, children: r.body.getNumChildren(), definitions: r.api.definitions, props: r.docProps._store, scriptProps: r.scriptProps._store, dirty: r.body._state.dirty });
}

// ================================================================ 1. the renderer

console.log('1. docTableToHtml_(): the value selected in a Status dropdown');
const r = report();
const text = report({ release: false });
{
  check('the report has its statuses as dropdowns, but the one the dropdown has no option for', TDOCS.map(([id]) => r.shown(id)),
    ['dropdown:agreed', 'dropdown:withdrawn', 'dropdown:noted', 'dropdown:available', 'text:not treated', 'dropdown:revised']);
  check('DocumentApp has no text for them (what made the Status of the e-mail empty)', TDOCS.map(([id]) => r.field(id, 'Status')), ['', '', '', '', 'not treated', '']);
  check('agreed, withdrawn, noted, available, revised: the Status cell of the e-mail has the selected value',
    ['S4aA269001', 'S4aA269002', 'S4aA269003', 'S4aA269004', 'S4aA269006'].map((id) => statusOf(r.html(id))), [[P('agreed')], [P('withdrawn')], [P('noted')], [P('available')], [P('revised')]]);
  r.api.pick('S4aA269003', 'Plenary');
  check('Plenary is written as the option has it, with its capital', statusOf(r.html('S4aA269003')), [P('Plenary')]);
  const all = STATUS_NAMES.map((name) => { r.api.pick('S4aA269003', name); return statusOf(r.html('S4aA269003'))[0]; });
  check('every one of the 14 Document Status values', all, STATUS_NAMES.map((name) => P(name)));
  r.api.pick('S4aA269003', 'noted');
}

console.log('2. the rest of the table, and a text Status');
{
  check('a report with text statuses (not a template report) has them as text', TDOCS.map(([id]) => text.shown(id)), TDOCS.map(([, status]) => 'text:' + status));
  check('the e-mail table of a TDoc with a dropdown is, byte for byte, that of the same TDoc with the same Status as text',
    TDOCS.map(([id]) => r.html(id) === text.html(id)), TDOCS.map(() => true));
  check('a text Status is rendered as before: a known one, and one the dropdown has no option for',
    [statusOf(text.html('S4aA269001')), statusOf(text.html('S4aA269005')), statusOf(r.html('S4aA269005'))], [[P('agreed')], [P('not treated')], [P('not treated')]]);
  check('no Docs API is asked for a report with text statuses', text.api.calls.length, 0);
  const before = r.html('S4aA269004');
  const labels = rowsOf(before).map((row) => row[0]);
  check('the table has its rows, in order, Status once', [labels.indexOf('TDoc'), labels.filter((l) => l === 'Status').length, labels.length === r.table('S4aA269004').getNumRows()], [0, 1, true]);
  // The cells that are not the value cell of the Status row: unchanged, whatever they hold.
  const own = r.table('S4aA269004');
  const statusRow = labels.indexOf('Status');
  const otherRow = labels.indexOf('Minutes') !== -1 ? labels.indexOf('Minutes') : labels.indexOf('Title');
  const dropdown = own.getRow(statusRow).getCell(1)._dropdown;
  const otherCell = own.getRow(otherRow).getCell(1);
  const keptText = otherCell._t;
  otherCell._t = '';
  otherCell._dropdown = Object.assign({}, dropdown);
  const withOther = rowsOf(r.exec(() => r.s.docTableToHtml_(own, URL)));
  // Two rows labelled Status: the Status row is the first of them, the one findStatusInDocTable_() and so readTdocStatus_()
  // read. Only that one is given the value; the other is left as DocumentApp has it.
  const otherLabel = own.getRow(otherRow).getCell(0);
  const keptLabel = otherLabel._t;
  otherLabel._t = 'Status';
  const withTwo = rowsOf(r.exec(() => r.s.docTableToHtml_(own, URL)));
  otherLabel._t = keptLabel;
  otherCell._t = keptText;
  otherCell._dropdown = null;
  check('a dropdown in another row is not given the status: only the value cell of the Status row is', [withOther[otherRow][1], withOther[statusRow][1]], [EMPTY, P('available')]);
  check('of two rows labelled Status, only the first -- the one readTdocStatus_() reads -- is given the value', [otherRow < statusRow, withTwo[otherRow], withTwo[statusRow]], [true, ['Status', P('available')], ['Status', EMPTY]]);
  check('the label cell of the Status row is the label', /<td style="[^"]*font-weight:bold[^"]*"><p style="margin:0 0 4px 0;">Status<\/p><\/td>/.test(before), true);
  check('the table is as before again', r.html('S4aA269004'), before);
}

console.log('3. text beside a dropdown, a dropdown that cannot be read, HTML in a value');
{
  const edge = report();
  edge.user(() => edge.api.statusCell('S4aA269003').setText('  see <minutes> & notes '));
  check('text beside a dropdown: the text is the status (the rule of readTdocStatus_()), rendered and escaped as text always was, spaces and all',
    [edge.exec(() => edge.s.readTdocStatus_(edge.table('S4aA269003'), edge.field('S4aA269003', 'Status'))), statusOf(edge.html('S4aA269003'))], ['see <minutes> & notes', [P('  see &lt;minutes&gt; &amp; notes ')]]);

  // A custom label of an option, as a user can give one in Google Docs.
  const definition = edge.api.definitions[Object.keys(edge.api.definitions)[0]];
  const option = definition.dropdownDefinitionProperties.options.filter((o) => o.displayValue === 'available')[0];
  option.displayValue = 'R&D <b>"open"</b>';
  check('a value with HTML in it is escaped', statusOf(edge.html('S4aA269004')), [P('R&amp;D &lt;b&gt;&quot;open&quot;&lt;/b&gt;')]);
  check('no tag of the value reaches the e-mail', /<b>"open"<\/b>|R&D </.test(edge.html('S4aA269004')), false);
  option.displayValue = 'available';

  edge.api.fail.get = () => 'synthetic read failure';
  let threw = null;
  let html = '';
  try { html = edge.html('S4aA269001'); } catch (e) { threw = e.message; }
  check('the API cannot be read: no exception, the Status cell stays empty, the rest of the table is there', [threw, statusOf(html), rowsOf(html).length === edge.table('S4aA269001').getNumRows()], [null, [EMPTY], true]);
  check('a text status is still its text then', statusOf(edge.html('S4aA269005')), [P('not treated')]);
  edge.api.fail.get = null;

  const noApi = report();
  noApi.s.Docs = undefined;
  let threwNoApi = null;
  let htmlNoApi = '';
  try { htmlNoApi = noApi.html('S4aA269001'); } catch (e) { threwNoApi = e.message; }
  check('no Docs API service at all: no exception, the Status cell stays empty', [threwNoApi, statusOf(htmlNoApi)], [null, [EMPTY]]);
}

// ================================================================ 2. the builders

console.log('4. the two e-mail builders');
{
  const meta = (id) => ({ tdoc: id, title: 'Synthetic title of ' + id, agendaItem: '4.3' });
  const args = ['list@example.org', null, null, null, '15 October 2026, 15:00 CEST', 'https://example.org/drafts', 'reporter@example.com', '26-10-15-1500CEST', 'FS_TEST', URL, 'list@example.org'];
  const single = (rep, id) => rep.exec(() => rep.s.buildEmailExportForTdocTable_(rep.table(id), meta(id), ...args));
  const one = single(r, 'S4aA269001');
  check('buildEmailExportForTdocTable_(): an agreed TDoc has "agreed" in its Status (the builder itself refuses nothing)', statusOf(bodyOf(one.eml)), [P('agreed')]);
  check('and a withdrawn one "withdrawn"', statusOf(bodyOf(single(r, 'S4aA269002').eml)), [P('withdrawn')]);
  check('subject and file name are what they are for the same TDoc with a text Status', [one.subject, one.fileName], [single(text, 'S4aA269001').subject, single(text, 'S4aA269001').fileName]);
  check('the whole e-mail body is that of the same TDoc with a text Status', bodyOf(one.eml) === bodyOf(single(text, 'S4aA269001').eml), true);

  const group = (rep, ids) => rep.exec(() => rep.s.buildEmailExportForGroup_(ids.map((id) => rep.table(id)), ids.map(meta), ...args));
  const ids = ['S4aA269003', 'S4aA269001', 'S4aA269005', 'S4aA269002'];
  const grouped = group(r, ids);
  check('buildEmailExportForGroup_(): every table of the group has its own Status, in the order of the group', statusOf(bodyOf(grouped.eml)), [P('noted'), P('agreed'), P('not treated'), P('withdrawn')]);
  check('and the group is, as a whole, that of the report with text statuses', [bodyOf(grouped.eml) === bodyOf(group(text, ids).eml), grouped.subject === group(text, ids).subject], [true, true]);
}

// ================================================================ 3. Generate

console.log('5. Generate (the dialog RPC), and what it leaves alone');
{
  const g = report();
  const before = snapshot(g);
  const calls = g.api.calls.length;
  const fetches = g.fetches.length;
  const saves = g.saves;
  const sheetOpens = g.sheetOpens;
  const result = g.generate(['S4aA269003', 'S4aA269004']);
  const emls = g.files.filter((b) => /\.eml$/.test(b.getName()));
  check('two e-mails are generated', [result.ok, result.error, emls.map((b) => b.getName())], [true, null, ['FS_6G_MED_S4aA269003.eml', 'FS_6G_MED_S4aA269004.eml']]);
  check('each has the value of its dropdown in the Status row', emls.map((b) => statusOf(bodyOf(blobText(b)))), [[P('noted')], [P('available')]]);
  check('the subject is as it was', /^Subject: \[4\.3\]\[26-10-15-1500CEST\]\[S4aA269003\] Discussion: /m.test(blobText(emls[0]).split('\r\n\r\n')[0].replace(/\r\n[ \t]/g, ' ')), true);

  check('the report is what it was: every cell, every dropdown and its value, the definition with its options and colours, the properties; nothing unsaved', snapshot(g) === before, true);
  check('the dropdowns still show what they showed', TDOCS.map(([id]) => g.shown(id)), ['dropdown:agreed', 'dropdown:withdrawn', 'dropdown:noted', 'dropdown:available', 'text:not treated', 'dropdown:revised']);
  const during = g.api.calls.slice(calls);
  check('the Docs API was read once and never written', [during.map((c) => c.call), g.api.calls.filter((c, i) => i >= calls && c.call === 'batchUpdate').length], [['get'], 0]);
  check('the document was not saved, the TDoc list was not looked at, nothing was fetched', [g.saves - saves, g.sheetOpens - sheetOpens, g.fetches.length - fetches, g.body._state.dirty, g.body._state.closed], [0, 0, 0, false, false]);

  g.files.length = 0;
  const again = g.generate(['S4aA269003', 'S4aA269004']);
  const emlsAgain = g.files.filter((b) => /\.eml$/.test(b.getName()));
  const stable = (eml) => eml.replace(/^(Date|Message-ID): .*$/gm, '$1:').replace(/boundary="[^"]*"/g, 'boundary=""');
  check('generating again gives the same e-mails', [again.ok, emlsAgain.map((b) => stable(blobText(b)))], [true, emls.map((b) => stable(blobText(b)))]);
  check('and the report is still what it was', snapshot(g) === before, true);

  // The eligibility rule is the one it was: it reads the status the same way.
  const refused = g.generate(['S4aA269001']);
  check('an agreed TDoc is still refused by Generate', [refused.ok, /S4aA269001 is already agreed and is excluded from discussion e-mail export/.test(refused.error)], [false, true]);
  g.api.pick('S4aA269004', 'reserved');
  const reserved = g.generate(['S4aA269004']);
  check('and so is a reserved one', [reserved.ok, /S4aA269004 is reserved/.test(reserved.error)], [false, true]);
}

// ================================================================ 4. the code

console.log('6. the code');
{
  const start = CODE.indexOf('\nfunction docTableToHtml_(');
  const fn = CODE.slice(start, CODE.indexOf('\n}\n', start) + 3);
  const code = fn.replace(/^\s*\/\/.*$/gm, '');
  check('docTableToHtml_() reads a dropdown through readTdocStatus_(), once, and escapes what it reads', [(code.match(/readTdocStatus_\(table, cellText\)/g) || []).length, /escapeHtmlForEmailExport_\(selected\)/.test(code)], [1, true]);
  check('it has no way of its own to read a dropdown, and writes nothing', /Docs\.|statusDropdownIndex_|statusDropdownEntryOfTable_|displayValue|setText|batchUpdate|saveAndClose|setProperty/.test(code), false);
  check('Code.js is version 2.22.0 (2.21.1 brought the dropdown status of a discussion e-mail; 2.22.0 the shared minutes)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.22.0');
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll e-mail status dropdown checks passed.');
process.exit(failures ? 1 : 0);
