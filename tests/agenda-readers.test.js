/**
 * The agenda readers after T-2026.10.20 (MBS SA4#138, 2026-11-11).
 *
 * With T-2026.10.20 deployed, diagnoseAgendaConversion() showed: the agenda
 * S4-261857 downloads, Drive converts it to a native Google Doc, and
 * DocumentApp.openById(...).getBody() fails on all four attempts. Converted
 * again outside Apps Script, the same Word file gives a sound Google Doc
 * (Drive exports it in full: one table of 101 rows headed "A.I.#"). Now the
 * converted agenda is read by DocumentApp, else the Docs API, else the Word
 * file itself -- each through agendaBodyView_() and the unchanged parser.
 *
 *   A. DocumentApp reads it: nothing else is asked;
 *   B. DocumentApp fails, the Docs API reads it (on the same document, before
 *      it is trashed);
 *   C. both APIs fail: the Word file is read, with no Google Doc;
 *   D. all three fail: one error naming each; the document is still trashed;
 *   E. temporary-file cleanup in every case; the log never holds agenda text;
 *   F. equivalence on the REAL agendas: the Word file reader gives exactly
 *      the items Google's own conversion gives (tests/fixtures/agenda-S4-
 *      261857.docx and agenda-S4-260868.docx against their .google-items.json:
 *      Drive's export of the converted Google Doc, read with the parser's
 *      rules); the three readers agree on SA4#138;
 *   G. TDoc association: every one of the 388 TDocs of SA4#138 falls under
 *      an agenda item read from the Word file; the MBS report's projection
 *      keeps its 46;
 *   H. the readers' semantics: the Body view, the Docs API blocks and the
 *      Word blocks (lists, headings, deleted/inserted text, fields, drawings,
 *      content controls, nested tables, namespaces, malformed input);
 *   I. diagnoseAgendaReaders(): every reader, compared, nothing written;
 *   J. what changed in Code.js: one section and two functions, against
 *      T-2026.10.20;
 *   K. an empty or malformed agenda is not a successful read: the next
 *      reader is asked; if none reads one, the error names each.
 *
 * Fakes only: no network, no Google account. Document IDs are synthetic.
 *
 * Run: node tests/agenda-readers.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeBlob, unzip } = require('./helpers/fake-zip.js');
const { codeWithoutTdocListDownload, codeWithoutAgendaReaders, AGENDA_READERS_SECTION_START, AGENDA_READERS_SECTION_END, AGENDA_READERS_SECTION_FUNCTIONS, AGENDA_READERS_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log('  ok   ' + name);
  } else {
    failures++;
    console.log('  FAIL ' + name + '\n         expected ' + e + '\n         actual   ' + a);
  }
}

const ROOT = path.join(__dirname, '..');
const FIX = path.join(__dirname, 'fixtures');
const DOCX_138 = fs.readFileSync(path.join(FIX, 'agenda-S4-261857.docx'));
const DOCX_136 = fs.readFileSync(path.join(FIX, 'agenda-S4-260868.docx'));
const GOOGLE_138 = JSON.parse(fs.readFileSync(path.join(FIX, 'agenda-S4-261857.google-items.json'), 'utf8')).items;
const GOOGLE_136 = JSON.parse(fs.readFileSync(path.join(FIX, 'agenda-S4-260868.google-items.json'), 'utf8')).items;
const TDOCS_138 = JSON.parse(fs.readFileSync(path.join(FIX, 'meeting-SA4-138-tdoc-agenda.json'), 'utf8')).tdocs;
const NAME_138 = 'S4-261857 Meeting agenda for SA4#138.docx';
const GDOC = 'application/vnd.google-apps.document';
const DOC_ID = '1convertedAgendaDocIdForReadersTest000000';
const UNEXPECTED = 'Unexpected error while getting the method or property openById on object DocumentApp.';
const HEADINGS = { NORMAL: 'NORMAL', TITLE: 'TITLE', SUBTITLE: 'SUBTITLE', HEADING1: 'H1', HEADING2: 'H2', HEADING3: 'H3', HEADING4: 'H4', HEADING5: 'H5', HEADING6: 'H6' };
const strip = (items) => items.map((i) => ({ number: i.number, title: i.title, level: i.level }));

/** A DocumentApp body as DocumentApp gives it for a table agenda: no item paragraphs, one "A.I.#" table. Independent of agendaBodyView_. */
function documentAppBody(items) {
  const rows = [['\tA.I.#', '\tAgenda Item']].concat(items.map((i) => ['\t' + i.number, '\t' + i.title]));
  const table = {
    getNumRows: () => rows.length,
    getRow: (r) => ({ getNumCells: () => rows[r].length, getCell: (c) => ({ getText: () => rows[r][c] }) }),
    getCell: (r, c) => ({ getText: () => rows[r][c] })
  };
  const paragraph = (text) => ({ getType: () => 'PARAGRAPH', asParagraph: function () { return this; }, getText: () => text, getHeading: () => 'NORMAL' });
  const children = [paragraph('\tTitle:\tMeeting agenda'), paragraph('Introduction'), { getType: () => 'TABLE' }];
  return { getNumChildren: () => children.length, getChild: (i) => children[i], getTables: () => [table] };
}

/** A Docs API document as Documents.get() returns it for the same agenda. */
function docsApiDocument(items) {
  const p = (text, extra) => ({ paragraph: Object.assign({ elements: [{ startIndex: 1, textRun: { content: text + '\n', textStyle: {} } }], paragraphStyle: { namedStyleType: 'NORMAL_TEXT' } }, extra || {}) });
  const cell = (text) => ({ content: [p(text)] });
  const rows = [['\tA.I.#', '\tAgenda Item']].concat(items.map((i) => ['\t' + i.number, '\t' + i.title]));
  return {
    documentId: DOC_ID, title: 'TEMP_Agenda',
    body: { content: [
      { sectionBreak: { sectionStyle: {} } },
      p('\tTitle:\tMeeting agenda', { paragraphStyle: { namedStyleType: 'TITLE' } }),
      p('1 A numbered list paragraph is not an agenda item', { bullet: { listId: 'kix.list1' } }),
      { table: { rows: rows.length, columns: 2, tableRows: rows.map((r) => ({ tableCells: r.map(cell) })) } },
      p('')
    ] }
  };
}

/**
 * A sandbox: Drive conversion, DocumentApp, the Docs API, DriveApp trash and
 * Utilities (newBlob/unzip on real ZIP bytes). o.documentApp: 'ok' | 'fail';
 * o.docs: a document | 'fail' | 'absent'; o.trashThrows.
 */
function world(o) {
  const opts = Object.assign({ documentApp: 'ok', docs: 'fail', items: GOOGLE_138 }, o || {});
  const { sandbox, docProps } = loadCode({ documentProperties: opts.props || {} });
  const w = { sandbox, docProps, logs: [], events: [], unzips: 0 };
  sandbox.Logger = { log: (m) => { w.logs.push(String(m)); } };
  sandbox.MimeType = { GOOGLE_DOCS: GDOC };
  sandbox.DocumentApp.ParagraphHeading = HEADINGS;
  sandbox.Drive = { Files: { create: (resource) => { w.events.push('create'); return { id: DOC_ID, name: resource.name, mimeType: GDOC }; } } };
  sandbox.DocumentApp.openById = (id) => {
    w.events.push('DocumentApp.openById');
    if (opts.documentApp !== 'ok') throw new Error(UNEXPECTED);
    return { getBody: () => documentAppBody(opts.documentAppItems || opts.items) };
  };
  if (opts.docs !== 'absent') {
    sandbox.Docs = { Documents: { get: (id) => {
      w.events.push('Docs.Documents.get ' + id);
      if (opts.docs === 'fail') throw new Error('API call to docs.documents.get failed with error: Internal error encountered.');
      return opts.docs;
    } } };
  }
  sandbox.DriveApp = { getFileById: (id) => ({ setTrashed: (v) => { w.events.push('trash ' + id); if (opts.trashThrows) throw new Error('Drive trash failed'); } }) };
  sandbox.Utilities.sleep = () => {};
  sandbox.Utilities.newBlob = makeBlob;
  sandbox.Utilities.unzip = (blob) => { w.unzips++; return unzip(blob); };
  w.run = (bytes) => {
    try { return { result: sandbox.parseAgendaWordBlob_(makeBlob(bytes || DOCX_138, 'application/octet-stream', NAME_138)) }; } catch (e) { return { error: e.message }; }
  };
  w.agendaLogs = () => w.logs.filter((l) => l.indexOf('[AGENDA]') === 0);
  w.leaks = () => w.logs.filter((l) => /Multicast-Broadcast-Streaming|Opening of the meeting|Approval of the agenda/.test(l));
  return w;
}

// ====================================================================== A
console.log('A. DocumentApp reads the agenda: nothing else is asked');
{
  const w = world({ documentApp: 'ok', docs: docsApiDocument(GOOGLE_138) });
  const out = w.run();
  check('the agenda items are those of the document', out.result && strip(out.result), strip(GOOGLE_138));
  check('converted, opened once, trashed; no Docs API call, no Word file read', [w.events, w.unzips], [['create', 'DocumentApp.openById', 'trash ' + DOC_ID], 0]);
  check('the log names the reader', w.agendaLogs().some((l) => l === '[AGENDA] read by DocumentApp: 100 agenda item(s)'), true);
  check('...and AGENDA_LAST_READER_', w.sandbox.AGENDA_LAST_READER_, 'DocumentApp');
}

// ====================================================================== B
console.log('B. DocumentApp fails, the Docs API reads the same document');
{
  const w = world({ documentApp: 'fail', docs: docsApiDocument(GOOGLE_138) });
  const out = w.run();
  check('the agenda items come from the Docs API, equal to Google\'s', out.result && strip(out.result), strip(GOOGLE_138));
  check('Docs.Documents.get on the converted document, BEFORE it is trashed; the Word file not read',
    [w.events.filter((e) => e !== 'DocumentApp.openById'), w.events.filter((e) => e === 'DocumentApp.openById').length, w.unzips],
    [['create', 'Docs.Documents.get ' + DOC_ID, 'trash ' + DOC_ID], 4, 0]);
  check('the log: DocumentApp\'s failure with its stack, the document\'s shape (counts only), the reader', [
    w.agendaLogs().some((l) => l.indexOf('[AGENDA] DocumentApp could not read the agenda (Google Doc ' + DOC_ID + '): Error: The agenda converted from Word') === 0 && /\n\s+at /.test(l)),
    w.agendaLogs().filter((l) => l.indexOf('[AGENDA] Docs.Documents.get(') === 0),
    w.agendaLogs().some((l) => l === '[AGENDA] read by Docs API (after: DocumentApp): 100 agenda item(s)')
  ], [true, ['[AGENDA] Docs.Documents.get(' + DOC_ID + '): 5 body element(s), 2 paragraph(s), 1 list item(s), 1 table(s) ["101x2"]'], true]);
  check('no agenda text in the log', w.leaks(), []);
}

// ====================================================================== C
console.log('C. both APIs fail: the Word file itself');
{
  const w = world({ documentApp: 'fail', docs: 'fail' });
  const out = w.run();
  check('the agenda items come from the Word file, equal to Google\'s conversion of it', out.result && strip(out.result), strip(GOOGLE_138));
  check('both APIs tried on the document, then it is trashed; the Word file unzipped once', [w.events.filter((e) => e !== 'DocumentApp.openById'), w.unzips],
    [['create', 'Docs.Documents.get ' + DOC_ID, 'trash ' + DOC_ID], 1]);
  check('each failure logged with its original error; the reader named', [
    w.agendaLogs().some((l) => l.indexOf('[AGENDA] Docs API could not read the agenda (Google Doc ' + DOC_ID + '): Error: API call to docs.documents.get failed with error: Internal error encountered.') === 0 && /\n\s+at /.test(l)),
    w.agendaLogs().some((l) => l === '[AGENDA] read by Word file (after: DocumentApp, Docs API): 100 agenda item(s)')
  ], [true, true]);
  const absent = world({ documentApp: 'fail', docs: 'absent' });
  check('without the Docs advanced service: a Docs API failure, then the Word file', [strip(absent.run().result || []).length, absent.agendaLogs().some((l) => /Docs API could not read the agenda .*the Docs advanced service is not available/.test(l))], [100, true]);
  check('no agenda text in the log', w.leaks(), []);
}

// ====================================================================== D
console.log('D. all three fail');
{
  const w = world({ documentApp: 'fail', docs: 'fail' });
  const out = w.run(Buffer.from('not a zip at all'));
  check('one error naming each reader and its reason', out.error && out.error.replace(/\(Google Doc \w+\)/, '(Google Doc X)'),
    'The agenda could not be read -- DocumentApp: The agenda converted from Word (Google Doc X) could not be opened after 4 attempts over 7 s (DocumentApp.openById): ' + UNEXPECTED +
    '; Docs API: API call to docs.documents.get failed with error: Internal error encountered.; Word file: the Word file is not a .docx (it cannot be unzipped): Could not unzip: not a ZIP file');
  check('the temporary document is still trashed, once, last', w.events[w.events.length - 1] === 'trash ' + DOC_ID && w.events.filter((e) => /^trash/.test(e)).length, 1);
  const noDocument = world({ documentApp: 'fail', docs: 'fail' });
  const zipWithoutDocument = Buffer.from(fs.readFileSync(path.join(FIX, 'agenda-S4-261857.docx')));
  // A ZIP that is not a .docx: rebuild the real one's bytes with every name changed so that word/document.xml is missing.
  const renamed = Buffer.from(zipWithoutDocument.toString('latin1').split('word/document.xml').join('word/documenX.xml'), 'latin1');
  check('a ZIP without word/document.xml: named as such', /Word file: the Word file has no word\/document\.xml$/.test(noDocument.run(renamed).error || ''), true);
}

// ====================================================================== E
console.log('E. temporary-file cleanup');
{
  const cases = [
    ['DocumentApp reads', { documentApp: 'ok' }],
    ['the Docs API reads', { documentApp: 'fail', docs: docsApiDocument(GOOGLE_138) }],
    ['the Word file reads', { documentApp: 'fail', docs: 'fail' }],
    ['nothing reads', { documentApp: 'fail', docs: 'fail', bytes: Buffer.from('x') }]
  ];
  check('trashed exactly once, after every reader that touches it', cases.map(([name, o]) => {
    const w = world(o); w.run(o.bytes);
    const trash = w.events.indexOf('trash ' + DOC_ID);
    return [name, w.events.filter((e) => /^trash/.test(e)).length, trash === w.events.length - 1];
  }), cases.map(([name]) => [name, 1, true]));
  const t = world({ documentApp: 'fail', docs: 'fail', trashThrows: true });
  const out = t.run();
  check('a trash that fails: the agenda is still returned, the failure logged', [strip(out.result || []).length, t.agendaLogs().some((l) => l.indexOf('[AGENDA] could not move the temporary agenda document ' + DOC_ID + ' to the trash: Error: Drive trash failed') === 0)], [100, true]);
}

// ====================================================================== F
console.log('F. equivalence on the real agendas');
{
  const w = world();
  const s = w.sandbox;
  const wordItems = (bytes) => s.parseAgendaStructureWithText_(s.wordFileAgendaBody_(Array.from(bytes)), '');
  const w138 = wordItems(DOCX_138), w136 = wordItems(DOCX_136);
  check('SA4#138 (S4-261857): the Word file gives Google\'s 100 items -- numbers, titles, levels, order', [w138.length, JSON.stringify(strip(w138)) === JSON.stringify(strip(GOOGLE_138))], [100, true]);
  check('SA4#136 (S4-260868): the Word file gives Google\'s 102 items', [w136.length, JSON.stringify(strip(w136)) === JSON.stringify(strip(GOOGLE_136))], [102, true]);
  check('both come from the "A.I.#" table: no descriptive text, NORMAL heading', [w138.every((i) => i.text === '' && i.heading === 'NORMAL'), w136.every((i) => i.text === '' && i.heading === 'NORMAL')], [true, true]);
  check('the MBS section of SA4#138: 8 and 8.1-8.10', w138.filter((i) => /^8(\.|$)/.test(i.number)).map((i) => i.number), ['8', '8.1', '8.2', '8.3', '8.4', '8.5', '8.6', '8.7', '8.8', '8.9', '8.10']);
  const viaDocumentApp = s.parseAgendaStructureWithText_(documentAppBody(GOOGLE_138), '');
  const viaDocsApi = s.parseAgendaStructureWithText_(s.docsApiAgendaBody_(docsApiDocument(GOOGLE_138)), '');
  const key = (items) => JSON.stringify(items.map((i) => [i.number, i.title, i.level, i.heading, i.text]));
  check('the three readers give identical items on SA4#138 (number, title, level, heading, text)', [key(viaDocsApi) === key(viaDocumentApp), key(w138) === key(viaDocumentApp)], [true, true]);
}

// ====================================================================== G
console.log('G. TDoc association on SA4#138');
{
  const w = world({ props: { MEETING_TYPE: 'main', REPORT_SUFFIX: 'MBS', MEETING_NUMBER: '138', MEETING_FOLDER: 'TSGS4_138_Calgary' } });
  const s = w.sandbox;
  const items = s.parseAgendaStructureWithText_(s.wordFileAgendaBody_(Array.from(DOCX_138)), '');
  const numbers = items.map((i) => i.number);
  check('all 388 TDocs of the meeting fall under an agenda item read from the Word file', [TDOCS_138.length, TDOCS_138.filter(([id, ai]) => numbers.indexOf(ai.trim()) === -1)], [388, []]);
  const projected = s.projectAgendaItems_(items, s.getMeetingContext_().report.agendaSelector).map((i) => i.number);
  const mbs = TDOCS_138.filter(([id, ai]) => projected.indexOf(ai.trim()) !== -1);
  check('the MBS report\'s projection: 8.1-8.10 (the parent 8 dropped), with the 46 MBS TDocs under them', [projected, mbs.length],
    [['8.1', '8.2', '8.3', '8.4', '8.5', '8.6', '8.7', '8.8', '8.9', '8.10'], 46]);
  const google = s.projectAgendaItems_(GOOGLE_138, s.getMeetingContext_().report.agendaSelector).map((i) => i.number);
  check('...exactly the projection of Google\'s items', JSON.stringify(projected) === JSON.stringify(google), true);
}

// ====================================================================== H
console.log('H. the readers\' semantics');
{
  const s = world().sandbox;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const doc = (body, ns) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + (ns || NS) + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body>' + body + '<w:sectPr/></w:body></w:document>';
  const styles = '<w:styles ' + NS + '><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Titre1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListNum"><w:name w:val="List Number"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListChild"><w:name w:val="List Child"/><w:basedOn w:val="ListNum"/></w:style></w:styles>';
  const p = (inner, pPr) => '<w:p>' + (pPr ? '<w:pPr>' + pPr + '</w:pPr>' : '') + inner + '</w:p>';
  const r = (t) => '<w:r><w:t xml:space="preserve">' + t + '</w:t></w:r>';
  const blocks = (body, st, ns) => s.wordFileAgendaBlocks_(doc(body, ns), st === undefined ? styles : st);
  const texts = (bs) => bs.map((b) => [b.type, b.text, b.heading]);

  check('Word runs: text, tabs, line breaks; page breaks, deleted text, field codes, drawings and text boxes are not text', texts(blocks(
    p('<w:r><w:tab/><w:t>8.1</w:t><w:br/><w:t>A</w:t><w:br w:type="page"/></w:r>' +
      '<w:del><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins><w:r><w:t> new</w:t></w:r></w:ins>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><mc:AlternateContent><mc:Choice><w:drawing><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice></mc:AlternateContent></w:r>' +
      '<w:hyperlink><w:r><w:t> link</w:t></w:r></w:hyperlink><w:r><w:noBreakHyphen/><w:t>x &amp; y &lt;z&gt; &#8211;</w:t></w:r>'))),
    [['PARAGRAPH', '\t8.1\rA new7 link-x & y <z> –', 'NORMAL']]);
  check('lists: own numbering, style numbering (inherited), numId 0 = not a list; heading by style name', texts(blocks(
    p(r('a'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>') + p(r('b'), '<w:pStyle w:val="ListChild"/>') +
    p(r('c'), '<w:pStyle w:val="ListNum"/><w:numPr><w:numId w:val="0"/></w:numPr>') + p(r('d'), '<w:pStyle w:val="Titre1"/>') + p(r('e')))),
    [['LIST_ITEM', 'a', 'NORMAL'], ['LIST_ITEM', 'b', 'NORMAL'], ['PARAGRAPH', 'c', 'NORMAL'], ['PARAGRAPH', 'd', 'HEADING1'], ['PARAGRAPH', 'e', 'NORMAL']]);
  check('content controls are unwrapped (block, row, cell); a table cell is its paragraphs joined by "\\n"', JSON.stringify(blocks(
    '<w:sdt><w:sdtPr/><w:sdtContent>' + p(r('in sdt')) + '</w:sdtContent></w:sdt>' +
    '<w:tbl><w:tblPr/><w:sdt><w:sdtContent><w:tr><w:tc>' + p(r('A.I.#')) + '</w:tc><w:sdt><w:sdtContent><w:tc>' + p(r('one')) + p(r('two')) + '</w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt></w:tbl>')),
    JSON.stringify([{ type: 'PARAGRAPH', text: 'in sdt', heading: 'NORMAL' }, { type: 'TABLE', rows: [[[{ type: 'PARAGRAPH', text: 'A.I.#', heading: 'NORMAL' }], [{ type: 'PARAGRAPH', text: 'one', heading: 'NORMAL' }, { type: 'PARAGRAPH', text: 'two', heading: 'NORMAL' }]]] }]));
  const view = s.agendaBodyView_(blocks('<w:tbl><w:tr><w:tc>' + p(r('outer')) + '<w:tbl><w:tr><w:tc>' + p(r('inner')) + '</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>' +
    '<w:tbl><w:tr><w:tc>' + p(r('A.I.#')) + '</w:tc><w:tc>' + p(r('x')) + p(r('y')) + '</w:tc></w:tr></w:tbl>'));
  check('the Body view: tables in document order, a nested table after its table; cell text; row cells', [view.getTables().length, view.getTables().map((t) => t.getCell(0, 0).getText()), view.getTables()[2].getRow(0).getNumCells(), view.getTables()[2].getCell(0, 1).getText()],
    [3, ['outer\ninner', 'inner', 'A.I.#'], 2, 'x\ny']);
  check('the Body view: element types as DocumentApp names them (others are skipped by the parser)', (() => { const v = s.agendaBodyView_([{ type: 'PARAGRAPH', text: 'p' }, { type: 'LIST_ITEM', text: 'l' }, { type: 'TABLE', rows: [] }, { type: 'OTHER' }]);
    return [0, 1, 2, 3].map((i) => v.getChild(i).getType()); })(), ['PARAGRAPH', 'LIST_ITEM', 'TABLE', null]);
  check('the paragraph path wins over the table, as in DocumentApp: "8.1 Title" paragraphs, list items skipped, text collected', strip(s.parseAgendaStructureWithText_(s.agendaBodyView_(blocks(
    p(r('8 MBS')) + p(r('8.1 Opening')) + p(r('Some text')) + p(r('9 Not numbered as a list'), '<w:numPr><w:numId w:val="2"/></w:numPr>') + p(r('8.2 IPR')) +
    '<w:tbl><w:tr><w:tc>' + p(r('A.I.#')) + '</w:tc><w:tc>' + p(r('x')) + '</w:tc></w:tr></w:tbl>')), '')),
    [{ number: '8', title: 'MBS', level: 1 }, { number: '8.1', title: 'Opening', level: 2 }, { number: '8.2', title: 'IPR', level: 2 }]);
  check('another prefix and the Strict namespace are WordprocessingML too; anything else is refused', [
    texts(s.wordFileAgendaBlocks_('<x:document xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main"><x:body><x:p><x:r><x:t>strict</x:t></x:r></x:p></x:body></x:document>', '')),
    (() => { try { s.wordFileAgendaBlocks_('<document xmlns="urn:other"><body/></document>', ''); return 'accepted'; } catch (e) { return e.message; } })(),
    (() => { try { s.wordFileAgendaBlocks_(doc(p(r('a')) + '</w:tbl>'), styles); return 'accepted'; } catch (e) { return e.message; } })()
  ], [[['PARAGRAPH', 'strict', 'NORMAL']], 'word/document.xml is not WordprocessingML', 'malformed XML: unexpected </w:tbl>']);
  check('without styles.xml: no headings, no style lists, nothing else lost', texts(blocks(p(r('a'), '<w:pStyle w:val="Titre1"/>'), '')), [['PARAGRAPH', 'a', 'NORMAL']]);

  const dp = (text, extra) => ({ paragraph: Object.assign({ elements: [{ textRun: { content: text } }] }, extra || {}) });
  check('Docs API blocks: the line end dropped, a soft break as "\\r", bullets are lists, named styles are headings, other elements skipped', JSON.stringify(s.docsApiAgendaBlocks_([
    { sectionBreak: {} }, dp('a\u000bb\n', { paragraphStyle: { namedStyleType: 'HEADING_2' } }), dp('c\n', { bullet: { listId: 'l' } }), { tableOfContents: {} },
    { paragraph: { elements: [{ textRun: { content: 'x' } }, { inlineObjectElement: {} }, { textRun: { content: 'y\n' } }] } }])),
    JSON.stringify([{ type: 'OTHER' }, { type: 'PARAGRAPH', text: 'a\rb', heading: 'HEADING2' }, { type: 'LIST_ITEM', text: 'c', heading: 'NORMAL' }, { type: 'OTHER' }, { type: 'PARAGRAPH', text: 'xy', heading: 'NORMAL' }]));
  check('a Docs API answer without a body is a failure', (() => { try { s.docsApiAgendaBody_({}); return 'accepted'; } catch (e) { return e.message; } })(), 'the Docs API returned no document body');
}

// ====================================================================== I
console.log('I. diagnoseAgendaReaders()');
{
  const PROPS = { AGENDA_TDOC: 'S4-261857', FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', MEETING_TYPE: 'main', REPORT_SUFFIX: 'MBS' };
  function withNetwork(w) {
    w.fetched = [];
    const outer = makeBlob(Buffer.from('outer zip'), 'application/zip', 'S4-261857.zip');
    outer._outer = true;
    w.sandbox.UrlFetchApp = { fetch: (url) => { w.fetched.push(url); return { getResponseCode: () => 200, getBlob: () => outer }; } };
    const inner = w.sandbox.Utilities.unzip;
    w.sandbox.Utilities.unzip = (blob) => (blob._outer ? [makeBlob(DOCX_138, 'application/octet-stream', NAME_138)] : inner(blob));
    return w;
  }
  const w = withNetwork(world({ props: PROPS, documentApp: 'fail', docs: 'fail' }));
  const before = JSON.stringify(w.docProps._store);
  const text = w.sandbox.diagnoseAgendaReaders();
  check('each reader, its result; the Word file reads the 100 items', text.replace(/\(Google Doc \w+\)/g, '(Google Doc X)'),
    'Agenda S4-261857 (Google Doc X): DocumentApp: FAILED -- The agenda converted from Word (Google Doc X) could not be opened after 4 attempts over 7 s (DocumentApp.openById): ' + UNEXPECTED +
    '; Docs API: FAILED -- API call to docs.documents.get failed with error: Internal error encountered.; Word file: 100 agenda item(s).');
  check('one download, one conversion, trashed once at the end; no property written; no agenda text logged',
    [w.fetched, w.events.filter((e) => e === 'create').length, w.events[w.events.length - 1], JSON.stringify(w.docProps._store) === before, w.leaks()],
    [[PROPS.FTP_BASE + 'S4-261857.zip'], 1, 'trash ' + DOC_ID, true, []]);
  const all = withNetwork(world({ props: PROPS, documentApp: 'ok', docs: docsApiDocument(GOOGLE_138) }));
  check('when the readers agree, it says so', all.sandbox.diagnoseAgendaReaders().replace(/\(Google Doc \w+\)/g, '(Google Doc X)'),
    'Agenda S4-261857 (Google Doc X): DocumentApp: 100 agenda item(s); Docs API: 100 agenda item(s), same as DocumentApp; Word file: 100 agenda item(s), same as DocumentApp.');
  const differ = withNetwork(world({ props: PROPS, documentApp: 'ok', docs: docsApiDocument(GOOGLE_138.slice(0, 99)), items: GOOGLE_138 }));
  check('...and when they do not', /Docs API: 99 agenda item\(s\), DIFFERENT from DocumentApp; Word file: 100 agenda item\(s\), same as DocumentApp\.$/.test(differ.sandbox.diagnoseAgendaReaders()), true);
  const none = withNetwork(world({ props: { MEETING_TYPE: 'main' } }));
  check('without an agenda TDoc: nothing fetched, nothing converted', [none.sandbox.diagnoseAgendaReaders(), none.fetched, none.events], ['No agenda TDoc configured.', [], []]);
}

// ====================================================================== J
console.log('J. what changed in Code.js: against T-2026.10.20, one section and two functions');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutAgendaReaders(FULL);
  // Later changes (the TDoc-list download, after T-2026.10.21) taken out, as for every earlier layer.
  const CODE = codeWithoutTdocListDownload(FULL).code;
  check('every change is in Code.js exactly once', view.problems, []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(AGENDA_READERS_SECTION_START), FULL.indexOf(AGENDA_READERS_SECTION_END) + AGENDA_READERS_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), AGENDA_READERS_SECTION_FUNCTIONS);
  check('it stands directly after the AGENDA WORD CONVERSION section', FULL.indexOf('// END OF AGENDA WORD CONVERSION\n// =========================================================\n\n' + AGENDA_READERS_SECTION_START) !== -1, true);
  check('it declares its constants and nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [],
    ['var AGENDA_READERS_', 'var AGENDA_LAST_READER_', 'var AGENDA_DOCS_API_HEADINGS_', 'var AGENDA_WORD_MAIN_NAMESPACES_']);
  check('Code.js is version 2.26.6 (2.26.5 released it as T-2026.10.21, 2.26.6 the TDoc-list download fallback; the ledgers take the headers out)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.6');
  let released = null;
  try {
    released = execFileSync('git', ['show', 'template-release/T-2026.10.20:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    released = null;
  }
  if (released === null) {
    console.log('  note: template-release/T-2026.10.20 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js is the Code.js of T-2026.10.20, byte for byte', view.code === released, true);
    check('the functions that changed: parseAgendaWordBlob_ (the chain) and diagnoseAgendaConversion (names the reader)',
      functionNames(released).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(CODE, n) !== functionSource(released, n)).sort(), AGENDA_READERS_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions are those of the section; none removed', [functionNames(CODE).filter((n) => functionNames(released).indexOf(n) === -1).sort(), functionNames(released).filter((n) => functionNames(CODE).indexOf(n) === -1)],
      [AGENDA_READERS_SECTION_FUNCTIONS.slice().sort(), []]);
    check('untouched: the parser, the conversion, the open with its retries, the cleanup, the build, the agenda choice, the download, the TDoc lists, the update, Close Report',
      ['parseAgendaStructureWithText_', 'parseAgendaFromTables_', 'parseAgendaFromHeadings_', 'convertWordBlobToGoogleDoc_', 'convertAgendaWordBlob_', 'openConvertedAgendaBody_', 'trashConvertedAgenda_',
        'buildSkeletonWithTdocTables', 'parseAgendaForReport_', 'downloadMeetingAgenda_', 'extractWordFromZip_', 'downloadAndGroupTdocs_', 'continuousUpdateCore_', 'closeReportCore_', 'getReportConfig_']
        .map((n) => functionSource(CODE, n) === functionSource(released, n)), Array(15).fill(true));
  }
}

// ====================================================================== K
console.log('K. an empty or malformed agenda is not a successful read');
{
  const NOT_NUMBERED = [{ number: 'Note', title: 'not an agenda' }, { number: 'Annex', title: 'neither' }];
  const da = world({ documentApp: 'ok', documentAppItems: [], docs: docsApiDocument(GOOGLE_138) });
  const daOut = da.run();
  check('DocumentApp opens a document with no agenda in it: the Docs API is asked and reads the 100 items',
    [strip(daOut.result || []).length, da.sandbox.AGENDA_LAST_READER_, da.agendaLogs().some((l) => l.indexOf('[AGENDA] DocumentApp could not read the agenda (Google Doc ' + DOC_ID + '): Error: no agenda items found') === 0)],
    [100, 'Docs API', true]);
  const daBad = world({ documentApp: 'ok', documentAppItems: NOT_NUMBERED, docs: docsApiDocument(GOOGLE_138) });
  check('DocumentApp reads an "A.I.#" table with no agenda numbers in it: the Docs API is asked',
    [strip(daBad.run().result || []).length, daBad.sandbox.AGENDA_LAST_READER_, daBad.agendaLogs().some((l) => /DocumentApp could not read the agenda .*none of the 2 agenda item\(s\) is numbered like an agenda item/.test(l))],
    [100, 'Docs API', true]);
  const emptyDocs = world({ documentApp: 'fail', docs: { body: { content: [] } } });
  check('the Docs API returns an empty body: the Word file is asked and reads the 100 items',
    [strip(emptyDocs.run().result || []).length, emptyDocs.sandbox.AGENDA_LAST_READER_, emptyDocs.unzips], [100, 'Word file', 1]);
  const headerOnly = world({ documentApp: 'fail', docs: docsApiDocument([]) });
  check('the Docs API returns an "A.I.#" table with no rows: the Word file is asked', [strip(headerOnly.run().result || []).length, headerOnly.sandbox.AGENDA_LAST_READER_], [100, 'Word file']);
  const nothing = world({ documentApp: 'ok', documentAppItems: [], docs: { body: { content: [] } } });
  const nothingOut = nothing.run(Buffer.from('x'));
  check('no reader finds an agenda: an error naming each reader, no empty agenda returned; trashed once',
    [nothingOut.result, (nothingOut.error || '').split(' -- ')[0], /DocumentApp: no agenda items found.*; Docs API: no agenda items found.*; Word file: the Word file is not a \.docx/.test(nothingOut.error || ''), nothing.events.filter((e) => /^trash/.test(e)).length],
    [undefined, 'The agenda could not be read', true, 1]);
  const s = nothing.sandbox;
  const NONE = 'no agenda items found (no numbered paragraphs, no "A.I.#" table)';
  check('agendaItemsProblem_: an agenda; empty; not an array; incomplete items; no agenda numbers', [
    s.agendaItemsProblem_(strip(GOOGLE_138)),
    s.agendaItemsProblem_([]),
    s.agendaItemsProblem_(null),
    s.agendaItemsProblem_([{ number: '8.1', title: 'MBS', level: 2 }, { number: '8.2', title: ' ', level: 2 }, { number: '', title: 'x', level: 1 }, { number: '9', title: 'y', level: 0 }]),
    s.agendaItemsProblem_(NOT_NUMBERED.map((i) => Object.assign({ level: 1 }, i)))
  ], [null, NONE, NONE, '3 of 4 agenda item(s) without a number, title or level', 'none of the 2 agenda item(s) is numbered like an agenda item']);
  check('...the real agendas pass: SA4#138 and SA4#136 as Google converts them', [s.agendaItemsProblem_(GOOGLE_138), s.agendaItemsProblem_(GOOGLE_136)], [null, null]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll agenda reader checks passed.');
process.exit(failures ? 1 : 0);
