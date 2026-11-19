/**
 * T-2026.10.32 (Code.js 2.27.0) -- the Undisposed Documents section and the
 * local document availability (docs/UNDISPOSED_AND_LOCAL_DOCS.md):
 *
 *   1. numbering: X.1.4 in a standard report, X.1.5 when the report has an
 *      "X.1.4 Documents", 11.0.5 in 6G Media (11.0.4 is its IPR reminder);
 *      no other heading is renamed; a section under a number that is no
 *      longer free is renumbered; no "Registration of Documents": no place;
 *   2. what it lists: Status Available, Parked or Reserved and Plenary No or
 *      blank; the count; "None.";
 *   3. the writer on a document: created, unchanged (nothing written),
 *      a status and a new TDoc changed row by row, text typed above its
 *      table refuses, text typed below it is kept, the other tables (TDoc
 *      tables, the four disposition tables) are not touched, the links;
 *   4. the read of the Docs API: an unchanged section is not read through
 *      DocumentApp; a read that does not fit the document is not trusted;
 *   5. the row plan;
 *   6. the Docs folder: its address, the exact-TDoc rule, the listing's
 *      checks (stale, another folder, missing, unreadable), nothing confirmed
 *      is taken back;
 *   7. Reserved -> Available: only a TDoc the list has as reserved; a final
 *      disposition in the report is never changed (the Status sync rule).
 *
 * All TDocs, titles and names are synthetic.
 *
 * Run: node tests/undisposed-local-docs.test.js
 */

const crypto = require('crypto');
const { loadCode } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');

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

const NOTE = 'TDocs whose Status is Available, Parked or Reserved and whose Plenary is No or blank, in the order of the report. Written by every update; text typed below the table is kept.';
const HEADER = ['Document', 'Title', 'Source', 'Status', 'Agenda Item'];
const LINK = (id) => 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138/Docs/' + id + '.zip';
const entry = (tdoc, label, plenary, extra) => Object.assign({ tdoc: tdoc, title: 'Synthetic title of ' + tdoc, source: 'ExampleCorp', status: label, label: label, plenary: plenary, agendaItem: '7.5', link: LINK(tdoc) }, extra || {});

function sandbox(props) {
  const loaded = loadCode({ documentProperties: Object.assign({ MEETING_TYPE: 'main', REPORT_SUFFIX: 'Audio' }, props || {}) });
  const s = loaded.sandbox;
  s.Logger = { log: () => {} };
  return { s: s, docProps: loaded.docProps };
}

/** A report body: [heading level (0 = normal), text] or ['T', rows]. */
function reportBody(s, layout) {
  const body = makeFakeDocumentBody(s);
  const H = s.DocumentApp.ParagraphHeading;
  const LEVEL = { 1: H.HEADING1, 2: H.HEADING2, 3: H.HEADING3 };
  layout.forEach((item) => {
    if (item[0] === 'T') { body.appendTable(item[1]); return; }
    const p = body.appendParagraph(item[1]);
    p.setHeading(item[0] ? LEVEL[item[0]] : H.NORMAL);
  });
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  return body;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
const texts = (body) => body._children.map((c) => (c.getType() === 'TABLE' ? '[' + rowsOf(c)[0][0] + ' x' + (c.getNumRows() - 1) + ']' : c.getText()));
const snapshot = (body) => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
/** Everything of the report but the section (from its heading to the next heading). */
const outsideSection = (body) => { const out = []; let inside = false; body._children.forEach((c) => { const table = c.getType() === 'TABLE'; if (!table && c.getHeading() !== 'NORMAL') inside = / Undisposed Documents$/.test(c.getText()); if (!inside) out.push(table ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()]); }); return JSON.stringify(out); };
const sectionTable = (body) => body._children.find((c) => c.getType() === 'TABLE' && rowsOf(c)[0][0] === 'Document');
const items = (layout) => layout.map((x) => (x[0] === 'T' ? { kind: 't', header: x[1][0] } : { kind: 'p', heading: x[0], text: x[1] }));

const STANDARD = [
  [2, '7.1 Opening of the session, registration of documents'], [3, '7.1.1 Opening of the session'], [0, 'The chair opened the session (typed by hand).'],
  [3, '7.1.2 Registration of Documents'], ['T', [['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4-261801', 'Synthetic title of S4-261801', 'ExampleCorp', '7.5']]], [0, ''],
  [2, '7.2 IPR, antitrust and consensus reminder'], [0, 'The reminder was given.'],
  [2, '7.5 IVAS'], ['T', [['TDoc', 'S4-261801'], ['Title', 'Synthetic title of S4-261801'], ['Status', 'available'], ['Minutes', 'Minutes typed by hand.']]],
  [1, 'Disposition Summary'], [0, 'Agreed, Endorsed, Approved (0)'], [0, 'None.']
];
const SIXG = [
  [2, '11.0 Opening of the session'], [3, '11.0.1 Opening of the session'], [0, 'Opened.'], [3, '11.0.2 Registration of Documents'], ['T', [['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4-261901', 't', 'ExampleCorp', '11.3']]],
  [3, '11.0.3 Document Reallocations'], [0, 'No document reallocations for this meeting.'], [3, '11.0.4 IPR and antitrust reminder'], [0, 'The reminder was given.'], [2, '11.1 Approval of the agenda']
];

// ================================================================ 1
console.log('1. numbering: the next free number of the opening item, never below 4; no other heading renamed');
{
  const { s } = sandbox();
  const at = (layout) => { const p = s.locateUndisposedSection_(items(layout)); return [p.wanted, p.at, p.headingAt]; };
  check('a standard report: 7.1.4, at the end of the opening item (before 7.2)', at(STANDARD), ['7.1.4', 6, -1]);
  const withDocs = STANDARD.slice(0, 6).concat([[3, '7.1.3 Document Reallocations'], [0, 'None.'], [3, '7.1.4 Documents'], ['T', [['TDoc', 'S4-261800']]]], STANDARD.slice(6));
  check('with "7.1.3 Document Reallocations" and "7.1.4 Documents": 7.1.5, after them', at(withDocs), ['7.1.5', 10, -1]);
  check('6G Media: 11.0.5 (11.0.4 is the IPR reminder; X.0.4 is not free), after the reminder, before 11.1', at(SIXG), ['11.0.5', 9, -1]);
  const mis = STANDARD.slice(0, 6).concat([[3, '7.1.4 Undisposed Documents'], [0, NOTE], [3, '7.1.4 Documents']], STANDARD.slice(6));
  const p = s.locateUndisposedSection_(items(mis));
  check('a section under a number that is no longer free: found, to be renumbered (7.1.4 -> 7.1.5)', [p.number, p.wanted, p.headingAt], ['7.1.4', '7.1.5', 6]);
  const styled = STANDARD.slice(0, 6).concat([[2, '7.1.3 Document Reallocations'], [0, 'None.']], STANDARD.slice(6));
  check('a numbered heading styled at level 2 inside the item goes by its number: 7.1.4 after it', at(styled), ['7.1.4', 8, -1]);
  check('no "Registration of Documents": no place', s.locateUndisposedSection_(items([[2, '7.1 Opening'], [0, 'x']])).error, 'the report has no "Registration of Documents" heading');
  check('the heading texts: "7.1.4 Undisposed Documents" is the section, "7.1.4 Documents" is not', [s.undisposedHeadingNumber_('7.1.4 Undisposed Documents'), s.undisposedHeadingNumber_('7.1.4 Documents'), s.undisposedParentOfRegistration_('11.0.2 Registration of Documents')], ['7.1.4', null, '11.0']);
}

// ================================================================ 2
console.log('2. what it lists');
{
  const { s } = sandbox();
  const listed = [entry('A', 'available'), entry('B', 'parked', false), entry('C', 'reserved', null), entry('D', 'available', true), entry('E', 'agreed'), entry('F', 'noted'),
    entry('G', 'revised'), entry('H', 'withdrawn'), entry('I', 'available', false, { conflict: true }), entry('J', 'endorsed')].filter((e) => s.isUndisposedEntry_(e)).map((e) => e.tdoc);
  check('Available, Parked, Reserved with Plenary No or blank; not Plenary Yes, not a final disposition, not a conflicting read', listed, ['A', 'B', 'C']);
  const content = s.undisposedContent_([entry('S4-261801', 'available'), entry('S4-261802', 'agreed'), entry('S4-261803', 'reserved')]);
  check('the note, the count, the table (Document first: never read as a TDoc table), the links', content, [['p', NOTE], ['p', 'Not yet disposed, not going to plenary (2)'],
    ['t', [HEADER, ['S4-261801', 'Synthetic title of S4-261801', 'ExampleCorp', 'available', '7.5'], ['S4-261803', 'Synthetic title of S4-261803', 'ExampleCorp', 'reserved', '7.5']], [LINK('S4-261801'), LINK('S4-261803')]]]);
  check('none: "None.", no table', s.undisposedContent_([entry('X', 'agreed')]), [['p', NOTE], ['p', 'Not yet disposed, not going to plenary (0)'], ['p', 'None.']]);
}

// ================================================================ 3
console.log('3. the writer on a document');
{
  const { s } = sandbox();
  const body = reportBody(s, STANDARD);
  const before = snapshot(body);
  const e1 = [entry('S4-261801', 'available'), entry('S4-261802', 'agreed'), entry('S4-261803', 'reserved'), entry('S4-261804', 'parked', true)];
  const w1 = s.writeUndisposedDocuments_(body, e1);
  check('created: 7.1.4, 2 TDocs', [w1.state, w1.number, w1.count], ['created', '7.1.4', 2]);
  check('where: after the registration table, before 7.2; heading 3', [texts(body).slice(4, 11), body.getChild(6).getHeading()],
    [['[TDoc x1]', '', '7.1.4 Undisposed Documents', NOTE, 'Not yet disposed, not going to plenary (2)', '[Document x2]', '7.2 IPR, antitrust and consensus reminder'], 'H3']);
  check('nothing else of the report changed: headings, minutes, TDoc tables, the summary', outsideSection(body), before);
  const t = sectionTable(body);
  check('the table: its rows, a link on each Document, the header bold, the first column not "TDoc"', [rowsOf(t).map((r) => r[0]), t.getRow(1).getCell(0).editAsText().getLinkUrl(0), s.isTDocTable_(t)],
    [['Document', 'S4-261801', 'S4-261803'], LINK('S4-261801'), false]);

  const snap = snapshot(body);
  const calls = body._state.calls;
  const w2 = s.writeUndisposedDocuments_(body, e1);
  check('the same again: unchanged, nothing written', [w2.state, w2.written, snapshot(body) === snap], ['unchanged', false, true]);
  const readCalls = body._state.calls - calls;

  // A status changes (S4-261801 agreed), a TDoc comes (S4-261802 reserved, between the others), one goes to plenary.
  const e2 = [entry('S4-261801', 'agreed'), entry('S4-261802', 'reserved'), entry('S4-261803', 'reserved'), entry('S4-261804', 'parked', true), entry('S4-261805', 'available')];
  const w3 = s.writeUndisposedDocuments_(body, e2);
  check('a status and two TDocs: replaced row by row, the count follows', [w3.state, rowsOf(sectionTable(body)).map((r) => r[0]), texts(body)[8]],
    ['replaced', ['Document', 'S4-261802', 'S4-261803', 'S4-261805'], 'Not yet disposed, not going to plenary (3)']);
  check('the same table object is kept (row by row, not written again), the new row has its link', [sectionTable(body) === t, t.getRow(1).getCell(0).editAsText().getLinkUrl(0)], [true, LINK('S4-261802')]);
  check('still nothing else changed', outsideSection(body), before);

  // Text typed below the table is kept; text typed above it refuses.
  const tableAt = body.getChildIndex(sectionTable(body));
  body.insertParagraph(tableAt + 1, 'Chair: these are discussed on Wednesday (typed by hand).');
  const w4 = s.writeUndisposedDocuments_(body, e2.concat([entry('S4-261806', 'available')]));
  check('text typed below the table: kept, directly below it; the table updated', [w4.state, texts(body)[body.getChildIndex(sectionTable(body)) + 1], sectionTable(body).getNumRows()],
    ['replaced', 'Chair: these are discussed on Wednesday (typed by hand).', 5]);
  body.insertParagraph(7, 'Typed into the section, above its table.');
  const snap2 = snapshot(body);
  const w5 = s.writeUndisposedDocuments_(body, e1);
  check('text typed above its table: refused, nothing written', [w5.state, w5.foreign, snapshot(body) === snap2], ['refused', ['Typed into the section, above its table.'], true]);
  body.getChild(7).removeFromParent();

  const w6 = s.writeUndisposedDocuments_(body, [entry('S4-261801', 'agreed')]);
  check('none left: "None.", the table gone, the typed text kept', [w6.state, texts(body).slice(6, 11)], ['replaced', ['7.1.4 Undisposed Documents', NOTE, 'Not yet disposed, not going to plenary (0)', 'None.', 'Chair: these are discussed on Wednesday (typed by hand).']]);
  const w7 = s.writeUndisposedDocuments_(body, e1);
  check('and back: the table again', [w7.state, rowsOf(sectionTable(body)).map((r) => r[0])], ['replaced', ['Document', 'S4-261801', 'S4-261803']]);

  // A "7.1.4 Documents" comes before it (a TDoc allocated before the agenda sections): the section becomes 7.1.5.
  const h = body.insertParagraph(6, '7.1.4 Documents');
  h.setHeading(s.DocumentApp.ParagraphHeading.HEADING3);
  body.insertTable(7, [['TDoc', 'S4-261800']]);
  const w8 = s.writeUndisposedDocuments_(body, e1);
  check('renumbered 7.1.5 when "7.1.4 Documents" comes before it; "7.1.4 Documents" is not renamed', [w8.state, texts(body)[6], texts(body)[8]], ['renumbered', '7.1.4 Documents', '7.1.5 Undisposed Documents']);
  console.log('  (an unchanged section read through DocumentApp: ' + readCalls + ' calls)');
}
{
  const { s } = sandbox({ REPORT_SUFFIX: '6G' });
  const body = reportBody(s, SIXG);
  const before = snapshot(body);
  const w = s.writeUndisposedDocuments_(body, [entry('S4-261901', 'available', false, { agendaItem: '11.3' })]);
  check('6G Media: 11.0.5, after the IPR reminder, before 11.1; nothing renamed', [w.state, texts(body).slice(7, 14), outsideSection(body) === before],
    ['created', ['11.0.4 IPR and antitrust reminder', 'The reminder was given.', '11.0.5 Undisposed Documents', NOTE, 'Not yet disposed, not going to plenary (1)', '[Document x1]', '11.1 Approval of the agenda'], true]);
}
{
  const { s } = sandbox();
  const body = reportBody(s, [[2, '7.1 Opening'], [0, 'x']]);
  const before = snapshot(body);
  check('no "Registration of Documents": no place, nothing written', [s.writeUndisposedDocuments_(body, [entry('A', 'available')]).state, snapshot(body) === before], ['no-place', true]);
}

// ================================================================ 4
console.log('4. the read of the Docs API');
{
  const { s } = sandbox();
  const body = reportBody(s, STANDARD);
  const e1 = [entry('S4-261801', 'available'), entry('S4-261803', 'reserved')];
  s.writeUndisposedDocuments_(body, e1);
  // The Docs API read of the same document (as the API gives it: heading styles, a section break, cells with a line end).
  const STYLE = { H1: 'HEADING_1', H2: 'HEADING_2', H3: 'HEADING_3', NORMAL: 'NORMAL_TEXT' };
  const run = (text, link) => ({ textRun: Object.assign({ content: text }, link ? { textStyle: { link: { url: link } } } : {}) });
  const cell = (text, link) => ({ content: [{ paragraph: { elements: [run(text, link), run('\n')] } }] });
  const api = (b) => [{ sectionBreak: {} }].concat(b._children.map((c) => (c.getType() === 'TABLE'
    ? { table: { tableRows: rowsOf(c).map((r, ri) => ({ tableCells: r.map((x, ci) => cell(x, ci === 0 && ri > 0 ? c.getRow(ri).getCell(0).editAsText().getLinkUrl(0) : '')) })) } }
    : { paragraph: { paragraphStyle: { namedStyleType: STYLE[c.getHeading()] }, elements: [run(c.getText() + '\n')] } })));
  const read = s.undisposedOwnedOfContent_(api(body));
  check('the API read gives the section as DocumentApp has it', [read.number, read.wanted, JSON.stringify(read.owned) === JSON.stringify(s.undisposedContent_(e1))], ['7.1.4', '7.1.4', true]);
  const fresh = api(body);
  const calls = body._state.calls;
  const w = s.writeUndisposedDocuments_(body, e1, fresh);
  check('unchanged by the read: no DocumentApp call at all', [w.state, w.number, body._state.calls - calls], ['unchanged', '7.1.4', 0]);
  const stale = api(body);
  s.writeUndisposedDocuments_(body, e1.concat([entry('S4-261805', 'available')]));
  const w2 = s.writeUndisposedDocuments_(body, e1.concat([entry('S4-261806', 'available')]), stale);
  check('a read that no longer fits the document (another number of rows): not trusted, the table is right', [w2.state, rowsOf(sectionTable(body)).map((r) => r[0])], ['replaced', ['Document', 'S4-261801', 'S4-261803', 'S4-261806']]);
  check('text typed into the section in the read: no decision from the read', s.undisposedOwnedOfContent_(api(body).slice(0, 8).concat([{ paragraph: { paragraphStyle: { namedStyleType: 'NORMAL_TEXT' }, elements: [run('typed\n')] } }], api(body).slice(8))), null);
}

// ================================================================ 4b
console.log('4b. an update without a read to share (text statuses, no Disposition Summary): one Docs API read, or the 15-minute guard');
{
  const STYLE = { H1: 'HEADING_1', H2: 'HEADING_2', H3: 'HEADING_3', NORMAL: 'NORMAL_TEXT' };
  const run = (text) => ({ textRun: { content: text } });
  const cell = (text) => ({ content: [{ paragraph: { elements: [run(text), run('\n')] } }] });
  const api = (b) => b._children.map((c) => (c.getType() === 'TABLE'
    ? { table: { tableRows: rowsOf(c).map((r) => ({ tableCells: r.map((x) => cell(x)) })) } }
    : { paragraph: { paragraphStyle: { namedStyleType: STYLE[c.getHeading()] }, elements: [run(c.getText() + '\n')] } }));
  const setup = (withApi) => {
    const { s } = sandbox();
    const body = reportBody(s, STANDARD);
    let reads = 0;
    let gets = 0;
    const realRead = s.readReportDispositions_;
    s.readReportDispositions_ = function () { reads++; return realRead.apply(this, arguments); };
    s.getMeetingContext_ = () => ({ meeting: { type: 'main' } });
    s.plenarySelectionEnabled_ = () => false;
    s.reportRunSnapshot_ = () => null;
    s.reportUpdateStageFailed_ = () => {};
    s.currentReportDocument_ = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000' });
    s.statusDropdownsAvailable_ = () => withApi;
    let fixed = null;
    s.plenarySnapshot_ = () => { gets++; const content = fixed || api(body); fixed = null; return { content: content }; };
    return { s: s, body: body, reads: () => reads, gets: () => gets, readNow: () => { fixed = api(body); } };
  };
  {
    const x = setup(true);
    const first = x.s.syncUndisposedDocumentsOnUpdate_(x.body);
    check('with the Docs API: created from one read of its own; the TDoc tables are not read through DocumentApp', [first.state, first.number, first.count, x.gets(), x.reads()], ['created', '7.1.4', 1, 1, 0]);
    x.readNow();
    const calls = x.body._state.calls;
    const second = x.s.syncUndisposedDocumentsOnUpdate_(x.body);
    check('the next update: unchanged from its read, no DocumentApp call', [second.state, x.gets(), x.reads(), x.body._state.calls - calls], ['unchanged', 2, 0, 0]);
    x.body._children.find((c) => c.getType() === 'TABLE' && rowsOf(c)[0][1] === 'S4-261801').getRow(2).getCell(1).setText('agreed');
    const third = x.s.syncUndisposedDocumentsOnUpdate_(x.body);
    check('a status typed by hand (Agreed): the row goes, from the read', [third.state, x.body._children.some((c) => c.getText && c.getText() === 'None.')], ['replaced', true]);
  }
  {
    const x = setup(false);
    const realNow = Date.now();
    const due = (mins) => x.s.undisposedDocumentAppReadDue_(x.body, new Date(realNow + mins * 60000));
    check('without the Docs API: the first update reads (no section yet)', [x.s.syncUndisposedDocumentsOnUpdate_(x.body).state, x.reads()], ['created', 1]);
    const calls = x.body._state.calls;
    const later = x.s.syncUndisposedDocumentsOnUpdate_(x.body);
    check('an update right after: left as it is, the TDoc tables not read', [later.state, x.reads(), x.body._state.calls - calls < 100], ['not-due', 1, true]);
    check('the guard: not due within 15 minutes, due after them', [due(14), due(16)], [false, true]);
    check('due when the clock went back, and while the report has no section', (() => { const back = due(-60); x.body._children.filter((c) => c.getType() !== 'TABLE' && / Undisposed Documents$/.test(c.getText())).forEach((h) => h.setText('removed by hand')); return [back, due(-59)]; })(), [true, true]);
  }
}

// ================================================================ 5
console.log('5. the row plan');
{
  const { s } = sandbox();
  const r = (id, st) => [[id, 't', 'src', st || 'available', '7.5'], LINK(id)];
  const ops = (have, want) => s.planUndisposedRowEdits_(have, want).map((e) => e.op + '@' + e.at);
  check('the same: all kept', ops([r('A'), r('B')], [r('A'), r('B')]), ['keep@0', 'keep@1']);
  check('one in the middle comes', ops([r('A'), r('C')], [r('A'), r('B'), r('C')]), ['keep@0', 'insert@1', 'keep@2']);
  check('one goes, one changes', ops([r('A'), r('B'), r('C')], [r('A'), r('C', 'reserved')]), ['keep@0', 'remove@1', 'set@1']);
  check('reordered', ops([r('A'), r('B')], [r('B'), r('A')]), ['insert@0', 'keep@1', 'remove@2']);
  check('all new', ops([r('A')], [r('B')]), ['remove@0', 'insert@0']);
}

// ================================================================ 6
console.log('6. the Docs folder of the meeting-room server');
const DOCS = 'http://10.10.10.10/ftp/SA/SA4/Docs/';
const BASE = 'http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/Audio/';
const sha = (t) => crypto.createHash('sha256').update(String(t), 'utf8').digest('hex');
const NOW = new Date('2026-11-19T17:00:00Z');
const minutesAgo = (m) => new Date(NOW.getTime() - m * 60000).toISOString();
const listing = (names, extra) => {
  const payload = '<html><body>' + names.map((n) => '<a href="' + n + '">' + n + '</a>').join('\n') + '</body></html>';
  return JSON.stringify(Object.assign({ format: 'sa4-local-listing', version: 1, group: 'Docs', folder: DOCS, fetchedAt: minutesAgo(2), scanId: sha(payload), agent: 'ROOM-LAPTOP', linkCount: names.length, sha256: sha(payload), payload: payload }, extra || {}));
};
{
  const { s } = sandbox();
  check('the Docs folder of a Local Revisions URL', [s.localDocsFolderOf_(BASE), s.localDocsFolderOf_('http://10.10.10.10/ftp/SA/SA4/Inbox'), s.localDocsFolderOf_('http://10.10.10.10/ftp/other/')], [DOCS, DOCS, '']);
  const found = s.localDocsPresentIn_('<a href="S4-261801.zip">x</a><a href="s4-261802.ZIP">x</a><a href="S4-261803_r1.zip">x</a><a href="S4-261804r1.docx">x</a><a href="S4-2618050.zip">x</a>' +
    '<a href="sub/S4-261806.zip">x</a><a href="http://10.10.10.10/ftp/SA/SA4/Inbox/S4-261807.zip">x</a><a href="/ftp/SA/SA4/Docs/S4-261808.zip">x</a><a href="../">up</a>', DOCS);
  check('exact TDoc numbers only, directly in the folder (not _r1, not r1, not 7 digits, not a subfolder or another folder)', Object.keys(found.present).sort(), ['S4-261801', 'S4-261802', 'S4-261808']);
}
function docsReport(props) {
  const { s, docProps } = sandbox(Object.assign({ LOCAL_LISTING_DRIVE_FOLDER_ID: 'FOLDERid0123456789' }, props || {}));
  const deps = (files, extra) => Object.assign({ now: NOW, localBase: BASE, sha256: sha, read: () => ({ files: files.map((raw) => ({ raw: raw, updated: NOW })), status: null }) }, extra || {});
  return { s: s, docProps: docProps, deps: deps };
}
{
  const { s, docProps, deps } = docsReport();
  const r1 = s.importLocalDocsListing_(null, ['S4-261801', 'S4-261802', 'S4-261803'], deps([listing(['S4-261801.zip', 'S4-261803_r1.zip'])]));
  check('a fresh listing: the TDocs of this report it has are confirmed and remembered', [r1.result, r1.added, docProps.getProperty('LOCAL_DOCS_CONFIRMED')], ['imported', ['S4-261801'], 'S4-261801']);
  const r2 = s.importLocalDocsListing_(null, ['S4-261801', 'S4-261802'], deps([listing(['S4-261802.zip'], { fetchedAt: minutesAgo(45) })]));
  check('a stale listing (agent stopped, server down): nothing new, nothing taken back', [r2.result, r2.added, Object.keys(r2.confirmed)], ['stale', [], ['S4-261801']]);
  const r3 = s.importLocalDocsListing_(null, ['S4-261802'], deps([listing(['S4-261802.zip'], { folder: 'http://10.10.10.10/ftp/SA/SA4/Inbox/Docs/' })]));
  check('a listing of another folder: not used', [r3.result, r3.added], ['invalid', []]);
  const r4 = s.importLocalDocsListing_(null, ['S4-261802'], deps([listing(['S4-261802.zip'], { sha256: sha('other') })]));
  check('an incomplete listing (checksum): not used', [r4.result, r4.added], ['invalid', []]);
  const r5 = s.importLocalDocsListing_(null, ['S4-261802'], deps([]));
  const r6 = s.importLocalDocsListing_(null, ['S4-261802'], deps([], { read: () => { throw new Error('Drive unavailable'); } }));
  check('no listing yet / the Drive folder cannot be read: nothing changed', [r5.result, r6.result, docProps.getProperty('LOCAL_DOCS_CONFIRMED')], ['missing', 'unavailable', 'S4-261801']);
  check('what the last check found, for Check Connections', /^Local Docs folder \(http:\/\/10\.10\.10\.10\/ftp\/SA\/SA4\/Docs\/\): unavailable at .* Confirmed present so far: 1 TDoc\(s\)\.$/.test(s.describeLocalDocsStatus_(docProps.getProperty('LOCAL_DOCS_AGENT_STATE'))), true);
}
{
  const { s } = sandbox();
  check('no agent folder configured: off, nothing read', s.importLocalDocsListing_(null, ['S4-261801'], { read: () => { throw new Error('must not be read'); } }), null);
}

// ================================================================ 7
console.log('7. Reserved -> Available');
{
  const { s, docProps, deps } = docsReport();
  const groups = () => ({ '7.5': { tdocs: ['reserved', 'reserved', 'available', 'agreed', 'withdrawn', 'reserved'].map((st, k) => ({ row: ['S4-26180' + (k + 1), 'Synthetic', st], tdocCol: 0, statusCol: 2 })) } });
  const g = groups();
  const files = listing(['S4-261801.zip', 'S4-261803.zip', 'S4-261804.zip', 'S4-261805.zip', 'S4-261806_r1.zip']);
  const r = s.applyLocalDocsAvailability_(g, null, deps([files]));
  check('reserved and present: available; reserved and absent, available, agreed, withdrawn: as the list says', [r.changed, g['7.5'].tdocs.map((t) => t.row[2])],
    [['S4-261801'], ['available', 'reserved', 'available', 'agreed', 'withdrawn', 'reserved']]);
  check('only this run\'s row is changed, and marked', [g['7.5'].tdocs[0].localDocAvailable, groups()['7.5'].tdocs[0].row[2]], [true, 'reserved']);
  const g2 = groups();
  const r2 = s.applyLocalDocsAvailability_(g2, null, deps([listing(['S4-261802.zip'], { fetchedAt: minutesAgo(90) })]));
  check('the server down later (stale listing): the confirmed TDoc stays available, nothing reverts', [r2.state, g2['7.5'].tdocs.map((t) => t.row[2]).slice(0, 2)], ['stale', ['available', 'reserved']]);
  check('without an agent folder: off, the list as it is', (() => { const x = sandbox(); const gg = groups(); return [x.s.applyLocalDocsAvailability_(gg, null, {}).state, gg['7.5'].tdocs[0].row[2]]; })(), ['off', 'reserved']);
  check('a failure changes nothing and never throws', (() => { const gg = groups(); s.importLocalDocsListing_ = () => { throw new Error('boom'); }; return [s.applyLocalDocsAvailabilitySafely_(gg), gg['7.5'].tdocs[0].row[2]]; })(), [null, 'reserved']);
  void docProps;
}
{
  // The Status sync takes the list's "available": it moves a report's Reserved, never a final disposition or the minutes.
  const { s } = sandbox();
  const body = reportBody(s, [[2, '7.5 IVAS'], ['T', [['TDoc', 'S4-261801'], ['Title', 't'], ['Status', 'Reserved'], ['Minutes', 'Typed minutes.']]], ['T', [['TDoc', 'S4-261802'], ['Title', 't'], ['Status', 'Agreed'], ['Minutes', 'Agreed with changes (typed).']]]]);
  const [t1, t2] = body._children.filter((c) => c.getType() === 'TABLE');
  s.applyTdocStatusUpdate_(t1, 'S4-261801', 'available');
  s.applyTdocStatusUpdate_(t2, 'S4-261802', 'available');
  check('Reserved -> Available; Agreed stays Agreed; the minutes untouched', [rowsOf(t1)[2][1], rowsOf(t2)[2][1], rowsOf(t1)[3][1], rowsOf(t2)[3][1]], ['available', 'Agreed', 'Typed minutes.', 'Agreed with changes (typed).']);
}

// ================================================================ 8
console.log('8. what changed in Code.js, and nothing else (T-2026.10.31 kept, with its concurrent-edit fixes)');
{
  const fs = require('fs');
  const path = require('path');
  const { execFileSync } = require('child_process');
  const ledger = require('./helpers/before-structure-fix.js');
  const ROOT = path.join(__dirname, '..');
  const FULL = fs.readFileSync(path.join(ROOT, 'Code.js'), 'utf8').replace(/\r/g, '');
  const view = ledger.codeWithoutUndisposedDocuments(FULL);
  check('every change is in Code.js exactly once', view.problems, []);
  const names = (src) => (src.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (src, name) => { const at = src.indexOf('\nfunction ' + name + '('); return at === -1 ? null : src.slice(at, src.indexOf('\n}\n', at)); };
  const section = FULL.slice(FULL.indexOf(ledger.UNDISPOSED_SECTION_START), FULL.indexOf(ledger.UNDISPOSED_SECTION_END) + ledger.UNDISPOSED_SECTION_END.length);
  check('the section: these functions, in this order', names(section), ledger.UNDISPOSED_SECTION_FUNCTIONS);
  check('it stands directly after Goes to Plenary', FULL.indexOf('// END OF GOES TO PLENARY\n// =========================================================\n\n' + ledger.UNDISPOSED_SECTION_START) !== -1, true);
  check('the released functions it changed', names(view.code).filter((n) => bodyOf(FULL, n) !== bodyOf(view.code, n)).sort(), ledger.UNDISPOSED_CHANGED_FUNCTIONS.slice().sort());
  check('the new functions are those of the section; none removed', [names(FULL).filter((n) => names(view.code).indexOf(n) === -1).sort(), names(view.code).filter((n) => names(FULL).indexOf(n) === -1)],
    [ledger.UNDISPOSED_SECTION_FUNCTIONS.slice().sort(), []]);
  check('Code.js is version 2.27.0 (T-2026.10.32)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.27.0');
  let released = null;
  try {
    released = execFileSync('git', ['show', 'af08ae1:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    released = null;
  }
  if (released === null) console.log('  note: commit af08ae1 (T-2026.10.31) is not available; the byte-for-byte comparison is skipped.');
  else check('without it, Code.js is that of T-2026.10.31 (af08ae1), byte for byte', view.code === released, true);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nall checks passed');
process.exit(failures ? 1 : 0);
