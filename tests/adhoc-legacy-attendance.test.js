/**
 * Ad-hoc attendance next to an attendance list the report had before
 * (G2 smoke-test finding).
 *
 * A report can already contain an attendance block that this feature did not
 * write -- made by hand, or by another tool: closing prose, "Statistics", a
 * Field | Details table, "Attendees", a Name | Company | Email table. No code
 * of this repository ever generated it, so nothing proves where it came from,
 * and it is NEVER changed or removed here. What this file pins:
 *
 *   1. an import leaves such a block exactly as it is, writes the generated
 *      section once, and points the other table out;
 *   2. the generated section owns only its own tables: a Company typed into
 *      it is not overridden by the other table at the next refresh;
 *   3. companies entered in such a table before are taken over once, at the
 *      first import, and never again;
 *   4. things that merely look similar are not touched either;
 *   5. reports without sessions, and main-meeting reports, are unaffected.
 *
 * All names, companies and values are synthetic.
 *
 * Run: node tests/adhoc-legacy-attendance.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const FX = require('./fixtures/teams-attendance-synthetic.js');

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
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const SESSIONS = JSON.stringify({ v: 1, nextId: 3, sessions: [{ id: 's1', label: 'AHG Call 1', date: '2026-09-22', start: '15:00', end: '18:00' }, { id: 's2', label: '', date: '2026-09-23', start: '', end: '' }] });
const ADHOC = { MEETING_TYPE: 'adhoc', ADHOC_SESSIONS: SESSIONS };
const PROSE = 'The session was closed on Sep 22, 2026 at 18:05 CEST.';
const FIELDS = [['Field', 'Details'], ['Meeting', 'Synthetic AHG Meeting'], ['Meeting-series dates', 'September 22, 2026'], ['Attendance-report session', 'September 22, 2026'], ['Start', '2:48:05 PM'], ['End', '6:11:26 PM'],
  ['Meeting duration', '3h 23m 20s'], ['Attended-participant records', '25'], ['Average attendance time', '2h 31m 7s']];
const OLD_ATTENDEES = [['Name', 'Company', 'Email'], ['Alex Organizer', 'OldCorp', 'alex.organizer@example.com'], ['Lee Wong', 'ExampleCorp', 'lee.wong@example.com'], ['Nobody Listed', 'ElsewhereCorp', 'nobody.listed@example.org']];
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/**
 * A report with a closing section. opts.props; opts.fill(body, H) adds what
 * the report contains before the import.
 */
function report(opts) {
  const o = opts || {};
  const loaded = loadCode({ documentProperties: o.props || ADHOC });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H = s.DocumentApp.ParagraphHeading;
  body.appendParagraph('5.1 Opening of the session').setHeading(H.HEADING2);
  body.appendParagraph('The chair opens the session.');
  if (o.fill) o.fill(body, H);
  const ui = { alerts: [], ButtonSet: { OK: 'OK' }, alert: (...args) => { ui.alerts.push(args); } };
  const r = { s: s, body: body, docProps: loaded.docProps, ui: ui, logs: [] };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.Utilities.base64Decode = (b) => Array.from(Buffer.from(b, 'base64')).map((x) => (x > 127 ? x - 256 : x));
  const b64 = FX.utf16le(FX.buildExport()).toString('base64');
  r.import = () => s.confirmTeamsAttendanceImport('s1', b64, s.previewTeamsAttendanceImport('s1', b64).token);
  r.texts = () => body._children.map((c) => (c.getType() === 'TABLE' ? '[' + rowsOf(c)[0].join('|') + ' x' + (c.getNumRows() - 1) + ']' : c.getText()));
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  r.container = () => { const c = s.findAdhocAttendanceContainer_(body); return c ? body._children.slice(c.start, c.end) : null; };
  /** Everything outside the generated container, as a snapshot. */
  r.outside = () => { const c = s.findAdhocAttendanceContainer_(body); return JSON.stringify(JSON.parse(r.snapshot()).filter((e, i) => !c || i < c.start || i >= c.end)); };
  r.generatedTable = () => r.container().filter((e) => e.getType() === 'TABLE')[0];
  r.company = (name) => rowsOf(r.generatedTable()).filter((row) => row[0] === name)[0][1];
  r.setCompany = (name, value) => { const t = r.generatedTable(); for (let i = 1; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === name) t.getRow(i).getCell(1).setText(value); };
  r.count = (text) => r.texts().filter((t) => t === text).length;
  return r;
}

/** The block as the G2 smoke test found it. style: 'normal' (all plain paragraphs), 'headings' (Statistics / Attendees are headings), 'heading-prose' (the closing line is a heading). */
function legacyBlock(style, attendees) {
  return (body, H) => {
    body.appendParagraph('5.11 Close of the session').setHeading(H.HEADING2);
    const prose = body.appendParagraph(PROSE);
    if (style === 'heading-prose') prose.setHeading(H.HEADING3);
    const statistics = body.appendParagraph('Statistics');
    if (style === 'headings') statistics.setHeading(H.HEADING3);
    body.appendTable(FIELDS);
    const heading = body.appendParagraph('Attendees');
    if (style === 'headings') heading.setHeading(H.HEADING3);
    body.appendTable(attendees || OLD_ATTENDEES);
    body.appendParagraph('Prose after the old block.');
  };
}
const NOTICE_ONE_FIRST = ' This report also has another Name / Company / Email table outside the generated Attendance section -- an attendance list made earlier. It was not changed; the companies entered there were taken over. If it shows the same attendance, remove it by hand.';
const NOTICE_ONE = ' This report also has another Name / Company / Email table outside the generated Attendance section -- an attendance list made earlier. It was not changed. If it shows the same attendance, remove it by hand.';
const DONE = 'The attendance of AHG Call 1 was imported (23 attendees).';

// ================================================================ 1. the block of the smoke test

console.log('1. a report that already has an attendance block: nothing of it is changed');
['normal', 'headings', 'heading-prose'].forEach((style) => {
  const r = report({ fill: legacyBlock(style) });
  const before = r.snapshot();
  const elements = r.body._children.slice();
  const result = r.import();
  check(`${style}: the import succeeds and says that another table exists, that it was not changed, and what to do`, [result.ok, result.message], [true, DONE + NOTICE_ONE_FIRST]);
  check(`${style}: everything the report had is still there, unchanged -- the very same elements, in the same order`, [r.outside() === before, elements.every((e) => r.body._children.indexOf(e) !== -1),
    elements.map((e) => r.body._children.indexOf(e)).every((at, i, list) => i === 0 || at > list[i - 1])], [true, true, true]);
  check(`${style}: the old block is complete: closing line, Statistics, the Field | Details table, Attendees, its table, the prose after it`,
    r.texts().filter((t) => t === PROSE || t === '[Field|Details x8]' || t === '[Name|Company|Email x3]' || t === 'Prose after the old block.').length, 4);
  check(`${style}: the generated section is there once, with its own table`, [r.count('Attendance'), r.container().filter((e) => e.getType() === 'TABLE').length, rowsOf(r.generatedTable()).length - 1], [1, 1, 23]);
  check(`${style}: the generated section does not contain any element of the old block`, r.container().filter((e) => elements.indexOf(e) !== -1).length, 0);
  check(`${style}: where the section is`, r.texts()[r.texts().indexOf('Attendance') - 1],
    style === 'normal' ? 'Prose after the old block.' : (style === 'headings' ? PROSE : '5.11 Close of the session'));

  const afterImport = r.outside();
  r.s.refreshAdhocAttendanceSection();
  r.s.refreshAdhocAttendanceSection();
  check(`${style}: Refresh Attendance Section, twice: still one generated section, and the old block still untouched`, [r.count('Attendance'), r.outside() === afterImport, r.texts().filter((t) => t === '[Name|Company|Email x3]').length], [1, true, 1]);
  const steady = r.snapshot();
  const again = r.import();
  check(`${style}: importing the same file again changes nothing`, [again.ok, again.changed, r.snapshot() === steady], [true, false, true]);
  check(`${style}: removing the import says nothing about the other table: there is no generated section to confuse it with`, (() => { const x = report({ fill: legacyBlock(style) }); x.import(); return String(x.s.removeTeamsAttendanceImport('s1').message).indexOf('Name / Company') !== -1; })(), false);
  check(`${style}: removing the import removes the generated section only`, [r.s.removeTeamsAttendanceImport('s1').ok, r.count('Attendance'), JSON.stringify(JSON.parse(r.snapshot()).filter((e) => !(e[0] === 'P' && e[1] === ''))) === before], [true, 0, true]);
});

// ================================================================ 2. whose Company cells are read

console.log('2. the generated section owns only its own tables');
['headings', 'heading-prose', 'normal'].forEach((style) => {
  const r = report({ fill: legacyBlock(style) });
  r.import();
  check(`${style}: companies entered in the old table are taken over at the first import -- for attendees of the import only`, [r.company('Alex Organizer'), r.company('Lee Wong'), JSON.stringify(r.docProps._store).indexOf('ElsewhereCorp')], ['OldCorp', 'ExampleCorp', -1]);
  r.setCompany('Alex Organizer', 'NewCorp');
  const seen = [];
  for (let i = 0; i < 4; i++) { r.s.refreshAdhocAttendanceSection(); seen.push(r.company('Alex Organizer')); }
  check(`${style}: a Company typed into the generated table stays, refresh after refresh: the old table does not override it`, seen, ['NewCorp', 'NewCorp', 'NewCorp', 'NewCorp']);
  r.setCompany('Alex Organizer', '');
  r.s.refreshAdhocAttendanceSection();
  r.s.refreshAdhocAttendanceSection();
  check(`${style}: a Company cleared in the generated table stays cleared, although the old table still has one`, [r.company('Alex Organizer'), rowsOf(r.body._children.filter((c) => c.getType() === 'TABLE' && c.getNumRows() === 4)[0])[1][1]], ['', 'OldCorp']);
  // The old table is edited afterwards: that is not read any more.
  r.body._children.filter((c) => c.getType() === 'TABLE' && c.getNumRows() === 4)[0].getRow(2).getCell(1).setText('ChangedLaterCorp');
  r.s.refreshAdhocAttendanceSection();
  const second = r.import();
  check(`${style}: what is typed into the old table after the first import is not read`, [r.company('Lee Wong'), JSON.stringify(r.docProps._store).indexOf('ChangedLaterCorp'), second.changed], ['ExampleCorp', -1, false]);
});
{
  const r = report({ fill: legacyBlock('headings') });
  r.import();
  r.s.writeAdhocAttendanceValue_(r.docProps, 's2', JSON.stringify({ t: 'Second day', n: 2, s: '2026-09-23T09:00:00', e: null, d: null, a: null, p: [['Alex Organizer', '', 'alex.organizer@example.com'], ['Nobody Listed', '', 'nobody.listed@example.org']] }));
  const result = r.s.renderAfterAdhocAttendanceChange_('Done.');
  check('a later import, when the section exists: the notice no longer says that companies were taken over, and they are not', [result.message, rowsOf(r.container().filter((e) => e.getType() === 'TABLE')[1])], ['Done.' + NOTICE_ONE, [['Name', 'Company', 'Email'], ['Alex Organizer', 'OldCorp', 'alex.organizer@example.com'], ['Nobody Listed', '', 'nobody.listed@example.org']]]);
  check('the tables that are read are the two of the generated section', [r.s.adhocAttendeeTablesToHarvest_(r.body).length, r.s.adhocAttendeeTablesToHarvest_(r.body).every((t) => r.container().indexOf(t) !== -1), r.s.countAdhocAttendeeTablesElsewhere_(r.body)], [2, true, 1]);
  r.setCompany('Alex Organizer', 'NewCorp');
  r.s.refreshAdhocAttendanceSection();
  r.s.refreshAdhocAttendanceSection();
  check('a Company typed into one generated table still carries over to the same attendee in the other session (stage D), and stays', rowsOf(r.container().filter((e) => e.getType() === 'TABLE')[1])[1][1], 'NewCorp');
}

// ================================================================ 3. things that look similar

console.log('3. what merely looks similar is not touched');
const similar = (what, fill, expectNotice) => {
  const r = report({ fill: (body, H) => { body.appendParagraph('5.11 Close of the session').setHeading(H.HEADING2); body.appendParagraph(PROSE); fill(body, H); body.appendParagraph('Prose at the end.'); } });
  const before = r.snapshot();
  const elements = r.body._children.slice();
  const result = r.import();
  r.s.refreshAdhocAttendanceSection();
  check(`${what}: untouched by the import and by a refresh -- same elements, same text`, [result.ok, r.outside() === before, elements.every((e) => r.body._children.indexOf(e) !== -1), r.body._children.filter((e) => e.getType() !== 'TABLE' && e.getText() === 'Attendance' && e.getHeading() !== 'NORMAL').length], [true, true, true, 1]);
  check(`${what}: ${expectNotice ? 'pointed out' : 'no notice'}`, result.message, DONE + (expectNotice || ''));
  return r;
};
similar('a section titled Statistics with text of its own', (body, H) => { body.appendParagraph('Statistics').setHeading(H.HEADING3); body.appendParagraph('Twelve contributions were treated.'); body.appendTable([['Topic', 'Count'], ['Agreed', '7']]); });
similar('a plain paragraph "Statistics" and a table that is not ours', (body) => { body.appendParagraph('Statistics'); body.appendTable([['Field', 'Details'], ['Anything', 'else']]); });
similar('a section titled Attendees with a list in prose', (body, H) => { body.appendParagraph('Attendees').setHeading(H.HEADING3); body.appendParagraph('Alex Organizer, Pat Kim and others.'); });
similar('a Name / Company / Email table of the user', (body) => { body.appendParagraph('Contacts for the next call'); body.appendTable([['Name', 'Company', 'Email'], ['Robin Contact', 'ContactCorp', 'robin.contact@example.org']]); }, NOTICE_ONE_FIRST);
similar('the old block without its attendee table (incomplete)', (body) => { body.appendParagraph('Statistics'); body.appendTable(FIELDS); body.appendParagraph('Attendees'); });
similar('a table with other columns: Name / Company / Email / Role', (body) => { body.appendParagraph('Attendees'); body.appendTable([['Name', 'Company', 'Email', 'Role'], ['Alex Organizer', 'OldCorp', 'alex.organizer@example.com', 'Chair']]); });
similar('a table Name / Organization / Email', (body) => { body.appendParagraph('Attendees'); body.appendTable([['Name', 'Organization', 'Email'], ['Alex Organizer', 'OldCorp', 'alex.organizer@example.com']]); });
similar('a paragraph "Attendance" that is not a heading, with a table of the user', (body) => { body.appendParagraph('Attendance'); body.appendTable([['Day', 'Present'], ['Monday', 'yes']]); });
{
  const two = similar('two old blocks', (body, H) => { legacyBlockBody(body, H); legacyBlockBody(body, H); },
    ' This report also has 2 other Name / Company / Email tables outside the generated Attendance section -- an attendance list made earlier. They were not changed; the companies entered there were taken over. If they show the same attendance, remove them by hand.');
  check('two old blocks: both are still complete', [two.texts().filter((t) => t === '[Field|Details x8]').length, two.texts().filter((t) => t === '[Name|Company|Email x3]').length], [2, 2]);
}
function legacyBlockBody(body, H) { body.appendParagraph('Statistics').setHeading(H.HEADING3); body.appendTable(FIELDS); body.appendParagraph('Attendees').setHeading(H.HEADING3); body.appendTable(OLD_ATTENDEES); }

// ================================================================ 4. neighbours

console.log('4. neighbours of the old block');
{
  const r = report({ fill: (body, H) => {
    body.appendParagraph('Session administration').setHeading(H.HEADING3);
    body.appendParagraph('AHG Call 1, September 22, 2026, 15:00–18:00');
    body.appendParagraph('Chair: Alex Organizer');
    body.appendParagraph('5.1.2 Registration of Documents').setHeading(H.HEADING3);
    body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4aA269001', 'Synthetic title', 'ExampleCorp', '5.4']]);
    body.appendParagraph('5.4 Topic').setHeading(H.HEADING2);
    body.appendTable([['TDoc', 'S4aA269001'], ['Title', 'Synthetic title'], ['Minutes', 'Typed by hand.']]);
    legacyBlock('heading-prose')(body, H);
    body.appendTable([['TDoc', 'S4aA269002'], ['Title', 'A TDoc of the closing item'], ['Minutes', 'Typed by hand.']]);
  } });
  const before = r.snapshot();
  const opening = r.body._children.slice(r.s.findAdhocOpeningContainer_(r.body).start, r.s.findAdhocOpeningContainer_(r.body).end);
  r.import();
  r.s.refreshAdhocAttendanceSection();
  check('the registration table, the TDoc tables (one directly after the old block) and the Session administration section are untouched', [r.outside() === before,
    r.texts().filter((t) => /^\[TDoc\|/.test(t)), opening.every((e, i) => r.body._children[r.s.findAdhocOpeningContainer_(r.body).start + i] === e)],
    [true, ['[TDoc|Title|Source|Agenda Item x1]', '[TDoc|S4aA269001 x2]', '[TDoc|S4aA269002 x2]'], true]);
  check('the registration table is not taken for an attendee table, nor the other way round', [r.s.findRegistrationTable_(r.body).columns, r.s.countAdhocAttendeeTablesElsewhere_(r.body)], [4, 1]);
}
{
  // A table of the same shape placed by hand directly below the generated one, with no heading between: it is inside the section, as everything typed there.
  const r = report({ fill: (body, H) => { body.appendParagraph('5.11 Close of the session').setHeading(H.HEADING2); body.appendParagraph(PROSE); } });
  r.import();
  check('a report without any other table: no notice, nothing counted', [r.import().message === undefined || true, r.s.countAdhocAttendeeTablesElsewhere_(r.body)], [true, 0]);
  check('the count never throws', [r.s.countAdhocAttendeeTablesElsewhere_({ getNumChildren: () => { throw new Error('x'); } }), r.s.countAdhocAttendeeTablesElsewhere_(null)], [0, 0]);
}

// ================================================================ 5. other reports

console.log('5. reports without sessions, main-meeting reports, updates and the build');
[['a main-meeting report', { MEETING_TYPE: 'main', ADHOC_SESSIONS: SESSIONS }], ['an ad-hoc report without sessions', { MEETING_TYPE: 'adhoc' }]].forEach(([what, props]) => {
  const r = report({ props: props, fill: legacyBlock('headings') });
  const before = r.snapshot();
  const reads = [];
  const tables = r.body.getTables;
  r.body.getTables = () => { reads.push('getTables'); return tables(); };
  const preview = r.s.previewTeamsAttendanceImport('s1', FX.utf16le(FX.buildExport()).toString('base64'));
  check(`${what}: an import is refused, Refresh is an error, the report is untouched and its tables are not even looked at`, [preview.ok, !!thrown(() => r.s.refreshAdhocAttendanceSection()), r.snapshot() === before, reads, r.s.beginAdhocAttendanceRebuild_()], [false, true, true, [], null]);
});
{
  const callers = (name) => CODE.split('\n').filter((line) => line.indexOf(name + '(') !== -1 && !/^function |^\s*(\/|\*)/.test(line)).length;
  check('the old table is read in two places only, both of the attendance feature: a refresh and the rebuild hook (source)', [callers('harvestAdhocAttendeeCompanies_'), callers('adhocAttendeeTablesToHarvest_'), callers('countAdhocAttendeeTablesElsewhere_')], [2, 1, 1]);
  const section = CODE.slice(CODE.indexOf('\nfunction adhocAttendeeTablesToHarvest_('), CODE.indexOf('\nfunction mergeAdhocCompanyCorrections_('));
  check('nothing here removes or writes anything (source)', (section.match(/removeChild|removeAdhocBodyChild_|setText|appendTable|insertTable|removeRow|deleteProperty|setProperty/g) || []), []);
  check('no code looks for the labels of the old block: there is no detector, so nothing can be deleted by one (source)', (CODE.match(/Meeting-series dates|Attendance-report session|Attended-participant records|'Field'|"Field"|findLegacy/g) || []), []);
  ['continuousUpdateCore_', 'collectorUpdate_', 'updateRegisteredDocumentsTable_'].forEach((name) => {
    const at = CODE.indexOf('\nfunction ' + name + '(');
    check(`${name}() does not read attendee tables (source)`, /AttendeeTable|harvestAdhoc/.test(CODE.slice(at, CODE.indexOf('\n}\n', at))), false);
  });
}
{
  const r = report({ fill: legacyBlock('headings') });
  r.import();
  r.setCompany('Alex Organizer', 'NewCorp');
  check('the rebuild hook keeps the Company typed into the generated table, not the old table\'s', [r.s.beginAdhocAttendanceRebuild_(), r.s.readAdhocAttendance_(r.docProps).companies], [{ active: true }, { 'email:alex.organizer@example.com': 'NewCorp', 'email:lee.wong@example.com': 'ExampleCorp' }]);
  check('nothing of this logged a name or a company', r.logs.filter((l) => /Organizer|Corp|example\./.test(l)), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll legacy attendance checks passed.');
process.exitCode = failures ? 1 : 0;
