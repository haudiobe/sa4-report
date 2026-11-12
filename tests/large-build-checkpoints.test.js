/**
 * Large-build checkpoints and the status-dropdown repair (after T-2026.10.22, GitHub issue #6).
 *
 * Audio SWG SA4#138: Full Report Build made every DocumentApp change of all its phases in one unsaved batch (10 agenda
 * items, 111 TDocs, 100 abstracts). The first save -- in the status-dropdown step -- was refused by Google Docs ("Too
 * many changes applied before saving document ..."); the step was reported as skipped and the execution failed at its
 * end, when Apps Script tried the same save again.
 *
 * The fake report used here models what that failure depends on:
 *   - Google's limit: a save with more than LIMIT unsaved changes is refused with Google's message, and the changes
 *     stay unsaved;
 *   - a closed document: after saveAndClose() the handle and the body it gave are closed;
 *   - DocumentApp.openById() of the report gives a NEW document: its body and every element in it are new objects.
 *     Every element of the closed document throws when it is used at all -- read or written -- so a stale
 *     reference anywhere in a build makes the build fail;
 *   - the end of the execution: Apps Script saves what is still unsaved, under the same limit.
 * The fake Docs API (tests/helpers/fake-docs-api.js) always describes the report as it is now, saved.
 *
 *    A. the live failure, reproduced with the released Code.js of T-2026.10.22;
 *    B. the same build with checkpoints: completed, everything saved, every element fresh;
 *    C. small builds are as before: no checkpoint, one save;
 *    D. failures: a checkpoint that cannot be saved or reopened stops the build and says what is saved;
 *       a status-dropdown step that cannot save fails the build instead of being skipped;
 *    E. outside a Full Build nothing of it runs;
 *    F. the repair of an existing report: dry run, repair, run again, what it never touches, what it refuses;
 *    G. what changed in Code.js: against T-2026.10.22, one section, six functions and the document access.
 *
 * All names and numbers are synthetic. No live service is used.
 *
 * Run: node tests/large-build-checkpoints.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { makeFakeDocsApi } = require('./helpers/fake-docs-api.js');
const { codeWithoutMeetingTimeZones, codeWithoutLargeBuildCheckpoints, LARGE_BUILD_SECTION_START, LARGE_BUILD_SECTION_END, LARGE_BUILD_SECTION_FUNCTIONS, LARGE_BUILD_CHANGED_FUNCTIONS, LARGE_BUILD_DOCUMENT_ACCESS_COUNT } =
  require('./helpers/before-structure-fix.js');

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

const ROOT = path.join(__dirname, '..');
const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
let RELEASED = null;
try {
  RELEASED = execFileSync('git', ['show', 'template-release/T-2026.10.22:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
} catch (e) {
  RELEASED = null;
}

const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const TOO_MANY = 'Too many changes applied before saving document. Please save changes in smaller batches using Document.saveAndClose(), then reopen the document with Document.openById().';
const NAMES = ['available', 'noted', 'agreed', 'revised', 'parked', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'other', 'replied', 'Plenary', 'postponed'];
const DEFINITION = { title: 'Document Status', options: NAMES.map((name) => ({ displayValue: name })) };
const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic meeting', MEETING_ID: '85999', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'November 9, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85999', AGENDA_TDOC: 'S4aA260090' };
/** Ten agenda items, as the Audio report has. */
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Audio topics'], ['4.1', 'Topic one'], ['4.2', 'Topic two'],
  ['4.3', 'Topic three'], ['4.4', 'Topic four'], ['4.5', 'Topic five'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
const STATES = ['agreed', 'available', 'reserved', 'noted', 'not treated', 'approved', 'replied to'];
/** n TDocs over the five topic items; every seventh has a status without an option ("not treated"). */
const tdocList = (n) => Array.from({ length: n }, (_, i) => ['S4aA26' + String(9001 + i), '4.' + (1 + (i % 5)), STATES[i % STATES.length]]);
/** What its Status is after a completed build. */
const expectedStatus = (t) => t[0] + ' ' + (t[2] === 'not treated' ? 'text:not treated' : 'dropdown:' + (t[2] === 'replied to' ? 'replied' : t[2]));

/** Every object of a fake body that has methods: the body, its paragraphs, its tables, their rows, cells and cell paragraphs. */
function elementsOf(body) {
  const out = [body];
  body._children.forEach((c) => {
    out.push(c);
    if (c.getType() !== 'TABLE') return;
    for (let i = 0; i < c.getNumRows(); i++) {
      const row = c.getRow(i);
      out.push(row);
      for (let k = 0; k < row.getNumCells(); k++) { const cell = row.getCell(k); out.push(cell, cell.getChild(0)); }
    }
  });
  return out;
}

/** The same content in a new fake body: what DocumentApp.openById() gives after a save. */
function copyBody(s, from) {
  const to = makeFakeDocumentBody(s);
  from._children.forEach((c) => {
    if (c.getType() !== 'TABLE') {
      const p = to.appendParagraph(c.getText());
      p.setHeading(c.getHeading());
      p._bold = c._bold;
      if (c._links) p._links = c._links.map((l) => l.slice());
      return;
    }
    const rows = [];
    for (let i = 0; i < c.getNumRows(); i++) { const row = c.getRow(i); const texts = []; for (let k = 0; k < row.getNumCells(); k++) texts.push(row.getCell(k).getText()); rows.push(texts); }
    const t = to.appendTable(rows);
    for (let i = 0; i < c.getNumRows(); i++) {
      const a = c.getRow(i);
      const b = t.getRow(i);
      b._minHeight = a._minHeight;
      for (let k = 0; k < a.getNumCells(); k++) {
        const x = a.getCell(k);
        const y = b.getCell(k);
        ['_links', '_bold', '_background', '_width', '_spacing', '_dropdown'].forEach((key) => { y[key] = x[key] === null || x[key] === undefined ? x[key] : JSON.parse(JSON.stringify(x[key])); });
      }
    }
  });
  to._state.dirty = false;
  to._state.changes = 0;
  to._state.nextDropdown = from._state.nextDropdown;
  to._state.onFirstTouch = from._state.onFirstTouch;
  return to;
}

/** Makes every method of every element of a closed body throw: nothing of a closed document can be used again. */
function poison(body) {
  elementsOf(body).forEach((el) => {
    Object.keys(el).forEach((key) => {
      if (typeof el[key] === 'function') el[key] = () => { throw new Error('STALE: an element of the document that was saved and closed was used after it was reopened (' + key + ')'); };
    });
  });
}

/**
 * One report, with the real build. options:
 *   code      another Code.js (the released one)
 *   tdocs     the TDoc list (default: 111)
 *   limit     Google's limit of unsaved changes per save (default 2000)
 *   docs      false: no Google Docs API service
 */
function report(options) {
  const o = options || {};
  const loaded = loadCode({ code: o.code, documentProperties: Object.assign({}, PROPS, o.props || {}) });
  const s = loaded.sandbox;
  const r = { s: s, docProps: loaded.docProps, tdocs: (o.tdocs || tdocList(111)).map((t) => t.slice()), logs: [], events: [], saves: [], limit: o.limit || 2000, reopenFails: 0, sleeps: [] };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.ElementType.TEXT = 'TEXT';
  s.DocumentApp.ElementType.UNSUPPORTED = 'UNSUPPORTED';
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  r.body = makeFakeDocumentBody(s);
  // The fake Docs API always looks at the report as it is now: the body of the document that is open.
  const bodyNow = new Proxy({}, { get: (_, key) => r.body[key], set: (_, key, value) => { r.body[key] = value; return true; } });
  r.api = makeFakeDocsApi({ body: bodyNow, reportId: REPORT_ID });
  r.api.addSource(TEMPLATE_ID, [DEFINITION]);

  /** A handle on the report, as DocumentApp gives one. It is closed by its saveAndClose(). */
  function handle(kind) {
    const h = { open: true, kind: kind };
    h.getId = () => REPORT_ID;
    h.getBody = () => { if (!h.open) throw new Error('Document is closed, its contents cannot be updated.'); return r.body; };
    h.saveAndClose = () => {
      if (!h.open) throw new Error('Document is closed, its contents cannot be updated.');
      const pending = r.body._state.changes || 0;
      if (pending > r.limit) { r.events.push('save refused (' + pending + ')'); throw new Error(TOO_MANY); }
      r.saves.push(pending);
      r.events.push('save');
      r.body._state.dirty = false;
      r.body._state.changes = 0;
      r.body._state.closed = true;
      h.open = false;
    };
    return h;
  }
  r.active = handle('active');
  s.DocumentApp.getActiveDocument = () => r.active;
  s.DocumentApp.openById = (id) => {
    if (id !== REPORT_ID) return { getBody: () => makeFakeDocumentBody(s) };
    r.events.push('openById');
    if (r.reopenFails > 0) { r.reopenFails--; throw new Error('synthetic: the document could not be opened'); }
    const old = r.body;
    r.body = copyBody(s, old);
    poison(old);
    return handle('reopened');
  };
  s.DocumentApp.getUi = () => { throw new Error('no UI in a build'); };
  s.Utilities.sleep = (ms) => r.sleeps.push(ms);
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => !r.locked, releaseLock: () => {} }) };
  s.SA4_RELEASE_ = { releaseId: 'T-2026.10.22', flavor: 'template', codeVersion: '2.26.6', templateDocumentId: TEMPLATE_ID };
  if (o.docs !== false) s.Docs = { Documents: { get: (id, opts) => { r.events.push('get ' + (id === REPORT_ID ? 'report' : 'template')); return r.api.Docs.Documents.get(id, opts); },
    batchUpdate: (resource, id) => { r.events.push('batchUpdate'); return r.api.Docs.Documents.batchUpdate(resource, id); } } };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
  s.collectorUpdate_ = () => ({ failures: [] });
  // The e-mail and revision phases: stand-ins that take the body the way the real collectors do, and write into it.
  s.collectEmailDiscussionCore_ = () => {
    s.getActiveDocumentBodyCounted_('checkRSSFeed_').getTables().forEach((t) => { if (t.getRow(0).getCell(0).getText() === 'TDoc' && t.getNumRows() > 6) t.getRow(6).getCell(1).setText('synthetic e-mail line'); });
  };
  s.collectRevisionsCore_ = () => { s.getActiveDocumentBodyCounted_('updateRevisions_').getTables().length; };
  // The abstracts: a token, and an answer for every TDoc (a row inserted as fetchAndAddAbstract_() inserts it).
  s.getReviewerApiTokenForRun_ = () => 'synthetic-token';
  s.fetchAndAddAbstract_ = (table, number) => { const row = table.insertTableRow(2); row.appendTableCell('Abstract'); row.appendTableCell('Synthetic abstract of ' + number); return true; };
  s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], getContentType: () => 'application/octet-stream', setName() { return this; } }) }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const rows = r.tdocs.concat([['S4aA260090', '', 'available']]);
    const values = [HEADER].concat(rows.map(([id, agenda, status]) => [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'pCR', 'Agreement', agenda, 'Synthetic item', status, '', '', '', '']));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }),
      getRange: (row, col, n) => ({ getDisplayValues: () => [['Uploaded']].concat(rows.map(() => [''])).slice(row - 1, row - 1 + n) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };

  /** One execution: the active document is open at its start; at its end Apps Script saves what is unsaved. */
  r.exec = (fn) => {
    r.active = handle('active');
    r.body._state.closed = false;
    if (s.resetStatusDropdownRun_) s.resetStatusDropdownRun_();
    const from = r.events.length;
    let out;
    let thrown = null;
    try { out = fn(); } catch (e) { thrown = e; }
    const pending = r.body._state.changes || 0;
    r.endOfExecution = pending === 0 ? 'nothing unsaved' : (pending > r.limit ? 'FAILED: ' + TOO_MANY : 'saved ' + pending);
    if (pending && pending <= r.limit) { r.body._state.dirty = false; r.body._state.changes = 0; }
    r.last = r.events.slice(from);
    if (thrown) throw thrown;
    return out;
  };
  r.build = () => r.exec(() => s.runFullReportBuildCore_());
  r.update = () => r.exec(() => s.continuousUpdateCore_());
  /** Somebody edits the report by hand, between two executions; the edit is saved. */
  r.user = (fn) => { r.body._state.closed = false; const out = fn(); r.body._state.dirty = false; r.body._state.changes = 0; r.api.edited(); return out; };
  r.tables = () => r.body._children.filter((c) => c.getType() === 'TABLE');
  r.tdocTable = (number) => r.tables().filter((t) => t.getRow(0).getNumCells() === 2 && t.getRow(0).getCell(0).getText() === 'TDoc' && t.getRow(0).getCell(1).getText() === number)[0];
  r.cell = (number, label) => { const t = r.tdocTable(number); for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === label) return t.getRow(i).getCell(1); return null; };
  r.all = () => r.tdocs.map((t) => t[0] + ' ' + r.api.shown(t[0]));
  r.phases = (built) => built.phases.map((p) => p.name + ' ' + p.status);
  return r;
}

// ====================================================================== A
console.log('A. the live failure, reproduced with the released Code.js of T-2026.10.22');
if (RELEASED === null) {
  console.log('  note: template-release/T-2026.10.22 is not available in this checkout; the reproduction is skipped.');
} else {
  const r = report({ code: RELEASED });
  const built = r.build();
  const dropdowns = built.phases.filter((p) => p.name === 'status dropdowns')[0];
  check('the build reports success although nothing could be saved: every phase "done", the dropdown step "skipped"',
    [built.ok, r.phases(built)], [true, ['skeleton done', 'e-mail done', 'revisions done', 'abstracts done', 'formatting done', 'status dropdowns skipped']]);
  check('the step was skipped because its save was refused with Google\'s message', dropdowns.detail, TOO_MANY);
  check('it was the first and only save of the build', [r.saves.length, r.last.filter((e) => /^save refused/.test(e)).length], [0, 1]);
  check('the execution then fails at its end: the same changes are still unsaved', r.endOfExecution, 'FAILED: ' + TOO_MANY);
  check('no dropdown, and the report is not marked', [r.all().filter((x) => /dropdown:/.test(x)).length, r.docProps.getProperty('STATUS_DROPDOWNS')], [0, null]);
}

// ====================================================================== B
console.log('\nB. the same build with checkpoints: completed, saved, with fresh references only');
{
  const r = report();
  const built = r.build();
  check('the build completes: every phase done', [built.ok, r.phases(built)], [true, ['skeleton done', 'e-mail done', 'revisions done', 'abstracts done', 'formatting done', 'status dropdowns done']]);
  check('it saved several times, never more than Google allows in one save', [r.saves.length > 5, Math.max.apply(null, r.saves) <= r.limit], [true, true]);
  check('every checkpoint is a save followed by opening the report again, by its ID', (() => { const ev = r.events.filter((e) => e === 'save' || e === 'openById'); return ev.slice(0, -1).every((e, i) => (i % 2 === 0 ? e === 'save' : e === 'openById')); })(), true);
  check('the last save is the status-dropdown step\'s, and it is followed by its one read of the report', r.last.slice(r.last.lastIndexOf('save'), r.last.lastIndexOf('save') + 2), ['save', 'get report']);
  check('the execution ends with nothing unsaved', r.endOfExecution, 'nothing unsaved');
  check('checkpoints were made inside the skeleton, the widths and the formatting, and after every phase but the last', ['skeleton: TDoc tables', 'formatting: widths', 'formatting', 'after skeleton', 'after e-mail', 'after revisions', 'after abstracts', 'after formatting', 'after status dropdowns']
    .map((label) => r.logs.some((l) => l.indexOf('[CHECKPOINT] ' + label + ':') === 0)), [true, true, true, true, true, true, true, true, false]);
  check('no checkpoint carried more than about the budget of estimated changes', r.logs.filter((l) => /^\[CHECKPOINT\]/.test(l)).map((l) => Number((l.match(/about (\d+) change/) || [])[1])).filter((n) => n > r.s.LARGE_BUILD_CHECKPOINT_BUDGET_ + 100), []);
  check('the total line counts them', /^\[FULLBUILD\] total: \d+ ms -- skeleton done \d+ ms, .* -- \d+ checkpoint\(s\)$/.test(r.logs.filter((l) => /^\[FULLBUILD\] total/.test(l))[0]), true);
  check('111 TDoc tables, each once, each with the status of its own row', [r.tdocs.filter((t) => r.tables().filter((x) => x.getRow(0).getNumCells() === 2 && x.getRow(0).getCell(1).getText() === t[0]).length === 1).length, r.all().filter((x, i) => x !== expectedStatus(r.tdocs[i]))], [111, []]);
  check('every TDoc table got its abstract, once', r.tdocs.filter((t) => { const tb = r.tdocTable(t[0]); let n = 0; for (let i = 0; i < tb.getNumRows(); i++) if (tb.getRow(i).getCell(0).getText() === 'Abstract') n++; return n !== 1; }), []);
  check('the e-mail phase wrote into the document that is open now', r.tdocs.filter((t) => r.cell(t[0], 'E-mail Discussion').getText() !== 'synthetic e-mail line'), []);
  check('the formatting reached every row of every table of the document that is open now', r.tables().filter((t) => { for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i)._minHeight !== 0) return true; return false; }).length, 0);
  check('and the widths every row of every TDoc table', r.tdocs.filter((t) => { const tb = r.tdocTable(t[0]); for (let i = 0; i < tb.getNumRows(); i++) if (tb.getRow(i).getCell(0)._width !== 108 || tb.getRow(i).getCell(1)._width !== 360) return true; return false; }), []);
  check('the registration table lists all 111', (() => { const reg = r.tables().filter((t) => t.getRow(0).getCell(0).getText() === 'TDoc' && t.getRow(0).getNumCells() > 2)[0]; return reg.getNumRows() - 1; })(), 111);
  check('the definition was created, the report is marked', [Object.keys(r.api.definitions).length, r.docProps.getProperty('STATUS_DROPDOWNS')], [1, '1']);
  check('after the build, checkpoints are off and the report is the active document again', [r.s.LARGE_BUILD_CHECKPOINTS_.enabled, r.s.LARGE_BUILD_CHECKPOINTS_.reopened, r.s.currentReportDocument_() === r.active], [false, null, true]);

  // With a smaller budget the abstracts loop makes checkpoints too: it takes its tables again, and every abstract lands.
  const fine = report();
  fine.s.LARGE_BUILD_CHECKPOINT_BUDGET_ = 300;
  const b2 = fine.build();
  check('a smaller budget: checkpoints inside the abstracts loop as well; the build completes, every abstract once, every status right',
    [b2.ok, fine.logs.filter((l) => l.indexOf('[CHECKPOINT] abstracts:') === 0).length > 1, fine.tdocs.filter((t) => fine.cell(t[0], 'Abstract') === null).length, fine.all().filter((x, i) => x !== expectedStatus(fine.tdocs[i]))],
    [true, true, 0, []]);

  // A rebuild of the built report: it is cleared first, so no table is doubled; what was selected is kept.
  r.api.pick('S4aA269001', 'parked');
  const again = r.build();
  check('a rebuild of it completes too, with no doubled table, and keeps a decision selected by hand', [again.ok, r.tables().filter((t) => t.getRow(0).getNumCells() === 2 && t.getRow(0).getCell(0).getText() === 'TDoc').length, r.api.shown('S4aA269001')],
    [true, 111, 'dropdown:parked']);
}

// ====================================================================== C
console.log('\nC. small builds are as before');
{
  const r = report({ tdocs: tdocList(9) });
  const built = r.build();
  check('nine TDocs: completed, no checkpoint, one save -- the status-dropdown step\'s', [built.ok, r.events.filter((e) => e === 'openById').length, r.saves.length, r.logs.filter((l) => /^\[CHECKPOINT\]/.test(l))], [true, 0, 1, []]);
  check('its total line is the one every release wrote', /^\[FULLBUILD\] total: \d+ ms -- skeleton done \d+ ms, e-mail done \d+ ms, revisions done \d+ ms, abstracts done \d+ ms, formatting done \d+ ms, status dropdowns done \d+ ms$/.test(r.logs.filter((l) => /^\[FULLBUILD\] total/.test(l))[0]), true);
  const plain = report({ tdocs: tdocList(9), docs: false });
  const b2 = plain.build();
  check('without the Docs API (five phases): no save at all, as before; Apps Script saves at the end', [b2.ok, plain.saves.length, plain.events.filter((e) => e === 'openById').length, /^saved \d+$/.test(plain.endOfExecution)], [true, 0, 0, true]);
  const big = report({ docs: false });
  const b3 = big.build();
  check('a large build without the Docs API still makes its checkpoints, and ends with nothing too large to save', [b3.ok, big.saves.length > 5, big.endOfExecution === 'nothing unsaved' || /^saved \d+$/.test(big.endOfExecution)], [true, true, true]);
}

// ====================================================================== D
console.log('\nD. failures stop the build and say what is saved');
{
  // A limit so low that even a checkpoint's share is refused.
  const r = report({ limit: 300 });
  const built = r.build();
  check('a checkpoint that cannot be saved fails its phase; later phases do not run', [built.ok, built.failedPhase, r.phases(built)], [false, 'skeleton', ['skeleton failed']]);
  check('the error names the checkpoint, Google\'s message, and that nothing of the build is saved yet',
    [/^The report could not be saved at the checkpoint "skeleton: TDoc tables" \(Too many changes applied before saving document\./.test(built.error), /Nothing this build made has been saved\./.test(built.error), /do not use it as it is/.test(built.error)], [true, true, true]);

  const later = report();
  later.limit = 2000;
  let saves = 0;
  const inner = later.s.DocumentApp.openById;
  later.s.DocumentApp.openById = (id) => { const h = inner(id); if (id === REPORT_ID && ++saves === 3) later.limit = 10; return h; };
  const b2 = later.build();
  check('a later checkpoint that cannot be saved names the last one that was', [b2.ok, /What the build made up to the checkpoint "[^"]+" is saved; what it made after that is not\./.test(b2.error)], [false, true]);

  const reopen = report();
  reopen.reopenFails = 3;
  const b3 = reopen.build();
  check('a report that cannot be opened again after its save: three attempts, with pauses of 1 and 2 s, then the build stops', [b3.ok, reopen.events.filter((e) => e === 'openById').length, reopen.sleeps, /^The report was saved at the checkpoint "[^"]+" but could not be opened again \(synthetic: the document could not be opened\)\. The build stops there\.$/.test(b3.error)],
    [false, 3, [1000, 2000], true]);
  const flaky = report();
  flaky.reopenFails = 2;
  check('two failed attempts and a third that works: the build goes on and completes', [flaky.build().ok, flaky.sleeps.slice(0, 2)], [true, [1000, 2000]]);

  // The status-dropdown step's own save: with no checkpoints at all (the budget set out of reach), it is refused.
  const unchecked = report();
  unchecked.s.LARGE_BUILD_CHECKPOINT_BUDGET_ = 1e9;
  const b4 = unchecked.build();
  check('a refused save in the status-dropdown step FAILS the build, in that step, with Google\'s message -- it is no longer "skipped"',
    [b4.ok, b4.failedPhase, b4.error, r.s.formatFullBuildSummary_(b4).split('\n')[0]], [false, 'status dropdowns', 'the document could not be saved (' + TOO_MANY + ')', 'Full report build failed: the document could not be saved (' + TOO_MANY + ')']);

  // A table that is gone after a checkpoint: the build stops instead of writing into the wrong one.
  const shrink = report();
  shrink.s.LARGE_BUILD_CHECKPOINT_BUDGET_ = 300;
  const open2 = shrink.s.DocumentApp.openById;
  let n = 0;
  shrink.s.DocumentApp.openById = (id) => { const h = open2(id); if (id === REPORT_ID && shrink.s.LARGE_BUILD_CHECKPOINTS_.last === 'abstracts' && n++ === 0) { const t = shrink.tables(); shrink.body._children.splice(shrink.body._children.indexOf(t[t.length - 1]), 1); } return h; };
  const b5 = shrink.build();
  check('a reopened report with another number of tables stops the loop: no write into the wrong table', [b5.ok, b5.failedPhase, /^After the checkpoint "abstracts" the report has \d+ tables instead of \d+\. The build stops there\.$/.test(b5.error)], [false, 'abstracts', true]);
}

// ====================================================================== E
console.log('\nE. outside a Full Build nothing of it runs');
{
  const r = report();
  r.build();
  const n = r.events.length;
  r.tdocs.push(['S4aA269999', '4.2', 'agreed']);
  const u = r.update();
  check('an update of the large report: no checkpoint, no reopening; it saves once, in its status-dropdown step', [u.success, r.events.slice(n).filter((e) => e === 'openById').length, r.events.slice(n).filter((e) => e === 'save').length, r.api.shown('S4aA269999')], [true, 0, 1, 'dropdown:agreed']);
  const S = loadCode().sandbox;
  check('the helpers do nothing while checkpoints are off', [S.largeBuildCheckpoint_('x'), S.largeBuildPhaseCheckpoint_('x'), S.largeBuildStep_('BODY', 'x', 5000), S.largeBuildTablesIfDue_(['T'], 'x'), (S.largeBuildCount_(5000), S.LARGE_BUILD_CHECKPOINTS_.pending)], [null, null, 'BODY', ['T'], 0]);
  S.DocumentApp.getActiveDocument = () => 'ACTIVE';
  check('currentReportDocument_() is the active document when no checkpoint reopened one', S.currentReportDocument_(), 'ACTIVE');
  const CODE_ONLY = FULL.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  // (Two counters of the performance log are named after the call; they are strings, not calls.)
  check('no code but currentReportDocument_() asks DocumentApp for the active document (source)', [(CODE_ONLY.match(/DocumentApp\.getActiveDocument\(\)(?!\.getBody\(\) calls \(total\)')/g) || []).length,
    /function currentReportDocument_\(\) \{\n  return LARGE_BUILD_CHECKPOINTS_\.reopened \|\| DocumentApp\.getActiveDocument\(\);\n\}/.test(FULL)], [1, true]);
  check('checkpoints are switched on in one place: the Full Build (source)', [(FULL.match(/beginLargeBuildCheckpoints_\(\)/g) || []).length - 1, /resetStatusDropdownRun_\(\);\n  \/\/ After T-2026\.10\.22: .*\n  beginLargeBuildCheckpoints_\(\);\n\n  phase\('skeleton'/.test(FULL)], [1, true]);
}

// ====================================================================== F
console.log('\nF. the repair of an existing report after a build whose status-dropdown step did not run');
/** The Audio report as the failed build left it, once its content is saved: text statuses, no definition, not marked; minutes typed in by hand. */
function failedAudioReport() {
  const r = report();
  const finalize = r.s.finalizeStatusDropdowns_;
  r.s.finalizeStatusDropdowns_ = () => ({ ran: false, inserted: 0, updated: 0, reason: 'switched off: the step of the failed build' });
  r.build();
  r.s.finalizeStatusDropdowns_ = finalize;
  r.user(() => { r.cell('S4aA269003', 'Minutes').setText('Minutes typed during the meeting.'); r.cell('S4aA269010', 'Status').setText('agreed'); });
  return r;
}
{
  const r = failedAudioReport();
  check('the starting point: 111 text statuses, no definition, not marked', [r.all().filter((x) => / text:/.test(x)).length, Object.keys(r.api.definitions).length, r.docProps.getProperty('STATUS_DROPDOWNS')], [111, 0, null]);
  const before = r.api.calls.length;
  const changesBefore = r.body._state.changes || 0;
  const preview = r.exec(() => r.s.repairStatusDropdowns_({ dryRun: true }));
  check('the dry run: 111 TDoc tables, none doubled, 95 to make dropdowns of, 16 left as text ("not treated")',
    [preview.ok, preview.tdocTables, preview.duplicates, preview.dropdownsBefore, preview.toConvert, preview.unmapped.length, preview.definition], [true, 111, [], 0, 95, 16, 'to be created']);
  check('it only read: the report once, the template once; no write, no DocumentApp change, no mark',
    [r.api.calls.slice(before).map((c) => c.call + ' ' + (c.id === REPORT_ID ? 'report' : 'template')), (r.body._state.changes || 0) - changesBefore, r.docProps.getProperty('STATUS_DROPDOWNS')], [['get report', 'get template'], 0, null]);
  check('previewStatusDropdownRepair() says the same, for the log', r.exec(() => r.s.previewStatusDropdownRepair()).split('\n').slice(0, 5),
    ['Status dropdown repair -- dry run, nothing was changed.', 'TDoc tables: 111', 'already a dropdown: 0', 'to be made a dropdown: 95', 'left as text: 16 (no option for: S4aA269005 "not treated", S4aA269012 "not treated", S4aA269019 "not treated", S4aA269026 "not treated", S4aA269033 "not treated", …)']);

  const minutes = r.tdocs.map((t) => r.cell(t[0], 'Minutes').getText());
  const abstracts = r.tdocs.map((t) => r.cell(t[0], 'Abstract').getText());
  const tables = r.tables().length;
  const text = r.exec(() => r.s.repairStatusDropdownsAfterBuild());
  check('the repair: the definition is created, 95 dropdowns made, everything else unchanged, the report marked',
    text.split('\n'), ['Status dropdown repair -- done.', 'TDoc tables: 111', 'already a dropdown: 0', 'made a dropdown: 95', 'left as text: 16 (no option for: S4aA269005 "not treated", S4aA269012 "not treated", S4aA269019 "not treated", S4aA269026 "not treated", S4aA269033 "not treated", …)',
      'definition: created; report marked before: false, after: true; everything else unchanged: true']);
  check('every Status is now what the build would have made -- including the one typed in by hand ("agreed")', r.all().filter((x, i) => x !== (r.tdocs[i][0] === 'S4aA269010' ? 'S4aA269010 dropdown:agreed' : expectedStatus(r.tdocs[i]))), []);
  check('nothing else: the same tables, the minutes, the abstracts; DocumentApp changed nothing', [r.tables().length === tables, r.tdocs.map((t) => r.cell(t[0], 'Minutes').getText()), r.tdocs.map((t) => r.cell(t[0], 'Abstract').getText()), r.body._state.changes || 0, r.saves.length],
    [true, minutes, abstracts, 0, r.saves.length]);
  check('the minutes typed by hand are there', r.cell('S4aA269003', 'Minutes').getText(), 'Minutes typed during the meeting.');

  // Run again: nothing to do.
  r.api.pick('S4aA269001', 'parked');
  const revision = r.api.revision();
  const n = r.api.calls.length;
  const second = r.exec(() => r.s.repairStatusDropdowns_({ dryRun: false }));
  check('run again: nothing to convert, nothing written, still marked, the document at the same revision', [second.ok, second.toConvert, second.converted, second.dropdownsBefore, r.api.calls.slice(n).filter((c) => c.call === 'batchUpdate').length, r.docProps.getProperty('STATUS_DROPDOWNS'), r.api.revision() === revision],
    [true, 0, 0, 95, 0, '1', true]);
  check('a value selected by hand is kept', r.api.shown('S4aA269001'), 'dropdown:parked');
  check('one definition only', Object.keys(r.api.definitions).length, 1);
  check('the next update treats it as a report with dropdowns: a new TDoc gets its dropdown', (() => { r.tdocs.push(['S4aA269998', '4.1', 'noted']); r.update(); return r.api.shown('S4aA269998'); })(), 'dropdown:noted');
}
{
  // A batch that is refused: what was sent is complete, the rest is text; not marked; a second run finishes it.
  const r = failedAudioReport();
  r.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange && r.api.calls.filter((c) => c.call === 'batchUpdate' && c.requests[0].deleteContentRange).length === 2 ? 'synthetic failure of the second batch' : null);
  const first = r.exec(() => r.s.repairStatusDropdowns_({ dryRun: false }));
  check('a refused batch: not complete, not marked; 50 made, 45 still text, nothing unexpected', [first.ok, first.marked, first.converted, first.notConverted.length, first.unexpected, /Run the repair again\.$/.test(first.error)], [false, false, 50, 45, [], true]);
  check('no Status cell is half written', r.all().filter((x) => / empty$|\+/.test(x)), []);
  r.api.fail.batchUpdate = null;
  const second = r.exec(() => r.s.repairStatusDropdowns_({ dryRun: false }));
  check('run again: the rest is made, the report marked', [second.ok, second.converted, second.dropdownsBefore, r.docProps.getProperty('STATUS_DROPDOWNS'), r.all().filter((x) => /dropdown:/.test(x)).length], [true, 45, 50, '1', 95]);
}
{
  // What it refuses.
  const twice = failedAudioReport();
  twice.user(() => twice.body.appendTable([['TDoc', 'S4aA269002'], ['Status', 'agreed']]));
  const calls = twice.api.calls.length;
  const refused = twice.exec(() => twice.s.repairStatusDropdowns_({ dryRun: false }));
  check('a report with a doubled TDoc table is refused: nothing written, not marked, the number named', [refused.ok, refused.duplicates, twice.api.calls.slice(calls).filter((c) => c.call === 'batchUpdate').length, twice.docProps.getProperty('STATUS_DROPDOWNS'), /^1 TDoc number\(s\) have more than one table \(S4aA269002\)\./.test(refused.error)],
    [false, ['S4aA269002'], 0, null, true]);
  const empty = report();
  const none = empty.exec(() => empty.s.repairStatusDropdowns_({ dryRun: false }));
  check('a report without TDoc tables is refused', [none.ok, none.error], [false, 'The report has no TDoc table: it is not a built report. Nothing was changed.']);
  const noApi = report({ docs: false });
  check('without the Docs API: refused, nothing changed', noApi.exec(() => noApi.s.repairStatusDropdowns_({ dryRun: false })).error, 'Status dropdowns are not available in this report: it needs the template runtime with the Google Docs API service. Nothing was changed.');
  const noTemplate = failedAudioReport();
  noTemplate.api.addSource(TEMPLATE_ID, []);
  check('a template without the definition: refused, nothing changed', [noTemplate.exec(() => noTemplate.s.repairStatusDropdowns_({ dryRun: false })).error, Object.keys(noTemplate.api.definitions).length], ['The template has no "Document Status" dropdown. Nothing was changed.', 0]);
  const locked = failedAudioReport();
  locked.locked = true;
  const n = locked.api.calls.length;
  check('while an update holds the lock: nothing is read or changed', [locked.exec(() => locked.s.repairStatusDropdownsAfterBuild()), locked.api.calls.length - n], ['Nothing was changed: an update of this report is running right now.', 0]);
  const repairSource = FULL.slice(FULL.indexOf('\nfunction planStatusDropdownRepair_('), FULL.indexOf('// END OF LARGE BUILD CHECKPOINTS')).replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  check('the repair changes nothing through DocumentApp, downloads nothing and clears nothing (source)',
    /\.setText\(|\.clear\(|appendTable|insertTable|appendParagraph|removeFromParent|saveAndClose|UrlFetchApp|downloadAndGroupTdocs_|buildSkeletonWithTdocTables|runFullReportBuild/.test(repairSource), false);
}

// ====================================================================== G
console.log('\nG. what changed in Code.js: against T-2026.10.22, one section, six functions and the document access');
{
  const view = codeWithoutLargeBuildCheckpoints(FULL);
  check('every change is in Code.js exactly once, and the document access is changed in exactly ' + LARGE_BUILD_DOCUMENT_ACCESS_COUNT + ' places', view.problems, []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(LARGE_BUILD_SECTION_START), FULL.indexOf(LARGE_BUILD_SECTION_END) + LARGE_BUILD_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), LARGE_BUILD_SECTION_FUNCTIONS);
  check('it stands directly after the TDOC LIST DOWNLOAD section', FULL.indexOf('// END OF TDOC LIST DOWNLOAD\n// =========================================================\n\n' + LARGE_BUILD_SECTION_START) !== -1, true);
  check('it declares its settings and its state, and nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [],
    ['var LARGE_BUILD_CHECKPOINT_BUDGET_', 'var LARGE_BUILD_UNITS_', 'var LARGE_BUILD_REOPEN_DELAYS_MS_', 'var LARGE_BUILD_CHECKPOINTS_']);
  check('Code.js is version 2.26.8 (2.26.7 released it as T-2026.10.23, 2.26.8 the meeting time zones; the ledgers take the headers out)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.8');
  if (RELEASED === null) {
    console.log('  note: template-release/T-2026.10.22 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js is the Code.js of T-2026.10.22, byte for byte', view.code === RELEASED, true);
    const RELEASED_CODE_ONLY = RELEASED.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
    check('the document access: every code use of DocumentApp.getActiveDocument() the release had, and no other', (RELEASED_CODE_ONLY.match(/DocumentApp\.getActiveDocument\(\)(?!\.getBody\(\) calls \(total\)')/g) || []).length, LARGE_BUILD_DOCUMENT_ACCESS_COUNT);
    // With the document access read back, only these functions differ from the release.
    // Later layers (the meeting time zones, after T-2026.10.23) taken out, as for every earlier layer.
    const LAYER = codeWithoutMeetingTimeZones(FULL).code;
    const accessOnly = LAYER.split('currentReportDocument_()').join('DocumentApp.getActiveDocument()');
    check('beyond the document access, the functions that changed: ' + LARGE_BUILD_CHANGED_FUNCTIONS.join(', '),
      functionNames(RELEASED).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(accessOnly, n) !== functionSource(RELEASED, n)).sort(), LARGE_BUILD_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions are those of the section; none removed', [functionNames(LAYER).filter((n) => functionNames(RELEASED).indexOf(n) === -1).sort(), functionNames(RELEASED).filter((n) => functionNames(LAYER).indexOf(n) === -1)],
      [LARGE_BUILD_SECTION_FUNCTIONS.slice().sort(), []]);
    check('untouched: the update, the conversion of an existing report, the planning and sending of the dropdowns, the agenda and the TDoc list',
      ['continuousUpdateCore_', 'migrateStatusFieldsToDropdowns_', 'planStatusDropdownRequests_', 'sendStatusDropdownInsertPairs_', 'statusDropdownHealCandidates_', 'buildStatusDocsIndex_', 'downloadAndGroupTdocs_', 'readAgendaWithFallbacks_', 'fetchAndAddAbstract_', 'collectRevisionsOnly']
        .map((n) => functionSource(accessOnly, n) === functionSource(RELEASED, n)), Array(10).fill(true));
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll large-build checkpoint checks passed.');
process.exit(failures ? 1 : 0);
