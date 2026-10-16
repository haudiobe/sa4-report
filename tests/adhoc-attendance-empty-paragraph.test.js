/**
 * Attendance rendering -- the production failure of 2026-10-16: the section
 * could not be written ("Cannot insert an empty text element.") and a later
 * Refresh Attendance failed the same way, in removeAdhocBodyChild_().
 *
 * Cause: the Attendance section ends the report, so its last body child is
 * the empty paragraph that Google Docs keeps after the last table. Replacing
 * the section "emptied" that paragraph with setText(''), which Google Docs
 * refuses on a paragraph without text. The fake paragraph of
 * tests/helpers/fake-document.js refuses it the same way.
 *
 * The real render path runs here (renderAdhocAttendanceSection_ with blocks,
 * and removeAdhocBodyChild_); only the body is fake. Synthetic data only.
 *
 * Run: node tests/adhoc-attendance-empty-paragraph.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`);
  else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}
/** The result of fn, or the message of what it throws. */
const attempt = (fn) => { try { return fn(); } catch (e) { return 'THROWS: ' + e.message; } };

const loaded = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
const s = loaded.sandbox;
s.Logger = { log: () => {} };
s.getConfig_ = () => ({});

const block = (n, attendees) => ({ sessionId: 's' + n, key: 's' + n, day: '2026-09-2' + n, statistics: [['Meeting', 'Synthetic ' + n], ['Attendance records', String(attendees.length)]],
  attendees: attendees.map((name) => ({ name: name, company: 'ExampleCorp', email: name.toLowerCase().replace(' ', '.') + '@example.com' })) });
const ONE = [block(1, ['Pat Kim', 'Robin Lee'])];
const TWO = ONE.concat([block(2, ['Sam Rivera'])]);

/** A report that ends with the Attendance section, or with prose after it; a strict body, so the document ends with a paragraph and tables are separated. */
function report(after) {
  const body = makeFakeDocumentBody(s, { tableSeparators: true, endParagraph: true });
  body.appendParagraph('5 Close of the meeting').setHeading(s.DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Closing prose typed by hand.');
  // The report can go on after the closing section: the Attendance section is then written before the next heading.
  if (after) { body.appendParagraph('6 Annex').setHeading(s.DocumentApp.ParagraphHeading.HEADING1); body.appendParagraph('Prose after the section.'); }
  return body;
}
const texts = (body) => body._children.map((c) => (c.getType() === 'TABLE' ? '[table x' + c.getNumRows() + ']' : c.getText()));
const blanks = (body) => texts(body).filter((t) => t === '').length;
const render = (body, blocks) => attempt(() => s.renderAdhocAttendanceSection_(body, blocks).container);

console.log('1. the section is the last thing in the document: the trailing empty paragraph is the last child of the container');
{
  const body = report(false);
  const before = texts(body).slice(0, 2);
  check('first render creates the section', render(body, ONE), 'created');
  const last = body._children[body._children.length - 1];
  check('the document ends with an empty paragraph, inside the container', [last.getText(), s.findAdhocAttendanceContainer_(body).end === body._children.length], ['', true]);
  check('replacing it (a second import, or Refresh Attendance) works', render(body, TWO), 'replaced');
  const once = texts(body);
  check('and again, with the same blocks: idempotent, no blank paragraph accumulates', [render(body, TWO), texts(body)], ['replaced', once]);
  check('repeated refresh leaves exactly one blank paragraph, the last one', [blanks(body), texts(body)[texts(body).length - 1]], [1, '']);
  check('content before the section is untouched', texts(body).slice(0, 2), before);
  check('one Attendance heading, one block per session', [texts(body).filter((t) => t === 'Attendance').length, texts(body).filter((t) => t === 'Statistics').length], [1, 2]);
}

console.log('2. no attendance rows: the section is removed, and created again');
{
  const body = report(false);
  const before = texts(body).slice(0, 2);
  render(body, ONE);
  check('removal with no blocks works', render(body, []), 'removed');
  check('nothing of the section is left; the report ends with the one empty paragraph Google Docs keeps', [texts(body).indexOf('Attendance'), texts(body).slice(0, 2), blanks(body)], [-1, before, 1]);
  check('removal when there is no section is a no-op', render(body, []), 'none');
  check('the section is created again', [render(body, TWO), texts(body).filter((t) => t === 'Attendance').length], ['created', 1]);
  check('remove and create, repeatedly: one section, the text before it unchanged', (() => { for (let i = 0; i < 3; i++) { render(body, []); render(body, ONE); } return [texts(body).filter((t) => t === 'Attendance').length, texts(body).slice(0, 2)]; })(), [1, before]);
}

console.log('3. a last paragraph that has text');
{
  const body = report(false);
  render(body, ONE);
  const last = body._children[body._children.length - 1];
  last.setText('stray text'); // a last paragraph with text inside the container is cleared, not removed
  check('a last paragraph with text is cleared (replace works)', [render(body, TWO), texts(body)[texts(body).length - 1]], ['replaced', '']);
}

console.log('4. the section is followed by prose: everything after it stays');
{
  const body = report(true);
  render(body, ONE);
  check('create, replace, remove, create with a final prose paragraph', [render(body, TWO), render(body, ONE), render(body, []), render(body, TWO)], ['replaced', 'replaced', 'removed', 'created']);
  check('the heading and prose after the section are untouched, and there is no blank paragraph', [texts(body).slice(-2), blanks(body)], [['6 Annex', 'Prose after the section.'], 0]);
}

console.log('5. the helper itself');
{
  const body = report(false);
  const empty = body.appendParagraph('');
  empty.setHeading(s.DocumentApp.ParagraphHeading.HEADING3);
  check('an empty last paragraph is not "emptied" again: no error, heading reset', [attempt(() => { s.removeAdhocBodyChild_(body, empty); return 'ok'; }), empty.getHeading(), body._children.indexOf(empty) !== -1], ['ok', s.DocumentApp.ParagraphHeading.NORMAL, true]);
  const middle = body._children[0];
  check('a paragraph that is not last is removed', [attempt(() => { s.removeAdhocBodyChild_(body, middle); return 'ok'; }), body._children.indexOf(middle)], ['ok', -1]);
}

if (failures) { console.log(`\n${failures} FAILED`); process.exit(1); }
console.log('\nall ok');
