/**
 * Issue #28 (2026-11-25): the AHG copy's update was refused -- "it has no row for S4aP260149, the agenda TDoc of this
 * meeting -- it may be the list of another meeting". The Portal served, for the same address, first the meeting's
 * own list (118 rows) and then another meeting's list (38 rows). Both captured responses are fixtures
 * (tests/fixtures/tdoc-list-86178/); the report's OWN identity check runs on them, unchanged: the first is accepted,
 * the second refused with the very words of the live run. The guard is the one of T-2026.10.33, not weakened.
 */
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { execFileSync } = require('child_process');
const { loadCode } = require('./helpers/load-code');

let failures = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) failures++;
  console.log((ok ? '  ok   ' : '  FAIL ') + name + (ok ? '' : '\n         expected ' + JSON.stringify(expected) + '\n         actual   ' + JSON.stringify(actual)));
}

/** One file of a zip, from its central directory. */
function zipFile(buf, name) {
  let end = buf.length - 22;
  while (end >= 0 && buf.readUInt32LE(end) !== 0x06054b50) end--;
  let at = buf.readUInt32LE(end + 16);
  for (let n = buf.readUInt16LE(end + 10); n > 0; n--) {
    const method = buf.readUInt16LE(at + 10), size = buf.readUInt32LE(at + 20), nameLen = buf.readUInt16LE(at + 28), extra = buf.readUInt16LE(at + 30), comment = buf.readUInt16LE(at + 32), local = buf.readUInt32LE(at + 42);
    if (buf.slice(at + 46, at + 46 + nameLen).toString() === name) {
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const data = buf.slice(start, start + size);
      return (method === 8 ? zlib.inflateRawSync(data) : data).toString('utf8');
    }
    at += 46 + nameLen + extra + comment;
  }
  return null;
}
const unxml = (s) => s.replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
/** The first sheet of an xlsx as rows of cell values (as Sheet.getDataRange().getValues() gives them: strings and numbers). */
function xlsxRows(file) {
  const buf = fs.readFileSync(file);
  const shared = (zipFile(buf, 'xl/sharedStrings.xml').match(/<si>[\s\S]*?<\/si>/g) || []).map(unxml);
  const sheet = zipFile(buf, 'xl/worksheets/sheet1.xml');
  const col = (ref) => ref.replace(/\d+/g, '').split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const rows = (sheet.match(/<row[^>]*>[\s\S]*?<\/row>/g) || []).map((row) => {
    const out = [];
    (row.match(/<c [^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || []).forEach((c) => {
      const ref = /r="([A-Z]+\d+)"/.exec(c)[1];
      const type = (/ t="(\w+)"/.exec(c) || [])[1];
      const v = (/<v>([\s\S]*?)<\/v>/.exec(c) || [])[1];
      out[col(ref)] = v === undefined ? '' : type === 's' ? shared[Number(v)] : type === 'str' || type === 'inlineStr' ? unxml(v) : Number(v);
    });
    for (let i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = '';
    return out;
  });
  const width = Math.max.apply(null, rows.map((r) => r.length));
  return rows.map((r) => r.concat(new Array(width - r.length).fill('')));
}

const FIX = path.join(__dirname, 'fixtures', 'tdoc-list-86178');
const good = xlsxRows(path.join(FIX, 'portal-1355utc-118rows.xlsx'));
const other = xlsxRows(path.join(FIX, 'portal-1515utc-38rows-other-meeting.xlsx'));
const { sandbox: S } = loadCode();
const tdocsOf = (rows) => rows.slice(1).map((r) => r[rows[0].indexOf('TDoc')]).filter(Boolean);
const CFG = { MEETING_ID: '86178', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178', AGENDA_TDOC: 'S4aP260149' };

console.log('A. the two responses');
check('the meeting\'s list: 118 rows, S4aP260067..S4aP260182, the agenda S4aP260149 in it', [good.length, tdocsOf(good)[0], tdocsOf(good).slice(-1)[0], tdocsOf(good).indexOf('S4aP260149') !== -1], [118, 'S4aP260067', 'S4aP260182', true]);
check('the other meeting\'s list: 38 rows, S4aP260008..S4aP260043, no TDoc of the meeting', [other.length, tdocsOf(other)[0], tdocsOf(other).slice(-1)[0], tdocsOf(other).filter((n) => tdocsOf(good).indexOf(n) !== -1)], [38, 'S4aP260008', 'S4aP260043', []]);

console.log('B. the report\'s own identity check on them (the live guard, unchanged)');
const ok = S.evaluateTdocListIdentity_(good, CFG, null, [], 'update');
check('the meeting\'s list: accepted by its agenda TDoc', [ok.ok, ok.via], [true, 'agenda']);
const refused = S.evaluateTdocListIdentity_(other, CFG, null, [], 'update');
check('the other list: refused, with the very words of the live run of 2026-11-25',
  [refused.ok, refused.error], [false, 'The TDoc list was not used: it could not be shown to be the list of this meeting (it has no row for S4aP260149, the agenda TDoc of this meeting -- it may be the list of another meeting)']);
const noAgenda = Object.assign({}, CFG, { AGENDA_TDOC: '' });
const anchor = JSON.stringify({ tdoc: 'S4aP260067', meetingId: '86178', tdocListUrl: CFG.TDOC_LIST_URL });
check('without an agenda TDoc but with the anchor the meeting\'s list gave: the other list refused too', [S.evaluateTdocListIdentity_(other, noAgenda, anchor, [], 'update').ok, /no row for S4aP260067, a TDoc of the list this meeting had before/.test(S.evaluateTdocListIdentity_(other, noAgenda, anchor, [], 'update').error)], [false, true]);
check('... without either, by the report\'s own TDoc tables: refused as well; the meeting\'s list accepted', [S.evaluateTdocListIdentity_(other, noAgenda, null, ['S4aP260067', 'S4aP260149', 'S4aP260150'], 'update').ok, S.evaluateTdocListIdentity_(good, noAgenda, null, ['S4aP260067', 'S4aP260149', 'S4aP260150'], 'update').ok], [false, true]);

console.log('C. the guard is that of production (T-2026.10.33), not weakened');
const ROOT = path.join(__dirname, '..');
const bodyOf = (code, n) => { const at = code.indexOf('\nfunction ' + n + '('); return at === -1 ? null : code.slice(at, code.indexOf('\n}\n', at)); };
let released = null;
const lf = (s) => s.split('\r\n').join('\n');
try { released = lf(execFileSync('git', ['-C', ROOT, 'show', 'b7299c1:Code.js'], { encoding: 'utf8', maxBuffer: 1 << 28, stdio: ['ignore', 'pipe', 'ignore'] })); } catch (e) { released = null; }
const now = lf(fs.readFileSync(path.join(ROOT, 'Code.js'), 'utf8'));
if (released) {
  check('evaluateTdocListIdentity_, tdocListRawNumbers_, tdocListAnchorFor_, assertTdocListIdentity_: byte for byte those of T-2026.10.33',
    ['evaluateTdocListIdentity_', 'tdocListRawNumbers_', 'tdocListAnchorFor_', 'assertTdocListIdentity_'].filter((n) => bodyOf(now, n) !== bodyOf(released, n)), []);
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
