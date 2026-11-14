/**
 * The agenda Word conversion after T-2026.10.19 (MBS SA4#138, 2026-11-11).
 *
 * Two Full Builds in a row failed in the skeleton phase with "Unexpected error
 * while getting the method or property openById on object DocumentApp." on
 * the Google Doc Drive had just made from the agenda TDoc's Word file
 * (S4-261857). Now (AGENDA WORD CONVERSION section of Code.js):
 *
 *   A. the Word file goes to Drive with a Word content type (by extension),
 *      whatever Utilities.unzip() gave it; v3 "name"; id/name/mimeType asked;
 *   B. success: opened once, parsed, trashed, every step logged with the ID;
 *   C. delayed availability: bounded retries (1 s, 2 s, 4 s), then success;
 *   D. permanent failure: four attempts, one error naming the step, the
 *      document and the original message; every attempt's stack in the log;
 *      the temporary document trashed;
 *   E. Drive returns something that is not a Google Doc: refused, trashed,
 *      never opened;
 *   F. Drive.Files.create fails: the step and the original message; no file
 *      to trash;
 *   G. a trash that fails is logged, never thrown;
 *   H. the whole download -> unzip -> convert -> open -> parse path, and
 *      diagnoseAgendaConversion(), which writes nothing;
 *   I. what changed in Code.js: one section and two functions, against
 *      T-2026.10.19.
 *
 * Fakes only: no network, no Google account. All IDs are synthetic.
 *
 * Run: node tests/agenda-word-conversion.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { codeWithoutAgendaReaders, codeWithoutAgendaWordConversion, AGENDA_CONVERSION_SECTION_START, AGENDA_CONVERSION_SECTION_END, AGENDA_CONVERSION_SECTION_FUNCTIONS, AGENDA_CONVERSION_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const GDOC = 'application/vnd.google-apps.document';
const NAME_138 = 'S4-261857 Meeting agenda for SA4#138.docx';
const UNEXPECTED = 'Unexpected error while getting the method or property openById on object DocumentApp.';
const DOC_ID = '1convertedAgendaDocIdForTestsOnly000000000';

/** A Word blob as Utilities.unzip() gives it: a name, a content type, bytes. */
function wordBlob(name, contentType) {
  const blob = {
    _name: name, _type: contentType === undefined ? 'application/octet-stream' : contentType, typeSets: [],
    getName: () => blob._name,
    setName: (n) => { blob._name = n; return blob; },
    getContentType: () => blob._type,
    setContentType: (t) => { blob.typeSets.push(t); blob._type = t; return blob; },
    getBytes: () => new Array(69447).fill(0)
  };
  return blob;
}

/**
 * A sandbox with fake Drive / DriveApp / DocumentApp / Utilities.sleep.
 * o.created: what Drive.Files.create returns (default: a Google Doc);
 * o.createThrows; o.openFailures: how many openById calls fail first
 * (Infinity: all); o.trashThrows.
 */
function world(o) {
  const opts = o || {};
  const { sandbox, docProps } = loadCode({ documentProperties: opts.props || {} });
  const w = { sandbox, docProps, logs: [], creates: [], opens: [], sleeps: [], trashed: [], parsedBodies: [] };
  sandbox.Logger = { log: (m) => { w.logs.push(String(m)); } };
  sandbox.MimeType = { GOOGLE_DOCS: GDOC };
  sandbox.Drive = {
    Files: {
      create: (resource, blob, options) => {
        w.creates.push({ resource: Object.assign({}, resource), blobType: blob.getContentType(), blobName: blob.getName(), options: options ? Object.assign({}, options) : null });
        if (opts.createThrows) throw new Error(opts.createThrows);
        return opts.created !== undefined ? opts.created : { id: DOC_ID, name: resource.name, mimeType: GDOC };
      }
    }
  };
  let openCalls = 0;
  const body = { _fake: 'converted agenda body' };
  sandbox.DocumentApp.openById = (id) => {
    openCalls++;
    w.opens.push(id);
    if (openCalls <= (opts.openFailures || 0)) throw new Error(UNEXPECTED);
    return { getBody: () => body };
  };
  sandbox.DriveApp = {
    getFileById: (id) => ({
      setTrashed: (v) => {
        if (opts.trashThrows) throw new Error('Drive trash failed');
        w.trashed.push([id, v]);
      }
    })
  };
  sandbox.Utilities.sleep = (ms) => { w.sleeps.push(ms); };
  // The parser is not under test here: it records the body it is given.
  sandbox.parseAgendaStructureWithText_ = (b, prefix) => { w.parsedBodies.push([b === body, prefix]); return [{ number: '1', title: 'Opening', level: 1 }, { number: '8.1', title: 'MBS', level: 2 }]; };
  w.run = (blob) => {
    try { return { result: sandbox.parseAgendaWordBlob_(blob) }; } catch (e) { return { error: e.message }; }
  };
  w.agendaLogs = () => w.logs.filter((l) => l.indexOf('[AGENDA]') === 0);
  return w;
}

// ====================================================================== A
console.log('A. what goes to Drive');
{
  const w = world();
  const blob = wordBlob(NAME_138, 'application/octet-stream');
  w.run(blob);
  const c = w.creates[0];
  check('one Drive.Files.create', w.creates.length, 1);
  check('the Word file is sent as a .docx, whatever unzip gave it', [c.blobType, blob.typeSets], [DOCX, [DOCX]]);
  check('target: a native Google Doc, v3 "name" (no v2 "title")', [c.resource.mimeType, /^TEMP_Agenda_\d+$/.test(c.resource.name), 'title' in c.resource], [GDOC, true, false]);
  check('Drive is asked for id, name and mimeType', c.options, { fields: 'id,name,mimeType' });
  check('the log names the file, its size, the type it came with and the type it was sent as', w.agendaLogs()[0],
    '[AGENDA] converting "' + NAME_138 + '" (69447 bytes, content type application/octet-stream, sent as ' + DOCX + ') to a Google Doc');

  const already = world(); const b2 = wordBlob(NAME_138, DOCX); already.run(b2);
  check('a blob that already is a .docx is left as it is', [b2.typeSets, already.creates[0].blobType], [[], DOCX]);
  const none = world(); const b3 = wordBlob(NAME_138, null); none.run(b3);
  check('a blob without a content type gets the .docx type; the log says "none"', [b3.typeSets, /content type none, sent as/.test(none.agendaLogs()[0])], [[DOCX], true]);
  const doc = world(); const b4 = wordBlob('S4-260001 agenda.DOC', 'application/octet-stream'); doc.run(b4);
  check('a .doc (any case) is sent as application/msword', b4.typeSets, ['application/msword']);
  const other = world(); const b5 = wordBlob('agenda.rtf', 'application/rtf'); other.run(b5);
  check('any other extension: the blob type is not touched', b5.typeSets, []);
  const s = w.sandbox;
  check('agendaWordContentType_: pure, by extension only ("#" in the name is no problem)',
    [s.agendaWordContentType_(NAME_138), s.agendaWordContentType_('a.doc'), s.agendaWordContentType_('a.docx.zip'), s.agendaWordContentType_(''), s.agendaWordContentType_(null)],
    [DOCX, 'application/msword', '', '', '']);
  const noDrive = world(); delete noDrive.sandbox.Drive;
  check('without the Advanced Drive service: the same error as before', /^Advanced Drive service is required to convert the Word agenda\./.test(noDrive.run(wordBlob(NAME_138)).error || ''), true);
}

// ====================================================================== B
console.log('B. successful conversion');
{
  const w = world();
  const out = w.run(wordBlob(NAME_138));
  check('the parsed agenda comes back', out, { result: [{ number: '1', title: 'Opening', level: 1 }, { number: '8.1', title: 'MBS', level: 2 }] });
  check('opened once, no wait; the body of that document is parsed with prefix ""', [w.opens, w.sleeps, w.parsedBodies], [[DOC_ID], [], [[true, '']]]);
  check('the temporary document is trashed once', w.trashed, [[DOC_ID, true]]);
  check('the log: convert, Drive result with ID and type, open, trash', w.agendaLogs().slice(1), [
    '[AGENDA] Drive.Files.create returned ' + DOC_ID + ' (' + GDOC + ', "' + w.creates[0].resource.name + '")',
    '[AGENDA] DocumentApp.openById(' + DOC_ID + ') succeeded on attempt 1 of 4',
    // After T-2026.10.20: the reader that read the agenda is named (tests/agenda-readers.test.js).
    '[AGENDA] read by DocumentApp: 2 agenda item(s)',
    '[AGENDA] temporary agenda document ' + DOC_ID + ' moved to the trash'
  ]);
}

// ====================================================================== C
console.log('C. delayed availability: bounded retries');
{
  const w = world({ openFailures: 2 });
  const out = w.run(wordBlob(NAME_138));
  check('the agenda is parsed after two failed opens', !!out.result && !out.error, true);
  check('three opens of the same document, waiting 1 s then 2 s', [w.opens, w.sleeps], [[DOC_ID, DOC_ID, DOC_ID], [1000, 2000]]);
  check('each failed attempt is logged with the original error and its stack', w.agendaLogs().filter((l) => /failed on attempt/.test(l)).map((l) => [/^\[AGENDA\] DocumentApp\.openById\(1converted\w+\)\.getBody\(\) failed on attempt [12] of 4: Error: Unexpected error while getting the method or property openById/.test(l), /\n\s+at /.test(l)]),
    [[true, true], [true, true]]);
  check('success is logged on attempt 3; trashed once', [w.agendaLogs().some((l) => l === '[AGENDA] DocumentApp.openById(' + DOC_ID + ') succeeded on attempt 3 of 4'), w.trashed], [true, [[DOC_ID, true]]]);
  check('the retry plan is bounded: four attempts, 7 s of waiting at most', [w.sandbox.AGENDA_CONVERSION_OPEN_DELAYS_MS_.length, w.sandbox.AGENDA_CONVERSION_OPEN_DELAYS_MS_.reduce((a, b) => a + b, 0)], [4, 7000]);
}

// ====================================================================== D
console.log('D. permanent failure');
{
  const w = world({ openFailures: Infinity });
  const out = w.run(wordBlob(NAME_138));
  // After T-2026.10.20 the Docs API and the Word file are tried next (here neither can be: tests/agenda-readers.test.js);
  // the DocumentApp part of the error is as before.
  check('one error naming the document, the step, the attempts and the original message', out.error.split('; Docs API:')[0],
    'The agenda could not be read -- DocumentApp: The agenda converted from Word (Google Doc ' + DOC_ID + ') could not be opened after 4 attempts over 7 s (DocumentApp.openById): ' + UNEXPECTED);
  check('four opens, waits 1, 2, 4 s; nothing parsed', [w.opens.length, w.sleeps, w.parsedBodies], [4, [1000, 2000, 4000], []]);
  check('every attempt in the log, with its stack', w.agendaLogs().filter((l) => /failed on attempt \d of 4: Error: Unexpected error/.test(l) && /\n\s+at /.test(l)).length, 4);
  check('the temporary document is still trashed (finally)', w.trashed, [[DOC_ID, true]]);
}

// ====================================================================== E
console.log('E. Drive returns something that is not a Google Doc');
{
  const w = world({ created: { id: 'notADoc1', name: 'TEMP_Agenda_1', mimeType: DOCX } });
  const out = w.run(wordBlob(NAME_138));
  check('refused, naming the source file, the file ID and its type', out.error,
    'Drive did not convert the agenda Word file "' + NAME_138 + '" to a Google Doc: file notADoc1 has the type ' + DOCX + '.');
  check('never opened; trashed once', [w.opens, w.trashed], [[], [['notADoc1', true]]]);
  const nomime = world({ created: { id: 'noMime1' } });
  check('no mimeType in the answer: refused as "unknown", trashed', [/has the type unknown\.$/.test(nomime.run(wordBlob(NAME_138)).error), nomime.trashed], [true, [['noMime1', true]]]);
  const noid = world({ created: {} });
  check('no ID: the same error as before; nothing to trash or open', [noid.run(wordBlob(NAME_138)).error, noid.trashed, noid.opens], ['Drive conversion did not return a file ID.', [], []]);
}

// ====================================================================== F
console.log('F. Drive.Files.create fails');
{
  const w = world({ createThrows: 'API call to drive.files.create failed with error: Bad Request' });
  const out = w.run(wordBlob(NAME_138));
  check('the error names the step, the file and the original message', out.error,
    'Drive could not convert the agenda Word file "' + NAME_138 + '" (Drive.Files.create): API call to drive.files.create failed with error: Bad Request');
  check('the original error and its stack are in the log', w.agendaLogs().some((l) => /^\[AGENDA\] Drive\.Files\.create failed for "S4-261857 Meeting agenda for SA4#138\.docx": Error: API call/.test(l) && /\n\s+at /.test(l)), true);
  check('no file: nothing opened, nothing trashed', [w.opens, w.trashed], [[], []]);
}

// ====================================================================== G
console.log('G. a trash that fails is logged, never thrown');
{
  const ok = world({ trashThrows: true });
  check('the agenda is still returned', !!ok.run(wordBlob(NAME_138)).result, true);
  check('the trash failure is logged with the document ID', ok.agendaLogs().some((l) => l.indexOf('[AGENDA] could not move the temporary agenda document ' + DOC_ID + ' to the trash: Error: Drive trash failed') === 0), true);
  const bad = world({ trashThrows: true, openFailures: Infinity });
  check('with an open that fails: the open error is the one thrown, not the trash error', /could not be opened after 4 attempts/.test(bad.run(wordBlob(NAME_138)).error), true);
}

// ====================================================================== H
console.log('H. the whole agenda path, and diagnoseAgendaConversion()');
{
  const PROPS = { AGENDA_TDOC: 'S4-261857', FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', MEETING_TYPE: 'main', REPORT_SUFFIX: 'MBS' };
  function withNetwork(w) {
    w.fetched = [];
    w.sandbox.UrlFetchApp = { fetch: (url) => { w.fetched.push(url); return { getResponseCode: () => 200, getBlob: () => wordBlob('S4-261857.zip', 'application/zip') }; } };
    w.sandbox.Utilities.unzip = () => [wordBlob('__MACOSX/._x.docx', null), wordBlob(NAME_138, 'application/octet-stream')];
    return w;
  }
  const w = withNetwork(world({ props: PROPS, openFailures: 1 }));
  const items = w.sandbox.downloadMeetingAgenda_('S4-261857', PROPS.FTP_BASE);
  check('download -> unzip -> convert -> open (one retry) -> parse', [w.fetched, w.creates.map((c) => [c.blobName, c.blobType]), w.opens.length, items.length, w.trashed.length],
    [[PROPS.FTP_BASE + 'S4-261857.zip'], [[NAME_138, DOCX]], 2, 2, 1]);

  const before = JSON.stringify(w.docProps._store);
  const d = withNetwork(world({ props: PROPS }));
  check('diagnoseAgendaConversion(): the same steps, a one-line answer (after T-2026.10.20: naming the reader)', d.sandbox.diagnoseAgendaConversion(), 'Agenda S4-261857: converted, opened and parsed -- 2 agenda item(s), read by DocumentApp.');
  check('...and it writes no property', JSON.stringify(d.docProps._store), before);
  const f = withNetwork(world({ props: PROPS, openFailures: Infinity }));
  const failed = f.sandbox.diagnoseAgendaConversion();
  check('a failure is answered, not thrown, with the full error', failed.split('; Docs API:')[0],
    'Agenda S4-261857: FAILED -- The agenda could not be read -- DocumentApp: The agenda converted from Word (Google Doc ' + DOC_ID + ') could not be opened after 4 attempts over 7 s (DocumentApp.openById): ' + UNEXPECTED);
  check('...logged with its stack; the temporary document trashed; no property written', [f.agendaLogs().some((l) => /^\[AGENDA\] diagnosis: FAILED for S4-261857: Error: /.test(l) && /\n\s+at /.test(l)), f.trashed.length, JSON.stringify(f.docProps._store)], [true, 1, before]);
  const n = withNetwork(world({ props: { MEETING_TYPE: 'main' } }));
  check('without an agenda TDoc: nothing fetched, nothing converted', [n.sandbox.diagnoseAgendaConversion(), n.fetched, n.creates], ['No agenda TDoc configured.', [], []]);
}

// ====================================================================== I
console.log('I. what changed in Code.js: against T-2026.10.19, one section and two functions of the agenda conversion');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutAgendaWordConversion(FULL);
  // After T-2026.10.20: the agenda readers came later and have their own ledger (tests/agenda-readers.test.js);
  // the comparisons below are made without them.
  const CODE = codeWithoutAgendaReaders(FULL).code;
  check('every change is in Code.js exactly once', view.problems, []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(AGENDA_CONVERSION_SECTION_START), FULL.indexOf(AGENDA_CONVERSION_SECTION_END) + AGENDA_CONVERSION_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), AGENDA_CONVERSION_SECTION_FUNCTIONS);
  check('it stands directly after convertWordBlobToGoogleDoc_()', FULL.indexOf('  return convertAgendaWordBlob_(wordBlob);\n}\n\n' + AGENDA_CONVERSION_SECTION_START) !== -1, true);
  check('it declares its constants and nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [],
    ['var AGENDA_GOOGLE_DOC_MIME_', 'var AGENDA_WORD_CONTENT_TYPES_', 'var AGENDA_CONVERSION_OPEN_DELAYS_MS_']);
  check('Code.js is version 2.26.9 (2.26.9 the session times in a main report; 2.26.4 released it as T-2026.10.20, 2.26.5 the agenda readers; 2.26.6 the TDoc-list download fallback, 2.26.7 the large-build checkpoints, 2.26.8 the meeting time zones; the ledgers take the headers out)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.9');
  let released = null;
  try {
    released = execFileSync('git', ['show', 'template-release/T-2026.10.19:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    released = null;
  }
  if (released === null) {
    console.log('  note: template-release/T-2026.10.19 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js is the Code.js of T-2026.10.19, byte for byte', view.code === released, true);
    check('the functions that changed: parseAgendaWordBlob_ and convertWordBlobToGoogleDoc_',
      functionNames(released).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(CODE, n) !== functionSource(released, n)).sort(), AGENDA_CONVERSION_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions are those of the section; none removed', [functionNames(CODE).filter((n) => functionNames(released).indexOf(n) === -1).sort(), functionNames(released).filter((n) => functionNames(CODE).indexOf(n) === -1)],
      [AGENDA_CONVERSION_SECTION_FUNCTIONS.slice().sort(), []]);
    check('untouched: the build, the agenda choice, the download, the unzip, the parser, the TDoc lists, the update, Close Report',
      ['buildSkeletonWithTdocTables', 'runFullReportBuildCore_', 'parseAgendaForReport_', 'downloadMeetingAgenda_', 'extractWordFromZip_', 'parseAgendaStructureWithText_', 'parseAgendaDocumentById',
        'downloadAndGroupTdocs_', 'continuousUpdateCore_', 'closeReportCore_', 'getReportConfig_'].map((n) => functionSource(CODE, n) === functionSource(released, n)), Array(11).fill(true));
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll agenda Word conversion checks passed.');
process.exit(failures ? 1 : 0);
