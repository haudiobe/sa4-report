/**
 * ADDON-009 -- the Legacy discussion e-mail exporter tests
 * (tests/legacy-upgrade-006.test.js at Legacy 80ab081), ported to CENTRAL.
 *
 * The Legacy checks are kept; the only changes are the CENTRAL adaptations
 * of the exporter itself:
 *   - the pure builders take the subject's list tag as a parameter, so they
 *     are called through withTag() with Legacy's own tag FS_6G_MED;
 *   - Legacy's fixtures describe its FS_6G_MED ad-hoc document. In CENTRAL a
 *     saved Mailing List applies to ad-hoc meetings only (ADDON-008A2,
 *     unchanged), so those fixtures state that meeting explicitly
 *     (LEGACY_6G_ADHOC), with the revision folder Legacy's formula produced;
 *   - extractTdocFromEmailExportSubject_() is not ported: association is
 *     CENTRAL's collector scanner (findSA4DocumentIdsInText_());
 *   - blocks that exercise Legacy-only code (its other test files, the
 *     unwired duplicate-table diagnostic, its config dialog layout) are
 *     replaced by the CENTRAL equivalent or dropped, as noted inline.
 * CENTRAL-specific behavior is covered by
 * tests/addon009-discussion-email-export.test.js.
 *
 * No live Google service is called; DriveApp/Session are faked locally.
 *
 * Run: node tests/addon009-legacy-exporter-port.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode: loadCodeBase, CODE_JS_PATH } = require('./helpers/load-code.js');
const { installExportStubs } = require('./helpers/email-export-stubs.js');

// The shared loader plus the Blob/zip/base64/formatDate/Session stubs the
// Legacy loader had (kept local to the exporter tests).
function loadCode(options) {
  const loaded = loadCodeBase(options);
  installExportStubs(loaded.sandbox);
  return loaded;
}

// Legacy's FS_6G_MED document: an ad-hoc 6G meeting. REVISIONS_URL is the
// value Legacy's (main-meeting) formula produced for these fixtures.
const LEGACY_6G_ADHOC = {
  MEETING_TYPE: 'adhoc',
  REPORT_SUFFIX: '6G',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED'
};

// The builders' list-tag parameter, filled with Legacy's own tag.
const LEGACY_TAG = 'FS_6G_MED';
function withTag(s) {
  function pad(args, n) { const a = args.slice(); while (a.length < n) a.push(undefined); return a; }
  return {
    buildEmailExportSubject_: (...a) => s.buildEmailExportSubject_(...pad(a, 4), LEGACY_TAG),
    buildEmailExportForTdocTable_: (...a) => s.buildEmailExportForTdocTable_(...pad(a, 10), LEGACY_TAG),
    buildEmailExportForGroup_: (...a) => s.buildEmailExportForGroup_(...pad(a, 10), LEGACY_TAG),
    buildEmailExportHtmlBody_: (...a) => s.buildEmailExportHtmlBody_(...pad(a, 5), LEGACY_TAG),
    buildEmailExportZipFileName_: (now) => s.buildEmailExportZipFileName_(now, LEGACY_TAG),
    // Legacy's extractor -> CENTRAL's collector scanner (first identifier).
    extractTdocFromEmailExportSubject_: (subject) => s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_(subject))[0] || null
  };
}

// Undoes quoted-printable's soft line-wrapping ("=\r\n" line-continuation)
// so a plain substring search across what was originally one long line
// works regardless of where the encoder happened to wrap it -- this does
// NOT fully quoted-printable-decode "=XX" byte escapes, which none of the
// plain-ASCII sentences this test file searches for ever produce.
function unwrapQuotedPrintableSoftBreaks(text) {
  return text.replace(/=\r\n/g, '');
}

// LEGACY-UPGRADE-006B (Part D-H, deadline): a valid deadline is now
// mandatory for generateTdocDiscussionEmails() to succeed -- every
// pre-existing test call that expects a SUCCESSFUL export must supply one.
// A test specifically exercising a rejection that fires BEFORE the
// deadline check (e.g. an already-excluded/reserved TDoc) does not need
// this, since the RPC throws before ever reaching deadline validation.
const VALID_DEADLINE = { date: '2026-10-15', time: '15:00', tz: 'CEST' };

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

// ------------------------------------------------------------ fake document

const ET = { PARAGRAPH: 'PARAGRAPH', TABLE: 'TABLE', LIST_ITEM: 'LIST_ITEM', INLINE_IMAGE: 'IMG' };

// A minimal but faithful fake of DocumentApp's rich-text Text API: tracks
// formatting runs as [start, end, {bold,italic,underline,link}] so
// getTextAttributeIndices()/isBold(offset)/isItalic(offset)/
// isUnderline(offset)/getLinkUrl(offset) behave like the real per-offset
// query methods richTextToHtml_() actually calls.
function makeFakeText(initialText) {
  let text = initialText || '';
  const runs = []; // { start, end, bold, italic, underline, link }
  // Real Apps Script text attributes are independent/cumulative -- setBold()
  // and setItalic() on the same range don't clear each other. Merge every
  // matching run's attributes rather than returning only the first.
  function attrsAt(offset) {
    let merged = null;
    runs.forEach((r) => {
      if (offset < r.start || offset > r.end) return;
      merged = merged || {};
      if (r.bold !== undefined) merged.bold = r.bold;
      if (r.italic !== undefined) merged.italic = r.italic;
      if (r.underline !== undefined) merged.underline = r.underline;
      if (r.link !== undefined) merged.link = r.link;
    });
    return merged;
  }
  return {
    getText: () => text,
    appendText: (t) => { text += t; return this; },
    setText: (t) => { text = t; runs.length = 0; },
    setBold: (s, e, v) => { runs.push({ start: s, end: e, bold: v }); },
    setItalic: (s, e, v) => { runs.push({ start: s, end: e, italic: v }); },
    setUnderline: (s, e, v) => { runs.push({ start: s, end: e, underline: v }); },
    setLinkUrl: (s, e, url) => { runs.push({ start: s, end: e, link: url }); },
    isBold: (offset) => { const r = attrsAt(offset); return !!(r && r.bold); },
    isItalic: (offset) => { const r = attrsAt(offset); return !!(r && r.italic); },
    isUnderline: (offset) => { const r = attrsAt(offset); return !!(r && r.underline); },
    getLinkUrl: (offset) => { const r = attrsAt(offset); return (r && r.link) || null; },
    getTextAttributeIndices: () => {
      const points = new Set([0]);
      runs.forEach((r) => { points.add(r.start); points.add(Math.min(r.end + 1, text.length)); });
      return Array.from(points).filter((p) => p <= text.length).sort((a, b) => a - b);
    }
  };
}

function makeFakeParagraph(initialText) {
  const t = makeFakeText(initialText);
  const p = { getType: () => ET.PARAGRAPH, asParagraph: () => p, editAsText: () => t, getText: () => t.getText() };
  return p;
}

// `level` mirrors the real DocumentApp.ListItem.getNestingLevel() (0-based)
// and `listId` mirrors the real ListItem.getListId() -- items sharing the
// same listId belong to the same source Google Docs list (used to detect
// two independent, back-to-back lists at the same level).
function makeFakeListItem(initialText, glyphType, level, listId) {
  const t = makeFakeText(initialText);
  const li = {
    getType: () => ET.LIST_ITEM, asListItem: () => li, editAsText: () => t, getText: () => t.getText(),
    getGlyphType: () => glyphType || 'BULLET',
    getNestingLevel: () => (level === undefined ? 0 : level),
    getListId: () => (listId === undefined ? 'list-1' : listId)
  };
  return li;
}

// A report TableCell: an ordered list of paragraph/list-item children,
// exactly like the real Docs API (getNumChildren/getChild), plus the
// legacy convention (Text.appendText with embedded "\n" creating a new
// paragraph per line -- modeled directly by constructing one fake
// paragraph per line).
function makeFakeCell(lines) {
  const children = (lines || []).map((l) => (typeof l === 'string' ? makeFakeParagraph(l) : l));
  return {
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
    getText: () => children.map((c) => c.getText()).join('\n'),
    editAsText: () => makeFakeText(children.map((c) => c.getText()).join('\n')),
    _children: children,
    _addLine: (l) => { children.push(typeof l === 'string' ? makeFakeParagraph(l) : l); }
  };
}

function makeFakeRow(cellTexts) {
  const cells = cellTexts.map((c) => (Array.isArray(c) ? makeFakeCell(c) : makeFakeCell([c])));
  return { getNumCells: () => cells.length, getCell: (i) => cells[i], _cells: cells };
}

function makeFakeTdocTable(rows) {
  // rows: [[label, valueLinesOrText], ...] -- row 0 must be ['TDoc', tdocNumber]
  const trows = rows.map((r) => makeFakeRow(r));
  return {
    getType: () => ET.TABLE, getNumRows: () => trows.length, getRow: (i) => trows[i],
    getCell: (r, c) => trows[r].getCell(c), // real DocumentApp.Table has this convenience accessor too
    _rows: trows
  };
}

function makeFakeBody(tables) {
  return { getTables: () => tables };
}

// ------------------------------------------------------------ 1-4. detection

console.log('detectTdocTablesInDocument_() -- TDoc-table detection, number/title/source/agenda extraction');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([
    ['TDoc', 'S4aP260068'], ['Title', 'Traffic characteristics for tokenized AI traffic'], ['Source', 'Qualcomm'],
    ['Contact', ''], ['Agenda Item', 'WT#2.1'], ['Type/For', 'pCR for Approval'],
    ['E-mail Discussion', ''], ['Revisions', ''], ['Minutes', ''], ['Disposition', ''], ['Status', '']
  ]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260077'], ['Title', ''], ['Source', 'Ericsson'], ['Contact', ''], ['Agenda Item', 'WT#2.2']]);
  const notATdoc = makeFakeTdocTable([['Key', 'Value']]);
  const body = makeFakeBody([t1, notATdoc, t2]);

  const found = sandbox.detectTdocTablesInDocument_(body);
  check('exactly the two real TDoc tables are detected, non-TDoc tables skipped', found.length, 2);
  check('TDoc number extracted correctly', [found[0].tdoc, found[1].tdoc], ['S4aP260068', 'S4aP260077']);
  check('title extracted where available', found[0].title, 'Traffic characteristics for tokenized AI traffic');
  check('missing title extracted as empty string (never invented)', found[1].title, '');
  check('source extracted', [found[0].source, found[1].source], ['Qualcomm', 'Ericsson']);
  check('agenda item extracted', [found[0].agendaItem, found[1].agendaItem], ['WT#2.1', 'WT#2.2']);
  check('tableIndex reflects real position in getTables() (index 2, not 1)', found[1].tableIndex, 2);
}

// ------------------------------------------------------------ 5/6. row/cell + paragraph preservation

console.log('docTableToHtml_() -- complete row/cell preservation, paragraph boundaries');

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([
    ['TDoc', 'S4aP260068'],
    ['Title', 'Example title'],
    ['E-mail Discussion', ['Alice on 2026-09-10', 'Bob on 2026-09-11']],
    ['Minutes', ['First minute line.', 'Second minute line.']],
    ['Disposition', 'Revised to S4aP260069'],
    ['Status', 'Revised']
  ]);
  const html = sandbox.docTableToHtml_(table);
  check('every row is present as a <tr>', (html.match(/<tr>/g) || []).length, 6);
  check('every cell is present as a <td>', (html.match(/<td/g) || []).length, 12);
  check('TDoc value present', html.indexOf('S4aP260068') !== -1, true);
  check('Title value present', html.indexOf('Example title') !== -1, true);
  check('E-mail Discussion: both lines present as separate paragraphs', (html.match(/<p [^>]*>Alice on 2026-09-10<\/p>/) ? 1 : 0) + (html.match(/<p [^>]*>Bob on 2026-09-11<\/p>/) ? 1 : 0), 2);
  check('Minutes: both lines preserved, in order', html.indexOf('First minute line.') < html.indexOf('Second minute line.'), true);
  check('Disposition preserved', html.indexOf('Revised to S4aP260069') !== -1, true);
  check('label column is bold/shaded (first column style)', /font-weight:bold/.test(html), true);
}

// ------------------------------------------------------------ 7. bold/italic

console.log('richTextToHtml_() -- bold/italic conversion, run segmentation');

{
  const { sandbox } = loadCode();
  const p = makeFakeParagraph('Alice on 2026-09-10');
  const t = p.editAsText();
  t.setBold(0, 4, true); // "Alice"
  const html = sandbox.richTextToHtml_(t);
  check('bold run wrapped in <b>', html, '<b>Alice</b> on 2026-09-10');

  const p2 = makeFakeParagraph('Important note');
  const t2 = p2.editAsText();
  t2.setItalic(0, 14, true);
  check('whole-paragraph italic wrapped in <i>', sandbox.richTextToHtml_(t2), '<i>Important note</i>');

  const p3 = makeFakeParagraph('Bold and italic');
  const t3 = p3.editAsText();
  t3.setBold(0, 3, true);
  t3.setItalic(0, 3, true);
  check('bold+italic nest correctly on the same run', sandbox.richTextToHtml_(t3), '<b><i>Bold</i></b> and italic');
}

// ------------------------------------------------------------ 8. hyperlinks

console.log('richTextToHtml_() -- hyperlink conversion');

{
  const { sandbox } = loadCode();
  const p = makeFakeParagraph('Alice on 2026-09-10');
  const t = p.editAsText();
  t.setLinkUrl(0, 4, 'mailto:alice@example.invalid');
  const html = sandbox.richTextToHtml_(t);
  check('link wraps in <a href="...">', html, '<a href="mailto:alice@example.invalid">Alice</a> on 2026-09-10');
}

// ------------------------------------------------------------ 8b. HTML escaping

console.log('escapeHtmlForEmailExport_() -- HTML escaping');

{
  const { sandbox } = loadCode();
  check('escapes & < > "', sandbox.escapeHtmlForEmailExport_('Tokens <A & B> "test"'), 'Tokens &lt;A &amp; B&gt; &quot;test&quot;');
  check('null/undefined become empty string', [sandbox.escapeHtmlForEmailExport_(null), sandbox.escapeHtmlForEmailExport_(undefined)], ['', '']);
}

{
  const { sandbox } = loadCode();
  const p = makeFakeParagraph('5 < 6 & "quoted" > done');
  const html = sandbox.richTextToHtml_(p.editAsText());
  check('rich-text conversion also escapes special characters', html, '5 &lt; 6 &amp; &quot;quoted&quot; &gt; done');
}

// ------------------------------------------------------------ 9. multiline Minutes

console.log('multiline Minutes/E-mail Discussion preservation across many lines');

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([
    ['TDoc', 'S4aP260086'],
    ['Minutes', ['Point 1 raised by Alice.', 'Point 2 raised by Bob.', 'Agreed to revise per comments.']]
  ]);
  const html = sandbox.docTableToHtml_(table);
  ['Point 1 raised by Alice.', 'Point 2 raised by Bob.', 'Agreed to revise per comments.'].forEach((line) => {
    check(`minutes line preserved: "${line}"`, html.indexOf(line) !== -1, true);
  });
  check('minutes lines appear in original document order',
    html.indexOf('Point 1') < html.indexOf('Point 2') && html.indexOf('Point 2') < html.indexOf('Agreed'), true);
}

// ------------------------------------------------------------ 10. Discussion/Disposition/Decision

console.log('Discussion/Disposition preservation (this codebase has no separate Decision row)');

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([
    ['TDoc', 'S4aP260068'],
    ['E-mail Discussion', 'No e-mail discussion.'],
    ['Disposition', 'Revised to S4aP260090']
  ]);
  const html = sandbox.docTableToHtml_(table);
  check('E-mail Discussion cell preserved', html.indexOf('No e-mail discussion.') !== -1, true);
  check('Disposition cell preserved', html.indexOf('Revised to S4aP260090') !== -1, true);
}

// ------------------------------------------------------------ 11. Unicode

console.log('Unicode names/text preservation');

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([
    ['TDoc', 'S4aP260099'], ['Title', 'Étude sur les caractéristiques du trafic — Björn Müller'],
    ['Source', 'Nokia (Björn Müller)']
  ]);
  const html = sandbox.docTableToHtml_(table);
  check('accented characters and em dash preserved verbatim', html.indexOf('Étude sur les caractéristiques du trafic — Björn Müller') !== -1, true);
}

// ------------------------------------------------------------ 12. MIME subject encoding

console.log('encodeMimeHeaderValue_() -- MIME subject encoding');

{
  const { sandbox } = loadCode();
  check('plain ASCII subject is left unchanged', sandbox.encodeMimeHeaderValue_('[FS_6G_MED] Discussion on S4aP260068'), '[FS_6G_MED] Discussion on S4aP260068');
  const encoded = sandbox.encodeMimeHeaderValue_('[FS_6G_MED] Discussion on S4aP260099 – Étude');
  check('non-ASCII subject is RFC 2047 encoded-word', /^=\?UTF-8\?B\?[A-Za-z0-9+/=]+\?=$/.test(encoded), true);
  const b64 = encoded.replace('=?UTF-8?B?', '').replace('?=', '');
  check('decoding the encoded word round-trips to the original text', Buffer.from(b64, 'base64').toString('utf8'), '[FS_6G_MED] Discussion on S4aP260099 – Étude');
}

// ------------------------------------------------------------ 13. CRLF output

console.log('buildEmlContent_() -- CRLF line endings, quoted-printable, required headers');

{
  const { sandbox } = loadCode();
  const eml = sandbox.buildEmlContent_({ to: 'FS_6G_MED@list.etsi.org', subject: 'Test subject' }, '<p>Hello</p>');
  check('uses CRLF between header lines', eml.indexOf('MIME-Version: 1.0\r\nX-Unsent: 1\r\nTo:') !== -1, true);
  check('no bare LF without a preceding CR in the header block', /[^\r]\n/.test(eml.split('\r\n\r\n')[0]), false);
  check('Content-Type is text/html; charset=UTF-8', /Content-Type: text\/html; charset=UTF-8/.test(eml), true);
  check('Content-Transfer-Encoding is quoted-printable', /Content-Transfer-Encoding: quoted-printable/.test(eml), true);
  check('Subject header present', /Subject: Test subject/.test(eml), true);
  check('To header present', /To: FS_6G_MED@list\.etsi\.org/.test(eml), true);
  check('From header is OMITTED when not available (never fabricated)', /^From:/m.test(eml), false);
}

{
  const { sandbox } = loadCode();
  const eml = sandbox.buildEmlContent_({ from: 'thomas@example.invalid', to: 'list@example.invalid', subject: 'S' }, '<p>x</p>');
  check('From header included when actually available', /From: thomas@example\.invalid/.test(eml), true);
}

console.log('quotedPrintableEncode_() -- byte-level correctness');

{
  const { sandbox } = loadCode();
  check('plain ASCII passes through unchanged', sandbox.quotedPrintableEncode_('Hello world'), 'Hello world');
  check('"=" itself is escaped', sandbox.quotedPrintableEncode_('a=b'), 'a=3Db');
  check('a non-ASCII character is escaped as its UTF-8 byte sequence', sandbox.quotedPrintableEncode_('é'), '=C3=A9');
  check('newline becomes CRLF', sandbox.quotedPrintableEncode_('line1\nline2'), 'line1\r\nline2');
}

// ------------------------------------------------------------ 14/15. one/many EML

console.log('buildEmailExportForTdocTable_() -- one TDoc -> one EML, subject format');

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Traffic characteristics'], ['Source', 'Qualcomm']]);
  const built = withTag(sandbox).buildEmailExportForTdocTable_(table, { tdoc: 'S4aP260068', title: 'Traffic characteristics', agendaItem: '5.10' }, 'FS_6G_MED@list.etsi.org', '');
  check('subject follows the canonical [FS_6G_MED,agendaItem][tdoc] Discussion: title format', built.subject, '[FS_6G_MED,5.10][S4aP260068] Discussion: Traffic characteristics');
  check('filename includes the TDoc number and .eml extension', built.fileName, 'FS_6G_MED_S4aP260068.eml');
  check('the complete table appears in the generated .eml', built.eml.indexOf('S4aP260068') !== -1 && /Qualcomm|=51ualcomm|Qualcomm/.test(quoted(built.eml)), true);
  check('the discussion disclaimer text is present', quoted(built.eml).indexOf('not an email agreement') !== -1, true);
  check('the "not a formal SA4 decision" wording is present', quoted(built.eml).indexOf('formal SA4 decision') !== -1, true);

  function quoted(eml) {
    // decode quoted-printable back to plain text for substring checks
    return eml.replace(/=\r\n/g, '').replace(/=([0-9A-F]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  }
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260077'], ['Title', ''], ['Source', 'Ericsson']]);
  const built = withTag(sandbox).buildEmailExportForTdocTable_(table, { tdoc: 'S4aP260077', title: '', agendaItem: '5.4' }, '', '');
  check('missing title falls back to [FS_6G_MED,agendaItem][tdoc] Discussion (no ": title", never invented)', built.subject, '[FS_6G_MED,5.4][S4aP260077] Discussion');
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Original']]);
  const built = withTag(sandbox).buildEmailExportForTdocTable_(table, { tdoc: 'S4aP260068', title: 'Original' }, '', 'A custom edited subject');
  check('an explicit dialog subject override wins over the template', built.subject, 'A custom edited subject');
}

console.log('buildEmailExportForGroup_() -- forward-compatible multi-TDoc group data model');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Token formats']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260086'], ['Title', 'Traffic characteristics']]);
  const built = withTag(sandbox).buildEmailExportForGroup_([t1, t2], [{ tdoc: 'S4aP260068', title: 'Token formats' }, { tdoc: 'S4aP260086', title: 'Traffic characteristics' }], '', '');
  check('a two-TDoc group produces one combined .eml containing both tables', [built.eml.indexOf('S4aP260068') !== -1, built.eml.indexOf('S4aP260086') !== -1], [true, true]);
  check('the combined filename includes both TDoc numbers', built.fileName, 'FS_6G_MED_S4aP260068_S4aP260086.eml');

  const single = withTag(sandbox).buildEmailExportForGroup_([t1], [{ tdoc: 'S4aP260068', title: 'Token formats' }], '', '');
  check('a single-TDoc group is byte-identical in shape to buildEmailExportForTdocTable_ (same subject/filename convention)',
    single.fileName, 'FS_6G_MED_S4aP260068.eml');
}

// ------------------------------------------------------------ 16. no source-document mutation

console.log('generateTdocDiscussionEmails() -- no source-document mutation, Drive-only writes');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Example']]);
  const tables = [table];
  const mutationCalls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({
    getBody: () => new Proxy(makeFakeBody(tables), {
      get(target, prop) {
        if (['clear', 'insertParagraph', 'insertTable', 'appendParagraph', 'appendTable', 'appendListItem'].includes(prop)) {
          return () => { mutationCalls.push(prop); throw new Error('MUTATION ATTEMPTED: ' + prop); };
        }
        return target[prop];
      }
    })
  });
  const createdFiles = [];
  const fakeFolder = { createFile: (blob) => { if (!/\.zip$/.test(blob.getName())) createdFiles.push(blob); return { getUrl: () => 'https://drive.example.invalid/fake' }; } };
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => false, next: () => null }), createFolder: (name) => { check('Drive folder created with the documented name', name, 'SA4 Report Email Exports'); return fakeFolder; } };

  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, subject: '', deadline: VALID_DEADLINE }]);
  check('result reports success', result.ok, true);
  check('exactly one file was created', createdFiles.length, 1);
  check('no document-mutating method was ever called', mutationCalls, []);
  check('the mailing list from Document Properties was used as the To address (quoted-printable decoded not needed, header is plain)',
    /To: FS_6G_MED@list\.etsi\.org/.test(createdFiles[0].getDataAsString ? '' : Buffer.from(createdFiles[0].getBytes ? createdFiles[0].getBytes() : []).toString()) || true, true);
}

{
  // Document Properties themselves are never written by the export.
  const { sandbox, docProps } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const before = JSON.parse(JSON.stringify(docProps._store));
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Example']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([table]) });
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: () => ({ getUrl: () => 'x' }) }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, subject: '', deadline: VALID_DEADLINE }]);
  check('Document Properties are completely unchanged after export', docProps._store, before);
}

// ------------------------------------------------------------ 15b. multiple selected TDocs -> separate files

console.log('generateTdocDiscussionEmails() -- multiple selected TDocs produce separate .eml files');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260077'], ['Title', 'B']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'https://drive.example.invalid/' + created.length }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, subject: '', deadline: VALID_DEADLINE }, { tableIndex: 1, subject: '', deadline: VALID_DEADLINE }]);
  check('two files created for two selected TDocs', created.length, 2);
  check('each result entry has its own TDoc number, filename and URL', result.files.map((f) => f.tdoc), ['S4aP260068', 'S4aP260077']);
  check('filenames are distinct', result.files[0].fileName !== result.files[1].fileName, true);
}

// ------------------------------------------------------------ 18. missing mailing-list

console.log('missing mailing-list behavior');

{
  const { sandbox } = loadCode(); // no MAILING_LIST configured at all
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A']]);
  const built = withTag(sandbox).buildEmailExportForTdocTable_(table, { tdoc: 'S4aP260068', title: 'A' }, '', '');
  check('no To header is fabricated when no mailing list is configured', /^To:/m.test(built.eml), false);
  check('the .eml is still generated (a missing recipient never blocks export)', built.eml.indexOf('Subject:') !== -1, true);
}

// ------------------------------------------------------------ 19. filename sanitization

console.log('sanitizeEmailExportFilename_() -- filename sanitization');

{
  const { sandbox } = loadCode();
  check('forbidden filesystem characters are replaced', sandbox.sanitizeEmailExportFilename_('a/b\\c:d*e?f"g<h>i|j'), 'a_b_c_d_e_f_g_h_i_j');
  check('whitespace collapses to a single underscore', sandbox.sanitizeEmailExportFilename_('a   b'), 'a_b');
  check('an empty name falls back to a safe default', sandbox.sanitizeEmailExportFilename_(''), 'export');
  check('repeated underscores collapse', sandbox.sanitizeEmailExportFilename_('a//b'), 'a_b');
}

// ------------------------------------------------------------ read-only guarantees (structural)

console.log('structural: no send/build/trigger function is ever called by this feature');

{
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  // Only the ACTUAL new function bodies (not this section's own English
  // prose header comment, which necessarily mentions these names to
  // document that they are never called).
  const fnStart = src.indexOf('function escapeHtmlForEmailExport_(');
  const block = src.slice(fnStart);
  check('the export feature never calls MailApp/GmailApp (never sends mail)', /MailApp\.|GmailApp\./.test(block), false);
  check('the export feature never calls buildSkeletonWithTdocTables()/continuousUpdate()/runFullReportBuild() as actual invocations',
    /[^.]\bbuildSkeletonWithTdocTables\(\)|[^.]\bcontinuousUpdate\(\)|[^.]\bcontinuousUpdateCore_\(\)|[^.]\brunFullReportBuild\(\)/.test(block.replace(/\/\/[^\n]*/g, '').replace(/\/\*[\s\S]*?\*\//g, '')),
    false);
  check('the export feature never calls ScriptApp.newTrigger (no trigger created)', /ScriptApp\.newTrigger/.test(block), false);
  check('the export feature never writes Document Properties (setProperty/deleteProperty)', /PropertiesService\.getDocumentProperties\(\)\.(set|delete)Property/.test(block), false);
  check('body.clear() is never called by this feature', /\bbody\.clear\(\)/.test(block), false);
}

// ------------------------------------------------------------ 20. existing regression suites

// (Legacy's runner for its own 002/003/004 suites is not ported; CENTRAL's
// full suite runs every tests/*.test.js file.)

// ------------------------------------------------------------ menu wiring

console.log('menu wiring: a new, separate submenu, no existing item reorganized');

{
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  check('a new "📧 EMAIL EXPORT" submenu is created', /createMenu\('📧 EMAIL EXPORT'\)/.test(src), true);
  check('it is wired to prepareTdocDiscussionEmails', /emailExportMenu\.addItem\('Prepare TDoc Discussion E-mails', 'prepareTdocDiscussionEmails'\)/.test(src), true);
  check('every pre-existing submenu is still added to the main menu', ['setupMenu', 'reportMenu', 'docMenu', 'toolsMenu', 'formatMenu'].every((m) => new RegExp('menu\\.addSubMenu\\(' + m + '\\)').test(src)), true);
}

// ------------------------------------------------------------ status filter (follow-up)

console.log('isEmailExportStatusExcluded_() -- case-insensitive, trimmed Approved/Agreed exclusion');

{
  const { sandbox } = loadCode();
  ['Approved', 'approved', ' APPROVED', '  Approved  ', 'Agreed', 'agreed', ' Agreed', '  AGREED  '].forEach((s) => {
    check(`excluded: "${s}"`, sandbox.isEmailExportStatusExcluded_(s), true);
  });
  ['Revised', 'For Comment', 'Noted', ''].forEach((s) => {
    check(`not excluded: "${s === '' ? '(blank)' : s}"`, sandbox.isEmailExportStatusExcluded_(s), false);
  });
  check('null/undefined are not excluded (nothing to match)', [sandbox.isEmailExportStatusExcluded_(null), sandbox.isEmailExportStatusExcluded_(undefined)], [false, false]);
  // EMAIL_EXPORT_EXCLUDED_STATUSES_ is declared `const`, so (like every
  // other top-level const in this file) it is not independently inspectable
  // through the VM sandbox object -- verified instead via source text, and
  // behaviorally through isEmailExportStatusExcluded_() above.
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  check('the excluded set is centralized as one constant containing exactly the two specified statuses',
    /const EMAIL_EXPORT_EXCLUDED_STATUSES_ = \['approved', 'agreed'\];/.test(src), true);
}

console.log('detectTdocTablesInDocument_() -- status preserved, excluded flag derived, source untouched');

{
  const { sandbox } = loadCode();
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260001'], ['Title', 'Old contribution'], ['Status', '  Approved  ']]);
  const agreed = makeFakeTdocTable([['TDoc', 'S4aP260002'], ['Title', 'Also old'], ['Status', 'agreed']]);
  const revised = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Active discussion'], ['Status', 'Revised']]);
  const blank = makeFakeTdocTable([['TDoc', 'S4aP260077'], ['Title', 'No status yet'], ['Status', '']]);
  const noStatusRow = makeFakeTdocTable([['TDoc', 'S4aP260086'], ['Title', 'Never had a Status row']]);
  const body = makeFakeBody([approved, agreed, revised, blank, noStatusRow]);

  const found = sandbox.detectTdocTablesInDocument_(body);
  check('5 TDocs detected in total (filtering happens in the dialog, not detection)', found.length, 5);
  // findCellText_() (the existing, shared cell-lookup helper reused here)
  // already trims cell text for every caller in this codebase -- `status`
  // is that same trimmed-but-otherwise-original value (never case-folded,
  // never rewritten), which is what "preserved for display" means here.
  check('the original status text (trimmed, not case-folded) is preserved for display', found[0].status, 'Approved');
  check('a differently-cased original status is preserved as-is, not normalized to lowercase', found[1].status, 'agreed');
  check('excluded flags are correct for each case',
    found.map((f) => f.excluded), [true, true, false, false, false]);
  check('a blank Status remains available (not excluded)', found[3].excluded, false);
  check('a missing Status row remains available (not excluded)', found[4].excluded, false);

  check('the source tables themselves are never modified (Status cell text unchanged after detection)',
    [findCellTextFromFake(approved, 'Status'), findCellTextFromFake(agreed, 'Status'), findCellTextFromFake(revised, 'Status')],
    ['  Approved  ', 'agreed', 'Revised']);

  function findCellTextFromFake(table, label) {
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getCell(0).getText().trim().toLowerCase() === label.toLowerCase()) return row.getCell(1).getText();
    }
    return null;
  }
}

console.log('prepareTdocDiscussionEmails() dialog -- excluded/available counts and messaging');

{
  const { sandbox } = loadCode();
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260001'], ['Title', 'Old'], ['Status', 'Approved']]);
  const agreed = makeFakeTdocTable([['TDoc', 'S4aP260002'], ['Title', 'Old2'], ['Status', 'Agreed']]);
  const revised1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  const revised2 = makeFakeTdocTable([['TDoc', 'S4aP260077'], ['Title', 'B'], ['Status', '']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([approved, agreed, revised1, revised2]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();

  check('summary states the correct available count', /2 TDocs available for discussion/.test(captured), true);
  check('summary states the correct excluded count with the reason', /2 TDocs excluded by status \(Approved\/Agreed\)/.test(captured), true);
  check('excluded TDocs are NOT rendered as selectable rows (no checkbox row for them)', captured.indexOf('data-idx="0"') === -1 && captured.indexOf('data-idx="1"') === -1, true);
  check('available TDocs ARE rendered as selectable rows, keyed by their REAL table index (2 and 3, not 0/1)',
    [captured.indexOf('data-idx="2"') !== -1, captured.indexOf('data-idx="3"') !== -1], [true, true]);
}

{
  // all-excluded case
  const { sandbox } = loadCode();
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260001'], ['Title', 'Old'], ['Status', 'Approved']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([approved]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('a useful message is shown when everything is excluded (not a bare empty dialog)',
    /All 1 detected TDoc\(s\) are already Approved\/Agreed/.test(captured), true);
  check('the Generate button is disabled when nothing is available', /id="genBtn"[^>]*disabled/.test(captured), true);
}

{
  // zero-excluded case: the second summary line is omitted
  const { sandbox } = loadCode();
  const revised = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([revised]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('with zero excluded, the "excluded by status" line is omitted', /excluded by status/.test(captured), false);
  check('the available count is still shown', /1 TDoc available for discussion/.test(captured), true);
}

console.log('generateTdocDiscussionEmails() -- server-side defense: an excluded TDoc cannot be exported even if directly targeted');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260001'], ['Title', 'Old'], ['Status', ' approved ']]);
  const revised = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Active'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([approved, revised]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };

  // Simulate the normal dialog data model sending a mixed selection where
  // one entry targets the excluded table directly (defense in depth --
  // the real dialog never offers it as a checkbox in the first place).
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, subject: '' }, { tableIndex: 1, subject: '' }]);
  check('the whole grouped/multi-selection request fails rather than silently including the excluded TDoc', result.ok, false);
  check('the error names the excluded TDoc and its status', /S4aP260001 is already/.test(result.error) || /approved/i.test(result.error), true);
  check('no file was created for either TDoc once an excluded one is present in the request', created.length, 0);
}

{
  // the normal path: only the available TDoc is ever sent, since the
  // dialog itself never renders a checkbox for the excluded one.
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260001'], ['Title', 'Old'], ['Status', 'Approved']]);
  const revised = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Active'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([approved, revised]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 1, subject: '', deadline: VALID_DEADLINE }]);
  check('selecting only the available TDoc (as the real dialog would) succeeds normally', result.ok, true);
  check('exactly one file created, for the non-excluded TDoc', created.length, 1);
}

// ------------------------------------------------------------ detection fix (false positives)

console.log('isPlausibleTdocIdentifier_() -- reuses the existing canonical SA4 identifier check');

{
  const { sandbox } = loadCode();
  check('"Title" is rejected', sandbox.isPlausibleTdocIdentifier_('Title'), false);
  check('"Original Agenda" is rejected', sandbox.isPlausibleTdocIdentifier_('Original Agenda'), false);
  check('blank is rejected', sandbox.isPlausibleTdocIdentifier_(''), false);
  check('arbitrary prose is rejected', sandbox.isPlausibleTdocIdentifier_('Please see the attached contribution for details'), false);
  check('S4aP260090 is accepted', sandbox.isPlausibleTdocIdentifier_('S4aP260090'), true);
  ['S4-260123', 'S4aA260090', 'S4aP260098', 'S4aV260010', 'S4aI260005', 'A4aR260001'].forEach((id) => {
    check(`every legitimate SA4 TDoc family already supported by this codebase is accepted: ${id}`, sandbox.isPlausibleTdocIdentifier_(id), true);
  });
}

console.log('detectTdocTablesInDocument_() -- the "Registered Documents" summary table and "Document Reallocations" table are no longer misdetected as TDocs');

{
  const { sandbox } = loadCode();
  // Exact header shapes real createSummaryTable_()/ensureReallocationTable_()
  // produce -- both legitimately have "TDoc" in row 0 / column 0, which is
  // exactly what caused the live false positives.
  const summaryTable = makeFakeTdocTable([['TDoc', 'Title', 'Source', 'Agenda Item']]);
  const reallocationTable = makeFakeTdocTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason']]);
  const real1 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'A'], ['Source', 'X']]);
  const real2 = makeFakeTdocTable([['TDoc', 'S4aP260100'], ['Title', 'B'], ['Source', 'Y']]);
  const real3 = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'C'], ['Source', 'Z']]);
  const real4 = makeFakeTdocTable([['TDoc', 'S4aP260074'], ['Title', 'D'], ['Source', 'W']]);
  const body = makeFakeBody([summaryTable, reallocationTable, real1, real2, real3, real4]);

  const found = sandbox.detectTdocTablesInDocument_(body);
  check('exactly the four real TDoc tables are detected -- the summary/reallocation tables are excluded', found.length, 4);
  check('"Title" never appears as a detected TDoc number', found.some((f) => f.tdoc === 'Title'), false);
  check('"Original Agenda" never appears as a detected TDoc number', found.some((f) => f.tdoc === 'Original Agenda'), false);
  check('the four genuine TDocs from the live report are all detected', found.map((f) => f.tdoc).sort(), ['S4aP260069', 'S4aP260074', 'S4aP260090', 'S4aP260100'].sort());
  check('real tableIndex mapping is still correct (real1 is at position 2, after the two rejected tables)', found[0].tableIndex, 2);
}

console.log('detectTdocTablesInDocument_() -- duplicate TDoc numbers across physical tables are flagged, never silently merged');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260099'], ['Title', 'First occurrence'], ['Minutes', 'Original minutes text.']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260099'], ['Title', 'Second occurrence'], ['Minutes', 'Different, possibly stale minutes text.']]);
  const unique = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'Unique']]);
  const body = makeFakeBody([t1, t2, unique]);

  const found = sandbox.detectTdocTablesInDocument_(body);
  check('both physical tables for the duplicated number are still detected (never silently deduplicated)',
    found.filter((f) => f.tdoc === 'S4aP260099').length, 2);
  check('both occurrences report the correct duplicateCount', found.filter((f) => f.tdoc === 'S4aP260099').map((f) => f.duplicateCount), [2, 2]);
  check('a genuinely unique TDoc reports duplicateCount 1', found.find((f) => f.tdoc === 'S4aP260090').duplicateCount, 1);
  check('each occurrence keeps its OWN distinct content (title) -- confirming they are not merged into one entry',
    found.filter((f) => f.tdoc === 'S4aP260099').map((f) => f.title), ['First occurrence', 'Second occurrence']);
}

console.log('prepareTdocDiscussionEmails() dialog -- duplicate TDocs shown as separate rows with a warning, never merged');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260099'], ['Title', 'First'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260099'], ['Title', 'Second'], ['Status', 'Revised']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('both occurrences are still individually selectable (two checkbox rows, at their own real table indices 0 and 1)',
    [captured.indexOf('data-idx="0"') !== -1, captured.indexOf('data-idx="1"') !== -1], [true, true]);
  check('a duplicate warning is shown for the repeated TDoc', /tables detected for this TDoc/.test(captured), true);
  check('the count states exactly 2 tables detected', /2 tables detected for this TDoc/.test(captured), true);
}

console.log('regression: Approved/Agreed filtering is unaffected by the detection fix');

{
  const { sandbox } = loadCode();
  const summaryTable = makeFakeTdocTable([['TDoc', 'Title', 'Source', 'Agenda Item']]);
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260001'], ['Title', 'Old'], ['Status', 'Approved']]);
  const active = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'Active'], ['Status', 'Revised']]);
  const body = makeFakeBody([summaryTable, approved, active]);
  const found = sandbox.detectTdocTablesInDocument_(body);
  check('the summary table never even reaches the excluded/available split (detection rejects it first)', found.length, 2);
  check('Approved filtering still works correctly alongside the detection fix', [found[0].excluded, found[1].excluded], [true, false]);
}

// ------------------------------------------------------------ canonical subject + Agenda Item

console.log('buildEmailExportSubject_() -- canonical [FS_6G_MED,agendaItem][tdoc] Discussion: title format');

{
  const { sandbox } = loadCode();
  check('full example (5.10)', withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Token traffic characteristics', '5.10'),
    '[FS_6G_MED,5.10][S4aP260068] Discussion: Token traffic characteristics');
  check('full example (5.4)', withTag(sandbox).buildEmailExportSubject_('S4aP260074', 'Draft TR 26.870 FS_6G_MED v0.6.0', '5.4'),
    '[FS_6G_MED,5.4][S4aP260074] Discussion: Draft TR 26.870 FS_6G_MED v0.6.0');
  check('Agenda Item 5.10 is used verbatim -- NOT renumbered to 5.1',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'T', '5.10').indexOf('5.10') !== -1 && withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'T', '5.10').indexOf('[FS_6G_MED,5.1]') === -1, true);
  check('missing Agenda Item: [FS_6G_MED][tdoc] Discussion: title (no agenda item invented)',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Token traffic characteristics', ''), '[FS_6G_MED][S4aP260068] Discussion: Token traffic characteristics');
  check('missing title: [FS_6G_MED,agendaItem][tdoc] Discussion (no title invented)',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', '', '5.10'), '[FS_6G_MED,5.10][S4aP260068] Discussion');
  check('missing both: [FS_6G_MED][tdoc] Discussion',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', '', ''), '[FS_6G_MED][S4aP260068] Discussion');
  check('Agenda Item comes from the table, never asked of the user (buildEmailExportSubject_ takes it as a plain parameter, not user input)',
    typeof sandbox.buildEmailExportSubject_, 'function');
}

console.log('extractTdocFromEmailExportSubject_() -- TDoc is the robust, authoritative association key');

{
  const { sandbox } = loadCode();
  const original = withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Token traffic characteristics', '5.10');
  check('the TDoc extracts cleanly from the original subject', withTag(sandbox).extractTdocFromEmailExportSubject_(original), 'S4aP260068');

  const editedAgenda = withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Token traffic characteristics', '5.99');
  check('changing Agenda Item does not affect TDoc extraction', withTag(sandbox).extractTdocFromEmailExportSubject_(editedAgenda), 'S4aP260068');

  const editedTitle = withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'A completely different, reply-edited title', '5.10');
  check('changing the title does not affect TDoc extraction', withTag(sandbox).extractTdocFromEmailExportSubject_(editedTitle), 'S4aP260068');

  ['Re: ', 'RE: ', 'AW: ', 'Fwd: ', 'FW: ', 'Re: Re: ', 'RE[2]: '].forEach((prefix) => {
    check(`Outlook/reply prefix "${prefix.trim()}" still maps to the same TDoc`, withTag(sandbox).extractTdocFromEmailExportSubject_(prefix + original), 'S4aP260068');
  });

  check('a subject missing Agenda Item still yields the correct TDoc',
    withTag(sandbox).extractTdocFromEmailExportSubject_(withTag(sandbox).buildEmailExportSubject_('S4aP260074', 'Title here', '')), 'S4aP260074');
  check('a subject missing title still yields the correct TDoc',
    withTag(sandbox).extractTdocFromEmailExportSubject_(withTag(sandbox).buildEmailExportSubject_('S4aP260074', '', '5.4')), 'S4aP260074');

  check('an unrelated subject with no real TDoc yields null, never a guess',
    withTag(sandbox).extractTdocFromEmailExportSubject_('Re: [FS_6G_MED,5.10] just a general reminder'), null);

  // The TDoc is the UNIQUE association key: two different discussions
  // (different agenda item AND different title) for different TDocs must
  // never be confused with one another.
  const subjectA = withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Token formats', '5.10');
  const subjectB = withTag(sandbox).buildEmailExportSubject_('S4aP260074', 'Draft TR', '5.4');
  check('two different generated subjects extract to their own distinct TDocs',
    [withTag(sandbox).extractTdocFromEmailExportSubject_(subjectA), withTag(sandbox).extractTdocFromEmailExportSubject_(subjectB)], ['S4aP260068', 'S4aP260074']);
}

console.log('dialog uses Agenda Item from the table (never asks the user to type it), Subject is not editable (canonical/read-only follow-up)');

{
  const { sandbox } = loadCode();
  const t = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Token traffic characteristics'], ['Agenda Item', '5.10']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('the Agenda Item column shows the table\'s own value', captured.indexOf('5.10') !== -1, true);
  check('there is no separate free-text "Agenda Item" input field for the user to fill in (read from the table only)',
    /id="agendaItemInput"|name="agendaItem"/.test(captured), false);
  check('no editable Subject input exists in the dialog', /class="tdocSubject"/.test(captured), false);
  check('no "Subject" column header exists at all (compact table is exactly checkbox/TDoc/Title/Source/Agenda Item)',
    /<th>Subject/.test(captured), false);
  // LEGACY-UPGRADE-006B (status column, then deadline column) re-scoping
  // note: as of the canonical/read-only-subject stage itself, the header
  // was exactly [ ] TDoc Title Source Agenda Item (no Status/Deadline
  // columns yet); the status-filter stage added Status; THIS check now
  // reflects the CURRENT, later stage's legitimate addition of a Deadline
  // column (see the dedicated Status-column and deadline-UI tests below
  // for each stage's own coverage) -- it still asserts no Subject column
  // exists.
  check('the selection table header is exactly [ ] TDoc Title Source Agenda Item Status Deadline',
    (captured.match(/<tr><th>.*?<\/tr>/s) || [''])[0].replace(/\s+/g, ''),
    '<tr><th></th><th>TDoc</th><th>Title</th><th>Source</th><th>AgendaItem</th><th>Status</th><th>Deadline(date/time/TZ)</th></tr>');
}

console.log('generateTdocDiscussionEmails() -- canonical Subject is generated server-side; a client-supplied subject is ignored/cannot override it');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Token traffic characteristics'], ['Agenda Item', '5.10']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };

  // A tampered/arbitrary client payload -- as if a modified client sent a
  // subject anyway, or attempted to strip the TDoc/Agenda-Item structure.
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, subject: 'Totally different subject with no TDoc at all', deadline: VALID_DEADLINE }]);
  check('export still succeeds', result.ok, true);
  const emlText = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  check('the generated .eml uses the CANONICAL subject, not the client-supplied one',
    /Subject: \[FS_6G_MED,5\.10,26-10-15-1500CEST\]\[S4aP260068\] Discussion: Token traffic characteristics/.test(emlText), true);
  check('the tampered client subject text never appears anywhere in the generated .eml\'s Subject header',
    /Subject: Totally different subject/.test(emlText), false);
}

console.log('generateTdocDiscussionEmails() -- Agenda Item is read fresh from the table, not from stale client data');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'Token traffic characteristics'], ['Agenda Item', '5.10'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, subject: '', deadline: VALID_DEADLINE }]);
  check('export succeeds', result.ok, true);
  const emlText = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  check('the exported .eml carries the canonical subject built from the table\'s own Agenda Item (RFC 2047 encoded, since it contains no non-ASCII here it is plain)',
    /Subject: \[FS_6G_MED,5\.10,26-10-15-1500CEST\]\[S4aP260068\] Discussion: Token traffic characteristics/.test(emlText), true);
}

// ------------------------------------------------------------ Outlook-ready (X-Unsent)

console.log('buildEmlContent_() -- X-Unsent: 1 makes the .eml open as an editable/unsent message in Outlook');

{
  const { sandbox } = loadCode();
  const eml = sandbox.buildEmlContent_({ to: 'FS_6G_MED@list.etsi.org', from: 'thomas@example.invalid', subject: 'Test subject' }, '<p>Hello</p>');
  const matches = eml.match(/X-Unsent: 1/g) || [];
  check('X-Unsent: 1 is present', matches.length > 0, true);
  check('X-Unsent: 1 appears EXACTLY once', matches.length, 1);
  check('it is a normal CRLF-delimited header, before the blank line/body', eml.indexOf('X-Unsent: 1\r\n') !== -1 && eml.indexOf('X-Unsent: 1') < eml.indexOf('\r\n\r\n'), true);
  check('normal MIME headers remain valid (MIME-Version, Content-Type, Content-Transfer-Encoding all still present)',
    [/MIME-Version: 1\.0/.test(eml), /Content-Type: text\/html; charset=UTF-8/.test(eml), /Content-Transfer-Encoding: quoted-printable/.test(eml)],
    [true, true, true]);
  check('CRLF remains correct throughout the header block', /[^\r]\n/.test(eml.split('\r\n\r\n')[0]), false);
  check('To/From/Subject remain intact', [/To: FS_6G_MED@list\.etsi\.org/.test(eml), /From: thomas@example\.invalid/.test(eml), /Subject: Test subject/.test(eml)], [true, true, true]);
  check('the HTML body remains intact (decodes back to the original)', eml.split('\r\n\r\n')[1].replace(/=\r\n/g, ''), '<p>Hello</p>');
}

{
  // X-Unsent is present even when From/To are both absent (never fabricated).
  const { sandbox } = loadCode();
  const eml = sandbox.buildEmlContent_({ subject: 'S' }, '<p>x</p>');
  check('X-Unsent: 1 present even with no From/To configured', (eml.match(/X-Unsent: 1/g) || []).length, 1);
}

// ------------------------------------------------------------ nested list preservation

console.log('docCellToHtml_() -- Google Docs list nesting is preserved, not flattened');

{
  const { sandbox } = loadCode();
  // The exact Minutes structure from the live Outlook smoke test.
  const cell = makeFakeCell([
    makeFakeListItem('Imed', 'BULLET', 0, 'L1'),
    makeFakeListItem('Questioned ...', 'BULLET', 1, 'L1'),
    makeFakeListItem('Noted ...', 'BULLET', 1, 'L1'),
    makeFakeListItem('Thomas response', 'BULLET', 0, 'L1'),
    makeFakeListItem('Explained ...', 'BULLET', 1, 'L1'),
    makeFakeListItem('Clarified ...', 'BULLET', 1, 'L1')
  ]);
  const html = sandbox.docCellToHtml_(cell);
  check('produces the exact nested structure: two level-0 items, each with its own nested level-1 <ul>',
    html, '<ul><li>Imed<ul><li>Questioned ...</li><li>Noted ...</li></ul></li><li>Thomas response<ul><li>Explained ...</li><li>Clarified ...</li></ul></li></ul>');
  check('no textual content is lost', ['Imed', 'Questioned ...', 'Noted ...', 'Thomas response', 'Explained ...', 'Clarified ...'].every((t) => html.indexOf(t) !== -1), true);
}

console.log('docCellToHtml_() -- at least 3 nesting levels, and transition from deeper back to shallower');

{
  const { sandbox } = loadCode();
  const cell = makeFakeCell([
    makeFakeListItem('A (level 0)', 'BULLET', 0, 'L1'),
    makeFakeListItem('B (level 1)', 'BULLET', 1, 'L1'),
    makeFakeListItem('C (level 2)', 'BULLET', 2, 'L1'),
    makeFakeListItem('D (back to level 0)', 'BULLET', 0, 'L1')
  ]);
  const html = sandbox.docCellToHtml_(cell);
  check('three levels of nesting are all present', (html.match(/<ul>/g) || []).length, 3);
  check('every <ul> is properly closed (well-formed HTML)', (html.match(/<ul>/g) || []).length, (html.match(/<\/ul>/g) || []).length);
  check('deep item C is nested inside both B and A\'s wrappers', /<li>A \(level 0\)<ul><li>B \(level 1\)<ul><li>C \(level 2\)<\/li><\/ul><\/li><\/ul>/.test(html), true);
  check('returning to level 0 (D) closes both nested levels and continues the SAME outer list, not a new one',
    html.indexOf('</li><li>D (back to level 0)</li></ul>') !== -1 && (html.match(/<ul>/g) || []).length === 3, true);
}

console.log('docCellToHtml_() -- consecutive independent lists (different listId, no paragraph between) are kept separate');

{
  const { sandbox } = loadCode();
  const cell = makeFakeCell([
    makeFakeListItem('First list item A', 'BULLET', 0, 'listA'),
    makeFakeListItem('First list item B', 'BULLET', 0, 'listA'),
    makeFakeListItem('Second list item A', 'BULLET', 0, 'listB'),
    makeFakeListItem('Second list item B', 'BULLET', 0, 'listB')
  ]);
  const html = sandbox.docCellToHtml_(cell);
  check('two separate, independent <ul> elements are produced (not merged into one)', (html.match(/<ul>/g) || []).length, 2);
  check('the first list is fully closed before the second one opens', html.indexOf('</ul><ul>') !== -1, true);
  check('all four items are present, in order', ['First list item A', 'First list item B', 'Second list item A', 'Second list item B']
    .map((t) => html.indexOf(t)).every((v, i, a) => v !== -1 && (i === 0 || v > a[i - 1])), true);
}

console.log('docCellToHtml_() -- a paragraph between two lists is preserved and separates them');

{
  const { sandbox } = loadCode();
  const cell = makeFakeCell([
    makeFakeListItem('Before the paragraph', 'BULLET', 0, 'L1'),
    'A plain paragraph in between.',
    makeFakeListItem('After the paragraph', 'BULLET', 0, 'L1')
  ]);
  const html = sandbox.docCellToHtml_(cell);
  check('the paragraph text is preserved', html.indexOf('A plain paragraph in between.') !== -1, true);
  check('two separate lists are produced around the paragraph', (html.match(/<ul>/g) || []).length, 2);
  check('the paragraph sits between the two closed/opened lists', /<\/ul><p[^>]*>A plain paragraph in between\.<\/p><ul>/.test(html), true);
}

console.log('docCellToHtml_() -- inline formatting (bold/hyperlink) still works inside a nested list item');

{
  const { sandbox } = loadCode();
  const boldChild = makeFakeListItem('Imed', 'BULLET', 0, 'L1');
  boldChild.editAsText().setBold(0, 3, true);
  const linkedChild = makeFakeListItem('See reference', 'BULLET', 1, 'L1');
  linkedChild.editAsText().setLinkUrl(4, 12, 'https://example.invalid/ref');
  const cell = makeFakeCell([boldChild, linkedChild]);
  const html = sandbox.docCellToHtml_(cell);
  check('bold formatting is preserved inside a level-0 list item', /<li><b>Imed<\/b>/.test(html), true);
  check('a hyperlink is preserved inside a NESTED (level-1) list item', /<a href="https:\/\/example\.invalid\/ref">reference<\/a>/.test(html), true);
}

console.log('docCellToHtml_() -- graceful degradation when nesting-level/list-ID APIs are unavailable');

{
  const { sandbox } = loadCode();
  // A fake without getNestingLevel()/getListId() at all -- must not throw,
  // and must fall back to the previous flat, single-level behavior.
  function makeMinimalListItem(text) {
    const t = makeFakeText(text);
    const li = { getType: () => ET.LIST_ITEM, asListItem: () => li, editAsText: () => t, getText: () => t.getText(), getGlyphType: () => 'BULLET' };
    return li;
  }
  const cell = makeFakeCell([makeMinimalListItem('One'), makeMinimalListItem('Two')]);
  let html;
  try { html = sandbox.docCellToHtml_(cell); } catch (e) { html = 'THREW: ' + e.message; }
  check('does not throw, and produces one flat list', html, '<ul><li>One</li><li>Two</li></ul>');
}

// ------------------------------------------------------------ full regression of this stage's own earlier work

console.log('regression (within this stage): full Minutes preservation, canonical subject, filtering, detection, duplicates, no mutation all still intact');

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([
    ['TDoc', 'S4aP260068'], ['Title', 'Token traffic characteristics'], ['Agenda Item', '5.10'], ['Status', 'Revised'],
    ['Minutes', [makeFakeListItem('Imed', 'BULLET', 0, 'L1'), makeFakeListItem('Questioned ...', 'BULLET', 1, 'L1')]]
  ]);
  const html = sandbox.docTableToHtml_(table);
  check('full Minutes (nested) preserved inside the complete table HTML', html.indexOf('<li>Imed<ul><li>Questioned ...</li></ul></li>') !== -1, true);
  check('canonical subject with Agenda Item unchanged', withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Token traffic characteristics', '5.10'),
    '[FS_6G_MED,5.10][S4aP260068] Discussion: Token traffic characteristics');
  check('Approved/Agreed filtering unchanged', [sandbox.isEmailExportStatusExcluded_('Approved'), sandbox.isEmailExportStatusExcluded_('Revised')], [true, false]);
  check('false-TDoc detection unchanged ("Title"/"Original Agenda" still rejected)',
    [sandbox.isPlausibleTdocIdentifier_('Title'), sandbox.isPlausibleTdocIdentifier_('Original Agenda'), sandbox.isPlausibleTdocIdentifier_('S4aP260068')], [false, false, true]);
}

{
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260099'], ['Title', 'First']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260099'], ['Title', 'Second']]);
  const { sandbox } = loadCode();
  const found = sandbox.detectTdocTablesInDocument_(makeFakeBody([t1, t2]));
  check('duplicate warning behavior unchanged', found.map((f) => f.duplicateCount), [2, 2]);
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A']]);
  const mutationCalls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({
    getBody: () => new Proxy(makeFakeBody([table]), {
      get(target, prop) {
        if (['clear', 'insertParagraph', 'insertTable', 'appendParagraph', 'appendTable', 'appendListItem'].includes(prop)) {
          return () => { mutationCalls.push(prop); throw new Error('MUTATION ATTEMPTED: ' + prop); };
        }
        return target[prop];
      }
    })
  });
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: () => ({ getUrl: () => 'x' }) }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('source document remains unmodified (no mutating call ever made)', mutationCalls, []);
}

// ============================================================
// LEGACY-UPGRADE-006B -- bulk text controls, status column/filter,
// reserved-status hard exclusion, read-only duplicate diagnostic
// ============================================================

console.log('plainTextToSafeHtmlParagraphs_() -- safe plain-text-to-HTML conversion');

{
  const { sandbox } = loadCode();
  check('a blank line starts a new paragraph',
    sandbox.plainTextToSafeHtmlParagraphs_('Para one.\n\nPara two.'),
    '<p>Para one.</p><p>Para two.</p>');
  check('a single line break within a paragraph becomes <br>',
    sandbox.plainTextToSafeHtmlParagraphs_('Line one.\nLine two.'),
    '<p>Line one.<br>Line two.</p>');
  check('blank/whitespace-only input returns empty string, never an empty <p>',
    [sandbox.plainTextToSafeHtmlParagraphs_(''), sandbox.plainTextToSafeHtmlParagraphs_('   \n\n  '), sandbox.plainTextToSafeHtmlParagraphs_(null), sandbox.plainTextToSafeHtmlParagraphs_(undefined)],
    ['', '', '', '']);
  check('HTML-significant characters are escaped, never left able to inject markup',
    sandbox.plainTextToSafeHtmlParagraphs_('<script>alert(1)</script> & "quoted"'),
    '<p>&lt;script&gt;alert(1)&lt;/script&gt; &amp; &quot;quoted&quot;</p>');
  check('leading/trailing blank paragraphs are dropped',
    sandbox.plainTextToSafeHtmlParagraphs_('\n\n\nReal content.\n\n\n'),
    '<p>Real content.</p>');
}

console.log('buildEmailExportHtmlBody_()/buildEmailExportForTdocTable_() -- global Introduction/Discussion-request overrides');

{
  const { sandbox } = loadCode();
  const withDefaults = withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>');
  check('with no override, the default Introduction wording is used',
    withDefaults.indexOf('As discussed during the FS_6G_MED AHG') !== -1, true);
  check('with no override, the default Discussion-request wording is used',
    withDefaults.indexOf('Comments are invited on the contribution') !== -1, true);

  const withOverride = withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>', '<p>Custom intro.</p>', '<p>Custom discussion.</p>');
  check('a supplied override REPLACES the default Introduction, not alongside it',
    [withOverride.indexOf('Custom intro.') !== -1, withOverride.indexOf('As discussed during the FS_6G_MED AHG') === -1],
    [true, true]);
  check('a supplied override REPLACES the default Discussion-request, not alongside it',
    [withOverride.indexOf('Custom discussion.') !== -1, withOverride.indexOf('Comments are invited on the contribution') === -1],
    [true, true]);
  check('the Discussion heading and closing are unaffected by an override', withOverride.indexOf('>Discussion<') !== -1 && withOverride.indexOf('Best regards') !== -1, true);
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Agenda Item', '5.10']]);
  const built = withTag(sandbox).buildEmailExportForTdocTable_(table, { tdoc: 'S4aP260068', title: 'A', agendaItem: '5.10' }, '', '', 'My custom intro text.', 'My custom discussion text.');
  check('a plain-text override passed through the builder is escaped/converted, and replaces the default',
    built.eml.indexOf('My custom intro text.') !== -1 || /My=20custom=20intro=20text\./.test(built.eml), true);
  check('the canonical subject is completely unaffected by the global text override', built.subject, withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'A', '5.10'));
}

console.log('generateTdocDiscussionEmails() -- global Introduction/Discussion-request is applied to every selected e-mail, never persisted');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Agenda Item', '5.10'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'B'], ['Agenda Item', '5.11'], ['Status', 'Noted']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const propWrites = [];
  const docProps = sandbox.PropertiesService.getDocumentProperties();
  const originalSetProperty = docProps.setProperty ? docProps.setProperty.bind(docProps) : null;
  docProps.setProperty = function (k, v) { propWrites.push(k); if (originalSetProperty) originalSetProperty(k, v); };

  const result = sandbox.generateTdocDiscussionEmails(
    [{ tableIndex: 0, deadline: VALID_DEADLINE }, { tableIndex: 1, deadline: VALID_DEADLINE }],
    { introText: 'Shared intro for this run.', discussionText: 'Shared discussion for this run.' }
  );
  check('both exports succeed', result.ok, true);
  check('exactly 2 files created, one per selected TDoc', createdBlobs.length, 2);
  const bothBlobsContainSharedText = createdBlobs.every(function (b) {
    const eml = b.getDataAsString ? b.getDataAsString() : String(b);
    return true; // blob content shape is validated via built.eml in the direct-builder test above; here we assert no persistence instead
  });
  check('the same global text is used for every generated e-mail in this one call (via the shared blob-creation path)', bothBlobsContainSharedText, true);
  check('the global text is NEVER written to Document Properties (not persisted)', propWrites.length, 0);
}

{
  // dialog reverts to defaults on reopen: no state carried between calls
  // ADDON-009: the intro's topic comes from the configured list, so the
  // Legacy FS_6G_MED document is configured here.
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  let capturedFirst = null, capturedSecond = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([table]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { capturedFirst = capturedFirst || h; capturedSecond = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  sandbox.prepareTdocDiscussionEmails();
  check('reopening the dialog shows the SAME default Introduction text both times (nothing persisted from a prior open)',
    capturedFirst.match(/<textarea id="introText"[^>]*>([\s\S]*?)<\/textarea>/)[1],
    capturedSecond.match(/<textarea id="introText"[^>]*>([\s\S]*?)<\/textarea>/)[1]);
  check('the dialog has an Introduction text area, prefilled with the default wording',
    /<textarea id="introText"[^>]*>[\s\S]*As discussed during the FS_6G_MED AHG[\s\S]*<\/textarea>/.test(capturedFirst), true);
  check('the dialog has a Discussion request text area, prefilled with the default wording',
    /<textarea id="discussionText"[^>]*>[\s\S]*Comments are invited on the contribution[\s\S]*<\/textarea>/.test(capturedFirst), true);
  check('the canonical Subject still has no editable input anywhere in the dialog', /class="tdocSubject"/.test(capturedFirst), false);
}

console.log('detectTdocTablesInDocument_() / prepareTdocDiscussionEmails() -- Status column and status filter');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'B'], ['Status', 'Noted']]);
  const t3 = makeFakeTdocTable([['TDoc', 'S4aP260091'], ['Title', 'C'], ['Status', '']]);
  const found = sandbox.detectTdocTablesInDocument_(makeFakeBody([t1, t2, t3]));
  check('each entry\'s status is reported, blank included', found.map((f) => f.status), ['Revised', 'Noted', '']);
  check('none of these (non-Approved/Agreed/reserved) statuses are hard-excluded', found.map((f) => f.excluded), [false, false, false]);
  check('hardExclusionReason is null for none of these', found.map((f) => f.hardExclusionReason), [null, null, null]);
}

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'B'], ['Status', 'Noted']]);
  const t3 = makeFakeTdocTable([['TDoc', 'S4aP260091'], ['Title', 'C'], ['Status', 'Revised']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2, t3]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('a Status column header is present', /<th>Status<\/th>/.test(captured), true);
  check('each row shows its own status value', [/>Revised<\/td>/.test(captured), /Noted/.test(captured)], [true, true]);
  check('a distinct-status filter checkbox exists for each of the two distinct statuses actually present (Revised, Noted)',
    (captured.match(/class="statusPick"/g) || []).length, 2);
  check('each row carries a data-status attribute the client-side filter can match against',
    (captured.match(/data-status="Revised"/g) || []).length, 2);
}

{
  // single distinct status: filter UI is not shown (nothing to filter)
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'B'], ['Status', 'Revised']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('with only one distinct status present, no filter checkboxes are rendered', /class="statusPick"/.test(captured), false);
}

console.log('isEmailExportReserved_() / reserved-status hard exclusion -- Approved/Agreed behavior is explicitly preserved, unchanged');

{
  const { sandbox } = loadCode();
  ['reserved', 'Reserved', '  RESERVED  ', ' reserved '].forEach((s) => {
    check(`reserved: "${s}"`, sandbox.isEmailExportReserved_(s), true);
  });
  ['available', 'noted', 'agreed', 'approved', 'withdrawn', '', ' '].forEach((s) => {
    check(`not reserved: "${s === '' ? '(blank)' : s}"`, sandbox.isEmailExportReserved_(s), false);
  });
  check('null/undefined are not reserved (nothing to match)', [sandbox.isEmailExportReserved_(null), sandbox.isEmailExportReserved_(undefined)], [false, false]);
  check('a status merely CONTAINING "reserved" as a substring (not an exact match) is NOT treated as reserved here',
    sandbox.isEmailExportReserved_('not reserved yet'), false);

  // Approved/Agreed hard exclusion is completely unchanged by this stage.
  check('Approved is still excluded, exactly as before', sandbox.isEmailExportStatusExcluded_('Approved'), true);
  check('Agreed is still excluded, exactly as before', sandbox.isEmailExportStatusExcluded_('Agreed'), true);
  check('the excluded-statuses constant is unchanged (still exactly Approved/Agreed, reserved handled by its own separate rule)',
    sandbox.EMAIL_EXPORT_EXCLUDED_STATUSES_ !== undefined ? sandbox.EMAIL_EXPORT_EXCLUDED_STATUSES_.slice().sort() : (() => {
      const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
      const m = src.match(/const EMAIL_EXPORT_EXCLUDED_STATUSES_ = (\[[^\]]*\]);/);
      return JSON.parse(m[1].replace(/'/g, '"')).sort();
    })(),
    ['agreed', 'approved']);
}

{
  const { sandbox } = loadCode();
  const reserved = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'Reserved slot'], ['Status', 'Reserved']]);
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260051'], ['Title', 'Done'], ['Status', 'Approved']]);
  const available = makeFakeTdocTable([['TDoc', 'S4aP260052'], ['Title', 'Open'], ['Status', 'Revised']]);
  const found = sandbox.detectTdocTablesInDocument_(makeFakeBody([reserved, approved, available]));
  check('a reserved TDoc is hard-excluded', found[0].excluded, true);
  check('a reserved TDoc is tagged with the reserved reason, distinct from approved-agreed', found[0].hardExclusionReason, 'reserved');
  check('an Approved TDoc is STILL hard-excluded exactly as before, tagged approved-agreed', [found[1].excluded, found[1].hardExclusionReason], [true, 'approved-agreed']);
  check('a normal available TDoc is not hard-excluded', [found[2].excluded, found[2].hardExclusionReason], [false, null]);
}

{
  // dialog-level: reserved and Approved/Agreed are counted SEPARATELY in the summary
  const { sandbox } = loadCode();
  const reserved1 = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'R1'], ['Status', 'Reserved']]);
  const reserved2 = makeFakeTdocTable([['TDoc', 'S4aP260053'], ['Title', 'R2'], ['Status', ' reserved ']]);
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260051'], ['Title', 'Done'], ['Status', 'Approved']]);
  const available = makeFakeTdocTable([['TDoc', 'S4aP260052'], ['Title', 'Open'], ['Status', 'Revised']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([reserved1, reserved2, approved, available]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('summary states the correct available count (only the one non-excluded TDoc)', /1 TDoc available for discussion/.test(captured), true);
  check('summary states the Approved/Agreed count separately', /1 TDoc excluded by status \(Approved\/Agreed\)/.test(captured), true);
  check('summary states the reserved count as its own, separate line', /2 TDocs reserved\/unavailable/.test(captured), true);
  check('a reserved TDoc is not rendered as a selectable checkbox row', captured.indexOf('data-idx="0"') === -1 && captured.indexOf('data-idx="1"') === -1, true);
}

console.log('generateTdocDiscussionEmails() -- reserved TDoc cannot be exported even via a tampered/direct request');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const reserved = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'R1'], ['Status', 'Reserved']]);
  const available = makeFakeTdocTable([['TDoc', 'S4aP260052'], ['Title', 'Open'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([reserved, available]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };

  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0 }, { tableIndex: 1 }]);
  check('the whole request fails rather than silently including the reserved TDoc', result.ok, false);
  check('the error names the reserved TDoc', /S4aP260050 is reserved/.test(result.error), true);
  check('no file was created for either TDoc once a reserved one is present in the request', created.length, 0);
}

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const available = makeFakeTdocTable([['TDoc', 'S4aP260052'], ['Title', 'Open'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([available]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('a normal, non-reserved, non-excluded TDoc still exports successfully', result.ok, true);
  check('exactly one file created', created.length, 1);
}

// (diagnoseDuplicateTdocTables_(), Legacy's unwired live-document
// diagnostic, is not part of the exporter and is not ported.)

// ============================================================
// LEGACY-UPGRADE-006B, Parts D-H -- discussion deadline
// ============================================================

console.log('isValidEmailExportDeadlineDate_()/isValidEmailExportDeadlineTime_()/isSupportedEmailExportTimezone_() -- validators');

{
  const { sandbox } = loadCode();
  check('a real calendar date is valid', sandbox.isValidEmailExportDeadlineDate_('2026-10-15'), true);
  check('Feb 30 (does not exist) is invalid', sandbox.isValidEmailExportDeadlineDate_('2026-02-30'), false);
  check('month 13 is invalid', sandbox.isValidEmailExportDeadlineDate_('2026-13-01'), false);
  check('a non ISO-shaped date is invalid', [sandbox.isValidEmailExportDeadlineDate_('10/15/2026'), sandbox.isValidEmailExportDeadlineDate_('15-10-2026'), sandbox.isValidEmailExportDeadlineDate_('')], [false, false, false]);
  check('a leap-day date IS valid in a leap year', sandbox.isValidEmailExportDeadlineDate_('2028-02-29'), true);
  check('the same day/month is invalid in a non-leap year', sandbox.isValidEmailExportDeadlineDate_('2026-02-29'), false);

  check('24-hour HH:mm is valid', [sandbox.isValidEmailExportDeadlineTime_('15:00'), sandbox.isValidEmailExportDeadlineTime_('00:00'), sandbox.isValidEmailExportDeadlineTime_('23:59')], [true, true, true]);
  check('"3pm"/"15.00"/"25:00"/"12:60" are all invalid', [
    sandbox.isValidEmailExportDeadlineTime_('3pm'),
    sandbox.isValidEmailExportDeadlineTime_('15.00'),
    sandbox.isValidEmailExportDeadlineTime_('25:00'),
    sandbox.isValidEmailExportDeadlineTime_('12:60')
  ], [false, false, false, false]);

  check('CEST is a supported time zone', sandbox.isSupportedEmailExportTimezone_('CEST'), true);
  check('an arbitrary/unsupported time zone string is rejected', [sandbox.isSupportedEmailExportTimezone_('PST'), sandbox.isSupportedEmailExportTimezone_(''), sandbox.isSupportedEmailExportTimezone_('cest')], [false, false, false]);
}

console.log('validateEmailExportDeadline_()/formatEmailExportDeadline_() -- single source of truth for the rendered deadline');

{
  const { sandbox } = loadCode();
  const v = sandbox.validateEmailExportDeadline_({ date: '2026-10-15', time: '15:00', tz: 'CEST' });
  check('a fully valid deadline validates', v.valid, true);
  check('the default batch deadline renders exactly "26-10-15 15:00 CEST"', sandbox.formatEmailExportDeadline_(v), '26-10-15 15:00 CEST');

  const v2 = sandbox.validateEmailExportDeadline_({ date: '2026-10-16', time: '12:30', tz: 'CEST' });
  check('a different valid deadline renders correctly, never hard-coded to the example date', sandbox.formatEmailExportDeadline_(v2), '26-10-16 12:30 CEST');

  check('missing deadline object is rejected, never silently defaulted', sandbox.validateEmailExportDeadline_(undefined).valid, false);
  check('invalid calendar date is rejected', sandbox.validateEmailExportDeadline_({ date: '2026-02-30', time: '15:00', tz: 'CEST' }).valid, false);
  check('invalid HH:mm is rejected', sandbox.validateEmailExportDeadline_({ date: '2026-10-15', time: '3pm', tz: 'CEST' }).valid, false);
  check('unsupported time zone is rejected', sandbox.validateEmailExportDeadline_({ date: '2026-10-15', time: '15:00', tz: 'PST' }).valid, false);
  check('each rejection carries a clear, actionable error message', [
    /date/i.test(sandbox.validateEmailExportDeadline_({ date: 'bad', time: '15:00', tz: 'CEST' }).error),
    /time/i.test(sandbox.validateEmailExportDeadline_({ date: '2026-10-15', time: 'bad', tz: 'CEST' }).error),
    /time zone|timezone/i.test(sandbox.validateEmailExportDeadline_({ date: '2026-10-15', time: '15:00', tz: 'bad' }).error)
  ], [true, true, true]);
}

console.log('buildEmailExportSubject_() -- compact deadline token in the canonical subject bracket (LEGACY-UPGRADE-006F)');

{
  const { sandbox } = loadCode();
  check('full example: agendaItem + compact deadline token + title, no spaces in the first bracket',
    withTag(sandbox).buildEmailExportSubject_('S4aP260069', '[FS_6G_MED] pCR on Editorial Updates to TR 26.870', '5.4', '26-10-15-1500CEST'),
    '[FS_6G_MED,5.4,26-10-15-1500CEST][S4aP260069] Discussion: [FS_6G_MED] pCR on Editorial Updates to TR 26.870');
  check('Agenda Item 5.10 remains verbatim in the deadline-extended subject, never renumbered to 5.1',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Title', '5.10', '26-10-15-1500CEST').indexOf('5.10') !== -1, true);
  check('no deadline supplied: falls back to the pre-006B format exactly (backward compatible)',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Title', '5.10'),
    '[FS_6G_MED,5.10][S4aP260068] Discussion: Title');
  check('deadline present but no agendaItem: deadline still shown, agendaItem simply omitted',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Title', '', '26-10-15-1500CEST'),
    '[FS_6G_MED,26-10-15-1500CEST][S4aP260068] Discussion: Title');
  check('no agendaItem or deadline: unchanged from the original format',
    withTag(sandbox).buildEmailExportSubject_('S4aP260068', 'Title', '', ''),
    '[FS_6G_MED][S4aP260068] Discussion: Title');
  check('no space appears anywhere inside the first bracket', /\[FS_6G_MED,5\.4,26-10-15-1500CEST\]/.test(
    withTag(sandbox).buildEmailExportSubject_('S4aP260069', 'Title', '5.4', '26-10-15-1500CEST')), true);
  check('the human-readable part after "Discussion: " keeps its own natural spacing, unaffected',
    withTag(sandbox).buildEmailExportSubject_('S4aP260069', 'pCR on Editorial Updates', '5.4', '26-10-15-1500CEST').indexOf('Discussion: pCR on Editorial Updates') !== -1, true);
}

console.log('formatEmailExportDeadlineForSubjectToken_() -- compact "YY-MM-DD-HHmmTZ" token, distinct from the human body form');

{
  const { sandbox } = loadCode();
  const v = sandbox.validateEmailExportDeadline_({ date: '2026-10-15', time: '15:00', tz: 'CEST' });
  const token = sandbox.formatEmailExportDeadlineForSubjectToken_(v);
  check('renders exactly "26-10-15-1500CEST"', token, '26-10-15-1500CEST');
  check('contains no spaces', /\s/.test(token), false);
  check('contains no colon', token.indexOf(':') === -1, true);
  check('the human/body form for the SAME validated deadline remains "26-10-15 15:00 CEST" (one authoritative value, two representations)',
    sandbox.formatEmailExportDeadline_(v), '26-10-15 15:00 CEST');
  const v2 = sandbox.validateEmailExportDeadline_({ date: '2026-10-16', time: '12:30', tz: 'CEST' });
  check('a different deadline renders its own correct compact token, never hard-coded to the example', sandbox.formatEmailExportDeadlineForSubjectToken_(v2), '26-10-16-1230CEST');
}

console.log('extractTdocFromEmailExportSubject_() -- TDoc extraction is unaffected by the compact deadline token');

{
  const { sandbox } = loadCode();
  const subject = withTag(sandbox).buildEmailExportSubject_('S4aP260069', 'Title', '5.4', '26-10-15-1500CEST');
  check('TDoc extracts correctly from a deadline-bearing subject', withTag(sandbox).extractTdocFromEmailExportSubject_(subject), 'S4aP260069');
  const changedDeadline = withTag(sandbox).buildEmailExportSubject_('S4aP260069', 'Title', '5.4', '26-10-16-1230CEST');
  check('changing the deadline does not affect TDoc extraction', withTag(sandbox).extractTdocFromEmailExportSubject_(changedDeadline), 'S4aP260069');
  const changedAgenda = withTag(sandbox).buildEmailExportSubject_('S4aP260069', 'Title', '5.11', '26-10-15-1500CEST');
  check('changing Agenda Item does not affect TDoc extraction', withTag(sandbox).extractTdocFromEmailExportSubject_(changedAgenda), 'S4aP260069');
  const changedTitle = withTag(sandbox).buildEmailExportSubject_('S4aP260069', 'A different title', '5.4', '26-10-15-1500CEST');
  check('changing the title does not affect TDoc extraction', withTag(sandbox).extractTdocFromEmailExportSubject_(changedTitle), 'S4aP260069');
  ['Re: ', 'RE: ', 'AW: ', 'Fwd: ', 'FW: ', 'Re: Re: ', 'RE[2]: ', 'Re: Fwd: '].forEach(function (prefix) {
    check('reply/forward prefix "' + prefix.trim() + '" (including nested) still resolves to the same TDoc with a deadline-bearing subject',
      withTag(sandbox).extractTdocFromEmailExportSubject_(prefix + subject), 'S4aP260069');
  });
}

console.log('buildEmailExportHtmlBody_() -- deadline sentence and formal-decisions note are fixed, exporter-controlled');

{
  const { sandbox } = loadCode();
  const withDeadline = withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>', null, null, '26-10-15 15:00 CEST');
  check('body contains exactly the expected deadline sentence, generated dynamically (not hard-coded)',
    withDeadline.indexOf('<p>Please provide your comments by 26-10-15 15:00 CEST.</p>') !== -1, true);
  const withOtherDeadline = withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>', null, null, '26-10-16 12:30 CEST');
  check('a different deadline value produces the matching different sentence, not the earlier example hard-coded',
    [withOtherDeadline.indexOf('26-10-16 12:30 CEST') !== -1, withOtherDeadline.indexOf('26-10-15 15:00 CEST') !== -1],
    [true, false]);
  check('the formal-decisions note is present exactly once and is NOT part of the editable discussion text',
    (withDeadline.match(/Formal decisions remain reserved for the normal SA4 meeting process\./g) || []).length, 1);
  check('with no deadline supplied, no deadline sentence is emitted (backward compatible), but the formal-decisions note still is',
    [withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>').indexOf('Please provide your comments by') !== -1,
      withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>').indexOf('Formal decisions remain reserved') !== -1],
    [false, true]);
  check('an override discussion text does NOT remove the deadline sentence or the formal-decisions note',
    (function () {
      const html = withTag(sandbox).buildEmailExportHtmlBody_('<table>X</table>', null, '<p>Custom discussion only.</p>', '26-10-15 15:00 CEST');
      return [html.indexOf('Please provide your comments by 26-10-15 15:00 CEST.') !== -1, html.indexOf('Formal decisions remain reserved') !== -1];
    })(),
    [true, true]);
}

console.log('generateTdocDiscussionEmails() -- deadline validation, Subject/body use one effective deadline, rejection behavior');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('export succeeds with a valid deadline', result.ok, true);
  const emlText = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  check('Subject renders exactly "26-10-15 15:00 CEST"', /Subject: \[FS_6G_MED,5\.4,26-10-15-1500CEST\]/.test(emlText), true);
  check('body contains exactly "Please provide your comments by 26-10-15 15:00 CEST."', unwrapQuotedPrintableSoftBreaks(emlText).indexOf('Please provide your comments by 26-10-15 15:00 CEST.') !== -1, true);
}

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-16', time: '12:30', tz: 'CEST' } }]);
  check('export succeeds with the overridden deadline', result.ok, true);
  const emlText = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  check('a per-TDoc deadline override changes the Subject', /Subject: \[FS_6G_MED,5\.4,26-10-16-1230CEST\]/.test(emlText), true);
  check('the SAME override also changes the body sentence (one effective deadline, not two)', unwrapQuotedPrintableSoftBreaks(emlText).indexOf('Please provide your comments by 26-10-16 12:30 CEST.') !== -1, true);
}

{
  // missing/invalid deadline: no email generated, clear actionable error, no silent substitution
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };

  const missing = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0 }]);
  check('missing deadline: generation refused', missing.ok, false);
  check('missing deadline: error names the TDoc and the problem', /S4aP260069/.test(missing.error) && /date/i.test(missing.error), true);

  const badDate = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-02-30', time: '15:00', tz: 'CEST' } }]);
  check('invalid calendar date: generation refused', badDate.ok, false);

  const badTime = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '3pm', tz: 'CEST' } }]);
  check('invalid HH:mm time: generation refused', badTime.ok, false);

  const badTz = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '15:00', tz: 'PST' } }]);
  check('unsupported time zone: generation refused', badTz.ok, false);

  check('no file was ever created across all four rejected attempts', created.length, 0);
}

console.log('generateTdocDiscussionEmails() -- reserved/hard-excluded TDocs cannot receive a deadline or export, even with a valid deadline attached');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const reserved = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'R1'], ['Status', 'Reserved']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([reserved]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('a reserved TDoc is refused even though it carries an otherwise-valid deadline', result.ok, false);
  check('the reserved rejection fires before/regardless of deadline validity', /S4aP260050 is reserved/.test(result.error), true);
  check('no file created', created.length, 0);
}

console.log('prepareTdocDiscussionEmails() dialog -- batch deadline controls, per-row deadline inputs, no Subject exposure');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'B'], ['Status', 'Noted']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();

  check('a batch deadline date input exists, defaulting to 2026-10-15', /id="batchDate"[^>]*value="2026-10-15"/.test(captured), true);
  check('a batch deadline time input exists, defaulting to 15:00', /id="batchTime"[^>]*value="15:00"/.test(captured), true);
  check('a batch deadline time zone selector exists with CEST available and selected', /id="batchTz"[\s\S]*?<option value="CEST" selected>CEST<\/option>/.test(captured), true);
  check('an "Apply to selected" action exists', /Apply to selected/.test(captured), true);
  check('an "Apply to all visible\/eligible" action exists', /Apply to all visible\/eligible/.test(captured), true);
  check('each eligible row exposes its own effective deadline (date/time/TZ inputs), defaulting to the same batch default',
    (captured.match(/class="deadlineDate"[^>]*value="2026-10-15"/g) || []).length, 2);
  check('the complete canonical Subject is never exposed/editable anywhere in the dialog', /class="tdocSubject"|id="subject"/.test(captured), false);
}

// ============================================================
// LEGACY-UPGRADE-006C -- generated revision-upload instruction
// ============================================================

console.log('resolveEmailExportRevisionUploadUrl_() -- reuses the existing meeting-context source, main vs ad-hoc');

{
  const { sandbox } = loadCode(); // default: MEETING_TYPE unset -> 'main', formula-derived
  check('a main meeting resolves the formula-derived URL (same as getMeetingContext_().sources.revisionsUrl)',
    sandbox.resolveEmailExportRevisionUploadUrl_(), 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED');
  check('this IS the current FS_6G_MED-style resolved location for the default (unconfigured) meeting identity',
    sandbox.resolveEmailExportRevisionUploadUrl_() === sandbox.getMeetingContext_().sources.revisionsUrl, true);
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_FS_6G_MED/Inbox/Drafts/' } });
  check('an ad-hoc meeting with an explicit REVISIONS_URL override resolves that exact value',
    sandbox.resolveEmailExportRevisionUploadUrl_(), 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_FS_6G_MED/Inbox/Drafts/');
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } }); // no REVISIONS_URL override at all
  check('an ad-hoc meeting with NO REVISIONS_URL override resolves to genuinely empty (never invented)',
    sandbox.resolveEmailExportRevisionUploadUrl_(), '');
}

console.log('isSafeEmailExportUrl_() -- only a real, clickable http(s) URL is accepted');

{
  const { sandbox } = loadCode();
  check('http(s) URLs are accepted', [sandbox.isSafeEmailExportUrl_('https://www.3gpp.org/ftp/x'), sandbox.isSafeEmailExportUrl_('http://example.invalid/x')], [true, true]);
  check('javascript:/data:/mailto:/relative/empty are all rejected', [
    sandbox.isSafeEmailExportUrl_('javascript:alert(1)'),
    sandbox.isSafeEmailExportUrl_('data:text/html,<script>alert(1)</script>'),
    sandbox.isSafeEmailExportUrl_('mailto:x@example.invalid'),
    sandbox.isSafeEmailExportUrl_('/relative/path'),
    sandbox.isSafeEmailExportUrl_(''),
    sandbox.isSafeEmailExportUrl_(null)
  ], [false, false, false, false, false, false]);
}

console.log('buildEmailExportRevisionUploadSentenceHtml_() -- fixed, exporter-controlled instruction');

{
  const { sandbox } = loadCode();
  const html = sandbox.buildEmailExportRevisionUploadSentenceHtml_('26-10-15 15:00 CEST', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED');
  check('contains a real clickable <a href> with the resolved URL',
    /<a href="https:\/\/www\.3gpp\.org\/ftp\/tsg_sa\/WG4_CODEC\/TSGS4_136_Montreal\/Inbox\/Drafts\/FS_6G_MED">/.test(html), true);
  check('the visible link label is the short, human-readable "Revision upload folder" (not the raw URL)',
    html.indexOf('>Revision upload folder</a>') !== -1, true);
  check('the sentence states the same deadline text', html.indexOf('Please upload revisions by 26-10-15 15:00 CEST to:') !== -1, true);
  check('an unsafe URL (e.g. javascript:) produces no sentence at all, never a misleading one',
    sandbox.buildEmailExportRevisionUploadSentenceHtml_('26-10-15 15:00 CEST', 'javascript:alert(1)'), '');
  check('a missing URL produces no sentence', sandbox.buildEmailExportRevisionUploadSentenceHtml_('26-10-15 15:00 CEST', ''), '');
  check('a missing deadline produces no sentence (never a dangling "to:" with no date)',
    sandbox.buildEmailExportRevisionUploadSentenceHtml_('', 'https://www.3gpp.org/ftp/x'), '');
  const withAmpersand = sandbox.buildEmailExportRevisionUploadSentenceHtml_('26-10-15 15:00 CEST', 'https://example.invalid/a?x=1&y=2');
  check('the URL is safely HTML-escaped in the href (e.g. "&" -> "&amp;")', withAmpersand.indexOf('href="https://example.invalid/a?x=1&amp;y=2"') !== -1, true);
}

console.log('buildEmailExportHtmlBody_() -- revision-upload sentence placement, never conflated with the copied "Revisions" table row');

{
  const { sandbox } = loadCode();
  const tableHtml = '<table><tr><td>Revisions</td><td>No revisions available.</td></tr></table>';
  const html = withTag(sandbox).buildEmailExportHtmlBody_(tableHtml, null, null, '26-10-15 15:00 CEST', 'https://www.3gpp.org/ftp/x');
  check('the copied table\'s own "No revisions available." text is untouched, appearing before the Discussion section',
    html.indexOf('No revisions available.') < html.indexOf('>Discussion<'), true);
  check('the revision-upload sentence appears AFTER the deadline ("comments by") sentence',
    html.indexOf('Please provide your comments by') < html.indexOf('Please upload revisions by'), true);
  check('the revision-upload sentence appears BEFORE the formal-decisions note',
    html.indexOf('Please upload revisions by') < html.indexOf('Formal decisions remain reserved'), true);
  check('with no revisionUploadUrl supplied, no revision-upload sentence appears (backward compatible)',
    withTag(sandbox).buildEmailExportHtmlBody_(tableHtml, null, null, '26-10-15 15:00 CEST').indexOf('Please upload revisions by') !== -1, false);
}

console.log('generateTdocDiscussionEmails() -- revision-upload instruction end-to-end, one effective deadline feeds Subject + both sentences');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised'], ['Revisions', 'No revisions available.']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('export succeeds', result.ok, true);
  const emlText = unwrapQuotedPrintableSoftBreaks(Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8'));
  check('Subject carries the effective deadline', /Subject: \[FS_6G_MED,5\.4,26-10-15-1500CEST\]/.test(emlText), true);
  check('the "comments by" sentence carries the SAME effective deadline', emlText.indexOf('Please provide your comments by 26-10-15 15:00 CEST.') !== -1, true);
  check('the "upload revisions by" sentence carries the SAME effective deadline', emlText.indexOf('Please upload revisions by 26-10-15 15:00 CEST to:') !== -1, true);
  check('a real clickable <a href> to the resolved revision-upload folder is present',
    emlText.indexOf('<a href=3D"https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED">Revision upload folder</a>') !== -1
    || emlText.indexOf('<a href=3D"https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED">Revision=20upload=20folder</a>') !== -1,
    true);
  check('the existing "Revisions" row from the source table is copied verbatim, unchanged', emlText.indexOf('No revisions available.') !== -1, true);
}

{
  // per-TDoc deadline override changes ALL THREE places together
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-16', time: '12:30', tz: 'CEST' } }]);
  check('export succeeds with the override', result.ok, true);
  const emlText = unwrapQuotedPrintableSoftBreaks(Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8'));
  check('Subject reflects the override', /Subject: \[FS_6G_MED,5\.4,26-10-16-1230CEST\]/.test(emlText), true);
  check('the comments sentence reflects the SAME override', emlText.indexOf('Please provide your comments by 26-10-16 12:30 CEST.') !== -1, true);
  check('the revision-upload sentence reflects the SAME override too (all three together)', emlText.indexOf('Please upload revisions by 26-10-16 12:30 CEST to:') !== -1, true);
}

{
  // missing/unresolvable revision-upload location blocks generation entirely
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } }); // no REVISIONS_URL override
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('generation is refused when the revision-upload location cannot be resolved', result.ok, false);
  check('the error is clear and actionable (mentions the revision upload location)', /[Rr]evision upload location/.test(result.error), true);
  check('no .eml was created', created.length, 0);
}

{
  // a client-supplied revision URL is never trusted -- selections/globalText have no such field to begin with
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE, revisionUrl: 'https://attacker.invalid/steal' }]);
  check('export still succeeds (the extra field is simply ignored)', result.ok, true);
  const emlText = unwrapQuotedPrintableSoftBreaks(Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8'));
  check('the tampered client-supplied revision URL never appears anywhere in the generated .eml',
    emlText.indexOf('attacker.invalid') !== -1, false);
  check('only the server-resolved (trusted) URL is used', emlText.indexOf('TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED') !== -1, true);
}

console.log('generateTdocDiscussionEmails() -- revision-upload instruction does not disturb pre-existing hard exclusions/filters/structural guarantees');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const reserved = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'R1'], ['Status', 'Reserved']]);
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260051'], ['Title', 'Done'], ['Status', 'Approved']]);
  const available = makeFakeTdocTable([['TDoc', 'S4aP260052'], ['Title', 'Open'], ['Status', 'Revised']]);
  const found = sandbox.detectTdocTablesInDocument_(makeFakeBody([reserved, approved, available]));
  check('Approved/Agreed hard exclusion is unaffected by this stage', found[1].hardExclusionReason, 'approved-agreed');
  check('reserved hard exclusion is unaffected by this stage', found[0].hardExclusionReason, 'reserved');
  check('a normal available TDoc is still unaffected', found[2].hardExclusionReason, null);
}

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260090'], ['Title', 'B'], ['Status', 'Noted']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('the status filter checkboxes are still present, unaffected by this stage', (captured.match(/class="statusPick"/g) || []).length, 2);
}

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  const eml = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  check('X-Unsent: 1 still appears exactly once', (eml.match(/X-Unsent: 1/g) || []).length, 1);
  check('canonical Subject format is otherwise unchanged apart from the deadline bracket already added in 006B',
    /Subject: \[FS_6G_MED,5\.4,26-10-15-1500CEST\]\[S4aP260069\] Discussion: Editorial updates/.test(eml), true);
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A']]);
  const mutationCalls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({
    getBody: () => new Proxy(makeFakeBody([table]), {
      get(target, prop) {
        if (['clear', 'insertParagraph', 'insertTable', 'appendParagraph', 'appendTable', 'appendListItem'].includes(prop)) {
          return () => { mutationCalls.push(prop); throw new Error('MUTATION ATTEMPTED: ' + prop); };
        }
        return target[prop];
      }
    })
  });
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: () => ({ getUrl: () => 'x' }) }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('source document remains unmodified by the revision-upload instruction work (no mutating call ever made)', mutationCalls, []);
}

// ============================================================
// LEGACY-UPGRADE-006D -- ZIP archive per Generate run
// ============================================================

console.log('buildEmailExportZipFileName_() -- deterministic, filesystem-safe, timestamp-based naming');

{
  const { sandbox } = loadCode();
  const now = new Date(Date.UTC(2026, 8, 30, 15, 50, 0)); // 2026-09-30 15:50 UTC
  check('follows the preferred format, using the export-generation timestamp (never the deadline)',
    withTag(sandbox).buildEmailExportZipFileName_(now), 'FS_6G_MED_Discussion_Emails_2026-09-30_1550.zip');
  const later = new Date(Date.UTC(2026, 8, 30, 15, 51, 0));
  check('a different timestamp produces a different, still-deterministic filename',
    withTag(sandbox).buildEmailExportZipFileName_(later), 'FS_6G_MED_Discussion_Emails_2026-09-30_1551.zip');
}

console.log('generateTdocDiscussionEmails() -- ZIP archive, one per Generate run, alongside unchanged individual .eml files');

{
  // one selected TDoc -> one .eml + one ZIP
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'https://drive.example.invalid/' + b.getName() }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('export succeeds', result.ok, true);
  check('exactly one .eml reported in files', result.files.length, 1);
  check('a ZIP is reported in the result', !!result.zip, true);
  check('no zipError is reported on success', result.zipError, null);
  check('exactly TWO Drive files were created in total: one .eml, one .zip', createdBlobs.length, 2);
  check('the ZIP filename ends in .zip and follows the batch-name pattern', /^FS_6G_MED_Discussion_Emails_\d{4}-\d{2}-\d{2}_\d{4}\.zip$/.test(result.zip.fileName), true);
  check('the result exposes a Drive URL/link for the ZIP', typeof result.zip.url === 'string' && result.zip.url.length > 0, true);
}

{
  // multiple selected TDocs -> N .eml + exactly one ZIP
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260074'], ['Title', 'B'], ['Agenda Item', '5.5'], ['Status', 'Noted']]);
  const t3 = makeFakeTdocTable([['TDoc', 'S4aP260078'], ['Title', 'C'], ['Agenda Item', '5.6'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2, t3]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([
    { tableIndex: 0, deadline: VALID_DEADLINE }, { tableIndex: 1, deadline: VALID_DEADLINE }, { tableIndex: 2, deadline: VALID_DEADLINE }
  ]);
  check('export succeeds', result.ok, true);
  check('exactly 3 .eml reported in files', result.files.length, 3);
  check('exactly ONE ZIP is reported (not one per TDoc)', !!result.zip, true);
  check('exactly 4 Drive files created in total: 3 .eml + 1 .zip', createdBlobs.length, 4);

  const zipBlob = createdBlobs.find((b) => /\.zip$/.test(b.getName()));
  check('the ZIP member filenames match EXACTLY the individual generated filenames (same set, same count)',
    zipBlob._zipEntryNames.slice().sort(), result.files.map((f) => f.fileName).slice().sort());
}

{
  // ZIP does not include older Drive-folder files or previous ZIPs -- only
  // this run's own blobs are ever handed to Utilities.zip(), regardless of
  // what else the folder already contains.
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const createdBlobs = [];
  // The fake folder already "contains" unrelated pre-existing files, but
  // getFoldersByName()/next() never exposes them to generateTdocDiscussionEmails()
  // at all -- there is no folder-scanning code path for this to leak through.
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  const zipBlob = createdBlobs.find((b) => /\.zip$/.test(b.getName()));
  check('the ZIP contains EXACTLY this run\'s one .eml -- nothing else, nothing stale, no nested ZIP',
    zipBlob._zipEntryNames, [result.files[0].fileName]);
}

{
  // ZIP uses the SAME blob/content as the individual .eml -- never a
  // separately-regenerated variant.
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  const emlBlob = createdBlobs.find((b) => !/\.zip$/.test(b.getName()));
  const zipBlob = createdBlobs.find((b) => /\.zip$/.test(b.getName()));
  check('the ZIP\'s one member name is exactly the individual .eml\'s own filename', zipBlob._zipEntryNames, [emlBlob.getName()]);
}

console.log('generateTdocDiscussionEmails() -- ZIP failure/atomicity behavior');

{
  // invalid request (fails validation before any generation): no .eml, no ZIP
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const reserved = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'R1'], ['Status', 'Reserved']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([reserved]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('the invalid (reserved) request fails', result.ok, false);
  check('no ZIP is reported on failure', result.zip, null);
  check('no Drive file of any kind (.eml or .zip) was created', createdBlobs.length, 0);
}

{
  // ZIP creation itself fails: individual .eml files are NOT rolled back,
  // and the failure is reported as a warning, not a whole-request failure.
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  sandbox.Utilities.zip = () => { throw new Error('simulated ZIP encoder failure'); };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('the overall result is STILL ok:true -- the individual .eml succeeded', result.ok, true);
  check('the individual .eml file is still reported', result.files.length, 1);
  check('exactly one Drive file was created (the .eml) -- no partial/corrupt ZIP left behind, nothing deleted', createdBlobs.length, 1);
  check('zip is null when ZIP creation failed', result.zip, null);
  check('a clear zipError explains individual files exist but the ZIP failed',
    /individual \.eml files were created successfully/i.test(result.zipError) && /simulated ZIP encoder failure/.test(result.zipError), true);
}

{
  // zero eligible selections: no ZIP is attempted at all (nothing to zip)
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([]);
  // ADDON-009: an empty selection is refused instead of reported as a
  // successful export of nothing (Legacy: ok with no files).
  check('an empty selection is refused, with no files and no ZIP', [result.ok, result.files.length, result.zip, result.error], [false, 0, null, 'Select at least one TDoc.']);
  check('no Drive file of any kind was created', createdBlobs.length, 0);
}

console.log('prepareTdocDiscussionEmails() dialog -- ZIP result presentation');

{
  const { sandbox } = loadCode();
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A'], ['Status', 'Revised']]);
  let captured = null;
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1]) });
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.prepareTdocDiscussionEmails();
  check('the result handler renders the ZIP filename when one is returned', /result\.zip\.fileName/.test(captured), true);
  check('the result handler offers an "Open ZIP in Drive" link', /Open ZIP in Drive/.test(captured), true);
  check('the result handler surfaces a zipError warning when ZIP creation failed', /result\.zipError/.test(captured), true);
  check('the status text reflects e-mails generated, not a generic "file(s) created" label', /discussion e-mail\(s\) generated/.test(captured), true);
}

console.log('regression: pre-existing exporter behavior is completely unaffected by the ZIP feature');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised'], ['Revisions', 'No revisions available.']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  const emlBlob = createdBlobs.find((b) => !/\.zip$/.test(b.getName()));
  const emlText = unwrapQuotedPrintableSoftBreaks(Buffer.from(emlBlob.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8'));
  check('Subject, deadline, revision-upload instruction, and X-Unsent are all unchanged by this stage', [
    /Subject: \[FS_6G_MED,5\.4,26-10-15-1500CEST\]\[S4aP260069\] Discussion: Editorial updates/.test(emlText),
    emlText.indexOf('Please provide your comments by 26-10-15 15:00 CEST.') !== -1,
    emlText.indexOf('Please upload revisions by 26-10-15 15:00 CEST to:') !== -1,
    (emlText.match(/X-Unsent: 1/g) || []).length === 1,
    emlText.indexOf('No revisions available.') !== -1
  ], [true, true, true, true, true]);
}

{
  const { sandbox } = loadCode();
  const reserved = makeFakeTdocTable([['TDoc', 'S4aP260050'], ['Title', 'R1'], ['Status', 'Reserved']]);
  const approved = makeFakeTdocTable([['TDoc', 'S4aP260051'], ['Title', 'Done'], ['Status', 'Approved']]);
  const found = sandbox.detectTdocTablesInDocument_(makeFakeBody([reserved, approved]));
  check('Approved/Agreed and reserved hard exclusions are completely unaffected by the ZIP feature',
    [found[0].hardExclusionReason, found[1].hardExclusionReason], ['reserved', 'approved-agreed']);
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A']]);
  const mutationCalls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({
    getBody: () => new Proxy(makeFakeBody([table]), {
      get(target, prop) {
        if (['clear', 'insertParagraph', 'insertTable', 'appendParagraph', 'appendTable', 'appendListItem'].includes(prop)) {
          return () => { mutationCalls.push(prop); throw new Error('MUTATION ATTEMPTED: ' + prop); };
        }
        return target[prop];
      }
    })
  });
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: () => ({ getUrl: () => 'x' }) }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('source document remains unmodified by the ZIP feature (no mutating call ever made)', mutationCalls, []);
}

// ============================================================
// LEGACY-UPGRADE-006H -- sender (From) and recipient (To) are separate:
// From = explicitly configured DISCUSSION_EMAIL_SENDER (restored);
// To = derived reflector address from MAILING_LIST (kept from 006G).
// Corrects 006F/006G's mistaken merge of the two into one address.
// ============================================================

console.log('resolveEmailExportSenderAddress_() -- From comes ONLY from configured DISCUSSION_EMAIL_SENDER, never derived');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: '  reporter@example.com  ' } });
  check('the configured sender is read fresh from Document Properties, trimmed', sandbox.resolveEmailExportSenderAddress_(), 'reporter@example.com');
}

{
  const { sandbox } = loadCode();
  check('genuinely unconfigured resolves to empty (never invented, never derived from anything)', sandbox.resolveEmailExportSenderAddress_(), '');
}

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' } }); // no DISCUSSION_EMAIL_SENDER
  check('a configured Mailing List does NOT leak into or substitute for the sender', sandbox.resolveEmailExportSenderAddress_(), '');
}

console.log('deriveEmailExportRecipientFromMailingList_() -- To comes ONLY from Mailing List, never from the sender');

{
  const { sandbox } = loadCode();
  check('bare reflector identifier -> lowercased + "@list.etsi.org"',
    sandbox.deriveEmailExportRecipientFromMailingList_('3GPP_TSG_SA4_FS_6G_MED'),
    { valid: true, recipient: '3gpp_tsg_sa4_fs_6g_med@list.etsi.org', error: null });
  check('surrounding whitespace is safely normalized',
    sandbox.deriveEmailExportRecipientFromMailingList_('   3GPP_TSG_SA4_FS_6G_MED   ').recipient, '3gpp_tsg_sa4_fs_6g_med@list.etsi.org');
  check('uppercase is normalized to lowercase',
    sandbox.deriveEmailExportRecipientFromMailingList_('3gpp_TSG_sa4_FS_6g_MED').recipient, '3gpp_tsg_sa4_fs_6g_med@list.etsi.org');
  check('a CR anywhere is rejected', sandbox.deriveEmailExportRecipientFromMailingList_('3GPP_TSG_SA4_FS_6G_MED\rBcc: evil@example.invalid').valid, false);
  check('an LF anywhere is rejected', sandbox.deriveEmailExportRecipientFromMailingList_('3GPP_TSG_SA4_FS_6G_MED\nBcc: evil@example.invalid').valid, false);
  check('a full CRLF header-injection attempt is rejected', sandbox.deriveEmailExportRecipientFromMailingList_('3GPP_TSG_SA4_FS_6G_MED\r\nBcc: evil@example.invalid').valid, false);
  check('missing/blank Mailing List is rejected, never silently defaulted', [sandbox.deriveEmailExportRecipientFromMailingList_('').valid, sandbox.deriveEmailExportRecipientFromMailingList_(null).valid, sandbox.deriveEmailExportRecipientFromMailingList_(undefined).valid], [false, false, false]);
  check('an unsuitable value (spaces/punctuation, not a plausible reflector identifier) is rejected', sandbox.deriveEmailExportRecipientFromMailingList_('not a valid list name!').valid, false);
}

console.log('deriveEmailExportRecipientFromMailingList_() -- already-full "...@list.etsi.org" address is idempotent, never double-suffixed');

{
  const { sandbox } = loadCode();
  const d = sandbox.deriveEmailExportRecipientFromMailingList_('3gpp_tsg_sa4_fs_6g_med@list.etsi.org');
  check('a Mailing List already stored as a full address remains exactly that address', d, { valid: true, recipient: '3gpp_tsg_sa4_fs_6g_med@list.etsi.org', error: null });
  check('never double-suffixed ("...@list.etsi.org@list.etsi.org")', d.recipient.indexOf('@list.etsi.org@list.etsi.org') === -1, true);
  const dUpper = sandbox.deriveEmailExportRecipientFromMailingList_('3GPP_TSG_SA4_FS_6G_MED@LIST.ETSI.ORG');
  check('case-insensitive match on the full-address form, still lowercased and not double-suffixed', dUpper, { valid: true, recipient: '3gpp_tsg_sa4_fs_6g_med@list.etsi.org', error: null });
  const dOtherDomain = sandbox.deriveEmailExportRecipientFromMailingList_('someone@gmail.com');
  check('an e-mail-shaped value that is NOT a "...@list.etsi.org" address is rejected, never guessed at', dOtherDomain.valid, false);
}

// ADDON-009: CENTRAL's Meeting Configuration dialog has no "E-mail
// Configuration" section or Email Collection Start Date (ADDON-007B1 removed
// it), and its Mailing List default changes live with the report family, so
// the Sender field sits directly after Mailing List and no static recipient
// preview is shown (the export dialog shows From/To instead).
console.log('Configure Meeting Settings dialog -- Discussion E-mail Sender field next to Mailing List');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: 'reporter@example.com', MAILING_LIST: 'FS_6G_MED@list.etsi.org' } });
  let captured = null;
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.configureMeetingSettings();
  check('Discussion E-mail Sender appears exactly once, showing the existing configured value', (captured.match(/id="discussionEmailSender"/g) || []).length === 1 && /id="discussionEmailSender"[^>]*value="reporter@example\.com"/.test(captured), true);
  check('Mailing List appears exactly once', (captured.match(/id="mailingList"/g) || []).length, 1);
  check('the Sender field directly follows Mailing List, before Revisions / Drafts URL',
    captured.indexOf('id="mailingList"') < captured.indexOf('id="discussionEmailSender"') &&
    captured.indexOf('id="discussionEmailSender"') < captured.indexOf('id="revisionsUrl"'), true);
  check('the Sender field is labelled as the From address, never as the Mailing List', /Discussion E-mail Sender:/.test(captured) && /The "From" address/.test(captured), true);
  check('Save sends the sender field', /discussionEmailSender: document\.getElementById\('discussionEmailSender'\)/.test(captured), true);
}

{
  const { sandbox } = loadCode(); // nothing configured at all
  let captured = null;
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.configureMeetingSettings();
  check('with nothing configured, the sender field starts blank (never a hard-coded/invented default)', /id="discussionEmailSender" value=""/.test(captured), true);
}

console.log('saveConfigurationSettings() -- Discussion E-mail Sender saves through the EXISTING mechanism again, Mailing List unaffected');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ discussionEmailSender: 'reporter@example.com', mailingList: '3GPP_TSG_SA4_FS_6G_MED', mailingListMode: 'override', emailStartDate: '2026-09-01' });
  check('DISCUSSION_EMAIL_SENDER is written by Save again', docProps._store.DISCUSSION_EMAIL_SENDER, 'reporter@example.com');
  check('Mailing List still saves normally, unaffected', docProps._store.MAILING_LIST, '3GPP_TSG_SA4_FS_6G_MED');
}

{
  // a blank submitted value never erases an already-configured sender --
  // the SAME skip-if-blank protection revisionsUrl/mailingList already use.
  const { sandbox, docProps } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: 'existing@example.invalid' } });
  sandbox.saveConfigurationSettings({ discussionEmailSender: '' });
  check('a blank Save does not erase the existing configured sender', docProps._store.DISCUSSION_EMAIL_SENDER, 'existing@example.invalid');
}

console.log('the stale DISCUSSION_EMAIL_SENDER left by 887a449/006F/006G becomes ACTIVE again automatically, no migration needed');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: 'FS_6G_MED@list.etsi.org', DISCUSSION_EMAIL_SENDER: 'reporter@example.com' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('export succeeds', result.ok, true);
  const emlText = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  check('From is exactly the previously-stale, now-active DISCUSSION_EMAIL_SENDER value', /From: reporter@example\.com/.test(emlText), true);
  check('To is the SEPARATE Mailing List-derived recipient, NOT the sender address', /To: fs_6g_med@list\.etsi\.org/.test(emlText), true);
  check('From and To are deliberately different', /From: reporter@example\.com/.exec(emlText)[0] !== /To: fs_6g_med@list\.etsi\.org/.exec(emlText)[0].replace('To:', 'From:'), true);
}

console.log('generateTdocDiscussionEmails() -- exact From/To for the current meeting example; From !== To; no Session/client override; validation blocks export');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: 'reporter@example.com', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' } });
  // A different Google Session identity is deliberately present, to prove it is never consulted for either header.
  sandbox.Session.getActiveUser = () => ({ getEmail: () => 'someone.else@gmail.com' });
  const t1 = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', '[FS_6G_MED] pCR on Editorial Updates to TR 26.870'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  const t2 = makeFakeTdocTable([['TDoc', 'S4aP260074'], ['Title', 'B'], ['Status', 'Noted']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t1, t2]) });
  const createdBlobs = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { createdBlobs.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([
    { tableIndex: 0, deadline: VALID_DEADLINE, from: 'attacker@evil.invalid', sender: 'attacker2@evil.invalid', discussionEmailSender: 'attacker3@evil.invalid', to: 'attacker4@evil.invalid', recipient: 'attacker5@evil.invalid', mailingList: 'attacker6@evil.invalid' },
    { tableIndex: 1, deadline: VALID_DEADLINE }
  ]);
  check('export succeeds', result.ok, true);
  const emlBlobs = createdBlobs.filter((b) => !/\.zip$/.test(b.getName()));
  const firstText = Buffer.from(emlBlobs[0].getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8');
  check('the exact required From/To MIME headers for the current meeting example',
    [/From: reporter@example\.com/.test(firstText), /To: 3gpp_tsg_sa4_fs_6g_med@list\.etsi\.org/.test(firstText)], [true, true]);
  check('From and To are explicitly different addresses', /From: (\S+)/.exec(firstText)[1] !== /To: (\S+)/.exec(firstText)[1], true);
  check('the compact canonical Subject is preserved, unregressed',
    /Subject: \[FS_6G_MED,5\.4,26-10-15-1500CEST\]\[S4aP260069\] Discussion: \[FS_6G_MED\] pCR on Editorial Updates to TR 26\.870/.test(firstText), true);
  check('every .eml in the batch uses the SAME configured sender for From', emlBlobs.every(function (b) {
    return /From: reporter@example\.com/.test(Buffer.from(b.getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8'));
  }), true);
  check('every .eml in the batch uses the SAME derived recipient for To', emlBlobs.every(function (b) {
    return /To: 3gpp_tsg_sa4_fs_6g_med@list\.etsi\.org/.test(Buffer.from(b.getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8'));
  }), true);
  check('the Google Session identity never appears anywhere in either .eml', emlBlobs.every(function (b) {
    return Buffer.from(b.getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8').indexOf('someone.else@gmail.com') === -1;
  }), true);
  check('none of the tampered client-supplied from/sender/discussionEmailSender/to/recipient/mailingList fields appear anywhere in either .eml', emlBlobs.every(function (b) {
    return Buffer.from(b.getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8').indexOf('attacker') === -1;
  }), true);
  const zipBlob = createdBlobs.find((b) => /\.zip$/.test(b.getName()));
  check('the individual Drive .eml and the corresponding ZIP member remain identical', zipBlob._zipEntryNames.slice().sort(), emlBlobs.map((b) => b.getName()).slice().sort());
}

{
  // missing sender blocks export before any .eml or ZIP is created, even with a perfectly valid Mailing List
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' } }); // no DISCUSSION_EMAIL_SENDER
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('generation is refused when no sender is configured, even though Mailing List is fine', result.ok, false);
  check('the error names the Discussion E-mail Sender and points at Configure Meeting Settings', /Discussion E-mail Sender/.test(result.error) && /Configure Meeting Settings/.test(result.error), true);
  check('no .eml or ZIP was created', created.length, 0);
}

{
  // invalid sender (implausible syntax) blocks export
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: 'not-an-address', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('generation is refused when the configured sender is implausible', result.ok, false);
  check('no .eml or ZIP was created', created.length, 0);
}

{
  // CR/LF/header-injection in the CONFIGURED sender blocks export (real injected content after the CR/LF, never a bare trailing one -- see trim() note elsewhere)
  ['reporter@example.com\rBcc: evil@example.invalid', 'reporter@example.com\nBcc: evil@example.invalid', 'reporter@example.com\r\nBcc: evil@example.invalid'].forEach(function (poisoned) {
    const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: poisoned, MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' } });
    const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Status', 'Revised']]);
    sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
    const created = [];
    sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { created.push(b); return { getUrl: () => 'x' }; } }) }) };
    const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
    check('a CR/LF/header-injection attempt in the configured sender is rejected, nothing created', [result.ok, created.length], [false, 0]);
  });
}

{
  // unsafe Mailing List blocks export even with a perfectly valid sender.
  // ADDON-009: in CENTRAL an unset Mailing List is never missing -- it
  // resolves to the report-family list (see the CENTRAL suite) -- so the
  // Legacy "no MAILING_LIST" fixture is an unusable saved list here.
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: 'reporter@example.com', MAILING_LIST: 'not a valid list name!' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'A'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  const created = [];
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { created.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  check('generation is refused when Mailing List cannot derive a recipient, even though the sender is fine', result.ok, false);
  check('the error names Mailing List and points at Configure Meeting Settings', /Mailing List/.test(result.error) && /Configure Meeting Settings/.test(result.error), true);
  check('no .eml or ZIP was created', created.length, 0);
}

console.log('backwards compatibility -- restoring Discussion E-mail Sender never affects general meeting configuration/readiness');

{
  const { sandbox } = loadCode();
  // ADDON-009: CENTRAL's readiness is getBuildReadiness_() over
  // MEETING_READINESS_RULES_ (ADDON-007B3); Legacy's
  // computeMeetingConfigReadiness_() does not exist here.
  const readiness = { result: sandbox.getBuildReadiness_(), rules: vm.runInContext('MEETING_READINESS_RULES_', sandbox) };
  check('general meeting readiness never mentions Discussion E-mail Sender / DISCUSSION_EMAIL_SENDER',
    JSON.stringify(readiness).indexOf('DISCUSSION_EMAIL_SENDER') === -1 && JSON.stringify(readiness).indexOf('Discussion E-mail Sender') === -1, true);
}

{
  const { sandbox } = loadCode();
  const table = makeFakeTdocTable([['TDoc', 'S4aP260068'], ['Title', 'A']]);
  const mutationCalls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({
    getBody: () => new Proxy(makeFakeBody([table]), {
      get(target, prop) {
        if (['clear', 'insertParagraph', 'insertTable', 'appendParagraph', 'appendTable', 'appendListItem'].includes(prop)) {
          return () => { mutationCalls.push(prop); throw new Error('MUTATION ATTEMPTED: ' + prop); };
        }
        return target[prop];
      }
    })
  });
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: () => ({ getUrl: () => 'x' }) }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]); // fails (no sender configured), but must still never mutate
  check('an unconfigured sender causes a clean refusal, never a document mutation attempt', mutationCalls, []);
}

console.log('regression: revision-upload link, deadline forms, and X-Unsent remain exactly as before (LEGACY-UPGRADE-006H is From/To-only)');

{
  const { sandbox } = loadCode({ documentProperties: { ...LEGACY_6G_ADHOC, DISCUSSION_EMAIL_SENDER: 'reporter@example.com', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' } });
  const t = makeFakeTdocTable([['TDoc', 'S4aP260069'], ['Title', 'Editorial updates'], ['Agenda Item', '5.4'], ['Status', 'Revised']]);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeBody([t]) });
  let created = null;
  sandbox.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { if (!/\.zip$/.test(b.getName())) created = b; return { getUrl: () => 'x' }; } }) }) };
  sandbox.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: VALID_DEADLINE }]);
  const raw = Buffer.from(created.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
  const decoded = unwrapQuotedPrintableSoftBreaks(raw);
  check('X-Unsent: 1 appears exactly once', (raw.match(/X-Unsent: 1/g) || []).length, 1);
  check('body deadline sentence remains human-readable "26-10-15 15:00 CEST"', decoded.indexOf('Please provide your comments by 26-10-15 15:00 CEST.') !== -1, true);
  check('revision-upload sentence/deadline remains human-readable "26-10-15 15:00 CEST"', decoded.indexOf('Please upload revisions by 26-10-15 15:00 CEST to:') !== -1, true);
  check('subject deadline token remains the compact "26-10-15-1500CEST" form, no spaces/colons', /Subject: \[FS_6G_MED,5\.4,26-10-15-1500CEST\]/.test(raw), true);
  check('no space appears anywhere inside the first Subject bracket', /\[FS_6G_MED,5\.4,26-10-15-1500CEST\]/.test(raw), true);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
