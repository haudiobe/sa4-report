/**
 * 2.15.2 -- Google Docs links to places inside the report, in exported
 * discussion e-mails.
 *
 * Live failure (6G report): the linked text
 *   "6G Media Minutes – SA4-e (AH) on FS_6G_MED"
 * points to another tab of the same report; Google Docs stores that link as
 * "?tab=t.hbqo7ydmsg74", and the exporter copied it verbatim, giving
 * <a href="?tab=t.hbqo7ydmsg74"> -- meaningless in an .eml. Such links are
 * now resolved against the report's own URL; absolute links are unchanged.
 *
 * A representative document ID is used -- never a live document.
 *
 * Run: node tests/email-export-doc-links.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
const { installExportStubs, blobText } = require('./helpers/email-export-stubs.js');

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

const DOC_ID = '1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-Xy';
const DOC_URL = 'https://docs.google.com/document/d/' + DOC_ID + '/edit';
const MINUTES_TEXT = '6G Media Minutes – SA4-e (AH) on FS_6G_MED';
const TAB_LINK = '?tab=t.hbqo7ydmsg74';

/** A paragraph of `parts`: plain strings or {text, link}. */
function paragraph(parts) {
  let full = '';
  const runs = [];
  parts.forEach((p) => {
    const t = typeof p === 'string' ? p : p.text;
    if (typeof p !== 'string') runs.push({ start: full.length, end: full.length + t.length - 1, link: p.link });
    full += t;
  });
  const at = (o) => runs.find((r) => o >= r.start && o <= r.end);
  const text = {
    getText: () => full,
    getTextAttributeIndices: () => Array.from(new Set([0].concat(...runs.map((r) => [r.start, r.end + 1])))).filter((i) => i < full.length).sort((a, b) => a - b),
    isBold: () => false, isItalic: () => false, isUnderline: (o) => !!at(o),
    getLinkUrl: (o) => (at(o) ? at(o).link : null)
  };
  const p = { getType: () => 'PARAGRAPH', asParagraph: () => p, editAsText: () => text, getText: () => full };
  return p;
}

function cell(paragraphs) {
  return { getText: () => paragraphs.map((p) => p.getText()).join('\n'), getNumChildren: () => paragraphs.length, getChild: (i) => paragraphs[i] };
}

function tdocTable(id, minutesParagraphs) {
  const rows = [['TDoc', [paragraph([id])]], ['Title', [paragraph(['pCR on Editorial Updates to TR 26.870'])]], ['Agenda Item', [paragraph(['5.4'])]],
    ['Minutes', minutesParagraphs], ['Status', [paragraph(['Revised'])]]];
  const trows = rows.map(([label, value]) => { const cells = [cell([paragraph([label])]), cell(value)]; return { getNumCells: () => 2, getCell: (i) => cells[i] }; });
  return { getType: () => 'TABLE', getNumRows: () => trows.length, getRow: (i) => trows[i], getCell: (r, c) => trows[r].getCell(c) };
}

/** Full quoted-printable decode (soft breaks and =XX bytes) of an .eml body. */
function qpDecode(text) {
  const unwrapped = text.replace(/=\r\n/g, '');
  const bytes = [];
  for (let i = 0; i < unwrapped.length; i++) {
    if (unwrapped[i] === '=' && /^[0-9A-F]{2}$/.test(unwrapped.substr(i + 1, 2))) { bytes.push(parseInt(unwrapped.substr(i + 1, 2), 16)); i += 2; } else bytes.push(unwrapped.charCodeAt(i));
  }
  return Buffer.from(bytes).toString('utf8');
}

const FS_6G_MED = {
  MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_FS_6G_MED/Inbox/Drafts/', DISCUSSION_EMAIL_SENDER: 'reporter@example.com'
};

function exportOne(minutesParagraphs, documentId) {
  const { sandbox: s } = loadCode({ documentProperties: FS_6G_MED });
  installExportStubs(s);
  const tables = [tdocTable('S4aP260069', minutesParagraphs)];
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => tables }), getId: () => documentId });
  const files = [];
  s.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { files.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = s.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '15:00', tz: 'CEST' } }], null);
  const eml = files.find((b) => /\.eml$/.test(b.getName()));
  return { result, raw: eml ? blobText(eml) : '', html: eml ? qpDecode(blobText(eml).split('\r\n\r\n').slice(1).join('\r\n\r\n')) : '' };
}

// ------------------------------------------------------------ resolver

console.log('resolveEmailExportLinkUrl_(): only document-relative links are resolved');
{
  const { sandbox: s } = loadCode();
  const r = (link) => s.resolveEmailExportLinkUrl_(link, DOC_URL);
  check('?tab=... -> document URL + query', r(TAB_LINK), DOC_URL + TAB_LINK);
  check('#heading=... -> document URL + fragment', r('#heading=h.4b2x9k1'), DOC_URL + '#heading=h.4b2x9k1');
  check('#bookmark=... -> document URL + fragment', r('#bookmark=id.abc123'), DOC_URL + '#bookmark=id.abc123');
  check('?tab=...#heading=... -> query and fragment kept', r('?tab=t.hbqo7ydmsg74#heading=h.4b2x9k1'), DOC_URL + '?tab=t.hbqo7ydmsg74#heading=h.4b2x9k1');
  check('https URL unchanged', r('https://www.3gpp.org/ftp/x/S4aP260069.zip'), 'https://www.3gpp.org/ftp/x/S4aP260069.zip');
  check('another Google Doc (absolute) unchanged', r('https://docs.google.com/document/d/OTHER/edit?tab=t.1'), 'https://docs.google.com/document/d/OTHER/edit?tab=t.1');
  check('http URL unchanged', r('http://example.org/a?b=c#d'), 'http://example.org/a?b=c#d');
  check('mailto unchanged', r('mailto:alice@example.invalid'), 'mailto:alice@example.invalid');
  check('empty / null / undefined -> empty (no link)', [r(''), r(null), r(undefined)], ['', '', '']);
  check('arbitrary text is not turned into a Docs URL', r('not a link'), 'not a link');
  check('a root-relative path is not guessed', r('/document/d/OTHER/edit'), '/document/d/OTHER/edit');
  check('a relative-looking value with whitespace/quotes is not resolved', [r('?tab=t.x y'), r('#a"b')], ['?tab=t.x y', '#a"b']);
  check('no known document URL -> left unchanged', s.resolveEmailExportLinkUrl_(TAB_LINK, ''), TAB_LINK);
  check('document URL from the ID', s.buildEmailExportDocumentUrl_(DOC_ID), DOC_URL);
  check('no / malformed ID -> no document URL', [s.buildEmailExportDocumentUrl_(null), s.buildEmailExportDocumentUrl_(''), s.buildEmailExportDocumentUrl_('a/b?c')], ['', '', '']);
}

// ------------------------------------------------------------ HTML

console.log('richTextToHtml_() -- the real 6G minutes link');
{
  const { sandbox: s } = loadCode();
  const p = paragraph([{ text: MINUTES_TEXT, link: TAB_LINK }]);
  check('exact HTML: absolute href, visible text unchanged',
    s.richTextToHtml_(p.editAsText(), DOC_URL),
    '<a href="' + DOC_URL + '?tab=t.hbqo7ydmsg74">' + MINUTES_TEXT + '</a>');
  check('without a document URL the link is left as it was (Legacy behavior)',
    s.richTextToHtml_(p.editAsText()), '<a href="?tab=t.hbqo7ydmsg74">' + MINUTES_TEXT + '</a>');
  const mixed = paragraph(['See ', { text: 'TDoc', link: 'https://www.3gpp.org/ftp/x.zip' }, ' & ', { text: 'Alice', link: 'mailto:alice@example.invalid' }, ' <done>']);
  check('absolute and mailto links unchanged, text escaped as before',
    s.richTextToHtml_(mixed.editAsText(), DOC_URL),
    'See <a href="https://www.3gpp.org/ftp/x.zip">TDoc</a> &amp; <a href="mailto:alice@example.invalid">Alice</a> &lt;done&gt;');
  check('plain text without a link is untouched', s.richTextToHtml_(paragraph(['?tab=t.hbqo7ydmsg74 as plain text']).editAsText(), DOC_URL), '?tab=t.hbqo7ydmsg74 as plain text');
}

// ------------------------------------------------------------ end to end

console.log('generateTdocDiscussionEmails() -- the .eml carries the absolute link');
{
  const out = exportOne([
    paragraph([{ text: MINUTES_TEXT, link: TAB_LINK }]),
    paragraph(['Heading: ', { text: 'Section 5.4', link: '#heading=h.4b2x9k1' }]),
    paragraph(['Contribution: ', { text: 'S4aP260069', link: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_FS_6G_MED/Docs/S4aP260069.zip' }])
  ], DOC_ID);
  check('export succeeds', out.result.ok, true);
  check('the minutes link is the complete Google Docs URL, visible text unchanged',
    out.html.indexOf('<a href="' + DOC_URL + '?tab=t.hbqo7ydmsg74">' + MINUTES_TEXT + '</a>') !== -1, true);
  check('no document-relative href remains', /href="[?#]/.test(out.html), false);
  check('no standalone href="?tab=t.hbqo7ydmsg74"', out.html.indexOf('href="?tab=t.hbqo7ydmsg74"') === -1, true);
  check('heading link resolved too', out.html.indexOf('<a href="' + DOC_URL + '#heading=h.4b2x9k1">Section 5.4</a>') !== -1, true);
  check('external contribution link unchanged', out.html.indexOf('<a href="https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_FS_6G_MED/Docs/S4aP260069.zip">S4aP260069</a>') !== -1, true);
  check('still quoted-printable with soft line breaks (lines <= 76)', out.raw.split('\r\n\r\n').slice(1).join('').split('\r\n').every((l) => l.length <= 76), true);
  check('subject (TEMPLATE-002C form: no tag)', /^Subject: \[5\.4\]\[26-10-15-1500CEST\]\[S4aP260069\] Discussion: pCR on Editorial Updates to TR 26\.870$/m.test(out.raw), true);
}

{
  // No document ID available (defensive): links are copied as before, export still works.
  const out = exportOne([paragraph([{ text: MINUTES_TEXT, link: TAB_LINK }])], undefined);
  check('without a document ID the export still succeeds, link left unchanged',
    [out.result.ok, out.html.indexOf('href="?tab=t.hbqo7ydmsg74"') !== -1], [true, true]);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
