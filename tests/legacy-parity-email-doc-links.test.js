/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-email-doc-links.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-email-doc-links.test.js
 */

/**
 * LEGACY -- Google Docs document-relative links in exported discussion
 * e-mails (port of CENTRAL 2.15.2, 007d24c).
 *
 * Live Legacy e-mail (6G report, S4aP260091): the E-mail Discussion cell
 * holds "Offline Discussion" -> ?tab=t.hbqo7ydmsg74#heading=h.ib1vi38y65rc,
 * which the exporter copied verbatim:
 *   <a href="?tab=t.hbqo7ydmsg74#heading=h.ib1vi38y65rc">Offline Discussion</a>
 * Such links are now resolved against the report's own URL,
 * https://docs.google.com/document/d/<id>/edit; absolute links are
 * unchanged.
 *
 * The integration part runs the real Generate RPC
 * (generateTdocDiscussionEmails()) and inspects the final .eml, raw and
 * quoted-printable-decoded. A representative document ID is used.
 *
 * Run: node tests/legacy-email-doc-links.test.js
 */

const { loadCode } = require('./helpers/legacy-load-code.js');

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

const LINK = '?tab=t.hbqo7ydmsg74#heading=h.ib1vi38y65rc';
const DOC_URL = 'https://docs.google.com/document/d/TEST_DOCUMENT_ID/edit';
const ABSOLUTE = DOC_URL + LINK;
const TITLE = '[FS_6G_MED] Clarification of AI-native Traffic Characteristics in Clause 6.3.3';
const ZIP_URL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/S4aP260091.zip';
const DRAFTS_URL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts';

/** A paragraph of plain strings and {text, link} runs, like DocumentApp's Paragraph/Text. */
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
const cell = (paragraphs) => ({ getText: () => paragraphs.map((p) => p.getText()).join('\n'), getNumChildren: () => paragraphs.length, getChild: (i) => paragraphs[i] });

/** The report's S4aP260091 table; `discussion` = the E-mail Discussion cell's paragraphs. */
function liveTable(discussion) {
  const rows = [
    ['TDoc', [paragraph([{ text: 'S4aP260091', link: ZIP_URL }])]],
    ['Title', [paragraph([TITLE])]],
    ['Source', [paragraph(['Huawei, HiSilicon'])]],
    ['Agenda Item', [paragraph(['5.6.1'])]],
    ['Type/For', [paragraph(['pCR / Agreement'])]],
    ['E-mail Discussion', discussion || [paragraph([{ text: 'Offline Discussion', link: LINK }])]],
    ['Revisions', [paragraph(['None'])]],
    ['Minutes', [paragraph([''])]],
    ['Disposition', [paragraph([''])]],
    ['Status', [paragraph(['Revised'])]]
  ];
  const trows = rows.map(([label, value]) => { const cells = [cell([paragraph([label])]), cell(value)]; return { getNumCells: () => 2, getCell: (i) => cells[i] }; });
  return { getType: () => 'TABLE', getNumRows: () => trows.length, getRow: (i) => trows[i], getCell: (r, c) => trows[r].getCell(c) };
}

const blobText = (b) => Buffer.from(b.getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8');
function qpDecode(text) {
  const unwrapped = text.replace(/=\r\n/g, '');
  const bytes = [];
  for (let i = 0; i < unwrapped.length; i++) {
    if (unwrapped[i] === '=' && /^[0-9A-F]{2}$/.test(unwrapped.substr(i + 1, 2))) { bytes.push(parseInt(unwrapped.substr(i + 1, 2), 16)); i += 2; } else bytes.push(unwrapped.charCodeAt(i));
  }
  return Buffer.from(bytes).toString('utf8');
}

/** Runs the real Generate RPC on one table in a document with `documentId`. */
function generate(table, documentId) {
  const { sandbox: s } = loadCode({ documentProperties: {
    MEETING_TYPE: 'adhoc', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', DISCUSSION_EMAIL_SENDER: 'reporter@example.com', REVISIONS_URL: DRAFTS_URL
  } });
  const tables = [table];
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => tables }), getId: () => documentId });
  const files = [];
  s.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { files.push(b); return { getUrl: () => 'x' }; } }) }) };
  const result = s.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '15:00', tz: 'CEST' } }], null);
  const eml = files.find((b) => /\.eml$/.test(b.getName()));
  const raw = eml ? blobText(eml) : '';
  const parts = raw.split('\r\n\r\n');
  const bodyRaw = parts.slice(1).join('\r\n\r\n');
  return { s, result, eml, head: parts[0], bodyRaw, html: qpDecode(bodyRaw), zips: files.filter((b) => /\.zip$/.test(b.getName())) };
}

// ------------------------------------------------ the live case, end to end

console.log('Generate on the live S4aP260091 table (E-mail Discussion -> "Offline Discussion")');
{
  const out = generate(liveTable(), 'TEST_DOCUMENT_ID');
  check('export succeeds, Legacy file name', [out.result.ok, out.eml.getName()], [true, 'FS_6G_MED_S4aP260091.eml']);
  check('Legacy headers unchanged', out.head.split('\r\n'), [
    'MIME-Version: 1.0', 'X-Unsent: 1', 'From: reporter@example.com', 'To: 3gpp_tsg_sa4_fs_6g_med@list.etsi.org',
    'Subject: [FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260091] Discussion: ' + TITLE,
    'Content-Type: text/html; charset=UTF-8', 'Content-Transfer-Encoding: quoted-printable'
  ]);
  check('decoded .eml: "Offline Discussion" links to the absolute Google Docs URL',
    out.html.indexOf('<a href="' + ABSOLUTE + '">Offline Discussion</a>') !== -1, true);
  check('decoded .eml: no document-relative href', ['href="?tab=', 'href="#heading=', 'href="#bookmark='].filter((x) => out.html.indexOf(x) !== -1), []);
  check('decoded .eml: absolute links unchanged', (out.html.match(/href="[^"]*"/g) || []).filter((h) => h.indexOf(ABSOLUTE) === -1),
    ['href="' + ZIP_URL + '"', 'href="' + DRAFTS_URL + '"']);
  // INTENTIONAL DIFFERENCE (decision 2026-10-01, CENTRAL 2.15.0): the
  // default introduction is meeting-neutral. The list tag comes from the
  // meeting's mailing list; Legacy's "the October meeting" is never used.
  check('[template] introduction is meeting-neutral: tag from the mailing list, "the upcoming meeting"',
    [out.html.indexOf('As discussed during the FS_6G_MED AHG') !== -1, out.html.indexOf('prepare a revision for the upcoming meeting.') !== -1, /October meeting/.test(out.html)],
    [true, true, false]);
  const unwrapped = out.bodyRaw.replace(/=\r\n/g, '');
  check('raw quoted-printable: href=3D"<absolute URL, = as =3D>"',
    unwrapped.indexOf('href=3D"' + ABSOLUTE.replace(/=/g, '=3D') + '">Offline Discussion</a>') !== -1, true);
  check('raw quoted-printable: no relative form href=3D"? / href=3D"#', /href=3D"[?#]/.test(unwrapped), false);
  check('raw quoted-printable: valid (lines <= 76, only =XX escapes / soft breaks)',
    out.bodyRaw.split('\r\n').every((l) => l.length <= 76) && !/=(?![0-9A-F]{2}|\r\n|$)/.test(out.bodyRaw), true);
  check('ZIP still holds exactly the .eml', out.zips.length === 1 && JSON.stringify(out.zips[0]._zipEntryNames) === JSON.stringify(['FS_6G_MED_S4aP260091.eml']), true);
}

{
  const out = generate(liveTable([
    paragraph([{ text: 'Offline Discussion', link: LINK }]),
    paragraph(['Tab: ', { text: 'Minutes', link: '?tab=t.hbqo7ydmsg74' }]),
    paragraph(['Heading: ', { text: 'Section', link: '#heading=h.4b2x9k1' }]),
    paragraph(['Bookmark: ', { text: 'Mark', link: '#bookmark=id.abc123' }]),
    paragraph(['Thread: ', { text: 'Alice', link: 'https://list.etsi.org/scripts/wa.exe?A2=ind2609&L=3GPP_TSG_SA4_FS_6G_MED&P=1' }]),
    paragraph(['Mail: ', { text: 'Bob', link: 'mailto:bob@example.invalid' }]),
    paragraph(['Plain http: ', { text: 'x', link: 'http://example.org/a?b=c#d' }]),
    paragraph(['Not a link: ?tab=t.hbqo7ydmsg74'])
  ]), 'TEST_DOCUMENT_ID');
  check('?tab / #heading / #bookmark all resolved; https, mailto, http unchanged; plain text untouched', [
    out.html.indexOf('<a href="' + DOC_URL + '?tab=t.hbqo7ydmsg74">Minutes</a>') !== -1,
    out.html.indexOf('<a href="' + DOC_URL + '#heading=h.4b2x9k1">Section</a>') !== -1,
    out.html.indexOf('<a href="' + DOC_URL + '#bookmark=id.abc123">Mark</a>') !== -1,
    out.html.indexOf('<a href="https://list.etsi.org/scripts/wa.exe?A2=ind2609&amp;L=3GPP_TSG_SA4_FS_6G_MED&amp;P=1">Alice</a>') !== -1,
    out.html.indexOf('<a href="mailto:bob@example.invalid">Bob</a>') !== -1,
    out.html.indexOf('<a href="http://example.org/a?b=c#d">x</a>') !== -1,
    out.html.indexOf('Not a link: ?tab=t.hbqo7ydmsg74</p>') !== -1
  ], [true, true, true, true, true, true, true]);
}

{
  const out = generate(liveTable(), undefined);
  check('no document ID available: export still succeeds, link left as stored', [out.result.ok, out.html.indexOf('href="' + LINK + '"') !== -1], [true, true]);
}

// --------------------------------------------------- the ported helpers

console.log('ported CENTRAL 2.15.2 helpers (same semantics)');
{
  const { sandbox: s } = loadCode();
  const r = (l) => s.resolveEmailExportLinkUrl_(l, DOC_URL);
  check('resolver: document-relative forms', [r(LINK), r('?tab=t.x'), r('#heading=h.x'), r('#bookmark=id.x')],
    [ABSOLUTE, DOC_URL + '?tab=t.x', DOC_URL + '#heading=h.x', DOC_URL + '#bookmark=id.x']);
  check('resolver: absolute / mailto / other values unchanged',
    [r('https://a.example/x?y#z'), r('http://a.example'), r('mailto:a@b.invalid'), r('not a link'), r('/document/d/X/edit'), r('?tab=t.x y'), r(''), r(null)],
    ['https://a.example/x?y#z', 'http://a.example', 'mailto:a@b.invalid', 'not a link', '/document/d/X/edit', '?tab=t.x y', '', '']);
  check('resolver: no document URL -> unchanged', s.resolveEmailExportLinkUrl_(LINK, ''), LINK);
  check('document URL from the ID; none for a missing/malformed ID',
    [s.buildEmailExportDocumentUrl_('TEST_DOCUMENT_ID'), s.buildEmailExportDocumentUrl_(null), s.buildEmailExportDocumentUrl_('a/b')], [DOC_URL, '', '']);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
