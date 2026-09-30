/**
 * 2.15.2 follow-up -- the live "Offline Discussion" link, through the real
 * Generate path.
 *
 * Live report (6G, S4aP260091): the E-mail Discussion cell holds
 *   "Offline Discussion" -> ?tab=t.hbqo7ydmsg74#heading=h.ib1vi38y65rc
 * This runs the SAME top-level call CENTRAL's dialog makes --
 * generateTdocDiscussionEmails() -> ... -> buildEmlContent_() -- with a
 * representative active document ID, and inspects the final .eml (raw and
 * quoted-printable-decoded). No helper is called directly.
 *
 * Run: node tests/email-export-live-path.test.js
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

const LINK = '?tab=t.hbqo7ydmsg74#heading=h.ib1vi38y65rc';
const ABSOLUTE = 'https://docs.google.com/document/d/TEST_DOCUMENT_ID/edit?tab=t.hbqo7ydmsg74#heading=h.ib1vi38y65rc';
const TITLE = '[FS_6G_MED] Clarification of AI-native Traffic Characteristics in Clause 6.3.3';

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

/** The report's S4aP260091 table, rows as built by the report generator. */
function liveTable() {
  const rows = [
    ['TDoc', [paragraph([{ text: 'S4aP260091', link: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/S4aP260091.zip' }])]],
    ['Title', [paragraph([TITLE])]],
    ['Source', [paragraph(['Huawei, HiSilicon'])]],
    ['Agenda Item', [paragraph(['5.6.1'])]],
    ['Type/For', [paragraph(['pCR / Agreement'])]],
    ['E-mail Discussion', [paragraph([{ text: 'Offline Discussion', link: LINK }])]],
    ['Revisions', [paragraph(['None'])]],
    ['Minutes', [paragraph([''])]],
    ['Disposition', [paragraph([''])]],
    ['Status', [paragraph(['Revised'])]]
  ];
  const trows = rows.map(([label, value]) => { const cells = [cell([paragraph([label])]), cell(value)]; return { getNumCells: () => 2, getCell: (i) => cells[i] }; });
  return { getType: () => 'TABLE', getNumRows: () => trows.length, getRow: (i) => trows[i], getCell: (r, c) => trows[r].getCell(c) };
}

function qpDecode(text) {
  const unwrapped = text.replace(/=\r\n/g, '');
  const bytes = [];
  for (let i = 0; i < unwrapped.length; i++) {
    if (unwrapped[i] === '=' && /^[0-9A-F]{2}$/.test(unwrapped.substr(i + 1, 2))) { bytes.push(parseInt(unwrapped.substr(i + 1, 2), 16)); i += 2; } else bytes.push(unwrapped.charCodeAt(i));
  }
  return Buffer.from(bytes).toString('utf8');
}

console.log('Generate (the dialog RPC) on the live S4aP260091 table');
const { sandbox: s } = loadCode({ documentProperties: {
  MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', DISCUSSION_EMAIL_SENDER: 'reporter@example.com',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts'
} });
installExportStubs(s);
const tables = [liveTable()];
s.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => tables }), getId: () => 'TEST_DOCUMENT_ID' });
const files = [];
s.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { files.push(b); return { getUrl: () => 'x' }; } }) }) };
const result = s.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '15:00', tz: 'CEST' } }], null);

const eml = files.find((b) => /\.eml$/.test(b.getName()));
const raw = blobText(eml);
const [head, ...bodyParts] = raw.split('\r\n\r\n');
const bodyRaw = bodyParts.join('\r\n\r\n');
const html = qpDecode(bodyRaw);

check('export succeeds', [result.ok, eml.getName()], [true, 'FS_6G_MED_S4aP260091.eml']);
check('subject as generated live', /^Subject: \[FS_6G_MED,5\.6\.1,26-10-15-1500CEST\]\[S4aP260091\] Discussion: \[FS_6G_MED\] Clarification of AI-native Traffic Characteristics in Clause 6\.3\.3$/m.test(head), true);
check('decoded .eml: "Offline Discussion" links to the absolute Google Docs URL',
  html.indexOf('<a href="' + ABSOLUTE + '">Offline Discussion</a>') !== -1, true);
check('decoded .eml: no document-relative href', ['href="?tab=', 'href="#heading=', 'href="#bookmark='].filter((x) => html.indexOf(x) !== -1), []);
check('decoded .eml: the other links are unchanged', (html.match(/href="[^"]*"/g) || []).filter((h) => h.indexOf(ABSOLUTE) === -1), [
  'href="https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/S4aP260091.zip"',
  'href="https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts"'
]);
// Raw quoted-printable: "=" is =3D and long lines are soft-wrapped, so the
// URL is checked with only the soft breaks removed (no byte decoding).
const unwrapped = bodyRaw.replace(/=\r\n/g, '');
check('raw quoted-printable: href=3D"<absolute URL, = as =3D>"',
  unwrapped.indexOf('href=3D"' + ABSOLUTE.replace(/=/g, '=3D') + '">Offline Discussion</a>') !== -1, true);
check('raw quoted-printable: no relative form href=3D"?tab / #heading', /href=3D"[?#]/.test(unwrapped), false);
check('raw quoted-printable: every body line within 76 characters', bodyRaw.split('\r\n').every((l) => l.length <= 76), true);
check('CENTRAL default introduction (the live file said "October meeting")',
  [html.indexOf('prepare a revision for the upcoming meeting.') !== -1, html.indexOf('October meeting') === -1], [true, true]);

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
