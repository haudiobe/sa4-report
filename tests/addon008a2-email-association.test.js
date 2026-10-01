/**
 * ADDON-008A2 -- incoming e-mail / TDoc association for every SA4 family.
 *
 * Real failure (86172): the list message
 *   [ AMD_ARCH_Ph2-MED] S4aI260082 26501-CR0124-B "Implementing stage-2 conclusions on slicing"
 * never reached TDoc S4aI260082:
 *   1. checkRSSFeed_() skipped every ad-hoc table (legacy TDOC_ID_REGEX
 *      default '^S4-\d{6}$');
 *   2. parseEmailSubject_() returns null for it (single-part first bracket,
 *      and its full-ID regex cannot read S4a* identifiers);
 *   3. the short-number fallback ("082") cannot tell families apart.
 *
 * Now: tables are identified with the registered SA4 families; a subject
 * naming any full identifier is associated with exactly those documents;
 * only a subject without one falls back to the legacy short number, for
 * main-meeting S4- tables only. The collector reads the configured list.
 *
 * checkRSSFeed_() runs for real; only the list fetch
 * (collectHybridListservMessages_) is replaced by local messages.
 *
 * Run: node tests/addon008a2-email-association.test.js
 */

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

// The exact, unmodified subject of the real message.
const REAL_SUBJECT = '[ AMD_ARCH_Ph2-MED] S4aI260082 26501-CR0124-B "Implementing stage-2 conclusions on slicing"';

// ============================================================ scanner ====

console.log('findSA4DocumentIdsInText_: every full identifier, canonical, whatever the layout');
{
  const { sandbox: s } = loadCode();
  const ids = (t) => s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_(t));
  [
    ['1 exact real subject', REAL_SUBJECT, ['S4aI260082']],
    ['2 no leading space', '[AMD_ARCH_Ph2-MED] S4aI260082 26501-CR0124-B "Implementing stage-2 conclusions on slicing"', ['S4aI260082']],
    ['3 Re:', 'Re: ' + REAL_SUBJECT, ['S4aI260082']],
    ['4a FW:', 'FW: ' + REAL_SUBJECT, ['S4aI260082']],
    ['4b Fwd: Re:', 'Fwd: Re: ' + REAL_SUBJECT, ['S4aI260082']],
    ['5 inside the first bracket', '[AMD_ARCH_Ph2-MED; S4aI260082; 30 Sep 2026] title', ['S4aI260082']],
    ['6 newer two-bracket form', '[FS_6G_MED,5.4,26-10-15-1500CEST][S4aP260069] Discussion: title', ['S4aP260069']],
    ['7 S4aA', '[ULBC-MED] S4aA260090 agenda', ['S4aA260090']],
    ['8 S4aV', 'Comments on S4aV260012', ['S4aV260012']],
    ['9 S4aP', 'S4aP260098: revised agenda', ['S4aP260098']],
    ['10 A4aR', '[DCE] A4aR260087 draft', ['A4aR260087']],
    ['11 legacy S4-', '[FS_6G_MED; S4-261483; 25 Aug 2026 1400 CEST] Some title', ['S4-261483']],
    ['12 revision suffix -> base document', 'S4aI260082r1 updated', ['S4aI260082']],
    ['13 several identifiers, in order', 'S4aI260089 and S4aI260082 (was S4aA260090)', ['S4aI260089', 'S4aI260082', 'S4aA260090']],
    ['14 repeated identifier once', 'S4aI260082 / S4aI260082r2 / s4ai260082', ['S4aI260082']],
    ['case canonicalized to the family spelling', 'S4AI260082', ['S4aI260082']],
    ['15 unknown family rejected', 'S4aX260001 and S4bI260082', []],
    ['16 look-alike typo not guessed (lowercase l)', 'S4al260064_BBC.docx', []],
    ['no match inside a longer word', 'XS4aI260082 and S4aI2600821 and S4aI260082abc', []],
    ['no false positive from a CR number', '26501-CR0124-B "Implementing stage-2 conclusions on slicing"', []],
    ['empty', '', []]
  ].forEach(([name, text, expected]) => check(name, ids(text), expected));
}

console.log('parseEmailSubject_ is unchanged (legacy layout parser)');
{
  const { sandbox: s } = loadCode();
  check('still null for the real subject (the new scanner handles it instead)', s.parseEmailSubject_(REAL_SUBJECT), null);
  check('legacy short number unchanged', s.parseEmailSubject_('[FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Some title').tdocShort, '1483');
}

// ========================================================= end to end ====

const MSG_DATE = 'Wed, 30 Sep 2026 10:00:00 +0000';
const ADHOC_MBS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', REPORT_SUFFIX: 'MBS', MEETING_ID: '86172' };

function detailTable(body, id) {
  return body.insertTable(body._children.length, [['TDoc', id], ['Title', 't'], ['Source', 's'], ['Agenda Item', '2.5'], ['Type/For', 'CR'], ['E-mail discussion', '']]);
}

function labelledCell(table, label) {
  for (let r = 0; r < table.getNumRows(); r++) if (table.getRow(r).getCell(0).getText() === label) return table.getRow(r).getCell(1).getText();
  return null;
}

/** Runs the real collector over local messages; returns per-TDoc results. */
function collect(tdocIds, subjects, props) {
  const loaded = loadCode({ documentProperties: props || ADHOC_MBS });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const tables = {};
  tdocIds.forEach(id => { tables[id] = detailTable(body, id); });
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  let fetchedList = null;
  s.collectHybridListservMessages_ = (cfg) => {
    fetchedList = cfg.LIST_NAME;
    return subjects.map((title, i) => ({ title, author: 'Delegate ' + i, date: MSG_DATE, link: 'https://list.etsi.org/m' + i, messageId: 'm' + i, preview: '' }));
  };
  s.checkRSSFeed_(s.getCollectorConfig_());
  const result = {};
  tdocIds.forEach(id => {
    const stored = loaded.docProps._store['DISCUSS_' + id];
    result[id] = {
      processed: stored !== undefined,
      matched: stored ? Object.keys(JSON.parse(stored)).length : 0,
      cell: labelledCell(tables[id], 'E-mail discussion')
    };
  });
  return { result, fetchedList };
}

console.log('end to end: the real 86172 message reaches S4aI260082');
{
  const r = collect(['S4aI260082'], [REAL_SUBJECT]);
  check('collector reads the MBS list', r.fetchedList, '3GPP_TSG_SA_WG4_MBS');
  check('table S4aI260082 is processed (was skipped by ^S4-\\d{6}$)', r.result.S4aI260082.processed, true);
  check('DISCUSS_S4aI260082 holds the message', r.result.S4aI260082.matched, 1);
  check('written into the E-mail discussion cell', /^Delegate 0 on 2026-09-30T10:00/.test(r.result.S4aI260082.cell), true);

  const noSpace = collect(['S4aI260082'], [REAL_SUBJECT.replace('[ ', '[')]);
  check('the no-space variant associates the same way', noSpace.result.S4aI260082.matched, 1);
  const reply = collect(['S4aI260082'], ['Re: ' + REAL_SUBJECT]);
  check('a reply associates the same way', reply.result.S4aI260082.matched, 1);
}

console.log('agenda items are not required and other TDocs are not touched');
{
  const r = collect(['S4aI260080', 'S4aI260082', 'S4aI260089'], [REAL_SUBJECT]);
  check('only S4aI260082 gets the message', [r.result.S4aI260080.matched, r.result.S4aI260082.matched, r.result.S4aI260089.matched], [0, 1, 0]);
  check('the others say "No e-mail discussion."', [r.result.S4aI260080.cell, r.result.S4aI260089.cell], ['No e-mail discussion.', 'No e-mail discussion.']);
}

console.log('multiple identifiers in one subject reach every named report TDoc');
{
  const r = collect(['S4aI260082', 'S4aI260089', 'S4aI260080'], ['[AMD_ARCH_Ph2-MED] S4aI260082 merged into S4aI260089']);
  check('both named TDocs, not the third', [r.result.S4aI260082.matched, r.result.S4aI260089.matched, r.result.S4aI260080.matched], [1, 1, 0]);
}

console.log('short-number collisions across families are impossible');
{
  const legacyOnly = collect(['S4aI260082', 'S4aA260082'], ['[AMD_ARCH_Ph2-MED, 082, 5 Oct 2026 1400 CEST] title']);
  check('bare "082" reaches neither ad-hoc TDoc', [legacyOnly.result.S4aI260082.matched, legacyOnly.result.S4aA260082.matched], [0, 0]);
  const one = collect(['S4aI260082', 'S4aA260082'], ['[AMD_ARCH_Ph2-MED] S4aI260082 comments']);
  check('explicit S4aI260082 reaches only S4aI260082', [one.result.S4aI260082.matched, one.result.S4aA260082.matched], [1, 0]);
  const both = collect(['S4aI260082', 'S4aA260082'], ['S4aI260082 vs S4aA260082']);
  check('naming both reaches both', [both.result.S4aI260082.matched, both.result.S4aA260082.matched], [1, 1]);
  const full = collect(['S4aI260082', 'S4-260082'], ['[AMD_ARCH_Ph2-MED; 082; 5 Oct 2026] see S4aI260082']);
  check('a full identifier disables the short-number fallback for that message', [full.result.S4aI260082.matched, full.result['S4-260082'].matched], [1, 0]);
}

console.log('main meetings: legacy subjects keep working');
{
  const MAIN = { REPORT_SUFFIX: '6G', MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136' };
  const shortNum = collect(['S4-261483', 'S4-261484'], ['[FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Some title'], MAIN);
  check('legacy short number -> S4-261483 only', [shortNum.result['S4-261483'].matched, shortNum.result['S4-261484'].matched], [1, 0]);
  const fullMain = collect(['S4-261483'], ['[FS_6G_MED; S4-261483; 25 Aug 2026 1400 CEST] Some title'], MAIN);
  check('legacy full S4- identifier', fullMain.result['S4-261483'].matched, 1);
  const hyphenless = collect(['S4-261483'], ['[FS_6G_MED; S4261483; 25 Aug 2026 1400 CEST] Some title'], MAIN);
  check('legacy hyphenless S4 form still accepted by the fallback', hyphenless.result['S4-261483'].matched, 1);
  check('a main 6G report reads its family list (2.17.2: the 6G list)', shortNum.fetchedList, '3GPP_TSG_SA4_FS_6G_MED');
  const unknown = collect(['S4-261483', 'not-a-tdoc'], [], MAIN);
  check('a malformed table id is still skipped', unknown.result['not-a-tdoc'].processed, false);
}

// ======================================================== mailing list ====

console.log('collector list: configured MAILING_LIST (ad-hoc), else the family list');
{
  const listFor = (props) => { const { sandbox: s } = loadCode({ documentProperties: props }); s.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeDocumentBody(s), getId: () => 'D' }); return s.getCollectorConfig_(); };
  const configured = listFor(Object.assign({}, ADHOC_MBS, { REPORT_SUFFIX: 'Video', MAILING_LIST: '3GPP_TSG_SA_WG4_MBS' }));
  check('configured MBS list wins over the Video family list', [configured.LIST_NAME, configured.RSS_URL_V2],
    ['3GPP_TSG_SA_WG4_MBS', 'https://list.etsi.org/scripts/wa.exe?RSS&L=3GPP_TSG_SA_WG4_MBS&v=2.0&LIMIT=2000']);
  check('absent -> family list (unchanged)', listFor(ADHOC_MBS).LIST_NAME, '3GPP_TSG_SA_WG4_MBS');
  check('absent, Audio -> Audio list', listFor(Object.assign({}, ADHOC_MBS, { REPORT_SUFFIX: 'Audio' })).LIST_NAME, '3GPP_TSG_SA_WG4_AUDIO');
  // TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-003): the reflector-address
  // form "<list>@list.etsi.org" -- which the exporter already accepts as a
  // Mailing List -- is normalized to the list name instead of being
  // rejected. Any other address stays invalid.
  const address = listFor(Object.assign({}, ADHOC_MBS, { REPORT_SUFFIX: 'Video', MAILING_LIST: '3GPP_TSG_SA_WG4_MBS@LIST.ETSI.ORG' }));
  check('"<list>@list.etsi.org" is read as that list', [address.LIST_NAME, address.RSS_URL_V2.indexOf('L=3GPP_TSG_SA_WG4_MBS&') !== -1], ['3GPP_TSG_SA_WG4_MBS', true]);
  ['3GPP TSG SA WG4 MBS', '3GPP_TSG_SA_WG4_MBS@example.org', 'https://list.etsi.org/x', 'A&L=OTHER'].forEach(bad => {
    const c = listFor(Object.assign({}, ADHOC_MBS, { REPORT_SUFFIX: 'Video', MAILING_LIST: bad }));
    check(`invalid "${bad}" is not used; family list read instead`, [c.LIST_NAME, c.RSS_URL_V2.indexOf('L=3GPP_TSG_SA_WG4_VIDEO&') !== -1], ['3GPP_TSG_SA_WG4_VIDEO', true]);
  });
  check('main meetings keep the family list (existing getMeetingContext_ rule)',
    listFor({ REPORT_SUFFIX: 'Audio', MAILING_LIST: '3GPP_TSG_SA_WG4_MBS' }).LIST_NAME, '3GPP_TSG_SA_WG4_AUDIO');

  const { sandbox: s } = loadCode({ documentProperties: Object.assign({}, ADHOC_MBS, { MAILING_LIST: '3GPP_TSG_SA_WG4_VIDEO' }) });
  const body = makeFakeDocumentBody(s);
  body.insertTable(0, [['Key', 'Value'], ['LIST_NAME', 'SOME_TABLE_LIST']]);
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'D' });
  check('a Collector Configuration table LIST_NAME still wins (as before)', s.getCollectorConfig_().LIST_NAME, 'SOME_TABLE_LIST');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
