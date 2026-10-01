/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-adhoc-email-collection.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-adhoc-email-collection.test.js
 */

/**
 * LEGACY -- incoming discussion e-mail collection for ad-hoc TDoc families
 * (fix for DIAGNOSIS #3; port of the relevant subset of CENTRAL ADDON-008A2
 * plus the A1 archive list and EMAIL_START_DATE).
 *
 * The accepted Legacy exporter is the contract: its generated subject
 *   [FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260091] Discussion: ...
 * and replies to it must reach S4aP260091 and nothing else, read from the
 * report's own list (3GPP_TSG_SA4_FS_6G_MED) via both RSS and A1.
 *
 * Real production code throughout: generateTdocDiscussionEmails(),
 * getCollectorConfig_(), collectHybridListservMessages_() (RSS + A1),
 * checkRSSFeed_(), collectorUpdate_(), continuousUpdate(). Only the
 * platform (UrlFetchApp, XmlService, DocumentApp body) is faked.
 *
 * Run: node tests/legacy-adhoc-email-collection.test.js
 */

const { loadCode } = require('./helpers/legacy-load-code.js');
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

const LIST_6G = '3GPP_TSG_SA4_FS_6G_MED';
const PROPS_6G = {
  MEETING_TYPE: 'adhoc',
  MEETING_NAME: 'SA4-e (AH) on FS_6G_MED',
  MEETING_ID: '86178',
  REPORT_SUFFIX: '6G',
  MAILING_LIST: LIST_6G,
  EMAIL_START_DATE: '2026-08-21',
  MEETING_FOLDER: 'TSGS4_136_Montreal',
  MEETING_NUMBER: '136',
  DISCUSSION_EMAIL_SENDER: 'reporter@example.com',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts'
};
const MAIN = { REPORT_SUFFIX: '6G', MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136' };
const props6G = (extra) => Object.assign({}, PROPS_6G, extra || {});
const TITLE = '[FS_6G_MED] Clarification of AI-native Traffic Characteristics in Clause 6.3.3';
const MSG_DATE = 'Thu, 01 Oct 2026 08:00:00 GMT';
const NOW = '2026-10-01T09:00:00Z';

// --------------------------------------------------------------- exporter

function paragraph(text) {
  const t = { getText: () => text, getTextAttributeIndices: () => [0], isBold: () => false, isItalic: () => false, isUnderline: () => false, getLinkUrl: () => null };
  const p = { getType: () => 'PARAGRAPH', asParagraph: () => p, editAsText: () => t, getText: () => text };
  return p;
}
function exportTable(tdocId, agendaItem) {
  const rows = [['TDoc', tdocId], ['Title', TITLE], ['Source', 'Qualcomm'], ['Agenda Item', agendaItem],
    ['Type/For', 'pCR / Agreement'], ['E-mail Discussion', ''], ['Revisions', 'None'], ['Status', 'Available']];
  const trows = rows.map(([l, v]) => {
    const cells = [l, v].map((x) => ({ getText: () => x, getNumChildren: () => 1, getChild: () => paragraph(x) }));
    return { getNumCells: () => 2, getCell: (i) => cells[i] };
  });
  return { getType: () => 'TABLE', getNumRows: () => trows.length, getRow: (i) => trows[i], getCell: (r, c) => trows[r].getCell(c) };
}
/** Subject header of the .eml produced by the real Legacy Generate RPC. */
function exportedSubject(tdocId, agendaItem) {
  const { sandbox: s } = loadCode({ documentProperties: PROPS_6G });
  const tables = [exportTable(tdocId, agendaItem)];
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => tables }), getId: () => 'DOC' });
  const files = [];
  s.DriveApp = { getFoldersByName: () => ({ hasNext: () => true, next: () => ({ createFile: (b) => { files.push(b); return { getUrl: () => 'x' }; } }) }) };
  const r = s.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '15:00', tz: 'CEST' } }], null);
  const eml = files.find((b) => /\.eml$/.test(b.getName()));
  if (!eml) throw new Error('export failed: ' + JSON.stringify(r));
  const raw = Buffer.from(eml.getBytes().map((x) => (x < 0 ? x + 256 : x))).toString('utf8');
  return raw.split('\r\n\r\n')[0].split('\r\n').find((h) => h.indexOf('Subject: ') === 0).slice(9);
}

// ------------------------------------------------------------- fake ETSI

const fakeXmlService = {
  getNamespace: () => ({}),
  parse: (xml) => {
    const items = (xml.match(/<item>[\s\S]*?<\/item>/g) || []).map((it) => ({
      getChildText: (tag) => { const m = it.match(new RegExp('<' + tag + '>([\\s\\S]*?)</' + tag + '>')); return m ? m[1] : null; }
    }));
    const channel = { getChildren: () => items };
    return { getRootElement: () => ({ getChild: () => channel, getChildren: () => items }) };
  }
};
const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const rssXml = (msgs) => '<rss><channel>' + msgs.map((m, i) =>
  `<item><title>${esc(m.title)}</title><author>${m.author || 'Delegate ' + i}</author><pubDate>${m.date || MSG_DATE}</pubDate><link>https://list.etsi.org/rss/m${i}</link><guid>rss-${i}</guid><description></description></item>`).join('') + '</channel></rss>';
const a1Html = (list, msgs) => '<html><table>' + msgs.map((m, i) =>
  `<tr><td><div><a href="/scripts/wa.exe?A2=ind2610A&L=${list}&P=${i}">${esc(m.title)}</a></div></td><td><div>${m.author || 'Delegate ' + i}</div></td><td><div>${m.a1Date || m.date || MSG_DATE}</div></td></tr>`).join('') + '</table></html>';

/**
 * Serves `rss` messages on the RSS feed and `a1` messages on the A1 week
 * pages -- only for list `onList`; every other list is empty. Records URLs.
 */
function installFakeEtsi(s, onList, rss, a1) {
  const fetched = [];
  s.XmlService = fakeXmlService;
  s.Utilities.Charset = { UTF_8: 'UTF-8' };
  s.UrlFetchApp = {
    fetch: (url) => {
      fetched.push(url);
      const hit = new RegExp('[?&]L=' + onList + '(&|$)').test(url);
      const text = url.indexOf('?RSS&') !== -1 ? rssXml(hit ? (rss || []) : []) : (hit && a1 && a1.length ? a1Html(onList, a1) : '<html></html>');
      return { getResponseCode: () => 200, getContentText: () => text };
    }
  };
  return fetched;
}
const listsOf = (urls) => Array.from(new Set(urls.map((u) => (u.match(/[?&]L=([^&]+)/) || [])[1])));

/** Pins "now" so the A1 archive window is deterministic. */
function pinNow(s) {
  const RealDate = Date;
  s.Date = class extends RealDate {
    constructor(...a) { if (a.length === 0) super(NOW); else super(...a); }
    static now() { return new RealDate(NOW).getTime(); }
  };
}

// --------------------------------------------------------- fake document

function detailTable(body, id) {
  return body.insertTable(body._children.length, [['TDoc', id], ['Title', TITLE], ['Source', 'Qualcomm'], ['Agenda Item', '5.6.1'], ['Type/For', 'pCR'], ['E-mail discussion', '']]);
}
function labelledCell(table, label) {
  for (let r = 0; r < table.getNumRows(); r++) if (table.getRow(r).getCell(0).getText() === label) return table.getRow(r).getCell(1).getText();
  return null;
}
function env(props, tdocIds, collectorTableRows) {
  const loaded = loadCode({ documentProperties: props });
  const s = loaded.sandbox;
  const logs = [];
  s.Logger = { log: (m) => logs.push(String(m)) };
  pinNow(s);
  const body = makeFakeDocumentBody(s);
  if (collectorTableRows) body.insertTable(0, [['Key', 'Value']].concat(collectorTableRows));
  const tables = {};
  (tdocIds || []).forEach((id) => { tables[id] = detailTable(body, id); });
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'DOC' });
  return { loaded, s, tables, logs };
}
function results(e, ids) {
  const out = {};
  ids.forEach((id) => {
    const stored = e.loaded.docProps._store['DISCUSS_' + id];
    out[id] = { processed: stored !== undefined, matched: stored ? Object.keys(JSON.parse(stored)).length : 0, cell: labelledCell(e.tables[id], 'E-mail discussion') };
  });
  return out;
}
const matched = (r, ids) => ids.map((id) => r[id].matched);

/** Real fetch + real checkRSSFeed_(); messages served on `onList`. */
function collect(ids, subjects, opts) {
  const o = opts || {};
  const e = env(o.props || PROPS_6G, ids, o.table);
  const msgs = subjects.map((t) => (typeof t === 'string' ? { title: t } : t));
  const fetched = installFakeEtsi(e.s, o.onList || LIST_6G, o.a1Only ? [] : msgs, o.rssOnly ? [] : (o.a1 || []));
  e.s.checkRSSFeed_(e.s.getCollectorConfig_());
  return Object.assign(results(e, ids), { _fetched: fetched, _e: e });
}

const COLLIDE = ['S4aP260091', 'S4aI260091', 'S4aA260091', 'S4-260091'];

// ============================================ exporter -> collector (1-5)

console.log('exporter -> collector round trip (S4aP260091)');
const SUBJ = exportedSubject('S4aP260091', '5.6.1');
// INTENTIONAL DIFFERENCE (decision 2026-10-01, TEMPLATE-002C): the subject has
// no list tag. Legacy: [FS_6G_MED,<agenda item>,<deadline>][<TDoc>]; here:
// [<agenda item>][<deadline>][<TDoc>]. The round trip below uses this form.
check('[template] exporter subject: agenda item and deadline bracketed, no tag', SUBJ, '[5.6.1][26-10-15-1500CEST][S4aP260091] Discussion: ' + TITLE);
[['1 unprefixed', ''], ['2 Re:', 'Re: '], ['3 RE:', 'RE: '], ['4 Fwd:', 'Fwd: '], ['5 Re: Re:', 'Re: Re: '], ['5b RE: Fwd: Re:', 'RE: Fwd: Re: ']].forEach(([name, prefix]) => {
  const r = collect(COLLIDE, [prefix + SUBJ]);
  check(`${name} -> S4aP260091 exactly once, no collision table`, matched(r, COLLIDE), [1, 0, 0, 0]);
});
{
  const r = collect(['S4aP260091'], ['Re: ' + SUBJ]);
  check('E-mail discussion cell written', /^Delegate 0 on 2026-10-01 08:00\n$/.test(r.S4aP260091.cell), true);
}

// ================================================ collisions / precedence

console.log('6-9 collisions, precedence, wrong IDs, multiple IDs');
{
  const r = collect(COLLIDE, ['[FS_6G_MED, 091, 15 Oct 2026 1500 CEST] title without a full ID']);
  check('6 bare short number 091 reaches only the main S4- table, never S4aP/S4aI/S4aA', matched(r, COLLIDE), [0, 0, 0, 1]);
  const each = COLLIDE.map((id) => matched(collect(COLLIDE, ['Re: ' + SUBJ.replace('S4aP260091', id)]), COLLIDE));
  check('6 each explicit ID reaches only its own table', each, [[1, 0, 0, 0], [0, 1, 0, 0], [0, 0, 1, 0], [0, 0, 0, 1]]);
}
{
  const r = collect(['S4aP260091', 'S4-260082'], ['[FS_6G_MED, 082, 5 Oct 2026 1400 CEST][S4aP260091] Discussion: t']);
  check('7 explicit full ID wins; short-number fallback not run', matched(r, ['S4aP260091', 'S4-260082']), [1, 0]);
}
{
  const ids = ['S4aP260091', 'S4aP260999', 'S4-260091'];
  const r = collect(['S4aP260091', 'S4-260091'], ['Re: [FS_6G_MED, 091, 15 Oct 2026 1500 CEST][S4aP260999] Discussion: t']);
  check('8 wrong full ID S4aP260999: no fallback to 091', matched(r, ['S4aP260091', 'S4-260091']), [0, 0]);
  const r2 = collect(ids, ['Re: [FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260999] Discussion: t']);
  check('8 ... it reaches S4aP260999 only when that table exists', matched(r2, ids), [0, 1, 0]);
  const r3 = collect(['S4aP260091'], ['Re: [FS_6G_MED,5.6.1,26-10-15-1500CEST] Discussion: no identifier']);
  check('8 missing full ID: ad-hoc table not associated', matched(r3, ['S4aP260091']), [0]);
  const r4 = collect(['S4aP260091'], ['Re: [FS_6G_MED][XS4aP260091][S4aP2600911][S4al260091] t']);
  check('8 look-alikes / embedded IDs rejected', matched(r4, ['S4aP260091']), [0]);
}
{
  const ids = ['S4aP260091', 'S4aI260082', 'S4aP260069'];
  const r = collect(ids, ['Re: [FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260091] Discussion: merged with S4aI260082']);
  check('9 multiple explicit IDs reach exactly those tables', matched(r, ids), [1, 1, 0]);
  const rev = collect(['S4aP260091'], ['Re: S4aP260091r2 updated']);
  check('9 revision suffix maps to the base document', matched(rev, ['S4aP260091']), [1]);
}

console.log('full-ID scanner (ported from CENTRAL ADDON-008A2, same cases)');
{
  const { sandbox: s } = loadCode();
  const ids = (t) => s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_(t));
  [
    ['exporter subject', 'Fwd: Re: ' + SUBJ, ['S4aP260091']],
    ['S4aI outside brackets', '[ AMD_ARCH_Ph2-MED] S4aI260082 26501-CR0124-B "Implementing stage-2 conclusions on slicing"', ['S4aI260082']],
    ['S4aA', '[ULBC-MED] S4aA260090 agenda', ['S4aA260090']],
    ['S4aV', 'Comments on S4aV260012', ['S4aV260012']],
    ['A4aR', '[DCE] A4aR260087 draft', ['A4aR260087']],
    ['main S4-', '[FS_6G_MED; S4-261483; 25 Aug 2026 1400 CEST] Some title', ['S4-261483']],
    ['revision suffix -> base document', 'S4aI260082r1 updated', ['S4aI260082']],
    ['several, in order of appearance', 'S4aI260089 and S4aI260082 (was S4aA260090)', ['S4aI260089', 'S4aI260082', 'S4aA260090']],
    ['repeated once, canonical spelling', 'S4aI260082 / S4aI260082r2 / s4ai260082 / S4AP260091', ['S4aI260082', 'S4aP260091']],
    ['unknown family rejected', 'S4aX260001 and S4bI260082', []],
    ['look-alike not guessed (lowercase l)', 'S4al260064_BBC.docx', []],
    ['not inside a longer word', 'XS4aI260082 and S4aI2600821 and S4aI260082abc', []],
    ['no false positive from a CR number or a short number', '[FS_6G_MED, 091, 15 Oct 2026] 26501-CR0124-B', []],
    ['empty', '', []]
  ].forEach(([name, text, expected]) => check(name, ids(text), expected));
}

// ============================================================= main (10)

console.log('10 main-meeting fallback unchanged');
{
  // INTENTIONAL DIFFERENCE (decision 2026-10-01, Code.js 2.17.2): the 6G report
  // family derives its own list, 3GPP_TSG_SA4_FS_6G_MED. Legacy derives the
  // general list 3GPP_TSG_SA_WG4 for 6G and relies on a saved Mailing List.
  // MAIN is a main 6G report, so it reads the 6G family list.
  const MAIN_LIST = LIST_6G;
  const run = (subj, ids) => matched(collect(ids, [subj], { props: MAIN, onList: MAIN_LIST }), ids);
  check('short number -> S4-261483 only', run('[FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Some title', ['S4-261483', 'S4-261484']), [1, 0]);
  check('Re: short number', run('Re: [FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Some title', ['S4-261483']), [1]);
  check('full S4- in the first bracket', run('[FS_6G_MED; S4-261483; 25 Aug 2026 1400 CEST] Some title', ['S4-261483', 'S4-261484']), [1, 0]);
  check('hyphenless S4261483 still via the legacy parser', run('[FS_6G_MED; S4261483; 25 Aug 2026 1400 CEST] Some title', ['S4-261483']), [1]);
  check('three-digit short number (S4-260099 -> 099)', run('[FS_6G_MED, 099, 25 Aug 2026 1400 CEST] t', ['S4-260099']), [1]);
  check('exporter format with S4- now round-trips too', run('Re: [FS_6G_MED,5.6.1,26-10-15-1500CEST][S4-261483] Discussion: t', ['S4-261483', 'S4-261484']), [1, 0]);
  const r = collect(['S4-261483'], [], { props: MAIN, onList: MAIN_LIST });
  check('[template] a main 6G report reads its family list, the 6G list (RSS + A1)', listsOf(r._fetched), [MAIN_LIST]);
  const bad = collect(['S4-261483', 'not-a-tdoc'], ['[FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] t'], { props: MAIN, onList: MAIN_LIST });
  check('a malformed table id is still skipped', bad['not-a-tdoc'].processed, false);
}

// ======================================================== sources (11-17)

console.log('11-12 RSS and A1 both read the resolved 6G list');
{
  const r = collect(['S4aP260091'], ['Re: ' + SUBJ]);
  const rss = r._fetched.filter((u) => u.indexOf('?RSS&') !== -1);
  const a1 = r._fetched.filter((u) => u.indexOf('?A1=') !== -1);
  check('11 RSS URL', rss, ['https://list.etsi.org/scripts/wa.exe?RSS&L=3GPP_TSG_SA4_FS_6G_MED&v=2.0&LIMIT=2000']);
  check('12 A1 URLs: Sep+Oct weeks A-E, all on the 6G list', [a1.length, a1.every((u) => /&L=3GPP_TSG_SA4_FS_6G_MED$/.test(u))], [10, true]);
  check('12 A1 first/last week page', [a1[0], a1[9]].map((u) => u.replace('https://list.etsi.org/scripts/wa.exe?', '')), ['A1=ind2609A&L=3GPP_TSG_SA4_FS_6G_MED', 'A1=ind2610E&L=3GPP_TSG_SA4_FS_6G_MED']);
  check('no other list fetched', listsOf(r._fetched), [LIST_6G]);
  const a1only = collect(['S4aP260091'], [], { a1: [{ title: 'Re: ' + SUBJ }] });
  check('12 a reply present only in the A1 archive is associated', matched(a1only, ['S4aP260091']), [1]);
}

console.log('13-17 mailing-list resolution');
{
  const cfgOf = (props, table) => { const e = env(props, [], table); return e.s.getCollectorConfig_(); };
  const pair = (c) => [c.LIST_NAME, c.RSS_URL_V2];
  const rssFor = (l) => 'https://list.etsi.org/scripts/wa.exe?RSS&L=' + l + '&v=2.0&LIMIT=2000';
  check('13 plain name', pair(cfgOf(PROPS_6G)), [LIST_6G, rssFor(LIST_6G)]);
  check('14 @list.etsi.org form (lower case) normalized', pair(cfgOf(props6G({ MAILING_LIST: '3gpp_tsg_sa4_fs_6g_med@list.etsi.org' }))), ['3gpp_tsg_sa4_fs_6g_med', rssFor('3gpp_tsg_sa4_fs_6g_med')]);
  check('14 @LIST.ETSI.ORG upper case normalized', cfgOf(props6G({ MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED@LIST.ETSI.ORG' })).LIST_NAME, LIST_6G);
  ['3GPP TSG SA4 FS 6G MED', 'fs6g@example.com', 'https://list.etsi.org/x', 'A&L=OTHER', '3GPP_TSG_SA4_FS_6G_MED\r\nBcc: x', '@list.etsi.org'].forEach((bad) => {
    const e = env(props6G({ MAILING_LIST: bad }), []);
    const c = e.s.getCollectorConfig_();
    // [template] the family list a malformed value falls back to is the 6G list (see the note in section 10).
    check(`15 malformed ${JSON.stringify(bad)} -> family list, logged`, [c.LIST_NAME, c.RSS_URL_V2 === rssFor(LIST_6G), e.logs.some((l) => l.indexOf('Ignoring configured mailing list') === 0 && !/[\r\n]/.test(l))], [LIST_6G, true, true]);
  });
  const unset = Object.assign({}, PROPS_6G); delete unset.MAILING_LIST;
  check('[template] 16 unset (ad-hoc 6G) -> family list, now the 6G list', pair(cfgOf(unset)), [LIST_6G, rssFor(LIST_6G)]);
  check('16 unset (ad-hoc MBS) -> MBS family list', cfgOf(Object.assign({}, unset, { REPORT_SUFFIX: 'MBS' })).LIST_NAME, '3GPP_TSG_SA_WG4_MBS');
  check('16 main meeting keeps the family list even with MAILING_LIST set', cfgOf({ REPORT_SUFFIX: 'Audio', MAILING_LIST: LIST_6G }).LIST_NAME, '3GPP_TSG_SA_WG4_AUDIO');

  check('17 table LIST_NAME override wins; RSS follows it', pair(cfgOf(PROPS_6G, [['LIST_NAME', 'SOME_TABLE_LIST']])), ['SOME_TABLE_LIST', rssFor('SOME_TABLE_LIST')]);
  check('17 table RSS_URL_V2 override wins; A1 list follows its L=', pair(cfgOf(PROPS_6G, [['RSS_URL_V2', rssFor('OTHER_LIST').replace('LIMIT=2000', 'LIMIT=50')]])), ['OTHER_LIST', rssFor('OTHER_LIST').replace('LIMIT=2000', 'LIMIT=50')]);
  check('17 both set in the table: both kept as given', pair(cfgOf(PROPS_6G, [['LIST_NAME', 'L_ONE'], ['RSS_URL_V2', 'https://example.invalid/rss']])), ['L_ONE', 'https://example.invalid/rss']);
  // The snapshot "Create Configuration Tables" writes (populateConfigTable_):
  // derived defaults, not an override -- the meeting's MAILING_LIST wins.
  // [template] The snapshot holds the family defaults, and the 6G family default
  // is now the 6G list. To keep the point of this check (a snapshot must not
  // mask the meeting's own Mailing List) the meeting's list is a different one.
  const OTHER_LIST = '3GPP_TSG_SA4_SOME_OTHER_WI';
  const snapshot = [['DEBUG', 'true'], ['LIST_NAME', LIST_6G], ['RSS_URL_V2', rssFor(LIST_6G)],
    ['RSS_URL_V1', rssFor(LIST_6G).replace('v=2.0', 'v=1.0')], ['TDOC_ID_REGEX', '^S4-\\d{6}$']];
  check('17 generated snapshot table (family defaults) does not mask MAILING_LIST', pair(cfgOf(props6G({ MAILING_LIST: OTHER_LIST }), snapshot)), [OTHER_LIST, rssFor(OTHER_LIST)]);
  // A snapshot written BEFORE 2.17.2 holds the old 6G default (the general
  // list). It now differs from the family default, so it counts as what it
  // looks like: an explicit table override. Stated here so it is not a surprise.
  const oldSnapshot = [['LIST_NAME', '3GPP_TSG_SA_WG4'], ['RSS_URL_V2', rssFor('3GPP_TSG_SA_WG4')]];
  check('[template] 17 a pre-2.17.2 snapshot (general list) in a 6G document acts as an explicit table override',
    pair(cfgOf(PROPS_6G, oldSnapshot)), ['3GPP_TSG_SA_WG4', rssFor('3GPP_TSG_SA_WG4')]);
  const snap = collect(['S4aP260091'], ['Re: ' + SUBJ], { table: snapshot });
  check('17 ... and its TDOC_ID_REGEX row no longer gates S4aP tables', matched(snap, ['S4aP260091']), [1]);
  const sameAsFamily = cfgOf(Object.assign({}, unset), [['LIST_NAME', LIST_6G]]);
  check('17 snapshot with no MAILING_LIST -> family list (unchanged)', sameAsFamily.LIST_NAME, LIST_6G);
}

// ============================================================ dates (18)

console.log('18 EMAIL_START_DATE and archive window');
{
  const d = (iso) => new Date(iso).toUTCString();
  const subj = (n) => ({ title: 'Re: ' + SUBJ, author: 'A' + n });
  const run = (props, msgs) => collect(['S4aP260091'], msgs, { props });
  check('default 2026-08-21: 08-20 excluded, 08-21 and later included',
    matched(run(PROPS_6G, [Object.assign(subj(1), { date: d('2026-08-20T23:00:00Z') }), Object.assign(subj(2), { date: d('2026-08-21T00:00:00Z') }), Object.assign(subj(3), { date: MSG_DATE })]), ['S4aP260091']), [2]);
  const later = props6G({ EMAIL_START_DATE: '2026-09-15' });
  check('saved EMAIL_START_DATE 2026-09-15 respected', env(later, []).s.getCollectorConfig_().EMAIL_START_DATE, '2026-09-15');
  check('saved start date: 09-10 excluded, 09-16 included',
    matched(run(later, [Object.assign(subj(1), { date: d('2026-09-10T08:00:00Z') }), Object.assign(subj(2), { date: d('2026-09-16T08:00:00Z') })]), ['S4aP260091']), [1]);
  const unsetDate = Object.assign({}, PROPS_6G); delete unsetDate.EMAIL_START_DATE;
  check('unset -> 2026-08-21', env(unsetDate, []).s.getCollectorConfig_().EMAIL_START_DATE, '2026-08-21');
  const badDate = env(props6G({ EMAIL_START_DATE: '15/09/2026' }), []);
  check('malformed -> 2026-08-21, logged', [badDate.s.getCollectorConfig_().EMAIL_START_DATE, badDate.logs.some((l) => l.indexOf('Ignoring EMAIL_START_DATE') === 0)], ['2026-08-21', true]);
  check('Collector Config table EMAIL_START_DATE still wins', env(later, [], [['EMAIL_START_DATE', '2026-10-01']]).s.getCollectorConfig_().EMAIL_START_DATE, '2026-10-01');
  check('unparsable message date still passes (unchanged)', matched(run(PROPS_6G, [Object.assign(subj(1), { date: 'not a date' })]), ['S4aP260091']), [1]);
  const { sandbox: s } = loadCode();
  pinNow(s);
  check('archive look-back unchanged: 14 days from 2026-10-01 -> Sep+Oct (10 pages); 14 days from mid-month -> 5',
    [s.buildArchiveIndexUrlsByDaysBack_(LIST_6G, 14, {}).length, (() => { const t = loadCode().sandbox; t.Date = class extends Date { constructor(...a) { if (a.length === 0) super('2026-10-20T09:00:00Z'); else super(...a); } }; return t.buildArchiveIndexUrlsByDaysBack_(LIST_6G, 14, {}).length; })()],
    [10, 5]);
  check('archive with no list falls back to the SA4 default list', /&L=3GPP_TSG_SA_WG4$/.test(s.buildArchiveIndexUrlsByDaysBack_('', 14, {})[0]), true);
}

// ======================================================= duplicates (19)

console.log('19 the same message via RSS and A1 is stored and shown once');
{
  const msg = { title: 'Re: ' + SUBJ, author: 'Jane Delegate', date: 'Thu, 01 Oct 2026 08:00:00 GMT', a1Date: '2026-10-01T08:00:00Z' };
  const r = collect(['S4aP260091'], [msg], { a1: [msg] });
  check('fetched from both sources', r._fetched.some((u) => u.indexOf('?RSS&') !== -1) && r._fetched.some((u) => u.indexOf('?A1=') !== -1), true);
  check('one stored message, one line', [r.S4aP260091.matched, r.S4aP260091.cell], [1, 'Jane Delegate on 2026-10-01 08:00\n']);
  // A second run over the same document/properties adds nothing.
  installFakeEtsi(r._e.s, LIST_6G, [msg], [msg]);
  r._e.s.checkRSSFeed_(r._e.s.getCollectorConfig_());
  const again = results(r._e, ['S4aP260091']).S4aP260091;
  check('second run: still one message, one line', [again.matched, again.cell], [1, 'Jane Delegate on 2026-10-01 08:00\n']);
}

// ================================================== continuousUpdate (20)

console.log('20 continuousUpdate() -> collectorUpdate_() -> corrected collector');
{
  const e = env(PROPS_6G, ['S4aP260091', 'S4aI260091']);
  const s = e.s;
  // Stages before/after the collector that need the network, the TDoc
  // list or the real document are stubbed; collectorUpdate_() and
  // everything under checkRSSFeed_() run for real.
  s.assertMeetingReadyToBuild_ = () => {};
  s.downloadAndGroupTdocs_ = () => ({});
  s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
  s.updateRevisions_ = () => {};
  const fetched = installFakeEtsi(s, LIST_6G, [{ title: 'Re: ' + SUBJ }]);
  s.continuousUpdate();
  const r = results(e, ['S4aP260091', 'S4aI260091']);
  check('no error logged by continuousUpdateCore_', e.logs.filter((l) => /^ERROR|checkRSSFeed_ failed/.test(l)), []);
  check('6G list fetched', listsOf(fetched), [LIST_6G]);
  check('reply associated with S4aP260091 only', matched(r, ['S4aP260091', 'S4aI260091']), [1, 0]);
  check('trigger handler is still continuousUpdate', /newTrigger\('continuousUpdate'\)/.test(s.createContinuousTrigger.toString()), true);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
