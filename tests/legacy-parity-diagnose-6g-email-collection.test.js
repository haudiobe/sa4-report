/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/diagnose-6g-email-collection.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-diagnose-6g-email-collection.test.js
 */

/**
 * DIAGNOSIS #3 -- why the 6G report (S4aP TDocs) collected no discussion
 * e-mails. Written against e564f70 to record exactly where a message
 * disappeared; the checks marked [BLOCKER n fixed] now assert the behavior
 * of fix/legacy-adhoc-email-collection. The full acceptance matrix is in
 * tests/legacy-adhoc-email-collection.test.js. Unmarked checks still
 * characterize known, deliberately unfixed behavior (deadline token,
 * hourly trigger) and unchanged helpers.
 *
 * Uses the real Legacy code throughout:
 *   exporter   generateTdocDiscussionEmails() -> .eml Subject header
 *   config     getCollectorConfig_() / getReportConfig_()
 *   fetch      collectHybridListservMessages_() -> collectRssItems_() /
 *              buildArchiveIndexUrlsByDaysBack_() / collectA1_()
 *              (only UrlFetchApp and XmlService are faked)
 *   collector  checkRSSFeed_() over a fake document body
 *
 * Run: node tests/diagnose-6g-email-collection.test.js
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
const note = (label, value) => console.log(`       ${label}: ${JSON.stringify(value)}`);

// The 6G report's documented configuration (LEGACY_UPGRADES.md / task brief).
const PROPS_6G = {
  MEETING_TYPE: 'adhoc',
  MEETING_NAME: 'SA4-e (AH) on FS_6G_MED',
  MEETING_ID: '86178',
  REPORT_SUFFIX: '6G',
  MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED',
  EMAIL_START_DATE: '2026-08-21',
  AGENDA_TDOC: 'S4aP260098',
  // Stale leftovers observed live on 86178 (LEGACY_UPGRADES.md).
  MEETING_FOLDER: 'TSGS4_136_Montreal',
  MEETING_NUMBER: '136',
  DISCUSSION_EMAIL_SENDER: 'reporter@example.com',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts'
};
const TITLE = '[FS_6G_MED] Clarification of AI-native Traffic Characteristics in Clause 6.3.3';
const MSG_DATE = 'Thu, 01 Oct 2026 08:00:00 GMT';

// ------------------------------------------------------------ exporter ----

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
/** The Subject header of the .eml the real Legacy Generate RPC produces. */
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
  const head = raw.split('\r\n\r\n')[0].split('\r\n');
  return { ok: r.ok, subject: head.find((h) => h.indexOf('Subject: ') === 0).slice(9), to: head.find((h) => h.indexOf('To: ') === 0).slice(4) };
}

// ---------------------------------------------------------- fake fetch ----

/** Minimal XmlService over RSS 2.0 <item>s (only what collectRssItems_ calls). */
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
const rss = (msgs) => '<rss><channel>' + msgs.map((m, i) =>
  `<item><title>${m.title}</title><author>${m.author || 'Delegate ' + i}</author><pubDate>${m.date || MSG_DATE}</pubDate><link>https://list.etsi.org/m${i}</link><guid>g${i}</guid><description></description></item>`).join('') + '</channel></rss>';

/**
 * Serves `msgs` only from the list named `onList` (RSS feed and A1 week
 * pages); every other list returns an empty feed. Records every URL fetched.
 */
function installFakeEtsi(s, onList, msgs) {
  const fetched = [];
  s.XmlService = fakeXmlService;
  s.Utilities.Charset = { UTF_8: 'UTF-8' };
  s.UrlFetchApp = {
    fetch: (url) => {
      fetched.push(url);
      const hit = url.indexOf('L=' + onList + '&') !== -1 || url.endsWith('L=' + onList);
      const isRss = url.indexOf('?RSS&') !== -1;
      const text = isRss ? rss(hit ? msgs : []) : '<html><table></table></html>';
      return { getResponseCode: () => 200, getContentText: () => text };
    }
  };
  return fetched;
}

// ------------------------------------------------------- fake document ----

function detailTable(body, id) {
  return body.insertTable(body._children.length, [['TDoc', id], ['Title', TITLE], ['Source', 'Qualcomm'], ['Agenda Item', '5.6.1'], ['Type/For', 'pCR'], ['E-mail discussion', '']]);
}
function labelledCell(table, label) {
  for (let r = 0; r < table.getNumRows(); r++) if (table.getRow(r).getCell(0).getText() === label) return table.getRow(r).getCell(1).getText();
  return null;
}
function sandboxWithDoc(props, tdocIds, collectorTableRows) {
  const loaded = loadCode({ documentProperties: props || PROPS_6G });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  if (collectorTableRows) body.insertTable(0, [['Key', 'Value']].concat(collectorTableRows));
  const tables = {};
  (tdocIds || []).forEach((id) => { tables[id] = detailTable(body, id); });
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'DOC' });
  return { loaded, s, body, tables };
}
function results(env, tdocIds) {
  const out = {};
  tdocIds.forEach((id) => {
    const stored = env.loaded.docProps._store['DISCUSS_' + id];
    out[id] = { processed: stored !== undefined, matched: stored ? Object.keys(JSON.parse(stored)).length : 0, cell: labelledCell(env.tables[id], 'E-mail discussion') };
  });
  return out;
}
/** Real checkRSSFeed_(), messages injected past the fetch layer. */
function collectInjected(tdocIds, subjects, props, collectorTableRows) {
  const env = sandboxWithDoc(props, tdocIds, collectorTableRows);
  env.s.collectHybridListservMessages_ = () => subjects.map((title, i) => ({ title, author: 'Delegate ' + i, date: MSG_DATE, link: 'https://list.etsi.org/m' + i, messageId: 'm' + i, preview: '' }));
  env.s.checkRSSFeed_(env.s.getCollectorConfig_());
  return results(env, tdocIds);
}

// ======================================================= 1. exporter =====

console.log('1. exporter: the real Legacy subject');
const EXP = exportedSubject('S4aP260091', '5.6.1');
check('export succeeds', EXP.ok, true);
check('exact Legacy subject', EXP.subject, '[FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260091] Discussion: ' + TITLE);
check('recipient', EXP.to, '3gpp_tsg_sa4_fs_6g_med@list.etsi.org');
const SUBJ = EXP.subject;

// ========================================================= 2. config =====

console.log('2. collector configuration for the 6G report (no Key/Value table in the doc)');
{
  const env = sandboxWithDoc(PROPS_6G, []);
  const cfg = env.s.getCollectorConfig_();
  note('LIST_NAME', cfg.LIST_NAME);
  note('RSS_URL_V2', cfg.RSS_URL_V2);
  check('[BLOCKER 1 fixed] collector list is MAILING_LIST, not the 6G family list', cfg.LIST_NAME, '3GPP_TSG_SA4_FS_6G_MED');
  check('[BLOCKER 1 fixed] RSS URL points at 3GPP_TSG_SA4_FS_6G_MED', cfg.RSS_URL_V2, 'https://list.etsi.org/scripts/wa.exe?RSS&L=3GPP_TSG_SA4_FS_6G_MED&v=2.0&LIMIT=2000');
  check('meeting context knows the configured list (now used by the collector)', env.s.getMeetingContext_().sources.mailingList, '3GPP_TSG_SA4_FS_6G_MED');
  check('TDOC_ID_REGEX default unchanged (no longer the collector gate)', cfg.TDOC_ID_REGEX, '^S4-\\d{6}$');
  check('EMAIL_START_DATE resolves to 2026-08-21', cfg.EMAIL_START_DATE, '2026-08-21');
  check('readCollectorConfigTable_() with no Key/Value table turns DEBUG on', cfg.DEBUG, 'true');

  const later = sandboxWithDoc(Object.assign({}, PROPS_6G, { EMAIL_START_DATE: '2026-09-15' }), []);
  check('[fixed] the saved EMAIL_START_DATE property is read by the collector', later.s.getCollectorConfig_().EMAIL_START_DATE, '2026-09-15');

  const withTable = sandboxWithDoc(PROPS_6G, [], [['LIST_NAME', '3GPP_TSG_SA4_FS_6G_MED']]);
  const tcfg = withTable.s.getCollectorConfig_();
  check('[fixed] a Collector Config LIST_NAME row also sets the RSS URL', tcfg.RSS_URL_V2.indexOf('L=3GPP_TSG_SA4_FS_6G_MED&') !== -1, true);
}

// ====================================================== 3. archive =======

console.log('3. A1 archive URLs (2026-10-01, ARCHIVE_DAYS_BACK 14)');
{
  const { sandbox: s } = loadCode();
  const RealDate = Date;
  s.Date = class extends RealDate { constructor(...a) { if (a.length === 0) super('2026-10-01T08:00:00Z'); else super(...a); } static now() { return new RealDate('2026-10-01T08:00:00Z').getTime(); } };
  const urls = s.buildArchiveIndexUrlsByDaysBack_('3GPP_TSG_SA4_FS_6G_MED', 14, {});
  note('urls', urls.map((u) => u.replace('https://list.etsi.org/scripts/wa.exe?', '')));
  check('Sep + Oct weeks A-E (10 pages) -- window covers October', urls.length, 10);
  check('[BLOCKER 2 fixed] every archive URL uses the given list', urls.every((u) => /&L=3GPP_TSG_SA4_FS_6G_MED$/.test(u)), true);
}

// ============================================== 4. fetch / discovery =====

console.log('4. fetch: a reply exists ONLY on 3GPP_TSG_SA4_FS_6G_MED');
{
  const env = sandboxWithDoc(PROPS_6G, ['S4aP260091']);
  const fetched = installFakeEtsi(env.s, '3GPP_TSG_SA4_FS_6G_MED', [{ title: 'Re: ' + SUBJ }]);
  const msgs = env.s.collectHybridListservMessages_(env.s.getCollectorConfig_());
  const lists = Array.from(new Set(fetched.map((u) => (u.match(/L=([^&]+)/) || [])[1])));
  note('lists fetched', lists);
  check('[BLOCKER 1+2 fixed] only the FS_6G_MED list is fetched', lists, ['3GPP_TSG_SA4_FS_6G_MED']);
  check('[STAGE A fixed] message discovered and passes the date gate', msgs.map((m) => m.title), ['Re: ' + SUBJ]);
  check('deadline token 26-10-15-1500CEST is not understood (secondary: "No deadline set")', msgs[0].deadline, null);
}

// =================================================== 5. subject parse =====

console.log('5. parseEmailSubject_ on the exported subject and replies');
{
  const { sandbox: s } = loadCode();
  [SUBJ, 'Re: ' + SUBJ, 'RE: ' + SUBJ, 'Fwd: ' + SUBJ, 'Re: Re: ' + SUBJ, 'AW: RE: ' + SUBJ]
    .forEach((t) => check(`parseEmailSubject_ (unchanged, now only the fallback) null for "${t.slice(0, 24)}..."`, s.parseEmailSubject_(t), null));
  check('reply prefixes ARE stripped correctly', s.stripReplyPrefixes_('Re: Re: ' + SUBJ), SUBJ);
  check('its full-ID regex cannot read S4aP (S4aP260091 alone -> null)', s.parseEmailSubject_('[x, S4aP260091, y] t'), null);
  check('... but it does read main S4-', s.parseEmailSubject_('[FS_6G_MED; S4-261483; 25 Aug 2026 1400 CEST] t').tdocFull, 'S4-261483');
  // INTENTIONAL DIFFERENCE (CENTRAL ADDON-009): Legacy's unused helper
  // extractTdocFromEmailExportSubject_() was not ported; the association
  // key is read by the collector's own scanner.
  check('[template] the collector scanner recovers the ID; the unused Legacy helper is not ported',
    [s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_('Re: ' + SUBJ)), typeof s.extractTdocFromEmailExportSubject_], [['S4aP260091'], 'undefined']);
}

// ====================================================== 6. table gate =====

console.log('6. table gate in checkRSSFeed_');
{
  const env = sandboxWithDoc(PROPS_6G, ['S4aP260091']);
  const s = env.s;
  check('isTDocTable_ passes (header "TDoc")', s.isTDocTable_(env.tables.S4aP260091), true);
  check('extractTdocId_ unchanged (no longer the collector gate)', s.extractTdocId_('S4aP260091', '^S4-\\d{6}$'), '');
  check('[BLOCKER 3 fixed] collector gate parseExactSA4DocumentId_ accepts S4aP260091', s.parseExactSA4DocumentId_('S4aP260091').raw, 'S4aP260091');
  check('[BLOCKER 4 fixed] full-ID scanner reads the exported subject', s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_('Re: Re: ' + SUBJ)), ['S4aP260091']);
  check('main S4- passes', s.extractTdocId_('S4-261483', '^S4-\\d{6}$'), 'S4-261483');
}

// ======================================== 7. round trip, end to end =====

console.log('7. exporter -> collector round trip (real checkRSSFeed_, messages injected past fetch)');
{
  const r = collectInjected(['S4aP260091'], ['Re: ' + SUBJ]);
  check('[fixed] S4aP260091 table processed', r.S4aP260091.processed, true);
  check('[fixed] reply associated', r.S4aP260091.matched, 1);
}
{
  console.log('7b. a Collector Config TDOC_ID_REGEX row no longer affects collector eligibility');
  const r = collectInjected(['S4aP260091'], ['Re: ' + SUBJ], PROPS_6G, [['TDOC_ID_REGEX', '^S4-\\d{6}$']]);
  check('table processed', r.S4aP260091.processed, true);
  check('[fixed] reply associated', r.S4aP260091.matched, 1);
}
{
  console.log('7c. full real path: fetch + collector, message on FS_6G_MED');
  const env = sandboxWithDoc(PROPS_6G, ['S4aP260091']);
  installFakeEtsi(env.s, '3GPP_TSG_SA4_FS_6G_MED', [{ title: 'Re: ' + SUBJ }]);
  env.s.checkRSSFeed_(env.s.getCollectorConfig_());
  check('[fixed] the reply reaches S4aP260091', results(env, ['S4aP260091']).S4aP260091.matched, 1);
}

// ============================================================ 8. matrix ===

console.log('8. matrix (exporter-format subject per ID; gate forced open to expose later stages)');
{
  const { sandbox: s } = loadCode();
  const IDS = ['S4aP260091', 'S4aP260069', 'S4aI260082', 'S4aA260090', 'S4-261483'];
  const OPEN = [['TDOC_ID_REGEX', '(?:S4-|S4a[APVI]|A4aR)\\d{6}']];
  // INTENTIONAL DIFFERENCE (decision 2026-10-01, CENTRAL ADDON-009): the
  // subject tag is never hard-coded. The builder takes the tag derived from
  // the meeting's mailing list (3GPP_TSG_SA4_FS_6G_MED -> FS_6G_MED) and
  // refuses to build a subject without one.
  check('[template] the subject tag is a parameter derived from the mailing list, never a built-in FS_6G_MED',
    [s.deriveEmailExportListTag_('3gpp_tsg_sa4_fs_6g_med@list.etsi.org', 'FS_6G_MED'),
      (() => { try { s.buildEmailExportSubject_('S4aP260091', 't', '5.6.1', '26-10-15-1500CEST'); return 'built'; } catch (e) { return e.message; } })()],
    ['FS_6G_MED', 'buildEmailExportSubject_: a list tag is required.']);
  const rows = IDS.map((id) => {
    const subj = 'Re: ' + s.buildEmailExportSubject_(id, 't', '5.6.1', '26-10-15-1500CEST', 'FS_6G_MED');
    const gate = !!s.extractTdocId_(id, '^S4-\\d{6}$');
    const parsed = s.parseEmailSubject_(subj);
    const realGate = collectInjected([id], [subj])[id].matched;
    const opened = collectInjected([id], [subj], PROPS_6G, OPEN)[id].matched;
    return { id, gateAccepts: gate, parsed: parsed ? parsed.tdocFull || parsed.tdocShort : null, associatedToday: realGate, associatedGateOpen: opened };
  });
  rows.forEach((r) => console.log('       ' + JSON.stringify(r)));
  check('old TDOC_ID_REGEX would accept only S4- (no longer the collector gate)', rows.map((r) => r.gateAccepts), [false, false, false, false, true]);
  check('legacy parser: no ID for any exporter-format subject (now only the fallback)', rows.map((r) => r.parsed), [null, null, null, null, null]);
  check('[fixed] associated: every family', rows.map((r) => r.associatedToday), [1, 1, 1, 1, 1]);
  check('[fixed] TDOC_ID_REGEX row irrelevant', rows.map((r) => r.associatedGateOpen), [1, 1, 1, 1, 1]);

  // The short-number fallback is now main-only: a legacy-format subject for
  // 091 reaches only the S4- table.
  const coll = collectInjected(['S4aP260091', 'S4aI260091', 'S4-260091'], ['[FS_6G_MED, 091, 15 Oct 2026 1500 CEST] t'], PROPS_6G, OPEN);
  check('[COLLISION fixed] bare short number 091 reaches only the main S4- table', [coll.S4aP260091.matched, coll.S4aI260091.matched, coll['S4-260091'].matched], [0, 0, 1]);
}

// ========================================================== 9. trigger ===

console.log('9. trigger');
{
  const { sandbox: s } = loadCode();
  const calls = [];
  const deleted = [];
  const existing = { getHandlerFunction: () => 'continuousUpdate' };
  s.ScriptApp = {
    getProjectTriggers: () => [existing],
    deleteTrigger: (t) => deleted.push(t.getHandlerFunction()),
    newTrigger: (fn) => ({ timeBased: () => ({
      everyMinutes: (n) => { calls.push('everyMinutes(' + n + ')'); if ([1, 5, 10, 15, 30].indexOf(n) === -1) throw new Error('The value you passed to everyMinutes was invalid.'); return { create: () => calls.push('create ' + fn) }; },
      everyHours: (n) => ({ create: () => calls.push('everyHours(' + n + ') ' + fn) })
    }) })
  };
  // INTENTIONAL DIFFERENCE (CENTRAL 2.15.1 / 2.15.2): Legacy's diagnostic
  // recorded a known Legacy defect here ("Every hour" -> everyMinutes(60)
  // throws after the running trigger was already deleted; "not included" in
  // the Legacy release). sa4-report fixed it: every hour is everyHours(1),
  // and an interval that is not offered is refused before anything changes.
  // 15/30 minutes are not offered by the add-on runtime; the template
  // runtime offers them again in TEMPLATE-002B.
  let err = null;
  try { s.createContinuousTrigger(60, false); } catch (e) { err = e.message; }
  check('[template] "Every hour" works: everyHours(1), handler continuousUpdate', [err, calls], [null, ['everyHours(1) continuousUpdate']]);
  check('[template] the previous trigger is replaced only by a valid request', deleted, ['continuousUpdate']);
  calls.length = 0; deleted.length = 0;
  let err30 = null;
  try { s.createContinuousTrigger(30, false); } catch (e) { err30 = e.message; }
  check('[template] an interval that is not offered is refused before the running trigger is touched',
    [/^Unsupported update interval: 30/.test(err30), calls, deleted], [true, [], []]);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll diagnostic checks passed.');
}
