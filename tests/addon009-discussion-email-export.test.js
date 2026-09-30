/**
 * ADDON-009 -- discussion e-mail exporter in CENTRAL.
 *
 * The Legacy exporter's own checks are ported in
 * tests/addon009-legacy-exporter-port.test.js. This file covers what is
 * CENTRAL-specific: the meeting-context mailing list (ad-hoc override vs.
 * report-family list, main meetings unchanged), the subject tag derived
 * from it, every registered SA4 family, the real MBS (86172, S4aI260082) and
 * Legacy 6G (S4aP260069) subjects, the round trip of a generated subject
 * through the ADDON-008A2 collector, and the refuse-before-writing
 * guarantee of Generate.
 *
 * Only DriveApp/DocumentApp/Session are faked; the list fetch of the
 * collector round trip is replaced by local messages. No network.
 *
 * Run: node tests/addon009-discussion-email-export.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
const { installExportStubs, blobText } = require('./helpers/email-export-stubs.js');
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

// ------------------------------------------------------------ fixtures

const SENDER = 'reporter@example.com';
const DEADLINE = { date: '2026-10-15', time: '15:00', tz: 'CEST' };

// Meeting 86172, "3GPPSA4-e (AH) MBS SWG post 137-e": no saved Mailing
// List, so the MBS report-family list applies.
const MBS_86172 = {
  MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', REPORT_SUFFIX: 'MBS', MEETING_ID: '86172',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Inbox/Drafts/',
  DISCUSSION_EMAIL_SENDER: SENDER
};

// The Legacy FS_6G_MED ad-hoc document.
const FS_6G_MED = {
  MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED',
  REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_FS_6G_MED/Inbox/Drafts/',
  DISCUSSION_EMAIL_SENDER: SENDER
};

const MBS_TITLE = '26501-CR0124-B "Implementing stage-2 conclusions on slicing"';
const MBS_SUBJECT = '[MBS,2.5,26-10-15-1500CEST][S4aI260082] Discussion: ' + MBS_TITLE;
const FS6G_TITLE = '[FS_6G_MED] pCR on Editorial Updates to TR 26.870';
const FS6G_SUBJECT = '[FS_6G_MED,5.4,26-10-15-1500CEST][S4aP260069] Discussion: ' + FS6G_TITLE;

// A report TDoc table with plain-text cells (label/value rows).
function table(rows) {
  const cell = (text) => {
    const lines = String(text).split('\n').map((line) => {
      const t = { getText: () => line, getTextAttributeIndices: () => [0], isBold: () => false, isItalic: () => false, isUnderline: () => false, getLinkUrl: () => null };
      const p = { getType: () => 'PARAGRAPH', asParagraph: () => p, editAsText: () => t, getText: () => line };
      return p;
    });
    return { getText: () => String(text), getNumChildren: () => lines.length, getChild: (i) => lines[i] };
  };
  const trows = rows.map((r) => { const cells = r.map(cell); return { getNumCells: () => cells.length, getCell: (i) => cells[i] }; });
  return { getType: () => 'TABLE', getNumRows: () => trows.length, getRow: (i) => trows[i], getCell: (r, c) => trows[r].getCell(c) };
}

function tdocTable(id, title, agendaItem, status) {
  return table([['TDoc', id], ['Title', title], ['Source', 'Qualcomm'], ['Agenda Item', agendaItem], ['Type/For', 'CR / Agreement'],
    ['E-mail Discussion', 'No e-mail discussion.'], ['Revisions', ''], ['Minutes', ''], ['Disposition', ''], ['Status', status === undefined ? 'Revised' : status]]);
}

/**
 * Loads CENTRAL with `props`, puts `tables` in the document, records every
 * Drive call, and returns the helpers a check needs.
 */
function setup(props, tables) {
  const loaded = loadCode({ documentProperties: props });
  const s = installExportStubs(loaded.sandbox);
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => tables }), getId: () => 'FAKE_DOC' });
  const drive = { calls: [], files: [] };
  const folder = { createFile: (blob) => { drive.calls.push('createFile'); drive.files.push(blob); return { getUrl: () => 'https://drive.example.invalid/' + drive.files.length }; } };
  s.DriveApp = {
    getFoldersByName: (name) => { drive.calls.push('getFoldersByName:' + name); return { hasNext: () => false, next: () => null }; },
    createFolder: (name) => { drive.calls.push('createFolder:' + name); return folder; }
  };
  const run = (selections) => s.generateTdocDiscussionEmails(selections, null);
  const emls = () => drive.files.filter((b) => /\.eml$/.test(b.getName())).map((b) => ({ name: b.getName(), text: blobText(b) }));
  const zips = () => drive.files.filter((b) => /\.zip$/.test(b.getName()));
  return { s, docProps: loaded.docProps, drive, run, emls, zips };
}

const sel = (tableIndex, deadline) => ({ tableIndex: tableIndex, deadline: deadline || DEADLINE });
const header = (eml, name) => { const m = new RegExp('^' + name + ': (.*)$', 'm').exec(eml.split('\r\n\r\n')[0]); return m ? m[1] : null; };

function dialogHtml(props, tables) {
  const env = setup(props, tables);
  let captured = null;
  env.s.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  env.s.HtmlService = { createHtmlOutput: (h) => { captured = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  env.s.prepareTdocDiscussionEmails();
  return captured;
}

// ================================================ C. MBS regression ====

console.log('C/I. real MBS case (86172): S4aI260082, agenda item 2.5');
{
  const env = setup(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  const result = env.run([sel(0)]);
  check('export succeeds', result.ok, true);
  const eml = env.emls()[0];
  check('exact subject', header(eml.text, 'Subject'), MBS_SUBJECT);
  check('subject tag is the MBS list name without its SA4 prefix', env.s.deriveEmailExportListTag_('3gpp_tsg_sa_wg4_mbs@list.etsi.org', 'MBS'), 'MBS');
  check('To is the MBS reflector (report-family list, no saved override)', header(eml.text, 'To'), '3gpp_tsg_sa_wg4_mbs@list.etsi.org');
  check('From is the configured sender', header(eml.text, 'From'), SENDER);
  check('file name', eml.name, 'MBS_S4aI260082.eml');
  check('no space inside the first bracket', /^\[[^\] ]+\]/.test(header(eml.text, 'Subject')), true);
}

{
  // A non-canonical spelling in the report table is exported canonically.
  const env = setup(MBS_86172, [tdocTable('s4ai260082', MBS_TITLE, '2.5')]);
  env.run([sel(0)]);
  check('lower-case table value -> canonical S4aI260082 in subject and file name',
    [header(env.emls()[0].text, 'Subject'), env.emls()[0].name], [MBS_SUBJECT, 'MBS_S4aI260082.eml']);
}

// ======================================= D. Legacy 6G regression ====

console.log('D/I. Legacy example: S4aP260069, agenda item 5.4, FS_6G_MED');
{
  const env = setup(FS_6G_MED, [tdocTable('S4aP260069', FS6G_TITLE, '5.4')]);
  const result = env.run([sel(0)]);
  check('export succeeds', result.ok, true);
  const eml = env.emls()[0];
  check('exact Legacy subject convention', header(eml.text, 'Subject'), FS6G_SUBJECT);
  check('To is the saved ad-hoc Mailing List reflector', header(eml.text, 'To'), '3gpp_tsg_sa4_fs_6g_med@list.etsi.org');
  check('file name matches Legacy', eml.name, 'FS_6G_MED_S4aP260069.eml');
  check('ZIP name matches Legacy', /^FS_6G_MED_Discussion_Emails_\d{4}-\d{2}-\d{2}_\d{4}\.zip$/.test(result.zip.fileName), true);
  check('default introduction keeps the Legacy wording',
    blobText(env.drive.files[0]).replace(/=\r\n/g, '').indexOf('As discussed during the FS_6G_MED AHG') !== -1, true);
}

{
  const env = setup(FS_6G_MED, [tdocTable('S4aP260068', 'T', '5.10')]);
  env.run([sel(0)]);
  check('agenda item "5.10" is kept textually (never 5.1)', header(env.emls()[0].text, 'Subject'), '[FS_6G_MED,5.10,26-10-15-1500CEST][S4aP260068] Discussion: T');
}

// ====================================== N. round trip via ADDON-008A2 ====

console.log('N. generated subject -> ADDON-008A2 association');
{
  const { sandbox: s } = loadCode();
  const ids = (subject) => s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_(subject));
  check('MBS subject -> [S4aI260082]', ids(MBS_SUBJECT), ['S4aI260082']);
  check('Re: MBS subject -> [S4aI260082]', ids('Re: ' + MBS_SUBJECT), ['S4aI260082']);
  check('AW: RE: MBS subject -> [S4aI260082]', ids('AW: RE: ' + MBS_SUBJECT), ['S4aI260082']);
  check('6G subject -> [S4aP260069]', ids(FS6G_SUBJECT), ['S4aP260069']);
  check('the first bracket (tag, agenda item, deadline) contributes no identifier', ids('[MBS,2.5,26-10-15-1500CEST]'), []);
}

/** Runs the real collector over the given subjects; returns DISCUSS_ counts. */
function collect(props, ids, subjects) {
  const loaded = loadCode({ documentProperties: props });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  ids.forEach((id) => body.insertTable(body._children.length, [['TDoc', id], ['Title', 't'], ['Source', 's'], ['Agenda Item', '2.5'], ['Type/For', 'CR'], ['E-mail discussion', '']]));
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  let list = null;
  s.collectHybridListservMessages_ = (cfg) => {
    list = cfg.LIST_NAME;
    return subjects.map((title, i) => ({ title, author: 'Delegate ' + i, date: 'Thu, 01 Oct 2026 10:00:00 +0000', link: 'https://list.etsi.org/m' + i, messageId: 'm' + i, preview: '' }));
  };
  s.checkRSSFeed_(s.getCollectorConfig_());
  const counts = {};
  ids.forEach((id) => { const st = loaded.docProps._store['DISCUSS_' + id]; counts[id] = st ? Object.keys(JSON.parse(st)).length : 0; });
  return { counts, list };
}

{
  // The subject the exporter just generated, read back by the real collector.
  const env = setup(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  env.run([sel(0)]);
  const generated = header(env.emls()[0].text, 'Subject');
  const r = collect(MBS_86172, ['S4aI260080', 'S4aI260082', 'S4aI260089', 'S4aA260082'], [generated, 'Re: ' + generated]);
  check('checkRSSFeed_: the generated subject and its reply reach only S4aI260082',
    r.counts, { S4aI260080: 0, S4aI260082: 2, S4aI260089: 0, S4aA260082: 0 });
  check('the collector reads the same MBS list the e-mail is sent To', r.list.toLowerCase() + '@list.etsi.org', header(env.emls()[0].text, 'To'));
}

{
  const env = setup(FS_6G_MED, [tdocTable('S4aP260069', FS6G_TITLE, '5.4')]);
  env.run([sel(0)]);
  const r = collect(FS_6G_MED, ['S4aP260069', 'S4aP260068'], [header(env.emls()[0].text, 'Subject')]);
  check('checkRSSFeed_: the generated 6G subject reaches only S4aP260069', r.counts, { S4aP260069: 1, S4aP260068: 0 });
}

// ================================= E/F. every family, main meetings ====

console.log('E. representative S4aA / S4aV / A4aR ad-hoc IDs');
[
  ['Audio', 'S4aA260090', 'AUDIO', '3gpp_tsg_sa_wg4_audio@list.etsi.org'],
  ['Video', 'S4aV260012', 'VIDEO', '3gpp_tsg_sa_wg4_video@list.etsi.org'],
  ['RTC', 'A4aR260005', 'RTC', '3gpp_tsg_sa_wg4_rtc@list.etsi.org']
].forEach(([family, id, tag, to]) => {
  const env = setup({ MEETING_TYPE: 'adhoc', REPORT_SUFFIX: family, REVISIONS_URL: 'https://www.3gpp.org/ftp/x/Inbox/Drafts/', DISCUSSION_EMAIL_SENDER: SENDER }, [tdocTable(id, 'Title ' + id, '3.1')]);
  env.run([sel(0)]);
  const eml = env.emls()[0];
  check(`${family}: subject, To and file name`, [header(eml.text, 'Subject'), header(eml.text, 'To'), eml.name],
    [`[${tag},3.1,26-10-15-1500CEST][${id}] Discussion: Title ${id}`, to, `${tag}_${id}.eml`]);
  check(`${family}: round trip -> ${id}`, env.s.findSA4DocumentIdsInText_(env.s.stripReplyPrefixes_('Re: ' + header(eml.text, 'Subject'))), [id]);
});

console.log('F. main-meeting S4- compatibility (Mailing List semantics unchanged)');
{
  // A main meeting keeps its report-family list; a saved MAILING_LIST is
  // not used for main meetings (unchanged ADDON-008A2 semantics).
  const env = setup({ MEETING_TYPE: 'main', REPORT_SUFFIX: 'Audio', MAILING_LIST: '3GPP_TSG_SA4_SOMETHING_ELSE', DISCUSSION_EMAIL_SENDER: SENDER }, [tdocTable('S4-261234', 'Main title', '7.3')]);
  env.run([sel(0)]);
  const eml = env.emls()[0];
  check('main Audio: family list, not the saved MAILING_LIST', header(eml.text, 'To'), '3gpp_tsg_sa_wg4_audio@list.etsi.org');
  check('main Audio: subject', header(eml.text, 'Subject'), '[AUDIO,7.3,26-10-15-1500CEST][S4-261234] Discussion: Main title');
  check('main Audio: file name', eml.name, 'AUDIO_S4-261234.eml');
  check('main Audio: revision folder is the main-meeting formula',
    eml.text.replace(/=\r\n/g, '').indexOf('href=3D"https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/Audio"') !== -1, true);
  const r = collect({ MEETING_TYPE: 'main', REPORT_SUFFIX: 'Audio' }, ['S4-261234', 'S4-261235'], [header(eml.text, 'Subject')]);
  check('main S4-: the generated subject reaches only S4-261234 through the collector', r.counts, { 'S4-261234': 1, 'S4-261235': 0 });
}

{
  // The general SA4 list names no topic: the report family's SWG name is used.
  const env = setup({ MEETING_TYPE: 'main', REPORT_SUFFIX: '6G', DISCUSSION_EMAIL_SENDER: SENDER }, [tdocTable('S4-260500', 'T', '9.1')]);
  env.run([sel(0)]);
  const eml = env.emls()[0];
  check('main 6G: To the general SA4 reflector', header(eml.text, 'To'), '3gpp_tsg_sa_wg4@list.etsi.org');
  check('main 6G: tag falls back to the SWG name FS_6G_MED', header(eml.text, 'Subject'), '[FS_6G_MED,9.1,26-10-15-1500CEST][S4-260500] Discussion: T');
}

{
  const { sandbox: s } = loadCode();
  check('tag: 3GPP_TSG_SA4_ prefix', s.deriveEmailExportListTag_('3gpp_tsg_sa4_fs_6g_med@list.etsi.org', 'FS_6G_MED'), 'FS_6G_MED');
  check('tag: 3GPP_TSG_SA_WG4_ prefix', s.deriveEmailExportListTag_('3gpp_tsg_sa_wg4_video@list.etsi.org', 'Video'), 'VIDEO');
  check('tag: general list -> SWG name', s.deriveEmailExportListTag_('3gpp_tsg_sa_wg4@list.etsi.org', 'Plenary'), 'PLENARY');
  check('tag: list outside the SA4 naming is used as-is', s.deriveEmailExportListTag_('fs_6g_med@list.etsi.org', 'FS_6G_MED'), 'FS_6G_MED');
  check('tag: nothing usable -> null (never invented)', s.deriveEmailExportListTag_('3gpp_tsg_sa_wg4@list.etsi.org', ''), null);
  let threw = false;
  try { s.buildEmailExportSubject_('S4aI260082', 'T', '2.5', '26-10-15-1500CEST'); } catch (e) { threw = true; }
  check('the subject builder refuses to run without a tag', threw, true);
}

// ============================================ A/B. From/To, lists ====

console.log('A. sender and recipient are independent');
{
  const env = setup(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  env.s.Session.getActiveUser = () => ({ getEmail: () => 'someone.else@gmail.com' });
  env.s.Session.getEffectiveUser = () => ({ getEmail: () => 'someone.else@gmail.com' });
  env.run([Object.assign(sel(0), { from: 'x@evil.invalid', to: 'y@evil.invalid', mailingList: 'z@evil.invalid' })]);
  const eml = env.emls()[0].text;
  check('From = DISCUSSION_EMAIL_SENDER, To = reflector, different', [header(eml, 'From'), header(eml, 'To')], [SENDER, '3gpp_tsg_sa_wg4_mbs@list.etsi.org']);
  check('neither the Google session identity nor client fields appear', /someone\.else|evil\.invalid/.test(eml), false);
}

{
  // The Mailing List never stands in for a missing sender.
  const props = Object.assign({}, FS_6G_MED);
  delete props.DISCUSSION_EMAIL_SENDER;
  const env = setup(props, [tdocTable('S4aP260069', 'T', '5.4')]);
  const result = env.run([sel(0)]);
  check('no sender configured -> refused, no fallback to the list', [result.ok, /Discussion E-mail Sender/.test(result.error)], [false, true]);
}

console.log('B. ad-hoc Mailing List resolution');
{
  const withList = Object.assign({}, MBS_86172, { MAILING_LIST: '3GPP_TSG_SA4_AMD_ARCH' });
  const env = setup(withList, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  env.run([sel(0)]);
  const eml = env.emls()[0];
  check('a saved ad-hoc Mailing List overrides the family list (To, tag, file name)',
    [header(eml.text, 'To'), header(eml.text, 'Subject').slice(0, 10), eml.name],
    ['3gpp_tsg_sa4_amd_arch@list.etsi.org', '[AMD_ARCH,', 'AMD_ARCH_S4aI260082.eml']);
}

{
  const withList = Object.assign({}, MBS_86172, { MAILING_LIST: '3gpp_tsg_sa_wg4_mbs@LIST.ETSI.ORG' });
  const env = setup(withList, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  env.run([sel(0)]);
  check('a saved full reflector address is idempotent (lower-cased, not re-suffixed)',
    [header(env.emls()[0].text, 'To'), header(env.emls()[0].text, 'Subject')], ['3gpp_tsg_sa_wg4_mbs@list.etsi.org', MBS_SUBJECT]);
}

{
  // The collector reads the family list when a saved value is not a list
  // name; outgoing mail refuses instead of silently using another list.
  const withList = Object.assign({}, MBS_86172, { MAILING_LIST: 'mbs list' });
  const env = setup(withList, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  const result = env.run([sel(0)]);
  check('an unusable saved Mailing List is refused, not replaced', [result.ok, /Mailing List/.test(result.error), env.drive.calls], [false, true, []]);
}

// ================================= L/M. failure safety, preflight ====

console.log('L. CR/LF injection');
[
  ['sender', { DISCUSSION_EMAIL_SENDER: SENDER + '\r\nBcc: evil@example.invalid' }],
  ['sender (LF)', { DISCUSSION_EMAIL_SENDER: SENDER + '\nBcc: evil@example.invalid' }],
  ['Mailing List', { MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED\r\nBcc: evil@example.invalid' }],
  ['Mailing List (CR)', { MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED\rBcc: evil@example.invalid' }]
].forEach(([what, override]) => {
  const env = setup(Object.assign({}, FS_6G_MED, override), [tdocTable('S4aP260069', 'T', '5.4')]);
  const result = env.run([sel(0)]);
  check(`${what}: CR/LF refused before any Drive call`, [result.ok, env.drive.calls], [false, []]);
});

{
  // A title can never add a header line: non-ASCII/control text is an
  // RFC 2047 encoded word.
  const env = setup(FS_6G_MED, [tdocTable('S4aP260069', 'Title\r\nBcc: evil@example.invalid', '5.4')]);
  env.run([sel(0)]);
  const head = env.emls()[0].text.split('\r\n\r\n')[0];
  check('a CR/LF in the title does not create a header', /^Bcc:/m.test(head), false);
  check('the subject header is a single encoded word', /^Subject: =\?UTF-8\?B\?[A-Za-z0-9+/=]+\?=$/m.test(head), true);
}

console.log('M. preflight: nothing is written when configuration or a selection is invalid');
{
  const props = Object.assign({}, MBS_86172);
  delete props.DISCUSSION_EMAIL_SENDER;
  const env = setup(props, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  const result = env.run([sel(0)]);
  check('missing sender: refused, Drive never touched', [result.ok, env.drive.calls, result.zip], [false, [], null]);
  check('missing sender: Legacy message', /Discussion E-mail Sender is not configured or is invalid\. Open Configure Meeting Settings/.test(result.error), true);
}

{
  const env = setup(Object.assign({}, MBS_86172, { DISCUSSION_EMAIL_SENDER: 'not-an-address' }), [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  check('invalid sender: refused, Drive never touched', [env.run([sel(0)]).ok, env.drive.calls], [false, []]);
}

{
  // CENTRAL's meeting context always yields a list (family fallback), so a
  // missing list can only come from the context itself.
  const env = setup(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  const real = env.s.getMeetingContext_;
  env.s.getMeetingContext_ = (c) => { const ctx = real(c); ctx.sources.mailingList = ''; return ctx; };
  const result = env.run([sel(0)]);
  check('missing mailing list: refused, Drive never touched', [result.ok, env.drive.calls], [false, []]);
  check('missing mailing list: Legacy message', /^Mailing List is not configured\. Open Configure Meeting Settings/.test(result.error), true);
}

{
  const props = Object.assign({}, MBS_86172);
  delete props.REVISIONS_URL;
  const env = setup(props, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  const result = env.run([sel(0)]);
  check('ad-hoc without a revision folder: refused, Drive never touched', [result.ok, /Revision upload location/.test(result.error), env.drive.calls], [false, true, []]);
}

{
  // One bad selection in a batch: nothing at all is written (Legacy wrote
  // the files before it and then failed).
  const env = setup(MBS_86172, [tdocTable('S4aI260080', 'A', '2.5'), tdocTable('S4aI260082', MBS_TITLE, '2.5'), tdocTable('S4aI260089', 'C', '2.5', 'Agreed')]);
  const result = env.run([sel(0), sel(1), sel(2)]);
  check('an excluded TDoc late in the batch: refused, Drive never touched', [result.ok, result.error, env.drive.calls], [false, 'S4aI260089 is already Agreed and is excluded from discussion e-mail export.', []]);
  const bad = setup(MBS_86172, [tdocTable('S4aI260080', 'A', '2.5'), tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  const r2 = bad.run([sel(0), sel(1, { date: '2026-02-30', time: '15:00', tz: 'CEST' })]);
  check('an invalid deadline late in the batch: refused, Drive never touched', [r2.ok, /S4aI260082: Invalid or missing deadline date/.test(r2.error), bad.drive.calls], [false, true, []]);
}

console.log('M. malformed TDoc identifiers');
['S4aI26008', 'S4aI260082r1', 'S4aX260082', 'S4aI260082 (rev)', ''].forEach((badId) => {
  const env = setup(MBS_86172, [tdocTable(badId, 'T', '2.5')]);
  check(`"${badId}" is not offered in the dialog`, env.s.detectTdocTablesInDocument_({ getTables: () => [tdocTable(badId, 'T', '2.5')] }).length, 0);
  const result = env.run([sel(0)]);
  check(`"${badId}" is refused by Generate, Drive never touched`, [result.ok, env.drive.calls], [false, []]);
});

// ========================================= G. eligibility / status ====

console.log('G. Approved / Agreed / Reserved are never exported');
{
  const statuses = ['Approved', 'agreed', ' AGREED ', 'Reserved', 'reserved', 'Revised', 'Noted', '', 'available', 'Postponed'];
  const tables = statuses.map((st, i) => tdocTable('S4aI2600' + String(50 + i), 'T' + i, '2.5', st));
  const env = setup(MBS_86172, tables);
  const offered = env.s.detectTdocTablesInDocument_({ getTables: () => tables }).filter((t) => !t.excluded).map((t) => t.status);
  check('offered statuses', offered, ['Revised', 'Noted', '', 'available', 'Postponed']);
  statuses.forEach((st, i) => {
    const one = setup(MBS_86172, [tables[i]]);
    const result = one.run([sel(0)]);
    const excluded = ['approved', 'agreed', 'reserved'].indexOf(st.trim().toLowerCase()) !== -1;
    check(`status "${st}": ${excluded ? 'refused, nothing written' : 'exported'}`, [result.ok, one.emls().length], excluded ? [false, 0] : [true, 1]);
  });
}

// ============================================ H. deadline formats ====

console.log('H. deadline formatting');
{
  const { sandbox: s } = loadCode();
  const v = s.validateEmailExportDeadline_({ date: '2026-10-05', time: '09:30', tz: 'CEST' });
  check('subject token YY-MM-DD-HHmmTZ', s.formatEmailExportDeadlineForSubjectToken_(v), '26-10-05-0930CEST');
  check('body form YY-MM-DD HH:mm TZ', s.formatEmailExportDeadline_(v), '26-10-05 09:30 CEST');
  check('invalid calendar date / time / zone are refused',
    [s.validateEmailExportDeadline_({ date: '2026-02-29', time: '15:00', tz: 'CEST' }).valid,
      s.validateEmailExportDeadline_({ date: '2026-10-15', time: '3pm', tz: 'CEST' }).valid,
      s.validateEmailExportDeadline_({ date: '2026-10-15', time: '15:00', tz: 'CET' }).valid], [false, false, false]);
  const env = setup(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  env.run([sel(0, { date: '2026-10-05', time: '09:30', tz: 'CEST' })]);
  const body = env.emls()[0].text.replace(/=\r\n/g, '');
  check('one deadline, both forms: subject token and body sentences',
    [header(env.emls()[0].text, 'Subject').indexOf('[MBS,2.5,26-10-05-0930CEST]') === 0,
      body.indexOf('Please provide your comments by 26-10-05 09:30 CEST.') !== -1,
      body.indexOf('Please upload revisions by 26-10-05 09:30 CEST to:') !== -1], [true, true, true]);
}

// ================================================ J. .eml structure ====

console.log('J. .eml structure');
{
  const env = setup(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  env.run([sel(0)]);
  const text = env.emls()[0].text;
  const head = text.split('\r\n\r\n')[0];
  check('header lines, in order', head.split('\r\n').map((l) => l.split(':')[0]),
    ['MIME-Version', 'X-Unsent', 'From', 'To', 'Subject', 'Content-Type', 'Content-Transfer-Encoding']);
  check('HTML, UTF-8, quoted-printable', [header(text, 'Content-Type'), header(text, 'Content-Transfer-Encoding')], ['text/html; charset=UTF-8', 'quoted-printable']);
  check('no Cc/Bcc/Date header', /^(Cc|Bcc|Date):/m.test(head), false);
  check('CRLF only (no bare LF)', /[^\r]\n/.test(text), false);
  const body = text.slice(head.length + 4).replace(/=\r\n/g, '').replace(/=3D/g, '=');
  check('body: the copied TDoc table, the revision folder link and the closing',
    [body.indexOf('S4aI260082') !== -1, body.indexOf('<a href="' + MBS_86172.REVISIONS_URL + '">Revision upload folder</a>') !== -1,
      body.indexOf('Best regards,<br>Thomas') !== -1, body.indexOf('As discussed during the MBS AHG') !== -1], [true, true, true, true]);
  check('quoted-printable lines stay within 76 characters', text.slice(head.length + 4).split('\r\n').every((l) => l.length <= 76), true);
}

// ========================================== K/O. ZIP, multiple TDocs ====

console.log('K/O. several eligible TDocs -> one .eml each + one ZIP');
{
  const tables = [tdocTable('S4aI260080', 'First', '2.4'), tdocTable('S4aI260082', MBS_TITLE, '2.5'), tdocTable('S4aI260089', 'Third', '3.1', 'Noted')];
  const env = setup(MBS_86172, tables);
  const result = env.run([sel(0), sel(1), sel(2)]);
  check('success, three files', [result.ok, result.files.map((f) => f.fileName)], [true, ['MBS_S4aI260080.eml', 'MBS_S4aI260082.eml', 'MBS_S4aI260089.eml']]);
  check('each subject names its own TDoc', env.emls().map((e) => env.s.findSA4DocumentIdsInText_(header(e.text, 'Subject'))), [['S4aI260080'], ['S4aI260082'], ['S4aI260089']]);
  const zips = env.zips();
  check('exactly one ZIP, holding exactly this run\'s .eml files in order', [zips.length, zips[0]._zipEntryNames], [1, ['MBS_S4aI260080.eml', 'MBS_S4aI260082.eml', 'MBS_S4aI260089.eml']]);
  check('ZIP name', /^MBS_Discussion_Emails_\d{4}-\d{2}-\d{2}_\d{4}\.zip$/.test(result.zip.fileName), true);
  check('written to the "SA4 Report Email Exports" folder, files before the ZIP',
    env.drive.calls, ['getFoldersByName:SA4 Report Email Exports', 'createFolder:SA4 Report Email Exports', 'createFile', 'createFile', 'createFile', 'createFile']);
  check('Document Properties unchanged by the export', Object.keys(env.docProps._store).sort(), Object.keys(MBS_86172).sort());
}

// ============================================== P. nothing eligible ====

console.log('P. zero eligible TDocs');
{
  const tables = [tdocTable('S4aI260080', 'A', '2.5', 'Approved'), tdocTable('S4aI260081', 'B', '2.5', 'Agreed'), tdocTable('S4aI260082', 'C', '2.5', 'Reserved')];
  const html = dialogHtml(MBS_86172, tables);
  check('dialog says nothing is available and disables Generate',
    [/All 3 detected TDoc\(s\) are already Approved\/Agreed or reserved/.test(html), /<button id="genBtn" onclick="generate\(\)" disabled>/.test(html)], [true, true]);
  const env = setup(MBS_86172, tables);
  const result = env.run([]);
  check('Generate with nothing selected: refused, Drive never touched', [result.ok, result.error, env.drive.calls], [false, 'Select at least one TDoc.', []]);
}

// ================================= meeting-neutral defaults (cleanup) ====

console.log('meeting-neutral defaults: introduction and deadline');
{
  const { sandbox: s } = loadCode();
  const intro = s.buildEmailExportDefaultIntroText_('MBS');
  check('default introduction, exact',
    intro,
    'Dear all,\n\n' +
    'As discussed during the MBS AHG, this email starts a technical discussion on the contribution below. ' +
    'The purpose is to collect comments, refine the proposal and, where appropriate, prepare a revision for the upcoming meeting.\n\n' +
    'This discussion is not an email agreement and does not constitute a formal SA4 decision.');
  check('no hard-coded "October meeting" (or any month) in the default introduction',
    /October|January|February|March|April|May|June|July|August|September|November|December/.test(intro + s.buildEmailExportDefaultIntroText_('FS_6G_MED')), false);
  check('the default body (no introduction supplied) is meeting-neutral too',
    /October meeting/.test(s.buildEmailExportHtmlBody_('<table></table>', null, null, '26-10-15 15:00 CEST', null, 'MBS')), false);
}

{
  // A new, unrelated future meeting -- even with a meeting date saved --
  // gets no prefilled deadline, least of all Legacy's 2026-10-15 15:00.
  const future = {
    MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Video', MEETING_ID: '99999', MEETING_NAME: 'SA4-e (AH) Video SWG 2027', MEETING_DATE: 'March 3, 2027',
    REVISIONS_URL: 'https://www.3gpp.org/ftp/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Video/Inbox/Drafts/', DISCUSSION_EMAIL_SENDER: SENDER
  };
  const tables = [tdocTable('S4aV270001', 'A', '2.1'), tdocTable('S4aV270002', 'B', '2.2')];
  const html = dialogHtml(future, tables);
  check('the old 2026-10-15 / 15:00 default appears nowhere in the dialog', /2026-10-15|value="15:00"/.test(html), false);
  check('batch and per-row date/time inputs start empty',
    [/id="batchDate" value=""/.test(html), /id="batchTime" value=""/.test(html),
      (html.match(/class="deadlineDate"[^>]*value=""/g) || []).length, (html.match(/class="deadlineTime"[^>]*value=""/g) || []).length],
    [true, true, 2, 2]);
  check('the batch deadline is marked required', /Batch deadline \(required\):/.test(html), true);
  check('CEST stays the (only) selectable time zone', /<option value="CEST" selected>CEST<\/option>/.test(html), true);

  // What the dialog sends when the user enters nothing: refused, no output.
  const env = setup(future, tables);
  const untouched = env.run([{ tableIndex: 0, deadline: { date: '', time: '', tz: 'CEST' } }, { tableIndex: 1, deadline: { date: '', time: '', tz: 'CEST' } }]);
  check('Generate with the untouched (empty) deadline: refused, Drive never touched',
    [untouched.ok, /^S4aV270001: Invalid or missing deadline date/.test(untouched.error), env.drive.calls], [false, true, []]);
  const noDeadline = setup(future, tables);
  const r = noDeadline.run([{ tableIndex: 0 }]);
  check('Generate without any deadline: refused, Drive never touched', [r.ok, noDeadline.drive.calls], [false, []]);
  const timeOnly = setup(future, tables);
  const r2 = timeOnly.run([{ tableIndex: 0, deadline: { date: '2027-03-01', time: '', tz: 'CEST' } }]);
  check('a date without a time is refused (no default time filled in)', [r2.ok, /deadline time/.test(r2.error), timeOnly.drive.calls], [false, true, []]);
  const entered = setup(future, tables);
  entered.run([{ tableIndex: 0, deadline: { date: '2027-03-01', time: '12:00', tz: 'CEST' } }]);
  check('an entered deadline is used as entered', header(entered.emls()[0].text, 'Subject'), '[VIDEO,2.1,27-03-01-1200CEST][S4aV270001] Discussion: A');
}

// ================================================= menu and dialog ====

console.log('menu and dialog');
{
  const { sandbox: s } = loadCode();
  const log = [];
  const menu = (name) => ({
    addItem: (label, fn) => { log.push([name, label, fn]); return menu(name); },
    addSeparator: () => menu(name),
    addSubMenu: (sub) => { log.push([name, 'sub', sub._name]); return menu(name); },
    addToUi: () => {},
    _name: name
  });
  s.DocumentApp.getUi = () => ({ createMenu: (name) => menu(name) });
  s.onOpen();
  check('"📧 EMAIL EXPORT" -> "Prepare TDoc Discussion E-mails"',
    log.filter((l) => l[0] === '📧 EMAIL EXPORT'), [['📧 EMAIL EXPORT', 'Prepare TDoc Discussion E-mails', 'prepareTdocDiscussionEmails']]);
  check('submenu order: EMAIL EXPORT after FORMATTING, before CENTRAL ADD-ON',
    log.filter((l) => l[1] === 'sub').map((l) => l[2]),
    ['📝 INITIAL SETUP', '🚀 REPORT OPERATIONS', '📋 DOCUMENT MANAGEMENT', '🔧 TOOLS & DIAGNOSTICS', '🎨 FORMATTING & FIXES', '📧 EMAIL EXPORT', '☁️ CENTRAL ADD-ON (hourly)']);
  check('the Generate RPC is callable from the dialog (no trailing underscore)', typeof s.generateTdocDiscussionEmails, 'function');
}

{
  const html = dialogHtml(MBS_86172, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  check('dialog shows the resolved From / To', /From: <b>reporter@example\.com<\/b> &nbsp; To: <b>3gpp_tsg_sa_wg4_mbs@list\.etsi\.org<\/b>/.test(html), true);
  check('dialog introduction names the meeting\'s tag', /As discussed during the MBS AHG/.test(html), true);
  const props = Object.assign({}, MBS_86172);
  delete props.DISCUSSION_EMAIL_SENDER;
  const warn = dialogHtml(props, [tdocTable('S4aI260082', MBS_TITLE, '2.5')]);
  check('an unconfigured sender is shown when the dialog opens', /Discussion E-mail Sender is not configured/.test(warn), true);
}

{
  const { sandbox: s, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS' } });
  s.saveConfigurationSettings({ meetingType: 'adhoc', reportType: 'MBS', discussionEmailSender: '  ' + SENDER + ' ', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('Save stores the trimmed sender and leaves MAILING_LIST unset', [docProps._store.DISCUSSION_EMAIL_SENDER, docProps._store.MAILING_LIST], [SENDER, undefined]);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
