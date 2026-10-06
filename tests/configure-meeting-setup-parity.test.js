/**
 * Configure Meeting and the setup of a new report give a report the same
 * (after T-2026.10.8).
 *
 * PART 1 -- THE START OF THE MEETING. A report created from the template
 * gets the start time and the time zone of its meeting from the Portal
 * (tests/meeting-start-opening.test.js). A report that is given its meeting
 * in Configure Meeting must get the same: Resolve already has the Portal's
 * answer, and Save now carries the start it found. The values are worked
 * out on the server by the functions the creator uses, from the Portal's
 * own StartDate and StartTimeZone -- no second reading of the Portal, and no
 * further request.
 *
 *   1. the same meeting, selected at creation or in Configure Meeting: the
 *      same stored start and the same opening sentence;
 *   2. no further Portal request; Resolve alone stores nothing;
 *   3. a report changed from meeting A to meeting B never keeps A's start;
 *   4. what the Save does not carry, and what is never guessed.
 *
 * PART 2 -- THE DRAFTS FOLDER, A SECOND CHANCE. The first run asks for the
 * derived drafts folder once (tests/adhoc-revisions-folder-setup.test.js).
 * When that one request fails, a later Save of Configure Meeting asks again
 * -- while the report has no folder, and for the candidate of the meeting it
 * has then. Nothing else asks.
 *
 *   5. the first request fails: the candidate remains;
 *   6. a later Save asks again and stores the folder;
 *   7. an update, a build, opening the dialog and Resolve do not ask;
 *   8. a stored folder is never asked for again and never replaced;
 *   9. a report that was given another meeting gets that meeting's candidate.
 *
 * The Portal answers are the captured ones of tests/fixtures; all else is synthetic.
 *
 * Run: node tests/configure-meeting-setup-parity.test.js
 */

const fs = require('fs');
const path = require('path');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { configurableReport } = require('./helpers/config-dialog.js');
const { createReport, firstRunFor, plain } = require('./helpers/template-creator.js');
const { templateReport } = require('./helpers/template-report.js');

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

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const withoutComments = (src) => src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
const GM = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86172-85916-getmeetings.json'), 'utf8'));
const meeting = (base, changes) => [Object.assign({}, base[0], changes)];
const ROWS_MBS = [{ id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' }];
const ROWS_AUDIO = [{ id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '2' }];
const START = ['MEETING_ID', 'MEETING_DATE', 'MEETING_START_TIME', 'MEETING_TIME_ZONE', 'MEETING_START_BASIS'];
const MBS_DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
const MBS_DRAFTS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/inbox/drafts/';
const AUDIO_DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const AUDIO_DRAFTS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/';
const CANDIDATE_KEY = 'SA4_REVISIONS_URL_CANDIDATE';
const LISTING = { status: 200, text: '<html><head><title>Index of /drafts</title></head><body><a href="x_r1.docx">x_r1.docx</a></body></html>' };
const isDrafts = (url) => /\/inbox\/drafts\/$/i.test(url);
const draftsAsked = (r) => r.requests.filter(isDrafts);
/** The opening sentence a build of this document would write. */
const sentenceOf = (s, docProps) => { const start = s.getMeetingStartForOpening_(); return s.buildMeetingOpeningSentence_(docProps.getProperty('MEETING_DATE'), start.time, start.zone); };

/** The report as "Create New SA4 Report" and its first run leave it. */
function byCreation(getMeetings, rows, id) {
  const created = createReport(getMeetings, rows, id, {});
  const run = firstRunFor(created.description);
  const setup = run.run();
  return { ok: created.created.ok && setup.status === 'configured', stored: START.map((k) => run.docProps.getProperty(k)), sentence: sentenceOf(run.s, run.docProps), props: plain(run.docProps._store) };
}
/** The report as Configure Meeting leaves a document that had no meeting: Resolve, then Save. */
function byConfigureMeeting(getMeetings, rows, id) {
  const r = configurableReport({ meetings: { [id]: getMeetings }, rows: rows });
  const page = r.open().resolve(id).save();
  return { ok: page.alerts.length === 1 && /Configuration saved/.test(page.alerts[0]), stored: r.props(START), sentence: sentenceOf(r.s, r.docProps), r: r, page: page };
}

// =====================================================================
console.log('PART 1 -- the start of the meeting');
console.log('1. the same meeting, at creation or in Configure Meeting');
[
  ['86172 (online, 15:30, UTC+2)', GM.getMeetings86172, ROWS_MBS, '86172', ['86172', 'October 1, 2026', '15:30', 'UTC+2', '86172|October 1, 2026'], '<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.'],
  ['85916 (Shanghai, 09:00, UTC+8)', GM.getMeetings85916, ROWS_AUDIO, '85916', ['85916', 'September 28, 2026', '09:00', 'UTC+8', '85916|September 28, 2026'], '<Chair> opens the session on September 28, 2026 at 09:00 UTC+8.'],
  ['a meeting after the change to winter time (UTC+1)', meeting(GM.getMeetings86172, { StartDate: '2026-11-05 15:30:00', StartTimeZone: '(GMT+01:00) Brussels, Copenhagen, Madrid, Paris' }), ROWS_MBS, '86172',
    ['86172', 'November 5, 2026', '15:30', 'UTC+1', '86172|November 5, 2026'], '<Chair> opens the session on November 5, 2026 at 15:30 UTC+1.'],
  ['a meeting whose time the Portal does not have (00:00:00)', meeting(GM.getMeetings86172, { StartDate: '2026-10-01 00:00:00' }), ROWS_MBS, '86172',
    ['86172', 'October 1, 2026', null, 'UTC+2', '86172|October 1, 2026'], '<Chair> opens the session on October 1, 2026 at <start> UTC+2.'],
  ['a meeting without a time zone', meeting(GM.getMeetings86172, { StartTimeZone: '' }), ROWS_MBS, '86172',
    ['86172', 'October 1, 2026', '15:30', null, '86172|October 1, 2026'], '<Chair> opens the session on October 1, 2026 at 15:30 <time zone>.'],
  ['a meeting with neither', meeting(GM.getMeetings86172, { StartDate: '2026-10-01 00:00:00', StartTimeZone: 'Brussels, Copenhagen, Madrid, Paris' }), ROWS_MBS, '86172',
    ['86172', 'October 1, 2026', null, null, null], '<Chair> opens the session on October 1, 2026 at <start> <time zone>.']
].forEach(([name, getMeetings, rows, id, stored, sentence]) => {
  const creation = byCreation(getMeetings, rows, id);
  const dialog = byConfigureMeeting(getMeetings, rows, id);
  check(name + ': both ways complete', [creation.ok, dialog.ok], [true, true]);
  check(name + ': creation stores this start', creation.stored, stored);
  check(name + ': Configure Meeting stores the same', dialog.stored, creation.stored);
  check(name + ': and a build writes the same opening sentence', [creation.sentence, dialog.sentence], [sentence, sentence]);
});
{
  // The whole way to the document, for the dialog path: the real build on what Configure Meeting stored.
  const dialog = byConfigureMeeting(GM.getMeetings85916, ROWS_AUDIO, '85916');
  const props = {}; START.forEach((k) => { props[k] = dialog.r.docProps.getProperty(k); });
  const report = templateReport({ release: false, token: false, props: props, tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'available' }] });
  report.build();
  check('Build Report from Scratch, after Configure Meeting, writes it into the report',
    report.body._children.filter((c) => c.getType() === 'PARAGRAPH').map((c) => c.getText()).filter((text) => /opens the session/.test(text)), ['<Chair> opens the session on September 28, 2026 at 09:00 UTC+8.']);
}

console.log('2. no further Portal request; Resolve alone stores nothing');
{
  const r = configurableReport({ meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS });
  const page = r.open();
  check('opening the dialog asks the Portal nothing', [r.portalRequests, r.requests], [[], []]);
  page.resolve('86172');
  check('Resolve asks the Portal once, and stores nothing', [r.portalRequests, r.requests, r.props(START)], [['86172'], [], [null, null, null, null, null]]);
  page.save();
  check('Save asks the Portal nothing more: it sends what Resolve got', [r.portalRequests, page.sent[0].resolvedMeetingStart],
    [['86172'], { meetingId: '86172', startDate: '2026-10-01 15:30:00', timeZone: '(GMT+02.00)  Brussels, Copenhagen, Madrid, Paris' }]);
  check('what it sends is the Portal\'s own StartDate and StartTimeZone, not a value worked out in the page', [page.sent[0].meetingStartTime, page.sent[0].meetingTimeZone], [undefined, undefined]);
  const persist = functionSource(CODE, 'persistConfigurationSettings_');
  const fromResolved = functionSource(CODE, 'meetingStartFromResolved_');
  check('the server works them out with the creator\'s two functions, and both ways store through one function (source)',
    [/computeMeetingStartTimeFromStartDate_\(resolved\.startDate\)/.test(fromResolved), /computeMeetingTimeZoneLabel_\(resolved\.timeZone\)/.test(fromResolved), (withoutComments(persist).match(/storeMeetingStart_\(/g) || []).length,
      /computeMeetingStartTimeFromStartDate_\(p\.startDateRaw\)/.test(CREATOR), /computeMeetingTimeZoneLabel_\(p\.startTimeZoneRaw\)/.test(CREATOR)], [true, true, 2, true, true]);
  check('there is one reader of each Portal value in the code: no second parser (source)',
    [(withoutComments(CODE).match(/\nfunction computeMeetingStartTimeFromStartDate_\(/g) || []).length, (withoutComments(CODE).match(/\nfunction computeMeetingTimeZoneLabel_\(/g) || []).length,
      /UrlFetchApp|fetchMeetingMetadataById_|resolveMeeting/.test(withoutComments(fromResolved + functionSource(CODE, 'storeMeetingStart_')))], [1, 1, false]);
  check('the dialog has no field for the chair, and nothing stores one', [/chair/i.test(page.html), /chair/i.test(JSON.stringify(page.sent[0])), r.docProps.getKeys().filter((k) => /CHAIR/i.test(k))], [false, false, []]);
}

console.log('3. a report changed from meeting A to meeting B');
{
  const A = () => { const created = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {}); const run = firstRunFor(created.description); run.run(); return plain(run.docProps._store); };
  const repoint = (meetingB, rowsB, idB, edit) => {
    const r = configurableReport({ props: A(), meetings: { 86172: GM.getMeetings86172, [idB]: meetingB }, rows: rowsB });
    const before = r.props(START);
    const page = r.open();
    if (edit) edit(page, r); else page.resolve(idB).save();
    return { before: before, stored: r.props(START), sentence: sentenceOf(r.s, r.docProps), page: page, r: r };
  };
  const toB = repoint(GM.getMeetings85916, ROWS_AUDIO, '85916');
  check('the report has the start of A', toB.before, ['86172', 'October 1, 2026', '15:30', 'UTC+2', '86172|October 1, 2026']);
  check('Resolve B and Save: the start of B, nothing of A', [toB.stored, toB.sentence], [['85916', 'September 28, 2026', '09:00', 'UTC+8', '85916|September 28, 2026'], '<Chair> opens the session on September 28, 2026 at 09:00 UTC+8.']);

  const toBare = repoint(meeting(GM.getMeetings85916, { StartDate: '2026-09-28 00:00:00', StartTimeZone: '' }), ROWS_AUDIO, '85916');
  check('B has no start time and no time zone: A\'s are removed, and the sentence has the placeholders',
    [toBare.stored, toBare.sentence], [['85916', 'September 28, 2026', null, null, null], '<Chair> opens the session on September 28, 2026 at <start> <time zone>.']);
  const toTimeOnly = repoint(meeting(GM.getMeetings85916, { StartTimeZone: 'Beijing' }), ROWS_AUDIO, '85916');
  check('B has a time and no usable zone: B\'s time, and A\'s zone is not kept', [toTimeOnly.stored, toTimeOnly.sentence],
    [['85916', 'September 28, 2026', '09:00', null, '85916|September 28, 2026'], '<Chair> opens the session on September 28, 2026 at 09:00 <time zone>.']);

  const unknown = repoint([], [], '99999');
  check('B is not known to the Portal: nothing of A is written for it', [unknown.stored[0], unknown.sentence], ['99999', '<Chair> opens the session on October 1, 2026 at <start> <time zone>.']);

  const typed = repoint(GM.getMeetings85916, ROWS_AUDIO, '85916', (page) => { page.fields.meetingId.value = '85916'; page.save(); });
  check('the id of B typed and saved without Resolve: A\'s start does not apply to B (it is stored with the meeting it belongs to)',
    [typed.page.sent[0].resolvedMeetingStart, typed.stored[0], typed.sentence], [undefined, '85916', '<Chair> opens the session on October 1, 2026 at <start> <time zone>.']);

  const stale = repoint(GM.getMeetings85916, ROWS_AUDIO, '85916', (page) => { page.resolve('85916'); page.fields.meetingId.value = '86172'; page.save(); });
  check('B resolved, then the id changed back before Save: the start Resolve found for B is not sent for another id',
    [stale.page.sent[0].resolvedMeetingStart, stale.page.sent[0].meetingId], [undefined, '86172']);

  const backToA = repoint(GM.getMeetings85916, ROWS_AUDIO, '85916', (page) => { page.resolve('85916').save(); page.resolve('86172').save(); });
  check('A -> B -> A: the start of A again, from the Portal', [backToA.stored, backToA.r.portalRequests], [['86172', 'October 1, 2026', '15:30', 'UTC+2', '86172|October 1, 2026'], ['85916', '86172']]);
}

console.log('4. what a Save does not carry, and what is never guessed');
{
  const A = () => { const created = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {}); const run = firstRunFor(created.description); run.run(); return plain(run.docProps._store); };
  const kept = configurableReport({ props: A(), meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS });
  kept.open().save();
  check('a Save without Resolve, for the same meeting, keeps the stored start', kept.props(START), ['86172', 'October 1, 2026', '15:30', 'UTC+2', '86172|October 1, 2026']);

  const dateTyped = configurableReport({ meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS });
  const page = dateTyped.open().resolve('86172');
  page.fields.meetingDate.value = 'October 8, 2026';
  page.save();
  check('the meeting date changed by hand after Resolve: the Portal\'s time is that of another day, so no start is stored',
    [dateTyped.props(START), sentenceOf(dateTyped.s, dateTyped.docProps)], [['86172', 'October 8, 2026', null, null, null], '<Chair> opens the session on October 8, 2026 at <start> <time zone>.']);

  const mainMeeting = meeting(GM.getMeetings86172, { Id: 60778, Type: 'OR', Title: '3GPPSA4#137', MtgDocURL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Synthetic/Docs/' });
  const main = configurableReport({ meetings: { 60778: mainMeeting } });
  const mainPage = main.open().resolve('60778');
  mainPage.fields.reportType.value = 'Video';
  mainPage.save();
  check('a main meeting: no start is stored, as at creation', [main.docProps.getProperty('MEETING_TYPE'), main.props(START).slice(2)], ['main', [null, null, null]]);

  const from = kept.s.meetingStartFromResolved_;
  const RESOLVED = { meetingId: '86172', startDate: '2026-10-01 15:30:00', timeZone: '(GMT+02.00)  Brussels, Copenhagen, Madrid, Paris' };
  check('the rule, case by case',
    [from(RESOLVED, '86172', 'adhoc', 'October 1, 2026'), from(RESOLVED, '85916', 'adhoc', 'October 1, 2026'), from(RESOLVED, '86172', 'main', 'October 1, 2026'), from(RESOLVED, '86172', 'adhoc', 'October 2, 2026'),
      from(RESOLVED, '86172', 'adhoc', ''), from(null, '86172', 'adhoc', 'October 1, 2026'), from('86172', '86172', 'adhoc', 'October 1, 2026'), from({ meetingId: '', startDate: RESOLVED.startDate }, '', 'adhoc', 'October 1, 2026'),
      from({ meetingId: '86172', startDate: '15:30', timeZone: 'CEST' }, '86172', 'adhoc', 'October 1, 2026'), from({ meetingId: 86172, startDate: RESOLVED.startDate, timeZone: 'CEST' }, '86172', 'ADHOC', 'October 1, 2026')],
    [{ time: '15:30', zone: 'UTC+2' }, null, null, { time: '', zone: '' }, { time: '', zone: '' }, null, null, null, { time: '', zone: '' }, { time: '15:30', zone: '' }]);
  // What the page sends is data from a browser: only a time and an offset in their forms can come of it.
  const hostile = configurableReport({ props: { MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_DATE: 'October 1, 2026' } });
  hostile.s.persistConfigurationSettings_({ meetingId: '86172', meetingType: 'adhoc', meetingDate: 'October 1, 2026', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'keep',
    resolvedMeetingStart: { meetingId: '86172', startDate: '2026-10-01 15:30:00<script>', timeZone: '(GMT+02:00)</td><script>alert(1)</script>' } });
  check('a start that is not in the Portal\'s form stores no time; a zone is read as its offset only', hostile.props(START).slice(2), [null, 'UTC+2', '86172|October 1, 2026']);
}

// =====================================================================
console.log('PART 2 -- the drafts folder: a second chance');
/** A new report for 86172 whose first run could not confirm the drafts folder (the server answered 403). */
function afterFailedSetup() {
  const created = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  const run = firstRunFor(created.description);
  const asked = [];
  run.s.UrlFetchApp = { fetch: (url) => { asked.push(url); return { getResponseCode: () => 403, getContentText: () => 'Forbidden' }; } };
  run.deps.validateRevisionsCandidate = (url) => run.s.validateRevisionsUrlCandidate_(url);
  const setup = run.run();
  return { props: plain(run.docProps._store), setup: setup, asked: asked };
}

console.log('5. the first request fails: the candidate remains');
{
  const first = afterFailedSetup();
  check('the report is set up, asked once, has no REVISIONS_URL, and keeps the candidate as not confirmed',
    [first.setup.status, first.asked, first.props.REVISIONS_URL, first.props[CANDIDATE_KEY], first.setup.revisionsFolder.status], ['configured', [MBS_DRAFTS], undefined, MBS_DRAFTS, 'unconfirmed']);
}

console.log('6. a later Save of Configure Meeting asks again');
{
  const r = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? LISTING : null) });
  const page = r.open();
  check('the dialog shows no folder (none is stored) and has asked for nothing', [page.fields.revisionsUrl.value, r.requests], ['', []]);
  page.save();
  check('Save asks for the candidate once, and the answer is a listing: it is stored', [draftsAsked(r), r.docProps.getProperty('REVISIONS_URL'), r.docProps.getProperty(CANDIDATE_KEY)], [[MBS_DRAFTS], MBS_DRAFTS, null]);
  check('the message of the Save says so', [page.alerts.length, /Configuration saved\./.test(page.alerts[0]), page.alerts[0].indexOf('✅ Drafts folder found and stored: ' + MBS_DRAFTS) !== -1], [1, true, true]);
  check('the report uses it', r.s.getMeetingContext_().sources.revisionsUrl, MBS_DRAFTS);

  const again = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? { status: 403, text: '' } : null) });
  const againPage = again.open().save();
  check('when it fails again: saved all the same, nothing stored as fact, the candidate kept, and the message says what to do',
    [draftsAsked(again), again.docProps.getProperty('REVISIONS_URL'), again.docProps.getProperty(CANDIDATE_KEY), /Configuration saved\./.test(againPage.alerts[0]), /Drafts folder still to set up\./.test(againPage.alerts[0]), againPage.alerts[0].indexOf(MBS_DRAFTS) !== -1],
    [[MBS_DRAFTS], null, MBS_DRAFTS, true, true, true]);
  again.open().save();
  check('each further Save is one further request -- an explicit action, never a loop', draftsAsked(again).length, 2);

  const down = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: () => null });
  const downPage = down.open().save();
  check('a request that fails outright does not fail the Save', [/Configuration saved\./.test(downPage.alerts[0]), down.docProps.getProperty('MEETING_ID'), down.docProps.getProperty('REVISIONS_URL'), down.docProps.getProperty(CANDIDATE_KEY)], [true, '86172', null, MBS_DRAFTS]);

  // Whatever goes wrong in that step, the configuration is saved and the Save does not fail.
  const broken = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: () => null });
  broken.s.confirmAdhocRevisionsFolderWith_ = () => { throw new Error('synthetic failure of the step'); };
  const brokenPage = broken.open();
  brokenPage.fields.meetingName.value = 'Renamed by hand';
  brokenPage.save();
  check('an error inside the step is logged, not raised: the Save succeeds and what was entered is stored',
    [brokenPage.alerts.length, /^✅ Configuration saved\./.test(brokenPage.alerts[0]), /Error saving/.test(brokenPage.alerts[0]), broken.docProps.getProperty('MEETING_NAME'),
      broken.logs.filter((l) => /^Configure Meeting: the drafts folder could not be looked at: synthetic failure of the step$/.test(l)).length], [1, true, false, 'Renamed by hand', 1]);

  // The user enters the folder by hand: that is what is stored, and nothing is asked.
  const typed = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? { status: 403, text: '' } : null) });
  const typedPage = typed.open();
  typedPage.fields.revisionsUrl.value = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Inbox/Drafts/';
  typedPage.save();
  check('a folder entered by hand is stored as entered; nothing is asked, and the "not confirmed" record goes',
    [typed.docProps.getProperty('REVISIONS_URL'), draftsAsked(typed), typed.docProps.getProperty(CANDIDATE_KEY), /Drafts folder/.test(typedPage.alerts[0])],
    ['https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Inbox/Drafts/', [], null, false]);
}

console.log('7. nothing else asks');
{
  // The same report in the build/update arrangement: no REVISIONS_URL, the candidate recorded.
  const props = { MEETING_ID: '86172', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', REPORT_SUFFIX: 'MBS', FTP_BASE: MBS_DOCS, [CANDIDATE_KEY]: MBS_DRAFTS, SA4_BOOTSTRAP_STATE: 'done', FETCH_ABSTRACTS_ON_UPDATE: 'true' };
  const report = templateReport({ props: props, tdocs: [{ id: 'S4aI269001', agenda: '4.3', status: 'available', uploaded: true }] });
  report.s.collectRevisionsCore_ = () => { report.s.updateRevisions_(report.s.getCollectorConfig_()); };   // the real revision collection, for what it asks
  const built = report.build();
  report.update();
  report.update();
  check('a build and two updates complete', built.ok, true);
  check('none of them asked for a drafts folder, stored one, or touched the record', [report.fetches.filter(isDrafts), report.docProps.getProperty('REVISIONS_URL'), report.docProps.getProperty(CANDIDATE_KEY)], [[], null, MBS_DRAFTS]);

  const r = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? LISTING : null) });
  const page = r.open();
  page.resolve('86172');
  check('opening Configure Meeting and Resolve ask for no drafts folder (the lookup of a meeting never does)', [draftsAsked(r), r.portalRequests, r.docProps.getProperty('REVISIONS_URL')], [[], ['86172'], null]);
  r.s.finishReportSetup();
  check('"Finish Report Setup" on a report that is set up asks nothing', draftsAsked(r), []);

  // Who can ask at all.
  const callersOf = (src, name) => withoutComments(src).split('\nfunction ').filter((f) => new RegExp('[^A-Za-z0-9_$]' + name + '\\(').test(f.slice(f.indexOf('(')))).map((f) => f.slice(0, f.indexOf('(')));
  check('the folder is asked for in two places: the first run of a report, and the step after a Save of Configure Meeting (source)',
    [callersOf(CREATOR, 'confirmAdhocRevisionsFolderWith_'), callersOf(CREATOR, 'retryAdhocRevisionsFolderWith_'), callersOf(CREATOR, 'afterTemplateConfigurationSaved_'), callersOf(CODE, 'afterTemplateConfigurationSaved_'),
      callersOf(CODE, 'confirmAdhocRevisionsFolderWith_').concat(callersOf(CODE, 'retryAdhocRevisionsFolderWith_'))],
    [['retryAdhocRevisionsFolderWith_', 'finishReportSetupWith_'], ['afterTemplateConfigurationSaved_'], [], ['saveConfigurationSettings'], []]);
  check('the update, the build and the lookup of a meeting do not lead there (source)',
    ['continuousUpdateCore_', 'continuousUpdate', 'runFullReportBuildCore_', 'buildSkeletonWithTdocTables', 'updateRevisions_', 'collectorUpdate_', 'resolveMeetingForConfigDialog_', 'resolveMeetingCoreById_', 'discoverAgendaForConfigDialog_', 'configureMeetingSettings']
      .filter((name) => /afterTemplateConfigurationSaved_|validateRevisionsUrlCandidate_|fetchRevisionsUrlCandidate_|saveConfigurationSettings\(/.test(withoutComments(functionSource(CODE, name)).replace(/\.saveConfigurationSettings\(config\);/, ''))), []);
  check('outside the template runtime a Save does not do it at all', [/const templateSave = templateRuntimeRelease_\(\) \? beforeTemplateConfigurationSaved_\(config\) : null;\n    persistConfigurationSettings_\(templateSave \? templateSave\.config : config\);\n    return templateSave \? afterTemplateConfigurationSaved_\(templateSave\) : undefined;/.test(functionSource(CODE, 'saveConfigurationSettings')),
    (() => { const central = configurableReport({ release: false, props: { MEETING_ID: '86172', MEETING_TYPE: 'adhoc', FTP_BASE: MBS_DOCS }, network: (url) => (isDrafts(url) ? LISTING : null) });
      const out = central.s.saveConfigurationSettings({ meetingId: '86172', meetingType: 'adhoc', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'keep' }); return [out, central.requests, central.docProps.getProperty('REVISIONS_URL')]; })()],
  [true, [undefined, [], null]]);
}

console.log('8. a stored folder is never asked for again and never replaced');
{
  const OWN = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Inbox/Drafts/';
  const props = Object.assign(afterFailedSetup().props, { REVISIONS_URL: OWN });
  const r = configurableReport({ props: props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? LISTING : null) });
  const page = r.open();
  check('the dialog shows the stored folder', page.fields.revisionsUrl.value, OWN);
  page.save();
  page.resolve('86172').save();
  r.open().save();
  check('Save, Resolve and Save, and Save again: never asked, never changed, although the derived candidate would be confirmed', [draftsAsked(r), r.docProps.getProperty('REVISIONS_URL'), page.alerts.map((a) => /Drafts folder/.test(a))], [[], OWN, [false, false]]);
  check('the record of an earlier, unconfirmed candidate is removed once a folder is stored', r.docProps.getProperty(CANDIDATE_KEY), null);

  // The folder stored by the retry itself is subject to the same rule.
  const retried = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? LISTING : null) });
  retried.open().save();
  retried.open().save();
  retried.open().save();
  check('once the retry has stored it, later Saves ask nothing', [draftsAsked(retried), retried.docProps.getProperty('REVISIONS_URL')], [[MBS_DRAFTS], MBS_DRAFTS]);
}

console.log('9. a report that was given another meeting');
{
  // The report was set up for 86172 (MBS), its folder not confirmed. It is now given 85916 (Audio); the suggested folder is not taken over in the dialog.
  const toAudio = (network) => {
    const r = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172, 85916: GM.getMeetings85916 }, rows: ROWS_AUDIO, network: network });
    const page = r.open().resolve('85916');
    page.fields.revisionsUrl.value = '';
    page.save();
    return { r: r, page: page };
  };
  const ok = toAudio((url) => (isDrafts(url) ? LISTING : null));
  check('the report is on the new meeting', [ok.r.docProps.getProperty('MEETING_ID'), ok.r.docProps.getProperty('FTP_BASE')], ['85916', AUDIO_DOCS]);
  check('the folder asked for and stored is that of the NEW meeting; the old candidate is never asked for', [draftsAsked(ok.r), ok.r.docProps.getProperty('REVISIONS_URL'), ok.r.docProps.getProperty(CANDIDATE_KEY)], [[AUDIO_DRAFTS], AUDIO_DRAFTS, null]);

  const refused = toAudio((url) => (isDrafts(url) ? { status: 403, text: '' } : null));
  check('not confirmed: the record is the new meeting\'s candidate, not the old one', [draftsAsked(refused.r), refused.r.docProps.getProperty('REVISIONS_URL'), refused.r.docProps.getProperty(CANDIDATE_KEY)], [[AUDIO_DRAFTS], null, AUDIO_DRAFTS]);

  // The old candidate would be confirmed, the new one not: the old one must not get in.
  const onlyOld = toAudio((url) => (url === MBS_DRAFTS ? LISTING : (isDrafts(url) ? { status: 404, text: '' } : null)));
  check('even when the OLD folder would answer: it is not asked for and not stored', [draftsAsked(onlyOld.r), onlyOld.r.docProps.getProperty('REVISIONS_URL'), onlyOld.r.docProps.getProperty(CANDIDATE_KEY)], [[AUDIO_DRAFTS], null, AUDIO_DRAFTS]);

  // To a main meeting: its folder comes from the formula; the record of the ad-hoc candidate goes.
  const mainMeeting = meeting(GM.getMeetings86172, { Id: 60778, Type: 'OR', Title: '3GPPSA4#137', MtgDocURL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Synthetic/Docs/' });
  const main = configurableReport({ props: afterFailedSetup().props, meetings: { 60778: mainMeeting }, network: (url) => (isDrafts(url) ? LISTING : null) });
  const mainPage = main.open().resolve('60778');
  mainPage.fields.revisionsUrl.value = '';
  mainPage.fields.reportType.value = 'Video';
  mainPage.save();
  check('changed to a main meeting: nothing is asked, nothing is stored, and the old record is removed',
    [main.docProps.getProperty('MEETING_TYPE'), draftsAsked(main), main.docProps.getProperty('REVISIONS_URL'), main.docProps.getProperty(CANDIDATE_KEY)], ['main', [], null, null]);

  // About This Report shows a recorded candidate only while it is the candidate of the document folder the report has.
  const current = ok.r.s.currentRevisionsCandidate_;
  check('a record of another meeting\'s candidate is not shown as this report\'s', [current(MBS_DRAFTS, MBS_DOCS), current(MBS_DRAFTS, AUDIO_DOCS), current(AUDIO_DRAFTS, AUDIO_DOCS), current('', MBS_DOCS), current(MBS_DRAFTS, ''), current(null, null)],
    [MBS_DRAFTS, '', AUDIO_DRAFTS, '', '', '']);
  check('About This Report reads it that way (source)', /revisionsCandidate: currentRevisionsCandidate_\(props\.getProperty\(TEMPLATE_STATE_KEYS_\.revisionsCandidate\), props\.getProperty\('FTP_BASE'\)\)/.test(functionSource(CREATOR, 'showTemplateInfo')), true);
  check('the retry derives its candidate from the document folder stored now, never from the record (source)',
    [/ftpBase: trimmed_\(props\.getProperty\('FTP_BASE'\)\)/.test(functionSource(CREATOR, 'retryAdhocRevisionsFolderWith_')), /revisionsCandidate\)/.test(functionSource(CREATOR, 'retryAdhocRevisionsFolderWith_').replace(/props\.deleteProperty\(TEMPLATE_STATE_KEYS_\.revisionsCandidate\)/, '')),
      /getProperty\(TEMPLATE_STATE_KEYS_\.revisionsCandidate\)/.test(functionSource(CREATOR, 'confirmAdhocRevisionsFolderWith_'))], [true, false, false]);
}

console.log('10. the folder suggested by Resolve, taken over by the user (as before T-2026.10.8 -- recorded here, not changed)');
{
  // TWO WAYS a folder becomes REVISIONS_URL, on purpose (decision of 2026-10-06):
  //   - automatically: only a derived folder that a request confirmed (parts 5 to 9 above);
  //   - by the user: Resolve puts the derived folder into the field, marked "suggested -- please review", and a Save with it
  //     in the field stores it WITHOUT a request. This is the manual way for when the 3GPP server refuses scripted requests.
  // Do not remove it in favour of the automatic rule.
  const r = configurableReport({ props: afterFailedSetup().props, meetings: { 86172: GM.getMeetings86172 }, rows: ROWS_MBS, network: (url) => (isDrafts(url) ? { status: 403, text: '' } : null) });
  const page = r.open().resolve('86172');
  check('Resolve shows the derived folder as a suggestion to review', [page.fields.revisionsUrl.value, /suggested/.test(page.fields.revisionsUrlBadge ? page.fields.revisionsUrlBadge.innerHTML : '')], [MBS_DRAFTS, true]);
  page.save();
  check('saved with it in the field, it is stored as the user\'s entry, without a request', [r.docProps.getProperty('REVISIONS_URL'), draftsAsked(r), r.docProps.getProperty(CANDIDATE_KEY)], [MBS_DRAFTS, [], null]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll Configure Meeting / setup parity checks passed.');
process.exitCode = failures ? 1 : 0;
