/**
 * The opening sentence of an ad-hoc report (after T-2026.10.8).
 *
 * A new ad-hoc report without sessions gets one sentence in its "Opening of
 * the session" sub-section. It was
 *
 *   <Chair> opens the session on October 1, 2026 at <start> CEST.
 *
 * -- the start time a placeholder although the Portal gives it, and "CEST"
 * written into the code, which is wrong for a meeting outside Central
 * European Summer Time. It is now
 *
 *   <Chair> opens the session on October 1, 2026 at 15:30 UTC+2.
 *
 * with the time of day of the Portal's StartDate and the offset of the
 * Portal's StartTimeZone. What the Portal does not give stays a visible
 * placeholder ("<start>", "<time zone>"); the chair is always "<Chair>".
 *
 *   1. the time of day of a Portal StartDate; midnight is unknown;
 *   2. the Portal time zone as an offset from UTC;
 *   3. the sentence, for every combination of known and unknown;
 *   4. the creator takes both from the Portal into the setup information;
 *   5. the first run stores them with the meeting they belong to;
 *   6. Build Report from Scratch writes the sentence from what is stored;
 *   7. the start of one meeting is never written for another;
 *   8. what is not affected: main meetings, reports with sessions, Configure Meeting.
 *
 * Portal answers are the captured ones of tests/fixtures; everything else is synthetic.
 *
 * Run: node tests/meeting-start-opening.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { templateReport } = require('./helpers/template-report.js');
const { createReport, firstRunFor, payloadOf, descriptionWith, plain } = require('./helpers/template-creator.js');

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
const S = loadCode().sandbox;
const ROWS_MBS = [{ id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' }];
const ROWS_AUDIO = [{ id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '2' }];
/** A captured meeting with other values. */
const meeting = (base, changes) => [Object.assign({}, base[0], changes)];
const START_KEYS = ['MEETING_START_TIME', 'MEETING_TIME_ZONE', 'MEETING_START_BASIS'];
const startOf = (docProps) => START_KEYS.map((k) => docProps.getProperty(k));
const openingOf = (r) => r.body._children.filter((c) => c.getType() === 'PARAGRAPH').map((c) => c.getText()).filter((text) => /opens the session/.test(text));
/** A built ad-hoc report with these document properties. */
const built = (props) => { const r = templateReport({ release: false, token: false, props: props, tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'available' }] }); const result = r.build(); return { r: r, ok: result.ok, opening: openingOf(r) }; };

// ====================================================================
console.log('1. the time of day of a Portal StartDate');
{
  const time = S.computeMeetingStartTimeFromStartDate_;
  check('the captured meetings: 86172 starts at 15:30, 85916 at 09:00', [time(GM.getMeetings86172[0].StartDate), time(GM.getMeetings85916[0].StartDate)], ['15:30', '09:00']);
  check('hours and minutes, as written; seconds are left out', [time('2026-10-06 08:05:00'), time('2026-10-06 23:59:59'), time('2026-10-06T14:00:00'), time('2026-10-06 14:00'), time('2026-10-06 00:01:00'), time('2026-10-06 14:00:00.000')],
    ['08:05', '23:59', '14:00', '14:00', '00:01', '14:00']);
  check('midnight is not a start time: the Portal writes 00:00:00 for a meeting whose time was not entered', [time('2026-10-06 00:00:00'), time('2026-10-06 00:00'), time('2026-10-06T00:00:00')], ['', '', '']);
  check('a date without a time, another form of date, an impossible time and nothing at all give no time',
    [time('2026-10-06'), time('20261006T133000Z'), time('October 6, 2026 15:30'), time('2026-10-06 24:00:00'), time('2026-10-06 12:60:00'), time('2026-10-06 15:30:00 CEST'), time(''), time(null), time(undefined), time(1530)],
    ['', '', '', '', '', '', '', '', '', '']);
  check('the date of the meeting still comes from the same StartDate, as before', S.computeMeetingDateFromStartDate_(GM.getMeetings86172[0].StartDate), 'October 1, 2026');
}

// ====================================================================
console.log('2. the Portal time zone as an offset from UTC');
{
  const zone = S.computeMeetingTimeZoneLabel_;
  check('the captured meetings (the Portal writes the offset with a dot or with a colon)',
    [zone(GM.getMeetings86172[0].StartTimeZone), zone(GM.getMeetings86172[0].EndTimeZone), zone(GM.getMeetings85916[0].StartTimeZone)], ['UTC+2', 'UTC+1', 'UTC+8']);
  check('west of Greenwich, half and quarter hours, and Greenwich itself',
    [zone('(GMT-05:00) Eastern Time (US & Canada)'), zone('(GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi'), zone('(GMT+05:45) Kathmandu'), zone('(GMT-03.30) Newfoundland'), zone('(GMT) Greenwich Mean Time : Dublin, Edinburgh, Lisbon, London'),
      zone('(UTC) Coordinated Universal Time'), zone('(GMT+00:00) Monrovia, Reykjavik'), zone('(UTC+09:00) Osaka, Sapporo, Tokyo'), zone('(GMT+14:00) Kiritimati Island'), zone('  (gmt+01:00) Brussels')],
    ['UTC-5', 'UTC+5:30', 'UTC+5:45', 'UTC-3:30', 'UTC', 'UTC', 'UTC', 'UTC+9', 'UTC+14', 'UTC+1']);
  check('nothing is read out of a city or a zone name; an impossible offset, another form and nothing give no zone',
    [zone('Brussels, Copenhagen, Madrid, Paris'), zone('CEST'), zone('Europe/Berlin'), zone('(GMT+15:00) Nowhere'), zone('(GMT+02:75) Nowhere'), zone('GMT+02:00'), zone('(GMT+2) Brussels'), zone(''), zone(null), zone(undefined)],
    ['', '', '', '', '', '', '', '', '', '']);
  check('what is stored is checked in the same forms', [['15:30', '00:01', '23:59'].map((v) => S.isValidMeetingStartTime_(v)), ['00:00', '24:00', '9:00', '15:30:00', '', null, '<start>'].map((v) => S.isValidMeetingStartTime_(v)),
    ['UTC', 'UTC+2', 'UTC-3:30', 'UTC+14'].map((v) => S.isValidMeetingTimeZoneLabel_(v)), ['CEST', 'UTC+15', 'UTC+2:5', 'GMT+2', 'utc+2', '', null, 'UTC+2 '].map((v) => S.isValidMeetingTimeZoneLabel_(v))],
  [[true, true, true], Array(7).fill(false), [true, true, true, true], Array(8).fill(false)]);
}

// ====================================================================
console.log('3. the sentence');
{
  const sentence = S.buildMeetingOpeningSentence_;
  check('date, start time and time zone known', sentence('October 1, 2026', '15:30', 'UTC+2'), '<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.');
  check('the start time not known', sentence('October 1, 2026', '', 'UTC+2'), '<Chair> opens the session on October 1, 2026 at <start> UTC+2.');
  check('the time zone not known', sentence('October 1, 2026', '15:30', ''), '<Chair> opens the session on October 1, 2026 at 15:30 <time zone>.');
  check('neither known', sentence('October 1, 2026', '', ''), '<Chair> opens the session on October 1, 2026 at <start> <time zone>.');
  check('nothing known at all', [sentence('', '', ''), sentence(undefined, undefined, undefined), sentence(null, null, null), sentence('  ', ' ', ' ')], Array(4).fill('<Chair> opens the session on <meeting date> at <start> <time zone>.'));
  check('a value that is not in its stored form is not written: the placeholder is',
    [sentence('October 1, 2026', '00:00', 'CEST'), sentence('October 1, 2026', '25:00', 'UTC+99'), sentence('October 1, 2026', '15:30 or so', 'Brussels'), sentence('October 1, 2026', 1530, 2)],
    Array(4).fill('<Chair> opens the session on October 1, 2026 at <start> <time zone>.'));
  check('the chair is the placeholder in every case, and no zone name is ever written',
    [['October 1, 2026', '15:30', 'UTC+2'], ['', '', ''], ['x', '15:30', '']].map((a) => { const text = sentence(a[0], a[1], a[2]); return [text.indexOf('<Chair> opens the session on ') === 0, /CEST|CET\b/.test(text)]; }),
    [[true, false], [true, false], [true, false]]);
  const build = functionSource(CODE, 'buildSkeletonWithTdocTables');
  check('no time zone is written into the code: not in the sentence, not in the build (source)',
    [/CEST|CET\b|Europe\//.test(withoutComments(functionSource(CODE, 'buildMeetingOpeningSentence_'))), /CEST|CET\b/.test(withoutComments(build)), /getScriptTimeZone/.test(functionSource(CODE, 'getMeetingStartForOpening_'))], [false, false, false]);
  check('the build writes the sentence in one place, from the stored start (source)',
    [(build.match(/buildMeetingOpeningSentence_\(/g) || []).length, /const meetingStart = getMeetingStartForOpening_\(\);\n\s+body\.appendParagraph\(buildMeetingOpeningSentence_\(cfg\.MEETING_DATE, meetingStart\.time, meetingStart\.zone\)\);/.test(build)], [1, true]);
  check('nothing has a field for the chair: no property, no setup value, no dialog field (source)',
    [/CHAIR_NAME|MEETING_CHAIR|meetingChair|chairName/.test(withoutComments(CODE)), /chair/i.test(withoutComments(CREATOR))], [false, false]);
}

// ====================================================================
console.log('4. the creator: from the Portal into the setup information');
{
  const mbs = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  check('86172 is created', [mbs.preview.ok, mbs.created.ok], [true, true]);
  const config = payloadOf(mbs.s, mbs.description).config;
  check('the setup information has the start time and the time zone of the Portal, next to the date', [config.meetingDate, config.meetingStartTime, config.meetingTimeZone], ['October 1, 2026', '15:30', 'UTC+2']);
  const audio = createReport(GM.getMeetings85916, ROWS_AUDIO, '85916', {});
  const audioConfig = payloadOf(audio.s, audio.description).config;
  check('85916 (Shanghai): its own time and zone -- nothing Central European is assumed', [audioConfig.meetingDate, audioConfig.meetingStartTime, audioConfig.meetingTimeZone], ['September 28, 2026', '09:00', 'UTC+8']);

  const noTime = createReport(meeting(GM.getMeetings86172, { StartDate: '2026-10-01 00:00:00' }), ROWS_MBS, '86172', {});
  const noZone = createReport(meeting(GM.getMeetings86172, { StartTimeZone: '' }), ROWS_MBS, '86172', {});
  const oddZone = createReport(meeting(GM.getMeetings86172, { StartTimeZone: 'Brussels, Copenhagen, Madrid, Paris' }), ROWS_MBS, '86172', {});
  const withoutZone = GM.getMeetings86172.map((m) => { const copy = Object.assign({}, m); delete copy.StartTimeZone; return copy; });
  const missingZone = createReport(withoutZone, ROWS_MBS, '86172', {});
  check('a meeting at 00:00:00, without a time zone, or with one that has no offset: the report is created, with the value left empty',
    [noTime, noZone, oddZone, missingZone].map((c) => { const cfg = payloadOf(c.s, c.description).config; return [c.created.ok, cfg.meetingDate, cfg.meetingStartTime, cfg.meetingTimeZone]; }),
    [[true, 'October 1, 2026', '', 'UTC+2'], [true, 'October 1, 2026', '15:30', ''], [true, 'October 1, 2026', '15:30', ''], [true, 'October 1, 2026', '15:30', '']]);

  const main = createReport(meeting(GM.getMeetings86172, { Id: 60778, Type: 'OR', Title: '3GPPSA4#137', MtgDocURL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Synthetic/Docs/' }), [], '60778', { reportFamily: 'Video' });
  const mainConfig = payloadOf(main.s, main.description).config;
  check('a main-meeting report (its opening is copied from the report template) carries neither', [main.created.ok, mainConfig.meetingType, mainConfig.meetingStartTime, mainConfig.meetingTimeZone], [true, 'main', '', '']);

  check('the preview keeps the Portal time zone as it is given, beside the raw start date',
    [mbs.s.computeResolvedMeetingPreview_({}, mbs.preview.resolved).startTimeZoneRaw, mbs.s.computeResolvedMeetingPreview_({}, mbs.preview.resolved).startDateRaw, mbs.s.computeResolvedMeetingPreview_({}, null).startTimeZoneRaw],
    ['(GMT+02.00)  Brussels, Copenhagen, Madrid, Paris', '2026-10-01 15:30:00', null]);
}

// ====================================================================
console.log('5. the first run: stored with the meeting they belong to');
{
  const mbs = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  const run = firstRunFor(mbs.description);
  const setup = run.run();
  check('the first run stores the three values', [setup.status, startOf(run.docProps), run.docProps.getProperty('MEETING_DATE'), run.docProps.getProperty('MEETING_ID')],
    ['configured', ['15:30', 'UTC+2', '86172|October 1, 2026'], 'October 1, 2026', '86172']);
  check('and they are what the build reads', plain(run.s.getMeetingStartForOpening_()), { time: '15:30', zone: 'UTC+2' });

  const noTime = createReport(meeting(GM.getMeetings86172, { StartDate: '2026-10-01 00:00:00' }), ROWS_MBS, '86172', {});
  const runNoTime = firstRunFor(noTime.description);
  runNoTime.run();
  check('a meeting without a start time: the zone is stored, no time is', [startOf(runNoTime.docProps), plain(runNoTime.s.getMeetingStartForOpening_())], [[null, 'UTC+2', '86172|October 1, 2026'], { time: '', zone: 'UTC+2' }]);

  const neither = createReport(meeting(GM.getMeetings86172, { StartDate: '2026-10-01 00:00:00', StartTimeZone: '' }), ROWS_MBS, '86172', {});
  const runNeither = firstRunFor(neither.description);
  runNeither.run();
  check('a meeting with neither: nothing is stored for it at all', startOf(runNeither.docProps), [null, null, null]);

  // The setup information is text on the file, which an editor can change: it is checked before anything is stored.
  const tampered = (changes) => { const p = payloadOf(mbs.s, mbs.description); Object.assign(p.config, changes); const r = firstRunFor(descriptionWith(mbs.s, p)); const out = r.run(); return [out.status, out.errors || [], startOf(r.docProps)]; };
  check('a start time or a time zone that is not in its form is refused, and nothing is stored',
    [tampered({ meetingStartTime: '15:30 CEST' }), tampered({ meetingStartTime: '00:00' }), tampered({ meetingTimeZone: 'CEST' }), tampered({ meetingTimeZone: '<script>' }), tampered({ meetingStartTime: 1530 })],
    [['refused', ['The meeting start time is not a time.'], [null, null, null]], ['refused', ['The meeting start time is not a time.'], [null, null, null]],
      ['refused', ['The meeting time zone is not a UTC offset.'], [null, null, null]], ['refused', ['The meeting time zone is not a UTC offset.'], [null, null, null]],
      ['refused', ['Setup field "meetingStartTime" is not text.', 'The meeting start time is not a time.'], [null, null, null]]]);
  check('a setup field for the chair does not exist: it is refused like any unknown field', tampered({ chair: 'Somebody' }).slice(0, 2), ['refused', ['Unexpected setup field "chair".']]);
}

// ====================================================================
console.log('6. Build Report from Scratch writes the sentence');
{
  const STORED = { MEETING_ID: '86172', MEETING_DATE: 'October 1, 2026', MEETING_START_TIME: '15:30', MEETING_TIME_ZONE: 'UTC+2', MEETING_START_BASIS: '86172|October 1, 2026' };
  const full = built(STORED);
  check('date, start time and time zone', [full.ok, full.opening], [true, ['<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.']]);
  check('where the line always was: directly below the heading of the opening sub-section',
    (() => { const texts = full.r.body._children.filter((c) => c.getType() === 'PARAGRAPH').map((c) => c.getText()); const at = texts.indexOf('<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.'); return [/Opening of the session$/.test(texts[at - 1]), /Registration of Documents$/.test(texts[at + 1])]; })(), [true, true]);
  check('no start time stored', built(Object.assign({}, STORED, { MEETING_START_TIME: null })).opening, ['<Chair> opens the session on October 1, 2026 at <start> UTC+2.']);
  check('no time zone stored', built(Object.assign({}, STORED, { MEETING_TIME_ZONE: null })).opening, ['<Chair> opens the session on October 1, 2026 at 15:30 <time zone>.']);
  check('nothing stored (a report configured by hand, or created by an earlier release): the placeholders, and no "CEST"',
    built({ MEETING_ID: '86172', MEETING_DATE: 'October 1, 2026' }).opening, ['<Chair> opens the session on October 1, 2026 at <start> <time zone>.']);
  check('no meeting date either', built({ MEETING_ID: '86172', MEETING_DATE: null }).opening, ['<Chair> opens the session on <meeting date> at <start> <time zone>.']);
  check('stored values that are not in their form are not written', built(Object.assign({}, STORED, { MEETING_START_TIME: '00:00', MEETING_TIME_ZONE: 'CEST' })).opening,
    ['<Chair> opens the session on October 1, 2026 at <start> <time zone>.']);
  const rebuilt = built(STORED);
  rebuilt.r.build();
  check('a rebuild writes it once again, the same', openingOf(rebuilt.r), ['<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.']);
  rebuilt.r.update();
  check('an update does not touch it', openingOf(rebuilt.r), ['<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.']);

  // From the Portal to the document, in one piece.
  const created = createReport(GM.getMeetings85916, ROWS_AUDIO, '85916', {});
  const run = firstRunFor(created.description);
  run.run();
  const stored = {}; ['MEETING_ID', 'MEETING_DATE'].concat(START_KEYS).forEach((k) => { stored[k] = run.docProps.getProperty(k); });
  check('creator -> first run -> build, for the captured meeting 85916', built(stored).opening, ['<Chair> opens the session on September 28, 2026 at 09:00 UTC+8.']);
}

// ====================================================================
console.log('7. the start of one meeting is never written for another');
{
  const STORED = { MEETING_ID: '86172', MEETING_DATE: 'October 1, 2026', MEETING_START_TIME: '15:30', MEETING_TIME_ZONE: 'UTC+2', MEETING_START_BASIS: '86172|October 1, 2026' };
  check('the report was given another meeting date since: placeholders', built(Object.assign({}, STORED, { MEETING_DATE: 'October 8, 2026' })).opening, ['<Chair> opens the session on October 8, 2026 at <start> <time zone>.']);
  check('the report was given another meeting id since: placeholders', built(Object.assign({}, STORED, { MEETING_ID: '86999' })).opening, ['<Chair> opens the session on October 1, 2026 at <start> <time zone>.']);
  check('a start without the meeting it belongs to: placeholders', built(Object.assign({}, STORED, { MEETING_START_BASIS: null })).opening, ['<Chair> opens the session on October 1, 2026 at <start> <time zone>.']);

  // Configure Meeting: its save carries no start. It leaves a stored one alone -- and when it changes the meeting, the stored one no longer applies.
  const mbs = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  const run = firstRunFor(mbs.description);
  run.run();
  const dialogSave = (changes) => run.s.persistConfigurationSettings_(Object.assign({ meetingId: '86172', meetingType: 'adhoc', meetingName: 'SA4-e (AH) MBS SWG post 137-e', meetingDate: 'October 1, 2026', reportType: 'MBS',
    ftpBase: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/', tdocUrl: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172', mailingListMode: 'derived', apiTokenAction: 'keep', showPreview: true }, changes || {}));
  dialogSave();
  check('saving Configure Meeting for the same meeting keeps the stored start', [startOf(run.docProps), plain(run.s.getMeetingStartForOpening_())], [['15:30', 'UTC+2', '86172|October 1, 2026'], { time: '15:30', zone: 'UTC+2' }]);
  dialogSave({ meetingDate: 'October 8, 2026' });
  check('saving it with another meeting date: the stored start no longer applies', [run.docProps.getProperty('MEETING_DATE'), plain(run.s.getMeetingStartForOpening_())], ['October 8, 2026', { time: '', zone: '' }]);
  dialogSave({ meetingDate: 'October 1, 2026' });
  check('and applies again when the date is the one it was stored with', plain(run.s.getMeetingStartForOpening_()), { time: '15:30', zone: 'UTC+2' });
}

// ====================================================================
console.log('8. what is not affected');
{
  const manual = loadCode({ documentProperties: {} });
  manual.sandbox.persistConfigurationSettings_({ meetingId: '86172', meetingType: 'adhoc', meetingName: 'x', meetingDate: 'October 1, 2026', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'keep' });
  check('a configuration without a start (Configure Meeting, CENTRAL, Legacy) stores nothing for it', startOf(manual.docProps), [null, null, null]);
  const mainSave = loadCode({ documentProperties: {} });
  mainSave.sandbox.persistConfigurationSettings_({ meetingId: '60778', meetingType: 'main', meetingFolder: 'TSGS4_137_Synthetic', meetingNumber: '137', reportType: 'Video', mailingListMode: 'derived', apiTokenAction: 'keep' });
  check('nor does a main-meeting configuration', startOf(mainSave.docProps), [null, null, null]);
  const build = functionSource(CODE, 'buildSkeletonWithTdocTables');
  const adhocBranch = build.slice(build.indexOf("if (context.meeting.type === 'adhoc') {"), build.indexOf('// Opening section for SWG reports - copy X.1 content from template'));
  check('the sentence is in the ad-hoc branch only, and only for a report without sessions; the main branch copies the report template as before (source)',
    [/if \(!adhocSessionsEnabled_\(context\)\) \{\n\s+const meetingStart = getMeetingStartForOpening_\(\);/.test(adhocBranch), (withoutComments(build).match(/getMeetingStartForOpening_\(/g) || []).length,
      /\} else \{\n\s+\/\/ Opening section for SWG reports - copy X\.1 content from template\n\s+const openingHeader = findHeading_\(sourceBody, \/\^X\\\.1\\s\+\/\);\n\s+if \(openingHeader\) \{\n\s+copySectionContentWithReplacement_\(openingHeader, body, \/\^X\\\.2\\s\+\/, 'X', agendaPrefixNum\);/.test(build)],
    [true, 1, true]);
  check('the stored start is read by the build and by nothing else', (withoutComments(CODE).match(/getMeetingStartForOpening_\(/g) || []).length, 2);
  check('the Session administration section of a report with sessions does not use it', /MEETING_START|getMeetingStartForOpening_|buildMeetingOpeningSentence_/.test(functionSource(CODE, 'buildAdhocOpeningLines_') + functionSource(CODE, 'adhocSessionWhenText_')), false);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll meeting start / opening sentence checks passed.');
process.exitCode = failures ? 1 : 0;
