/**
 * SA4-ARCH-002 — Pure-logic regression tests.
 *
 * Extends the VM-sandbox approach already used by tests/revision-order.test.js
 * to a second batch of pure/near-pure functions in Code.js: status
 * normalization, report-type -> agenda-prefix/mailing-list mapping, title
 * generation, e-mail subject/reply-prefix parsing, deadline parsing
 * (including the extended-deadline override), and TDoc number
 * normalization/extraction.
 *
 * These values were captured by RUNNING the current, unmodified Code.js
 * against representative inputs and recording its actual output -- they are
 * not hand-derived from reading the regexes. The point of this file is to
 * freeze CURRENT behavior (including any of its quirks) so a later
 * MeetingContext/profile refactor can be checked against it, not to assert
 * what the "correct" behavior should be.
 *
 * Run: node tests/pure-logic.test.js
 */

const { loadCode } = require('./helpers/load-code.js');

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

// =========================================================== normalizeStatus_

{
  const { sandbox } = loadCode();
  console.log('normalizeStatus_');
  check('"Available" -> reserved/available bucket', sandbox.normalizeStatus_('Available'), 'available');
  check('"Reserved" -> reserved/available bucket', sandbox.normalizeStatus_('Reserved'), 'reserved');
  check('"revised to S4-261599" -> other (only "reserved"/"available" are recognized, not "revised")',
    sandbox.normalizeStatus_('revised to S4-261599'), 'other');
  check('"Agreed" -> other', sandbox.normalizeStatus_('Agreed'), 'other');
  check('"Noted" -> other', sandbox.normalizeStatus_('Noted'), 'other');
  check('"endorsed" (lowercase) -> other', sandbox.normalizeStatus_('endorsed'), 'other');
  check('"WITHDRAWN" (uppercase) -> other', sandbox.normalizeStatus_('WITHDRAWN'), 'other');
  check('empty string -> other', sandbox.normalizeStatus_(''), 'other');
  check('null -> other', sandbox.normalizeStatus_(null), 'other');
  check('undefined -> other', sandbox.normalizeStatus_(undefined), 'other');
}

// ================================================ getAgendaPrefixForReportType_

{
  const { sandbox } = loadCode();
  console.log('getAgendaPrefixForReportType_ (report-type -> agenda-prefix map)');
  const expected = {
    Liaison: '5.', Audio: '7.', MBS: '8.', Video: '9.', RTC: '10.', '6G': '11.', New: '18.'
  };
  Object.keys(expected).forEach(type => {
    check(`"${type}" -> "${expected[type]}"`, sandbox.getAgendaPrefixForReportType_(type), expected[type]);
  });
  check('unknown report type falls back to "11." (the 6G prefix)',
    sandbox.getAgendaPrefixForReportType_('Unknown'), '11.');
  check('undefined report type falls back to "11."',
    sandbox.getAgendaPrefixForReportType_(undefined), '11.');
}

// ======================================= report-type -> mailing list / drafts folder
// MAILING_LISTS / DRAFTS_FOLDERS are top-level `const` objects, so (unlike
// `function`-declared symbols) they are not exposed as properties on the vm
// sandbox object directly -- only code running *inside* Code.js's own lexical
// scope can see them. getReportConfig_() closes over both maps, so reading
// its output per REPORT_SUFFIX is how this suite observes them from outside.

{
  console.log('report type -> mailing list / drafts folder (observed via getReportConfig_)');
  const expected = {
    '6G':      { mailingList: '3GPP_TSG_SA_WG4',       draftsFolder: 'FS_6G_MED' },
    Audio:     { mailingList: '3GPP_TSG_SA_WG4_AUDIO',  draftsFolder: 'Audio' },
    Video:     { mailingList: '3GPP_TSG_SA_WG4_VIDEO',  draftsFolder: 'Video' },
    MBS:       { mailingList: '3GPP_TSG_SA_WG4_MBS',    draftsFolder: 'MBS' },
    RTC:       { mailingList: '3GPP_TSG_SA_WG4_RTC',    draftsFolder: 'RTC' },
    Liaison:   { mailingList: '3GPP_TSG_SA_WG4',        draftsFolder: 'Plenary' },
    New:       { mailingList: '3GPP_TSG_SA_WG4',        draftsFolder: 'Plenary' }
  };
  Object.keys(expected).forEach(type => {
    const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: type } });
    const cfg = sandbox.getReportConfig_();
    check(`"${type}" mailing list`, cfg.LIST_NAME, expected[type].mailingList);
    check(`"${type}" drafts folder`, cfg.DRAFTS_FOLDER, expected[type].draftsFolder);
  });
}

// ===================================================== generateReportTitle_

{
  const { sandbox } = loadCode();
  console.log('generateReportTitle_');
  check('6G, meeting number with "-e" suffix extracted from TDOC_LIST_URL',
    sandbox.generateReportTitle_({
      REPORT_SUFFIX: '6G',
      TDOC_LIST_URL: 'https://www.3gpp.org/ftp/.../TDoc_List_Meeting_SA4%23137-e.xlsx',
      MEETING_ID: '60777'
    }),
    '6G Media Minutes SA4#137-e');
  check('Audio, plain meeting number extracted from TDOC_LIST_URL',
    sandbox.generateReportTitle_({
      REPORT_SUFFIX: 'Audio',
      TDOC_LIST_URL: 'https://www.3gpp.org/ftp/.../TDoc_List_Meeting_SA4%23136.xlsx',
      MEETING_ID: '60777'
    }),
    'Audio SWG Minutes SA4#136');
  check('Video, no TDOC_LIST_URL -> falls back to SA4#{MEETING_ID}',
    sandbox.generateReportTitle_({ REPORT_SUFFIX: 'Video', TDOC_LIST_URL: '', MEETING_ID: '60777' }),
    'Video SWG Minutes SA4#60777');
  check('unrecognized report type is used verbatim as the topic name',
    sandbox.generateReportTitle_({ REPORT_SUFFIX: 'Unknown', TDOC_LIST_URL: '', MEETING_ID: '99999' }),
    'Unknown Minutes SA4#99999');
}

// ===================================================== stripReplyPrefixes_

{
  const { sandbox } = loadCode();
  console.log('stripReplyPrefixes_');
  check('single "Re: " prefix stripped',
    sandbox.stripReplyPrefixes_('Re: [FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Title here'),
    '[FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Title here');
  check('"RE[2]: " counted-reply prefix stripped',
    sandbox.stripReplyPrefixes_('RE[2]: [FS_6G_MED, 1483] Title'),
    '[FS_6G_MED, 1483] Title');
  check('stacked "AW: Re: " prefixes stripped',
    sandbox.stripReplyPrefixes_('AW: Re: [x] y'),
    '[x] y');
  check('no prefix -> unchanged (trimmed)',
    sandbox.stripReplyPrefixes_('[No prefix] Title'),
    '[No prefix] Title');
  check('empty string -> empty string',
    sandbox.stripReplyPrefixes_(''),
    '');
}

// ===================================================== parseEmailSubject_

{
  const { sandbox } = loadCode();
  console.log('parseEmailSubject_');
  check('short TDoc + comma-separated deadline',
    sandbox.parseEmailSubject_('[FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Some title'),
    { tdocShort: '1483', tdocFull: null, deadline: { dateStr: '2026-08-25T14:00:00', timezone: 'CEST', day: 25, month: 8, year: 2026, hour: 14, minute: 0 } });
  check('full TDoc + semicolon-separated deadline',
    sandbox.parseEmailSubject_('[FS_6G_MED; S4-261483; 25 Aug 2026 1400 CEST] Some title'),
    { tdocShort: '1483', tdocFull: 'S4-261483', deadline: { dateStr: '2026-08-25T14:00:00', timezone: 'CEST', day: 25, month: 8, year: 2026, hour: 14, minute: 0 } });
  check('"Re: " prefix does not defeat parsing (2.2.0 fix)',
    sandbox.parseEmailSubject_('Re: [FS_6G_MED, 1483, 25 Aug 2026 1400 CEST] Some title'),
    { tdocShort: '1483', tdocFull: null, deadline: { dateStr: '2026-08-25T14:00:00', timezone: 'CEST', day: 25, month: 8, year: 2026, hour: 14, minute: 0 } });
  check('no brackets -> null', sandbox.parseEmailSubject_('No brackets here'), null);
  check('bracket with only one part -> null', sandbox.parseEmailSubject_('[OnlyOnePart]'), null);
}

// ================================================== extractDeadlineFromTitle_

{
  const { sandbox } = loadCode();
  console.log('extractDeadlineFromTitle_');

  // Patterns 3, 4, 4b, 5 and 7 have no year in their input and fall back to
  // `new Date().getFullYear()` inside the production code (Code.js ~2740).
  // Hardcoding a year here would make this test fail every January 1st, so
  // these expectations are built from the real current year instead -- the
  // behavior under test is "falls back to current year", not "falls back to
  // 2026" specifically.
  const currentYear = new Date().getFullYear();

  check('"DD Month YYYY HHMM TZ"',
    sandbox.extractDeadlineFromTitle_('25 Aug 2026 1400 CEST'),
    { dateStr: '2026-08-25T14:00:00', timezone: 'CEST', day: 25, month: 8, year: 2026, hour: 14, minute: 0 });
  check('"DD Month YYYY HHam/pm TZ"',
    sandbox.extractDeadlineFromTitle_('26 August 2026 12pm CEST'),
    { dateStr: '2026-08-26T12:00:00', timezone: 'CEST', day: 26, month: 8, year: 2026, hour: 12, minute: 0 });
  check('"DD Month, HHMM TZ" (no year -> current year)',
    sandbox.extractDeadlineFromTitle_('26 August, 1300 CEST'),
    { dateStr: `${currentYear}-08-26T13:00:00`, timezone: 'CEST', day: 26, month: 8, year: currentYear, hour: 13, minute: 0 });
  check('"DD Month HHam/pm TZ" (no year -> current year)',
    sandbox.extractDeadlineFromTitle_('26 August 12pm CEST'),
    { dateStr: `${currentYear}-08-26T12:00:00`, timezone: 'CEST', day: 26, month: 8, year: currentYear, hour: 12, minute: 0 });
  check('"Month DDth HHam/pm TZ" (no year -> current year)',
    sandbox.extractDeadlineFromTitle_('August 26th 12pm CEST'),
    { dateStr: `${currentYear}-08-26T12:00:00`, timezone: 'CEST', day: 26, month: 8, year: currentYear, hour: 12, minute: 0 });
  check('"DDth Month HHMMTZ" (time stuck to timezone, no year -> current year)',
    sandbox.extractDeadlineFromTitle_('28th August 1200CEST'),
    { dateStr: `${currentYear}-08-28T12:00:00`, timezone: 'CEST', day: 28, month: 8, year: currentYear, hour: 12, minute: 0 });
  check('"DD Month YYYY" without time -> defaults to 23:59 CEST',
    sandbox.extractDeadlineFromTitle_('25 Aug 2026'),
    { dateStr: '2026-08-25T23:59:00', timezone: 'CEST', day: 25, month: 8, year: 2026, hour: 23, minute: 59 });
  check('"Month DD, HHMM TZ" (no year -> current year)',
    sandbox.extractDeadlineFromTitle_('Aug 26, 1400 CEST'),
    { dateStr: `${currentYear}-08-26T14:00:00`, timezone: 'CEST', day: 26, month: 8, year: currentYear, hour: 14, minute: 0 });
  check('unparseable text -> null',
    sandbox.extractDeadlineFromTitle_('nothing parseable here'), null);
}

// =================================================== parseExtendedDeadline_

{
  const { sandbox } = loadCode();
  console.log('parseExtendedDeadline_');
  check('"DD Mon YYYY HH:MM TZ"',
    sandbox.parseExtendedDeadline_('27 Aug 2026 23:59 CEST'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('"DD Mon YYYY HHMM TZ"',
    sandbox.parseExtendedDeadline_('27 Aug 2026 2359 CEST'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('"Mon DD, YYYY HH:MM TZ"',
    sandbox.parseExtendedDeadline_('Aug 27, 2026 23:59 CEST'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('ISO "YYYY-MM-DD HH:MM" (no timezone -> defaults to CEST)',
    sandbox.parseExtendedDeadline_('2026-08-27 23:59'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('ISO "YYYY-MM-DDTHH:MM TZ"',
    sandbox.parseExtendedDeadline_('2026-08-27T23:59 CEST'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('"DD Mon YYYY" date-only -> end of day',
    sandbox.parseExtendedDeadline_('27 Aug 2026'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('ISO date-only "YYYY-MM-DD" -> end of day',
    sandbox.parseExtendedDeadline_('2026-08-27'),
    { dateStr: '2026-08-27T23:59:00', timezone: 'CEST', day: 27, month: 8, year: 2026, hour: 23, minute: 59 });
  check('empty string -> null', sandbox.parseExtendedDeadline_(''), null);
}

// =============================================== TDoc normalization/extraction

{
  const { sandbox } = loadCode();
  console.log('normalizeTdoc_ / extractTdocId_ / computeShortNumber_');

  check('normalizeTdoc_ accepts "S4-NNNNNN"', sandbox.normalizeTdoc_('S4-261480'), 'S4-261480');
  check('normalizeTdoc_ upper-cases lowercase input', sandbox.normalizeTdoc_('s4-261480'), 'S4-261480');
  check('normalizeTdoc_ extracts from surrounding text', sandbox.normalizeTdoc_('foo S4-261480 bar'), 'S4-261480');
  check('normalizeTdoc_ rejects a 5-digit number (only 6 digits match)', sandbox.normalizeTdoc_('S4-12345'), '');
  check('normalizeTdoc_ empty string -> ""', sandbox.normalizeTdoc_(''), '');

  check('extractTdocId_ default regex finds "S4-NNNNNN"', sandbox.extractTdocId_('S4-261480 text'), 'S4-261480');
  check('extractTdocId_ no match -> ""', sandbox.extractTdocId_('no tdoc here'), '');
  check('extractTdocId_ empty input -> ""', sandbox.extractTdocId_(''), '');

  check('computeShortNumber_ drops a leading zero of the last 4 digits down to 3',
    sandbox.computeShortNumber_('S4-261480'), '1480');
  check('computeShortNumber_ "S4-260087" -> "087" (3rd digit is 0)',
    sandbox.computeShortNumber_('S4-260087'), '087');
  check('computeShortNumber_ "S4-260007" -> "007"',
    sandbox.computeShortNumber_('S4-260007'), '007');
  check('computeShortNumber_ no 6-digit run -> ""', sandbox.computeShortNumber_('no digits'), '');
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll pure-logic tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
