/**
 * A SYNTHETIC Microsoft Teams attendance export, for
 * tests/teams-attendance-parser.test.js.
 *
 * Every person, company, address, telephone number and meeting in this file
 * is invented -- every time and duration too -- and the domains are the
 * reserved example domains. Nothing here
 * was copied from a real export. What it reproduces is the STRUCTURE of one:
 *
 *   - four numbered sections: Summary, Participants, In-Meeting Activities,
 *     Meeting Engagement;
 *   - tab-separated cells, CR LF line ends, and (as bytes) UTF-16 LE with a
 *     byte-order mark;
 *   - a quoted meeting title with a line break inside it;
 *   - a Participants header of 15 columns (plus one this parser has never
 *     seen), and rows with fewer cells than the header;
 *   - United States dates with a 12-hour clock and no time zone;
 *   - durations with their zero components left out;
 *   - a Summary whose average and duration are NOT what the rows give.
 *
 * It is built here, as text, instead of being stored as a .csv file: a file
 * with tabs, CR LF and UTF-16 bytes would not survive Git's line-end
 * handling, and could not be read in a review.
 */

const DAY = '9/22/26';
const at = (time) => (time ? DAY + ', ' + time : '');

/** A cell that has to be quoted, written as Teams writes it. */
const quoted = (value) => '"' + String(value).replace(/"/g, '""') + '"';

const TITLE_LINES = ['Synthetic AHG Meeting ', '(September 22 - 24, 2026) on FS_EXAMPLE'];

const SUMMARY = {
  attendedParticipants: '25',
  startTime: at('2:48:05 PM'),
  endTime: at('6:11:26 PM'),
  meetingDuration: '3h 23m 20s',          // end minus start is 3:23:21
  averageAttendanceTime: '2h 31m 7s'     // not the mean of the rows below
};

const HEADER = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'Participant ID (UPN)', 'Role',
  'Engagement: Reaction-Applause', 'Engagement: Reaction-Laugh', 'Engagement: Reaction-Like', 'Engagement: Reaction-Love',
  'Engagement: Reaction-Surprised', 'Engagement: Camera On', 'Engagement: Raise Hands', 'Engagement: Unmute',
  'Synthetic Future Column'];

// The telephone number of the dial-in row and the Participant IDs. None of
// them may appear anywhere in what the parser returns.
const DIAL_IN_NUMBER = '15550100123';
const UPN = {
  jane: 'upn-jdoe-7731@tenant.example.org',
  zoe: 'upn-zoc-0042@tenant.example.org',
  lee: 'upn-lwong-5512@tenant.example.com',
  kai: 'upn-kai-9001@tenant.example.com',
  avery: 'upn-aquinn-3310@tenant.example.org'
};

/**
 * name (as Teams shows it), join, leave, duration, email, upn, role, and
 * `cells`: how many cells the row has (16 = full, 12 = short).
 * `raw: true` means the name is already written as a quoted cell.
 */
const PARTICIPANTS = [
  { name: 'Alex Organizer', join: '2:48:09 PM', leave: '6:11:26 PM', duration: '3h 23m 17s', email: 'alex.organizer@example.com', upn: 'alex.organizer@example.com', role: 'Organizer', engagement: ['', '', '', '', '', '', '3', '2'], future: 'x' },
  { name: 'Doe, Jane (External)', join: '3:00:00 PM', leave: '6:00:00 PM', duration: '3h', email: 'jane.doe@example.org', upn: UPN.jane, role: 'Presenter' },
  { name: '[ExampleCorp] Sam Rivera (Unverified)', join: '3:01:00 PM', leave: '5:01:00 PM', duration: '2h', email: '', upn: '', role: 'Presenter' },
  { name: 'Taro Yamada (山田 太郎) (External)', join: '3:00:10 PM', leave: '5:30:15 PM', duration: '2h 30m 5s', email: 'taro.yamada@example.net', upn: 'taro.yamada@example.net', role: 'Presenter' },
  { name: '张 三（示例）', join: '3:05:00 PM', leave: '4:05:07 PM', duration: '1h 7s', email: 'zhang.san@example.com', upn: 'zhang.san@example.com', role: 'Presenter' },
  { name: 'Zoë O\'Connor-Müller', join: '3:15:00 PM', leave: '4:00:00 PM', duration: '45m', email: 'zoe.oconnor-mueller@example.org', upn: UPN.zoe, role: 'Presenter' },
  { name: 'Pat Kim (ExampleCorp) (External)', join: '3:00:00 PM', leave: '6:06:40 PM', duration: '3h 6m 40s', email: 'pat.kim@example.com', upn: 'pat.kim@example.com', role: 'Presenter' },
  { name: 'Robin Lee - ExampleCorp (Unverified)', join: '3:02:00 PM', leave: '6:02:00 PM', duration: '3h', email: '', upn: '', role: 'Presenter' },
  { name: 'Morgan Diaz ExampleCorp/Research', join: '3:00:00 PM', leave: '5:45:00 PM', duration: '2h 45m', email: 'morgan.diaz@example.net', upn: 'morgan.diaz@example.net', role: 'Presenter' },
  { name: quoted(' ' + DIAL_IN_NUMBER + ' (Unverified)'), raw: true, join: '3:07:20 PM', leave: '5:41:02 PM', duration: '2h 33m 42s', email: '', upn: '', role: 'Attendee' },
  { name: 'Casey Guest (Unverified)', join: '3:01:30 PM', leave: '3:44:42 PM', duration: '43m 12s', email: '', upn: '', role: 'Presenter' },
  { name: 'Dana Twin (Unverified)', join: '3:00:00 PM', leave: '4:00:00 PM', duration: '1h', email: '', upn: '', role: 'Presenter' },
  { name: 'Dana Twin (Unverified)', join: '3:30:00 PM', leave: '5:00:00 PM', duration: '1h 30m', email: '', upn: '', role: 'Presenter' },
  { name: 'Jamie Sample', join: '3:00:00 PM', leave: '5:00:00 PM', duration: '2h', email: 'jamie.sample@example.com', upn: 'jamie.sample@example.com', role: 'Presenter' },
  { name: 'Jamie Sample (Unverified)', join: '5:10:00 PM', leave: '5:20:00 PM', duration: '10m', email: '', upn: '', role: 'Presenter' },
  { name: 'Lee Wong', join: '3:00:00 PM', leave: '4:00:00 PM', duration: '1h', email: 'lee.wong@example.com', upn: UPN.lee, role: 'Presenter' },
  { name: quoted('Wong, Lee  (External)'), raw: true, join: '4:30:00 PM', leave: '6:00:00 PM', duration: '1h 30m', email: 'LEE.WONG@example.com', upn: UPN.lee, role: 'Presenter' },
  { name: 'Kai Invalid', join: '3:00:00 PM', leave: '6:00:00 PM', duration: '3h', email: 'kai.invalid(at)example', upn: UPN.kai, role: 'Presenter' },
  { name: 'Noor Late', join: '', leave: '6:00:00 PM', duration: '30m', email: 'noor.late@example.org', upn: 'noor.late@example.org', role: 'Presenter' },
  { name: 'Omar Open', join: '3:10:00 PM', leave: '', duration: '', email: 'omar.open@example.com', upn: 'omar.open@example.com', role: 'Presenter' },
  { name: quoted('Sky "Ace" Tab\tName'), raw: true, join: '3:00:00 PM', leave: '6:00:00 PM', duration: '3h', email: 'sky.tab@example.net', upn: 'sky.tab@example.net', role: 'Presenter' },
  { name: 'Casey Guest (Unverified)', join: '3:45:30 PM', leave: '6:11:20 PM', duration: '2h 25m 49s', email: '', upn: '', role: 'Presenter' },
  // Left and joined again: Teams has already merged the intervals into one row.
  { name: 'Rita Rejoin', join: '3:00:00 PM', leave: '6:00:00 PM', duration: '1h 30m', email: 'rita.rejoin@example.com', upn: 'rita.rejoin@example.com', role: 'Presenter' },
  { name: 'Quinn, Avery', join: '', leave: '', duration: '', email: 'avery.quinn@example.org', upn: UPN.avery, role: 'Presenter', cells: 12 },
  { name: 'Riley Short (ExampleCorp) (External)', join: '', leave: '', duration: '', email: 'riley.short@example.com', upn: 'riley.short@example.com', role: 'Presenter', cells: 12 }
];

/** One participant row. In a comma-separated export a cell with a comma or a quote is quoted. */
function participantLine(p, d) {
  const cell = (v) => (d === ',' && /[",]/.test(v) ? quoted(v) : v);
  const cells = [p.raw ? p.name : cell(p.name), cell(at(p.join)), cell(at(p.leave)), p.duration, p.email, p.upn, p.role]
    .concat(p.engagement || ['', '', '', '', '', '', '', ''], [p.future || '']);
  return cells.slice(0, p.cells || HEADER.length).join(d);
}

// Sections the parser never reads. Their content only has to be there.
const ACTIVITIES = [
  ['Name', 'Join Time', 'Leave Time', 'Duration', 'Email', 'Role'],
  ['Alex Organizer', at('2:48:09 PM'), at('6:11:26 PM'), '3h 23m 17s', 'alex.organizer@example.com', 'Organizer'],
  ['Rita Rejoin', at('3:00:00 PM'), at('3:45:00 PM'), '45m', 'rita.rejoin@example.com', 'Presenter'],
  ['Rita Rejoin', at('5:15:00 PM'), at('6:00:00 PM'), '45m', 'rita.rejoin@example.com', 'Presenter']
];
const ENGAGEMENT = [
  ['Name', 'Engagement Type', 'Time'],
  ['Alex Organizer', 'Unmute', at('3:09:11 PM')],
  ['Alex Organizer', 'Raise Hands', at('3:20:00 PM')]
];

/**
 * The export as text.
 * options: { eol ('\r\n'), delimiter ('\t'), participants, summary,
 *            titleLines, withOtherSections (true) }
 */
function buildExport(options) {
  const o = options || {};
  const eol = o.eol || '\r\n';
  const d = o.delimiter || '\t';
  const summary = Object.assign({}, SUMMARY, o.summary || {});
  const cell = (v) => (d === ',' && /[",]/.test(v) ? quoted(v) : v);
  const lines = [
    '1. Summary',
    'Meeting title' + d + quoted((o.titleLines || TITLE_LINES).join(eol)),
    'Attended participants' + d + summary.attendedParticipants,
    'Start time' + d + cell(summary.startTime),
    'End time' + d + cell(summary.endTime),
    'Meeting duration' + d + summary.meetingDuration,
    'Average attendance time' + d + summary.averageAttendanceTime,
    '',
    '2. Participants',
    HEADER.join(d)
  ];
  (o.participants || PARTICIPANTS).forEach((p) => lines.push(participantLine(p, d)));
  if (o.withOtherSections !== false) {
    lines.push('', '3. In-Meeting Activities');
    ACTIVITIES.forEach((r) => lines.push(r.map(cell).join(d)));
    lines.push('', '4. Meeting Engagement');
    ENGAGEMENT.forEach((r) => lines.push(r.map(cell).join(d)));
  }
  return lines.join(eol) + eol;
}

/** The bytes of a real export: UTF-16 LE with a byte-order mark. */
function utf16le(text) {
  return Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from(text, 'utf16le')]);
}

module.exports = {
  DAY, SUMMARY, HEADER, PARTICIPANTS, TITLE_LINES, DIAL_IN_NUMBER, UPN,
  PRIVATE_STRINGS: [DIAL_IN_NUMBER, 'tenant.example', 'upn-'],
  buildExport, utf16le, quoted, at
};
