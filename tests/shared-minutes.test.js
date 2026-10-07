/**
 * Shared Minutes (Code.js 2.22.0, template reports): Configure Meeting >
 * Create Shared Minutes makes ONE separate Google Doc for a report, lets
 * anyone with the link edit it, reads that permission back, and writes
 * "Link to the shared minutes: <name>" into the report.
 *
 *   1. the pure rules: name, address, content, marker, query, the stored value
 *   2. Create: what is sent to Drive, what is stored, what the report shows
 *   3. where the document goes
 *   4. pressing the button again
 *   5. the permission: set, read back, refused
 *   6. every way it can stop half way, and what the next press does
 *   7. Verify and Forget
 *   8. the line in the report: main meeting, ad-hoc with sessions, ad-hoc without
 *   9. a real report: build, update, trigger, e-mail, rebuild -- nothing is created or shared there
 *  10. a copied report
 *  11. Configure Meeting: the section, its buttons, Save
 *  12. where the code is, and what did not change (CENTRAL, Legacy, the manifest)
 *
 * All ids, names and addresses are synthetic. No network, no Google account.
 *
 * Run: node tests/shared-minutes.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { makeSharedMinutesDrive, queryFor, DOC_MIME } = require('./helpers/fake-shared-minutes-drive.js');
const { templateReport, REPORT_ID } = require('./helpers/template-report.js');
const { configurableReport } = require('./helpers/config-dialog.js');
const { installExportStubs } = require('./helpers/email-export-stubs.js');

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

const ROOT = path.join(__dirname, '..');
const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));

const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const COPY_ID = 'COPYofREPORTdoc0000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.10', flavor: 'template', gitCommit: '03ec9fc0000000000000000000000000000000aa', gitTag: 'template-release/T-2026.10.10', templateDocumentId: TEMPLATE_ID };
const LABEL = 'Link to the shared minutes:';
const NAME = 'Audio SWG Shared Minutes – Synthetic ad-hoc';
const LINE = LABEL + ' ' + NAME;
const KEY = 'SHARED_DOCUMENT';
const PENDING = 'SHARED_DOCUMENT_PENDING';
const urlOf = (id) => 'https://docs.google.com/document/d/' + id + '/edit';
const H2 = 'H2';
const H3 = 'H3';
const T0 = Date.parse('2026-10-07T09:00:00.000Z');
const PROPS = { MEETING_ID: '85916', MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: 'Audio' };
// An ad-hoc report without sessions, as a build leaves its opening.
const ADHOC_BODY = [['1 Opening of the meeting', H2], ['1.1 Opening of the session', H3], ['<Chair> opens the session on September 22, 2026 at 15:00 UTC+2.'], ['1.2 Registration of Documents', H3], ['4 Topic', H2]];

/**
 * One report with its Drive, and the functions under test with every
 * service injected: the REAL Drive seam (sharedMinutesDrive_()) over the
 * fake Drive, the real property store, a fake document.
 * options: { props, paragraphs, reportId, drive, folder }
 */
function world(options) {
  const o = options || {};
  const props = Object.assign({}, PROPS, o.props || {});
  Object.keys(props).forEach((k) => { if (props[k] === null) delete props[k]; });
  const loaded = loadTemplateRuntime({ release: RELEASE, documentProperties: props });
  const s = loaded.sandbox;
  const w = { s: s, docProps: loaded.docProps, reportId: o.reportId || REPORT_ID, clock: { ms: T0 }, logs: [], bodyFails: false, storeFails: null, storeAttempts: 0 };
  w.drive = (o.drive || makeSharedMinutesDrive({ reportId: w.reportId, folder: o.folder })).install(s);
  w.body = makeFakeDocumentBody(s);
  (o.paragraphs || ADHOC_BODY).forEach(([text, heading]) => { const p = w.body.appendParagraph(text); if (heading) p.setHeading(heading); });
  s.Logger = { log: (m) => w.logs.push(String(m)) };
  // Writing the record can be made to fail: storeFails(n) -> true for the n-th attempt.
  const realSet = loaded.docProps.setProperty;
  loaded.docProps.setProperty = (k, v) => {
    if (k === KEY) { w.storeAttempts++; if (w.storeFails && w.storeFails(w.storeAttempts)) throw new Error('Properties storage is not available (synthetic)'); }
    return realSet(k, v);
  };
  w.deps = () => ({
    release: RELEASE, activeDocumentId: () => w.reportId, nowIso: () => new Date(w.clock.ms).toISOString(), nowMs: () => w.clock.ms, documentProperties: loaded.docProps,
    drive: s.sharedMinutesDrive_(), describe: () => ({ name: NAME, meetingName: loaded.docProps.getProperty('MEETING_NAME'), meetingDate: loaded.docProps.getProperty('MEETING_DATE') }),
    body: () => { if (w.bodyFails) throw new Error('The document is not available (synthetic)'); return w.body; }
  });
  w.create = () => JSON.parse(JSON.stringify(s.createSharedMinutesWith_(w.deps())));
  w.verify = () => JSON.parse(JSON.stringify(s.verifySharedMinutesWith_(w.deps())));
  w.forget = () => JSON.parse(JSON.stringify(s.forgetSharedMinutesWith_(w.deps())));
  w.record = () => { const raw = loaded.docProps.getProperty(KEY); return raw ? JSON.parse(raw) : null; };
  w.view = () => JSON.parse(JSON.stringify(s.sharedMinutesView_(s.currentSharedMinutesWith_(w.deps()), false)));
  w.minutes = () => w.drive.files.filter((f) => f.mimeType === DOC_MIME);
  return w;
}
/** The text of every paragraph of a body, a heading with '# '. */
const texts = (body) => body._children.filter((c) => c.getType() === 'PARAGRAPH').map((p) => (p.getHeading() !== 'NORMAL' ? '# ' : '') + p.getText());
/** The lines of a body that begin with the label: [text, the link at its first character, the link at the first character of the name]. */
const lines = (body) => body._children.filter((c) => c.getType() === 'PARAGRAPH' && c.getHeading() === 'NORMAL' && c.getText().toLowerCase().indexOf('link to the shared minutes') === 0)
  .map((p) => [p.getText(), p.editAsText().getLinkUrl(0), p.editAsText().getLinkUrl(LABEL.length + 1)]);
/** What was asked of Drive, by name, without the look-ups of the personal settings. */
const asked = (drive) => drive.names().filter((n) => n !== 'settings-list');
const MUTATIONS = ['create', 'share'];
const mutations = (drive) => drive.names().filter((n) => MUTATIONS.indexOf(n) !== -1);

// ------------------------------------------------------------------ 1
console.log('1. The pure rules');
{
  const s = loadTemplateRuntime({ release: RELEASE }).sandbox;
  check('the name: the first "Minutes" of the report title becomes "Shared Minutes" -- main meeting, ad-hoc meeting, a title without the word, no title',
    [s.sharedMinutesName_('Video SWG Minutes SA4#137-e'), s.sharedMinutesName_('6G Media Minutes – SA4-e (AH) on FS_6G_MED'), s.sharedMinutesName_('Minutes of the Minutes'), s.sharedMinutesName_('Report of SA4#137'), s.sharedMinutesName_('')],
    ['Video SWG Shared Minutes SA4#137-e', '6G Media Shared Minutes – SA4-e (AH) on FS_6G_MED', 'Shared Minutes of the Minutes', 'Shared Minutes – Report of SA4#137', 'Shared Minutes']);
  check('the name follows the title a build gives the report, for every family of a main meeting',
    ['6G', 'Audio', 'Video', 'MBS', 'RTC'].map((family) => s.sharedMinutesName_(s.generateReportTitle_({ REPORT_SUFFIX: family, TDOC_LIST_URL: 'https://www.3gpp.org/ftp/x/TDoc_List_Meeting_SA4%23137-e.xlsx', MEETING_ID: '60777' }))),
    ['6G Media Shared Minutes SA4#137-e', 'Audio SWG Shared Minutes SA4#137-e', 'Video SWG Shared Minutes SA4#137-e', 'MBS SWG Shared Minutes SA4#137-e', 'RTC SWG Shared Minutes SA4#137-e']);
  check('the address has one form', s.sharedMinutesUrl_('abc'), 'https://docs.google.com/document/d/abc/edit');
  check('the content is the title and one line naming the meeting -- nothing else, and nothing in it is markup',
    [s.sharedMinutesHtml_('A <b> & C', 'Meeting <1>', 'September 22, 2026'), s.sharedMinutesHtml_('Name', '', ''), s.sharedMinutesHtml_('Name', 'Meeting', null)],
    ['<html><head><meta charset="utf-8"></head><body><h1>A &lt;b&gt; &amp; C</h1><p>Meeting &lt;1&gt;, September 22, 2026</p></body></html>', '<html><head><meta charset="utf-8"></head><body><h1>Name</h1></body></html>',
      '<html><head><meta charset="utf-8"></head><body><h1>Name</h1><p>Meeting</p></body></html>']);
  check('the marker names the purpose, the report and the meeting', JSON.parse(JSON.stringify(s.sharedMinutesAppProperties_(REPORT_ID, '85916'))), { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '85916' });
  check('the query asks for the marked Google Docs of one report that are not in the trash', s.sharedMinutesQuery_(REPORT_ID), queryFor(REPORT_ID));
  const file = (over) => Object.assign({ id: 'SHAREDMINUTESdoc0000000000000000000000000001', mimeType: DOC_MIME, trashed: false, appProperties: { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '85916' } }, over || {});
  check('a file is the shared minutes of a report only when it is a Google Doc, not in the trash, and marked for this report and this meeting',
    [file(), file({ trashed: true }), file({ trashed: undefined }), file({ mimeType: 'application/pdf' }), file({ appProperties: undefined }), file({ appProperties: { sa4Purpose: 'other', sa4ReportId: REPORT_ID, sa4MeetingId: '85916' } }),
      file({ appProperties: { sa4Purpose: 'shared-minutes', sa4ReportId: COPY_ID, sa4MeetingId: '85916' } }), file({ appProperties: { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '86178' } }), file({ id: 'short' }), null]
      .map((meta) => s.isSharedMinutesFile_(meta, REPORT_ID, '85916')),
    [true, false, false, false, false, false, false, false, false, false]);
  check('"anyone with the link can edit" is a permission of type anyone and role writer -- no other',
    [[{ type: 'anyone', role: 'writer' }], [{ type: 'user', role: 'owner' }, { type: 'anyone', role: 'writer' }], [{ type: 'anyone', role: 'reader' }], [{ type: 'anyone', role: 'commenter' }], [{ type: 'user', role: 'writer' }],
      [{ type: 'domain', role: 'writer' }], [{ type: 'group', role: 'writer' }], [], null, undefined].map((p) => s.hasAnyoneWriterPermission_(p)),
    [true, true, false, false, false, false, false, false, false, false]);

  const good = { v: 1, fileId: 'SHAREDMINUTESdoc0000000000000000000000000001', url: urlOf('SHAREDMINUTESdoc0000000000000000000000000001'), name: NAME, reportDocumentId: REPORT_ID, meetingId: '85916', createdAt: '2026-10-07T09:00:00.000Z', permissionVerifiedAt: null };
  const read = (value, doc, meeting) => s.readSharedMinutes_(typeof value === 'string' || value === null ? value : JSON.stringify(value), doc || REPORT_ID, meeting === undefined ? '85916' : meeting).status;
  check('the stored value is used only when it is what this code stores, for this document and this meeting',
    [read(good), read(Object.assign({}, good, { permissionVerifiedAt: '2026-10-07T09:00:05.000Z' })), read(null), read(''), read('not json'), read('[]'), read(Object.assign({}, good, { v: 2 })), read(Object.assign({}, good, { v: 0 })),
      read(good, COPY_ID), read(good, REPORT_ID, '86178'), read(good, REPORT_ID, '')],
    ['ok', 'ok', 'absent', 'absent', 'invalid', 'invalid', 'unsupported', 'invalid', 'other-report', 'other-meeting', 'other-meeting']);
  check('a value with a field missing or of another kind is not used',
    [{ fileId: 'x' }, { url: 'https://example.org/' }, { url: urlOf('SHAREDMINUTESdoc0000000000000000000000000002') }, { name: '' }, { name: 7 }, { reportDocumentId: '' }, { meetingId: 85916 }, { meetingId: 'abc' }, { createdAt: '' }, { permissionVerifiedAt: true },
      { permissionVerifiedAt: undefined }].map((over) => read(Object.assign({}, good, over))),
    ['invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid']);
  check('the stored form has these fields, in this order', Object.keys(JSON.parse(s.serializeSharedMinutes_(good))), ['v', 'fileId', 'url', 'name', 'reportDocumentId', 'meetingId', 'createdAt', 'permissionVerifiedAt']);
}

// ------------------------------------------------------------------ 2
console.log('2. Create: what is sent to Drive, what is stored, what the report shows');
{
  const w = world();
  const result = w.create();
  const made = w.minutes()[0];
  check('one press: Drive is searched, the folder of the report is looked at, ONE document is created, its permissions are read, it is shared, and the permissions are read again',
    asked(w.drive), ['search', 'report-parents', 'folder', 'create', 'permissions', 'share', 'permissions']);
  const create = w.drive.calls.filter((c) => c[0] === 'create')[0][2];
  check('the creation is one request: name, type Google Doc, the folder of the report, and the marker',
    create.resource, { name: NAME, mimeType: DOC_MIME, appProperties: { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '85916' }, parents: ['FOLDERid00000000000000000000000000'] });
  check('with the content in it (HTML, converted by Drive): the title and the meeting', [create.mimeType, create.html],
    ['text/html', '<html><head><meta charset="utf-8"></head><body><h1>' + NAME + '</h1><p>Synthetic ad-hoc, September 22, 2026</p></body></html>']);
  check('the search is the one query, for this report, and looks in shared drives too',
    (({ q, includeItemsFromAllDrives, supportsAllDrives }) => [q, includeItemsFromAllDrives, supportsAllDrives])(w.drive.calls.filter((c) => c[0] === 'search')[0][2]), [queryFor(REPORT_ID), true, true]);
  check('the permission that is created: type anyone, role writer, not discoverable', w.drive.calls.filter((c) => c[0] === 'share')[0][2].resource, { type: 'anyone', role: 'writer', allowFileDiscovery: false });
  check('it is created on the new document, and on no other file', w.drive.calls.filter((c) => c[0] === 'share').map((c) => c[1]), [made.id]);
  check('the result: complete, created, with the name', [result.ok, result.state, result.created, result.adopted, result.message],
    [true, 'complete', true, false, 'The Shared Minutes document "' + NAME + '" was created. Anyone with the link can edit it. The link is in the report.']);
  check('anyone with the link can edit the document', w.drive.anyoneCanEdit(made.id), true);
  check('SHARED_DOCUMENT is stored: the file, its address, its name, this report, this meeting, when it was created and when the permission was verified', w.record(),
    { v: 1, fileId: made.id, url: urlOf(made.id), name: NAME, reportDocumentId: REPORT_ID, meetingId: '85916', createdAt: '2026-10-07T09:00:00.000Z', permissionVerifiedAt: '2026-10-07T09:00:00.000Z' });
  check('the marker that a creation is under way is gone', w.docProps.getProperty(PENDING), null);
  check('the report has one line, below the "Opening of the session" heading: the label as text, the name as the link', [lines(w.body), texts(w.body).slice(0, 4)],
    [[[LINE, null, urlOf(made.id)]], ['# 1 Opening of the meeting', '# 1.1 Opening of the session', LINE, '<Chair> opens the session on September 22, 2026 at 15:00 UTC+2.']]);
  check('the link is exactly the name', w.body._children[2]._links.filter((l) => l[2]), [[LABEL.length + 1, LINE.length - 1, urlOf(made.id)]]);
  check('what Configure Meeting shows afterwards: the name, the sharing with its tick, Open, Verify, Forget -- and no Create',
    w.view(), { lines: [{ text: NAME, strong: true }, { text: 'Anyone with the link can edit ✓ (verified 2026-10-07)', strong: false }], warning: '', createLabel: '', createDisabled: false, openUrl: urlOf(made.id), canVerify: true, canForget: true });

  const before = world();
  check('before: "Status: Not created", and the one button', before.view(),
    { lines: [{ text: 'Status: Not created', strong: false }], warning: '', createLabel: 'Create Shared Minutes', createDisabled: false, openUrl: '', canVerify: false, canForget: false });
  const noMeeting = world({ props: { MEETING_ID: null } });
  check('without a meeting the button is off, and pressing it creates nothing', [noMeeting.view().createDisabled, noMeeting.view().lines.map((l) => l.text), noMeeting.create().ok, asked(noMeeting.drive), noMeeting.record()],
    [true, ['Status: Not created', 'Resolve the meeting and save the configuration first.'], false, [], null]);
  const master = world({ reportId: TEMPLATE_ID });
  check('in the template itself nothing is created', [master.create().message, asked(master.drive)], ['This is the SA4 Report Template itself: shared minutes belong to a report.', []]);
}

// ------------------------------------------------------------------ 3
console.log('3. Where the document goes');
{
  const place = (w) => { const r = w.create(); const made = w.minutes()[0]; return [r.ok, made.parents, / It is in your My Drive: the folder of the report could not be used\.$/.test(r.message)]; };
  check('the folder of the report, when the user can add to it', place(world()), [true, ['FOLDERid00000000000000000000000000'], false]);
  check('a folder the user cannot add to: My Drive, and the message says so', place(world({ folder: { id: 'FOLDERid00000000000000000000000000', name: 'Read-only folder', writable: false } })), [true, [], true]);
  check('a report without a folder the user can see: My Drive, with the notice', place(world({ folder: null })), [true, [], true]);
  const failing = world();
  failing.drive.failFolder = true;
  check('a folder that cannot be looked at: My Drive, with the notice -- never a reason to fail', place(failing), [true, [], true]);
  check('the creation without a folder names no parent at all', failing.drive.calls.filter((c) => c[0] === 'create')[0][2].resource.parents, undefined);
}

// ------------------------------------------------------------------ 4
console.log('4. Pressing the button again');
{
  const w = world();
  w.create();
  const first = w.record();
  w.drive.calls.length = 0;
  w.clock.ms += 3600000;
  const again = w.create();
  check('a second press creates nothing and shares nothing: it looks at the document and its permissions', [asked(w.drive), mutations(w.drive)], [['meta', 'permissions'], []]);
  check('and says so', [again.ok, again.state, again.created, again.adopted, again.message], [true, 'complete', false, false, 'The Shared Minutes document "' + NAME + '" is set up already; nothing was created. Anyone with the link can edit it.']);
  check('there is still one document, one line, the same record -- only the time of the verification is new',
    [w.minutes().length, lines(w.body).length, Object.assign({}, w.record(), { permissionVerifiedAt: null }), w.record().permissionVerifiedAt],
    [1, 1, Object.assign({}, first, { permissionVerifiedAt: null }), '2026-10-07T10:00:00.000Z']);
  const writes = w.body._children[2]._links.length;
  w.create();
  w.create();
  check('ten presses: one document, one permission of anyone, one line, and the line is not written again', (() => { for (let i = 0; i < 7; i++) w.create(); return [w.minutes().length, w.minutes()[0].permissions.filter((p) => p.type === 'anyone').length, lines(w.body).length,
    w.body._children[2]._links.length === writes, w.drive.count('create'), w.drive.count('share')]; })(), [1, 1, 1, true, 0, 0]);
}

// ------------------------------------------------------------------ 5
console.log('5. The permission: set, read back, refused');
{
  const refused = world();
  refused.drive.failShare = true;
  const r = refused.create();
  const made = refused.minutes()[0];
  check('Drive refuses "anyone, writer": the setup is NOT reported as done', [r.ok, r.state, r.created],  [false, 'sharing-failed', true]);
  check('the message says that sharing failed, why, and that the document is not shared', r.message,
    'SHARING FAILED. The Shared Minutes document "' + NAME + '" was created and is recorded in this report, but "anyone with the link can edit" could not be set (Sharing outside the organization is not allowed (synthetic)). ' +
    'It is NOT shared, and no link was written into the report. Press "Finish sharing" to try again.');
  check('the document is kept (not in the trash) and recorded, with no verified permission', [made.trashed, refused.record().fileId, refused.record().permissionVerifiedAt, refused.drive.anyoneCanEdit(made.id)], [false, made.id, null, false]);
  check('no link is in the report', [lines(refused.body), texts(refused.body)], [[], ADHOC_BODY.map(([t, h]) => (h ? '# ' : '') + t)]);
  check('Configure Meeting shows it as not shared, and the button is "Finish sharing"', (({ lines: l, createLabel, openUrl, canVerify }) => [l.map((x) => x.text), createLabel, openUrl, canVerify])(refused.view()),
    [[NAME, 'NOT shared: "anyone with the link can edit" is not confirmed. The link is not in the report.'], 'Finish sharing', urlOf(made.id), true]);
  refused.drive.calls.length = 0;
  const still = refused.create();
  check('the next press, the policy unchanged: the same document, sharing tried again, failed again, nothing created', [still.state, still.message.indexOf('"' + NAME + '" exists and is recorded in this report') !== -1, asked(refused.drive), refused.minutes().length],
    ['sharing-failed', true, ['meta', 'permissions', 'share', 'permissions'], 1]);
  refused.drive.failShare = false;
  refused.drive.calls.length = 0;
  const later = refused.create();
  check('the next press, the policy lifted: the same document is shared, verified and linked; none is created', [later.ok, later.state, later.created, asked(refused.drive), refused.minutes().length, refused.drive.anyoneCanEdit(made.id),
    !!refused.record().permissionVerifiedAt, lines(refused.body)], [true, 'complete', false, ['meta', 'permissions', 'share', 'permissions'], 1, true, true, [[LINE, null, urlOf(made.id)]]]);

  const ignored = world();
  ignored.drive.shareIgnored = true;
  const dropped = ignored.create();
  check('Drive answers "done" and sets nothing: believed is only what is read back -- sharing failed, no link', [dropped.ok, dropped.state, /could not be confirmed\. It is NOT shared/.test(dropped.message), ignored.record().permissionVerifiedAt, lines(ignored.body)],
    [false, 'sharing-failed', true, null, []]);
  const lowered = world();
  lowered.drive.shareRole = 'reader';
  const viewOnly = lowered.create();
  check('Drive sets "anyone can VIEW" instead: that is not the permission -- sharing failed, no link', [viewOnly.ok, viewOnly.state, lowered.record().permissionVerifiedAt, lines(lowered.body)], [false, 'sharing-failed', null, []]);
  const unreadable = world();
  unreadable.drive.failPermissions = true;
  const blind = unreadable.create();
  check('the permissions cannot be read: nothing is shared blindly, sharing failed, no link', [blind.ok, blind.state, mutations(unreadable.drive), unreadable.record().permissionVerifiedAt, lines(unreadable.body)], [false, 'sharing-failed', ['create'], null, []]);
  const there = world();
  const marked = there.drive.add({ name: NAME, appProperties: { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '85916' }, permissions: [{ id: 'owner', type: 'user', role: 'owner' }, { id: 'anyoneWithLink', type: 'anyone', role: 'writer' }] });
  const found = there.create();
  check('a document that is shared already is not shared a second time', [found.ok, found.adopted, mutations(there.drive), there.record().fileId, there.minutes()[0].permissions.length], [true, true, [], marked.id, 2]);
}

// ------------------------------------------------------------------ 6
console.log('6. Every way it can stop half way, and what the next press does');
{
  // (a) created and shared, but the report could not record it at all.
  const a = world();
  a.storeFails = (n) => n === 1;
  const a1 = a.create();
  const aDoc = a.minutes()[0];
  check('(a) the document is created, the record cannot be stored: reported, and nothing is shared -- the report does not know the document yet',
    [a1.ok, a1.state, a1.message, mutations(a.drive), a.record(), lines(a.body)],
    [false, 'refused', 'The Shared Minutes document "' + NAME + '" exists, but could not be recorded in this report (Properties storage is not available (synthetic)), and was not shared. Press the button again: it is found and used, not created a second time.',
      ['create'], null, []]);
  a.drive.calls.length = 0;
  const a2 = a.create();
  check('(a) the next press finds it by its marker and goes on: recorded, shared, verified, linked -- and no second document',
    [a2.ok, a2.created, a2.adopted, asked(a.drive), a.minutes().length, a.record().fileId, lines(a.body), a2.message],
    [true, false, true, ['search', 'permissions', 'share', 'permissions'], 1, aDoc.id, [[LINE, null, urlOf(aDoc.id)]],
      'The Shared Minutes document "' + NAME + '" existed already for this report and is used; no second one was created. Anyone with the link can edit it. The link is in the report.']);

  // (b) created, shared and verified, but the verification could not be recorded.
  const b = world();
  b.storeFails = (n) => n === 2;
  const b1 = b.create();
  const bDoc = b.minutes()[0];
  check('(b) created and shared, the verification cannot be stored: not reported as done, no link; the record names the document, unverified',
    [b1.ok, b1.state, /is shared, but that could not be recorded/.test(b1.message), b.drive.anyoneCanEdit(bDoc.id), b.record().fileId, b.record().permissionVerifiedAt, lines(b.body)], [false, 'sharing-failed', true, true, bDoc.id, null, []]);
  b.drive.calls.length = 0;
  const b2 = b.create();
  check('(b) the next press reads the permission, finds it, shares nothing, records and links', [b2.ok, asked(b.drive), mutations(b.drive), !!b.record().permissionVerifiedAt, lines(b.body).length, b.minutes().length], [true, ['meta', 'permissions'], [], true, 1, 1]);

  // (c) everything stored, the report could not be written.
  const c = world();
  c.bodyFails = true;
  const c1 = c.create();
  check('(c) the link cannot be written into the report: the setup is done and stored, and the message says what is missing',
    [c1.ok, c1.state, !!c.record().permissionVerifiedAt, lines(c.body), /^The Shared Minutes document ".*" was created\. Anyone with the link can edit it\. The link could not be written into the report \(The document is not available \(synthetic\)\)\./.test(c1.message)],
    [true, 'complete', true, [], true]);
  c.bodyFails = false;
  c.drive.calls.length = 0;
  c.create();
  check('(c) the next press writes it; nothing is created or shared', [lines(c.body).length, mutations(c.drive), c.minutes().length], [1, [], 1]);
  const cBuild = world();
  cBuild.bodyFails = true;
  cBuild.create();
  cBuild.s.DocumentApp.getActiveDocument = () => ({ getId: () => REPORT_ID, getBody: () => cBuild.body });
  cBuild.drive.calls.length = 0;
  check('(c) and so does the end of a build, without asking Drive', [cBuild.s.finishSharedMinutesRebuild_(cBuild.body), lines(cBuild.body).length, asked(cBuild.drive)], ['', 1, []]);

  // (d) the request that creates the document times out: Drive has it, the report knows nothing.
  const d = world();
  d.drive.createThenThrow = true;
  const d1 = d.create();
  check('(d) a timeout directly after the creation: reported as not created, nothing shared or recorded -- but the marker of the attempt is stored, and Drive has the document',
    [d1.ok, d1.state, /^The Shared Minutes document could not be created \(Exceeded maximum execution time \(synthetic\)\)\. Nothing was shared or recorded\./.test(d1.message), d.record(), JSON.parse(d.docProps.getProperty(PENDING)), d.minutes().length, mutations(d.drive)],
    [false, 'refused', true, null, { startedAt: T0 }, 1, ['create']]);
  d.drive.createThenThrow = false;
  d.drive.calls.length = 0;
  d.clock.ms += 5000;
  const d2 = d.create();
  check('(d) the next press, seconds later: the document is found by its marker and used -- one document, shared, verified, linked',
    [d2.ok, d2.adopted, d2.created, mutations(d.drive), d.minutes().length, d.record().fileId === d.minutes()[0].id, d.docProps.getProperty(PENDING), lines(d.body).length], [true, true, false, ['share'], 1, true, null, 1]);

  // (d') the same, and Drive does not find the new document yet.
  const lag = world();
  lag.drive.createThenThrow = true;
  lag.drive.searchLag = true;
  lag.create();
  lag.drive.createThenThrow = false;
  lag.drive.calls.length = 0;
  lag.clock.ms += 5000;
  const lag2 = lag.create();
  check("(d') Drive does not find the document yet: the answer \"none\" is not trusted so soon after a creation -- nothing is created",
    [lag2.ok, lag2.state, lag2.message, asked(lag.drive), lag.minutes().length],
    [false, 'refused', 'Nothing was created: a creation was started 5 seconds ago, and Google Drive may not show its document yet. Press the button again in 115 seconds: the document is used if it exists, and created if it does not.', ['search'], 1]);
  lag.clock.ms += 60000;
  lag.drive.calls.length = 0;
  check("(d') still not after a minute: still nothing is created", [lag.create().state, mutations(lag.drive), lag.minutes().length], ['refused', [], 1]);
  lag.drive.settle();
  lag.clock.ms += 1000;
  lag.drive.calls.length = 0;
  const lag3 = lag.create();
  check("(d') once Drive finds it, it is used -- within the two minutes too", [lag3.ok, lag3.adopted, mutations(lag.drive), lag.minutes().length], [true, true, ['share'], 1]);
  const gone = world();
  gone.drive.failCreate = true;
  gone.create();
  gone.drive.failCreate = false;
  gone.clock.ms += 119000;
  check('a creation that really failed: for two minutes the button creates nothing ...', [gone.create().state, gone.minutes().length], ['refused', 0]);
  gone.clock.ms += 2000;
  const fresh = gone.create();
  check('... and after them it creates the document', [fresh.ok, fresh.created, gone.minutes().length, lines(gone.body).length], [true, true, 1, 1]);

  // (e) the request that shares the document times out.
  const e = world();
  e.drive.shareThenThrow = true;
  const e1 = e.create();
  check('(e) a timeout directly after the permission was set: the permission is read back, found, and the setup is complete', [e1.ok, e1.state, e.drive.anyoneCanEdit(e.minutes()[0].id), !!e.record().permissionVerifiedAt, lines(e.body).length], [true, 'complete', true, true, 1]);
  const e2 = world();
  let reads = 0;
  e2.drive.failPermissions = () => ++reads === 2;
  const e2r = e2.create();
  check('(e) the execution stops after the permission was set, before it is read back: not reported as done, recorded unverified, no link',
    [e2r.ok, e2r.state, e2.drive.anyoneCanEdit(e2.minutes()[0].id), e2.record().permissionVerifiedAt, lines(e2.body)], [false, 'sharing-failed', true, null, []]);
  e2.drive.failPermissions = false;
  e2.drive.calls.length = 0;
  const e2n = e2.create();
  check('(e) the next press finds the permission and sets none: complete, linked, one document, one permission of anyone',
    [e2n.ok, asked(e2.drive), e2.minutes().length, e2.minutes()[0].permissions.filter((p) => p.type === 'anyone').length, lines(e2.body).length], [true, ['meta', 'permissions'], 1, 1, 1]);

  // (f) two documents marked for this report and meeting.
  const f = world();
  const marker = { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '85916' };
  const one = f.drive.add({ name: NAME, appProperties: marker });
  const two = f.drive.add({ name: NAME, appProperties: marker });
  const f1 = f.create();
  check('(f) two marked documents: FAIL CLOSED -- nothing created, nothing shared, nothing recorded, no link, and the message lists both',
    [f1.ok, f1.state, mutations(f.drive), f.record(), f.docProps.getProperty(PENDING), lines(f.body), f.drive.anyoneCanEdit(one.id), f.drive.anyoneCanEdit(two.id), f1.message],
    [false, 'ambiguous', [], null, null, [], false, false, 'Nothing was created, shared or recorded: Google Drive has 2 Shared Minutes documents for this report and meeting, and none is picked for you:\n' +
      '• ' + NAME + ' — ' + urlOf(one.id) + '\n• ' + NAME + ' — ' + urlOf(two.id) + '\nMove all but one to the trash, then press the button again.']);
  check('(f) the next press: the same, as often as it is pressed', [f.create().state, f.create().state, mutations(f.drive), f.record()], ['ambiguous', 'ambiguous', [], null]);
  two.trashed = true;
  const f2 = f.create();
  check('(f) with one of them in the trash, the other is the document', [f2.ok, f2.adopted, f.record().fileId, f.drive.count('create'), f.drive.anyoneCanEdit(one.id), f.drive.anyoneCanEdit(two.id)], [true, true, one.id, 0, true, false]);
  const leaky = world();
  leaky.drive.add({ name: 'Minutes of another meeting', appProperties: { sa4Purpose: 'shared-minutes', sa4ReportId: REPORT_ID, sa4MeetingId: '86178' } });
  leaky.drive.add({ name: 'Not a document', mimeType: 'application/pdf', appProperties: marker });
  const l1 = leaky.create();
  check('a marked document of another meeting, and a marked file of another kind, are not the document: a new one is created, and they are not touched',
    [l1.ok, l1.created, leaky.minutes().length, leaky.drive.files.filter((x) => x.permissions.some((p) => p.type === 'anyone')).map((x) => x.name)], [true, true, 2, [NAME]]);

  // (g) the search itself fails.
  const g = world();
  g.drive.failSearch = true;
  const g1 = g.create();
  check('(g) Drive cannot be searched: nothing is created -- there may be a document already', [g1.ok, g1.message, mutations(g.drive), g.record(), g.docProps.getProperty(PENDING)],
    [false, 'Nothing was created: Google Drive could not be searched for an existing Shared Minutes document (Drive is not available (synthetic)). Try again.', [], null, null]);

  // (h) the recorded document is in the trash, or cannot be opened.
  const h = world();
  h.create();
  const old = h.minutes()[0];
  old.trashed = true;
  h.drive.calls.length = 0;
  const h1 = h.create();
  check('(h) the recorded document is in the trash: nothing is created or shared; the user restores it or forgets it', [h1.ok, /is in the trash, or is not the document that was created for this report\. Restore it, or use "Forget" and create a new one\.$/.test(h1.message), asked(h.drive), h.minutes().length],
    [false, true, ['meta'], 1]);
  h.forget();
  const h2 = h.create();
  check('(h) after Forget a new one is created, shared and linked; the one in the trash is left there', [h2.ok, h2.created, h.minutes().length, old.trashed, h.record().fileId !== old.id, lines(h.body).map((x) => x[2])], [true, true, 2, true, true, [urlOf(h.record().fileId)]]);
  const h3 = world();
  h3.create();
  h3.drive.failMeta = true;
  h3.drive.calls.length = 0;
  const h3r = h3.create();
  check('(h) the recorded document cannot be opened: nothing is created or shared, the record stays', [h3r.ok, asked(h3.drive), !!h3.record().permissionVerifiedAt, h3.minutes().length], [false, ['meta'], true, 1]);

  // (i) the report is configured for another meeting now.
  const i = world();
  i.create();
  const first = i.minutes()[0];
  i.docProps.setProperty('MEETING_ID', '86178');
  i.drive.calls.length = 0;
  const i1 = i.create();
  check('(i) the report was configured for another meeting: the recorded document is not used, and nothing is created until the user says so',
    [i1.ok, i1.message, asked(i.drive), i.view().createDisabled, i.view().canForget, i.view().lines[0].text],
    [false, 'Nothing was created: this report has the shared minutes of meeting 85916 recorded ("' + NAME + '"), and is configured for meeting 86178 now. Use "Forget" first; the document of meeting 85916 is not changed by it.', [], true, true,
      'Status: recorded for meeting 85916 ("' + NAME + '"); this report is configured for meeting 86178 now. It is not used.']);
  i.s.DocumentApp.getActiveDocument = () => ({ getId: () => REPORT_ID, getBody: () => i.body });
  const shownBefore = lines(i.body);
  i.body.insertParagraph(0, 'x');
  check('(i) a build does not write the link of the other meeting', [i.s.finishSharedMinutesRebuild_(makeFakeDocumentBody(i.s)), asked(i.drive)], ['', []]);
  i.forget();
  const i2 = i.create();
  check('(i) after Forget: a new document for the new meeting; the old one is as it was', [i2.ok, i2.created, i.minutes().length, i.record().meetingId, i.record().fileId !== first.id, first.trashed, first.permissions.length, shownBefore.length], [true, true, 2, '86178', true, false, 2, 1]);

  // (j) a value that is not ours.
  const j = world({ props: { SHARED_DOCUMENT: '{"v":1,"fileId":"x"}' } });
  const jView = j.view();
  const j1 = j.create();
  check('(j) a stored value that cannot be read is not used: the button works as if nothing were stored', [jView.lines[0].text, jView.canForget, j1.ok, j1.created, j.record().name], ['Status: Not created', true, true, true, NAME]);
  const newer = world({ props: { SHARED_DOCUMENT: '{"v":2,"fileId":"x","future":true}' } });
  check('(j) a value of a newer release is not used and not changed', [newer.create().ok, asked(newer.drive), newer.docProps.getProperty(KEY), newer.forget().ok, newer.docProps.getProperty(KEY), newer.view().createLabel],
    [false, [], '{"v":2,"fileId":"x","future":true}', false, '{"v":2,"fileId":"x","future":true}', '']);
}

// ------------------------------------------------------------------ 7
console.log('7. Verify and Forget');
{
  const w = world();
  w.create();
  const doc = w.minutes()[0];
  w.drive.calls.length = 0;
  w.clock.ms += 86400000;
  const v = w.verify();
  check('Verify asks Drive for the document and its permissions, and changes neither', [v.ok, v.message, asked(w.drive), mutations(w.drive), w.record().permissionVerifiedAt],
    [true, 'Verified: "' + NAME + '" exists, and anyone with the link can edit it.', ['meta', 'permissions'], [], '2026-10-08T09:00:00.000Z']);
  doc.permissions = doc.permissions.filter((p) => p.type !== 'anyone');
  w.drive.calls.length = 0;
  const v2 = w.verify();
  check('somebody took the sharing away: Verify says so, records it, and does NOT share', [v2.ok, v2.message, mutations(w.drive), w.record().permissionVerifiedAt, w.drive.anyoneCanEdit(doc.id)],
    [false, 'NOT VERIFIED: "' + NAME + '" is not shared so that anyone with the link can edit. Nothing was changed in Drive; press "Finish sharing" to share it.', [], null, false]);
  check('the link that is in the report stays; Configure Meeting shows "Finish sharing"', [lines(w.body).length, w.view().createLabel], [1, 'Finish sharing']);
  const fixed = w.create();
  check('"Finish sharing" shares the same document again', [fixed.ok, fixed.created, w.drive.anyoneCanEdit(doc.id), w.minutes().length, !!w.record().permissionVerifiedAt], [true, false, true, 1, true]);
  w.drive.failMeta = true;
  const before = w.record();
  check('Drive cannot be asked: Verify changes nothing', [w.verify().ok, w.record()], [false, before]);
  w.drive.failMeta = false;
  doc.trashed = true;
  check('the document is in the trash: not verified, and recorded as such', [w.verify().message, w.record().permissionVerifiedAt],
    ['NOT VERIFIED: the Shared Minutes document "' + NAME + '" is in the trash, or is not the document that was created for this report. Nothing was changed in Drive.', null]);
  const none = world();
  check('without a recorded document there is nothing to verify, and Drive is not asked', [none.verify(), asked(none.drive)], [{ ok: false, message: 'There is no Shared Minutes document of this report to verify.' }, []]);
  const unverified = world();
  unverified.drive.failShare = true;
  unverified.create();
  unverified.minutes()[0].permissions.push({ id: 'anyoneWithLink', type: 'anyone', role: 'writer' });
  unverified.drive.calls.length = 0;
  check('a document the user shared by hand: Verify finds the permission, records it and writes the link -- and sets nothing', [unverified.verify().ok, mutations(unverified.drive), !!unverified.record().permissionVerifiedAt, lines(unverified.body).length], [true, [], true, 1]);

  const f = world();
  f.create();
  const kept = f.minutes()[0];
  f.drive.calls.length = 0;
  const forgot = f.forget();
  check('Forget: the record is gone and the link is out of the report; Drive is not asked, the document is not changed', [forgot.ok, f.record(), f.docProps.getProperty(PENDING), lines(f.body), texts(f.body), asked(f.drive), kept.trashed, f.drive.anyoneCanEdit(kept.id)],
    [true, null, null, [], ADHOC_BODY.map(([t, h]) => (h ? '# ' : '') + t), [], false, true]);
  check('and says so', forgot.message, 'This report no longer names a Shared Minutes document. Its link was taken out of the report. The document "' + NAME + '" itself was not changed: it exists and is shared as before. ' +
    'Create Shared Minutes finds and uses it again, unless it is in the trash.');
  const back = f.create();
  check('Create after Forget finds the same document again: one document, never two', [back.ok, back.adopted, f.minutes().length, f.record().fileId], [true, true, 1, kept.id]);
  const typed = world();
  typed.create();
  typed.body._children[2].setText(LABEL + ' see the e-mail of the chair');
  typed.body._children[2].editAsText().setLinkUrl(0, 200, null);
  typed.forget();
  check('a line somebody rewrote by hand (it links nowhere) is not taken out by Forget', lines(typed.body), [[LABEL + ' see the e-mail of the chair', null, null]]);
  check('nothing recorded: Forget does nothing', [world().forget(), world().record()], [{ ok: false, message: 'This report has no Shared Minutes document recorded.' }, null]);
}

// ------------------------------------------------------------------ 8
console.log('8. The line in the report');
{
  const link = { name: NAME, url: urlOf('SHAREDMINUTESdoc0000000000000000000000000001') };
  const other = { name: NAME, url: urlOf('SHAREDMINUTESdoc0000000000000000000000000002') };
  const bodyOf = (paragraphs) => { const w = world({ paragraphs: paragraphs }); return w; };

  // Main meeting: the line the build copies from the meeting report template.
  const MAIN = [['9.1 Opening of the session', H2], ['Meeting name: SA4#137-e'], ['Drafts Folder: https://www.3gpp.org/ftp/x/Inbox/Drafts/Video'], [LABEL], ['Scribes: to be assigned'], ['9.1.1 Registration of Documents', H3], ['9.2 IPR and antitrust reminder', H2]];
  const main = bodyOf(MAIN);
  check('main meeting: the "Link to the shared minutes:" line of the report is filled in where it stands', [main.s.renderSharedMinutesLink_(main.body, link).line, texts(main.body)],
    ['replaced', ['# 9.1 Opening of the session', 'Meeting name: SA4#137-e', 'Drafts Folder: https://www.3gpp.org/ftp/x/Inbox/Drafts/Video', LINE, 'Scribes: to be assigned', '# 9.1.1 Registration of Documents', '# 9.2 IPR and antitrust reminder']]);
  check('the label is text, the name is the link', lines(main.body), [[LINE, null, link.url]]);
  const filled = bodyOf(MAIN.map(([t, h]) => (t === LABEL ? [LABEL + ' https://docs.google.com/document/d/typed-by-hand/edit'] : [t, h])));
  filled.body._children[3].editAsText().setLinkUrl(0, 80, 'https://docs.google.com/document/d/typed-by-hand/edit');
  check('a line that was filled in by hand, and linked as a whole, becomes the label and the link of the document', [filled.s.renderSharedMinutesLink_(filled.body, link).line, lines(filled.body), texts(filled.body).length], ['replaced', [[LINE, null, link.url]], MAIN.length]);
  const variant = bodyOf(MAIN.map(([t, h]) => (t === LABEL ? ['link to the Shared Minutes'] : [t, h])));
  check('the line is found whatever its capitals, and without the colon', [variant.s.renderSharedMinutesLink_(variant.body, link).line, lines(variant.body).length, texts(variant.body).length], ['replaced', 1, MAIN.length]);
  const noLine = bodyOf(MAIN.filter(([t]) => t !== LABEL));
  check('main meeting without that line: a new line directly below the "Opening of the session" heading', [noLine.s.renderSharedMinutesLink_(noLine.body, link).line, texts(noLine.body).slice(0, 3)],
    ['created', ['# 9.1 Opening of the session', LINE, 'Meeting name: SA4#137-e']]);
  const sixG = bodyOf([['11.0 Opening of the session, registration of documents', H2], ['11.0.1 Opening of the session', H3], ['Text of the template'], ['11.0.2 Registration of Documents', H3]]);
  check('the 6G report of a main meeting: below "11.0.1 Opening of the session", not below the section above it', [sixG.s.renderSharedMinutesLink_(sixG.body, link).line, texts(sixG.body).slice(0, 4)],
    ['created', ['# 11.0 Opening of the session, registration of documents', '# 11.0.1 Opening of the session', LINE, 'Text of the template']]);

  // Ad-hoc with sessions: below the Online information.
  const SESSIONS = [['1 Opening of the meeting', H2], ['1.1 Opening of the session', H3], ['Online information', H3], ['Meeting name: Synthetic ad-hoc'], ['Docs Folder: https://ftp.3gpp.org/x/Docs/'], ['Session administration', H3], ['A01: Kick-off'], ['1.2 Registration of Documents', H3]];
  const adhoc = bodyOf(SESSIONS);
  check('ad-hoc report with sessions: the line is the last of the Online information, directly above the Session administration', [adhoc.s.renderSharedMinutesLink_(adhoc.body, link).line, texts(adhoc.body)],
    ['created', ['# 1 Opening of the meeting', '# 1.1 Opening of the session', '# Online information', 'Meeting name: Synthetic ad-hoc', 'Docs Folder: https://ftp.3gpp.org/x/Docs/', LINE, '# Session administration', 'A01: Kick-off', '# 1.2 Registration of Documents']]);
  const block = adhoc.s.findAdhocMeetingInformationContainer_(adhoc.body);
  check('the Online information block itself ends before the line: the block is replaced without it', [block.start, block.end, adhoc.body._children[block.end].getText()], [2, 5, LINE]);
  adhoc.s.renderAdhocMeetingInformation_(adhoc.body, [{ text: 'Meeting name: Renamed', link: '' }, { text: 'Start Date: September 22, 2026', link: '' }, { text: 'Docs Folder: https://ftp.3gpp.org/y/Docs/', link: '' }], false);
  check('when the Online information is written again (the sessions changed), the line stays below it, once', [texts(adhoc.body).slice(2, 8), lines(adhoc.body).length],
    [['# Online information', 'Meeting name: Renamed', 'Start Date: September 22, 2026', 'Docs Folder: https://ftp.3gpp.org/y/Docs/', LINE, '# Session administration'], 1]);
  check('and it is found again there: nothing is written a second time', [adhoc.s.renderSharedMinutesLink_(adhoc.body, link).line, lines(adhoc.body).length], ['unchanged', 1]);

  // Ad-hoc without sessions.
  const plain = bodyOf(ADHOC_BODY);
  check('ad-hoc report without sessions: directly below the "Opening of the session" heading', [plain.s.renderSharedMinutesLink_(plain.body, link).line, texts(plain.body).slice(1, 4)],
    ['created', ['# 1.1 Opening of the session', LINE, '<Chair> opens the session on September 22, 2026 at 15:00 UTC+2.']]);
  check('a new line is ordinary text, not bold', [plain.body._children[2].getHeading(), plain.body._children[2]._bold], ['NORMAL', false]);

  // Again and again; a changed link; no place.
  const written = plain.body._children[2]._links.length;
  check('written again: unchanged -- the text is not set and no link is set', [[1, 2, 3].map(() => plain.s.renderSharedMinutesLink_(plain.body, link).line), plain.body._children[2]._links.length === written, lines(plain.body).length], [['unchanged', 'unchanged', 'unchanged'], true, 1]);
  check('another document with the same name: the same line, the new link, nothing left of the old one', [plain.s.renderSharedMinutesLink_(plain.body, other).line, lines(plain.body)], ['replaced', [[LINE, null, other.url]]]);
  check('another name: the same line, rewritten', [plain.s.renderSharedMinutesLink_(plain.body, { name: 'Renamed Shared Minutes', url: other.url }).line, lines(plain.body).map((l) => l[0]), texts(plain.body).length],
    ['replaced', [LABEL + ' Renamed Shared Minutes'], ADHOC_BODY.length + 1]);
  const empty = bodyOf([['4 Topic', H2], ['Text']]);
  check('a report without an opening section: nothing is written, and that is said', [empty.s.renderSharedMinutesLink_(empty.body, link).line, texts(empty.body)], ['no-place', ['# 4 Topic', 'Text']]);
  const heading = bodyOf([['1.1 Opening of the session', H3], ['Text']]);
  heading.body.insertParagraph(1, LABEL + ' a heading').setHeading(H3);
  check('a heading that begins with the label is not the line', [heading.s.renderSharedMinutesLink_(heading.body, link).line, lines(heading.body).length], ['created', 1]);
  const table = bodyOf(ADHOC_BODY);
  table.body.appendTable([['TDoc', 'S4aA269001'], ['Minutes', LABEL + ' in a table']]);
  table.s.renderSharedMinutesLink_(table.body, link);
  check('text in a TDoc table is never the line', table.body.getTables()[0].getRow(1).getCell(1).getText(), LABEL + ' in a table');

  // The end of a build.
  const build = (record, id) => {
    const w = world({ props: record ? { SHARED_DOCUMENT: typeof record === 'string' ? record : JSON.stringify(record) } : {} });
    w.s.DocumentApp.getActiveDocument = () => ({ getId: () => id || REPORT_ID, getBody: () => w.body });
    const scanned = { n: 0 };
    const real = w.body.getNumChildren;
    w.body.getNumChildren = () => { scanned.n++; return real(); };
    return [w.s.finishSharedMinutesRebuild_(w.body), lines(w.body).length, scanned.n > 0, asked(w.drive)];
  };
  const stored = { v: 1, fileId: 'SHAREDMINUTESdoc0000000000000000000000000001', url: link.url, name: NAME, reportDocumentId: REPORT_ID, meetingId: '85916', createdAt: '2026-10-07T09:00:00.000Z', permissionVerifiedAt: '2026-10-07T09:00:05.000Z' };
  check('the end of a build writes the link of a verified record -- without asking Drive', build(stored), ['', 1, true, []]);
  check('nothing stored: the document is not even looked at', build(null), ['', 0, false, []]);
  check('a record whose sharing is not verified: no link, the document is not looked at', build(Object.assign({}, stored, { permissionVerifiedAt: null })), ['', 0, false, []]);
  check('a record that names another document (a copied value): no link, the document is not looked at', build(stored, COPY_ID), ['', 0, false, []]);
  check('a value that cannot be read: no link', build('not json'), ['', 0, false, []]);
  const nowhere = world({ props: { SHARED_DOCUMENT: JSON.stringify(stored) }, paragraphs: [['4 Topic', H2]] });
  nowhere.s.DocumentApp.getActiveDocument = () => ({ getId: () => REPORT_ID, getBody: () => nowhere.body });
  check('a build of a report without an opening section says that the link could not be written', nowhere.s.finishSharedMinutesRebuild_(nowhere.body), '\n\n⚠️ The link to the shared minutes could not be written: the report has no opening section.');
  check('a failure while writing is a note of the build, never an error', nowhere.s.finishSharedMinutesRebuild_({ getNumChildren: () => { throw new Error('synthetic'); } }),
    '\n\n⚠️ The link to the shared minutes could not be written (synthetic). Use Configure Meeting > Verify to write it.');
}

// ------------------------------------------------------------------ 9
console.log('9. A real report: build, update, trigger, e-mail, rebuild');
{
  const TDOCS = [{ id: 'S4aA269001', agenda: '4.3', status: 'available', uploaded: true }, { id: 'S4aA269002', agenda: '4.4', status: 'agreed', uploaded: true }, { id: 'S4aA269003', agenda: '4.4', status: 'noted', uploaded: true }];
  const drive = makeSharedMinutesDrive({ reportId: REPORT_ID });
  const r = templateReport({ tdocs: TDOCS, drive: drive });
  installExportStubs(r.s);
  const statuses = () => TDOCS.map((t) => [t.id, r.shown(t.id) || r.field(t.id, 'TDoc Status') || r.field(t.id, 'Status')]);

  r.build();
  check('a build of a report without shared minutes: no line, and Drive is not asked about them', [lines(r.body), asked(drive)], [[], []]);
  const outline = r.outline();
  const status = statuses();
  const tables = r.tdocTables().length;
  const bodyBefore = texts(r.body);

  const created = JSON.parse(JSON.stringify(r.exec(() => r.s.createSharedMinutes())));
  const doc = drive.files.filter((f) => f.mimeType === DOC_MIME)[0];
  check('Create Shared Minutes (the button, with the live services): created, named after the title of the report, shared, linked',
    [created.ok, created.state, doc.name, drive.anyoneCanEdit(doc.id), JSON.parse(r.docProps.getProperty(KEY)).fileId, lines(r.body)], [true, 'complete', NAME, true, doc.id, [[LINE, null, urlOf(doc.id)]]]);
  check('the document says its title and the meeting', doc.html, '<html><head><meta charset="utf-8"></head><body><h1>' + NAME + '</h1><p>Synthetic ad-hoc, September 22, 2026</p></body></html>');
  check('the answer carries what the section shows now', [created.view.lines.map((l) => l.text), created.view.openUrl, created.view.createLabel], [[NAME, 'Anyone with the link can edit ✓ (verified ' + new Date().toISOString().slice(0, 10) + ')'], urlOf(doc.id), '']);
  check('the report is the report it was, with one line more: the same headings and TDoc tables, in the same order, with the same statuses', [r.outline(), statuses(), r.tdocTables().length, texts(r.body).filter((t) => t !== LINE)], [outline, status, tables, bodyBefore]);

  const record = r.docProps.getProperty(KEY);
  drive.calls.length = 0;
  r.update();
  r.update();
  check('Update Report, twice: nothing is created, nothing is shared, Drive is not asked about the shared minutes', [mutations(drive), asked(drive)], [[], []]);
  check('and the report has the line once, the record is untouched, the TDoc tables and statuses are as they were', [lines(r.body), r.docProps.getProperty(KEY) === record, r.outline(), statuses()], [[[LINE, null, urlOf(doc.id)]], true, outline, status]);
  r.exec(() => r.s.continuousUpdate());
  check('the scheduled trigger (continuousUpdate()): the same -- nothing created, nothing shared, Drive not asked', [mutations(drive), asked(drive), lines(r.body).length], [[], [], 1]);
  r.exec(() => r.s.addAbstractsOnly());
  check('the abstract sweep: the same', [asked(drive), lines(r.body).length], [[], 1]);
  const meta = (id) => ({ tdoc: id, title: 'Synthetic title of ' + id, agendaItem: '4.3' });
  const mailArgs = ['list@example.org', null, null, null, '15 October 2026, 15:00 CEST', 'https://example.org/drafts', 'reporter@example.com', '26-10-15-1500CEST', 'FS_TEST', urlOf(REPORT_ID), 'list@example.org'];
  const mail = r.exec(() => r.s.buildEmailExportForTdocTable_(r.table('S4aA269001'), meta('S4aA269001'), ...mailArgs));
  const groupMail = r.exec(() => r.s.buildEmailExportForGroup_(['S4aA269002', 'S4aA269003'].map((id) => r.table(id)), ['S4aA269002', 'S4aA269003'].map(meta), ...mailArgs));
  check('discussion e-mails are generated (one TDoc, a group): nothing created, nothing shared, Drive not asked',
    [mail.subject.indexOf('S4aA269001') !== -1, mail.fileName.slice(-4), groupMail.fileName.slice(-4), mail.eml.length > 500, mutations(drive), asked(drive)], [true, '.eml', '.eml', true, [], []]);

  // No shared-minutes function runs in any of them: they throw here, and nothing notices.
  const guarded = ['createSharedMinutes', 'createSharedMinutesWith_', 'verifySharedMinutesWith_', 'forgetSharedMinutesWith_', 'sharedMinutesDrive_', 'runSharedMinutesAction_'];
  const reals = {};
  guarded.forEach((name) => { reals[name] = r.s[name]; r.s[name] = () => { throw new Error('must not run: ' + name); }; });
  let ran = '';
  try {
    r.update();
    r.exec(() => r.s.continuousUpdate());
    r.build();
  } catch (e) { ran = e.message; }
  guarded.forEach((name) => { r.s[name] = reals[name]; });
  check('with every function that creates, shares or asks Drive made to throw, an update, the trigger and a build run as before', [ran, r.logs.filter((l) => /must not run/.test(l))], ['', []]);

  drive.calls.length = 0;
  r.build();
  check('Build Report from Scratch writes the link again from what is stored -- once, in its place, without asking Drive, creating or sharing', [lines(r.body), asked(drive), mutations(drive), texts(r.body).indexOf(LINE) === texts(r.body).indexOf('# 1.1 Opening of the session') + 1],
    [[[LINE, null, urlOf(doc.id)]], [], [], true]);
  check('and the rebuilt report has the same headings, TDoc tables and statuses', [r.outline(), statuses(), drive.files.filter((f) => f.mimeType === DOC_MIME).length], [outline, status, 1]);
  r.build();
  r.update();
  check('built and updated again: still one line, one document', [lines(r.body).length, drive.files.filter((f) => f.mimeType === DOC_MIME).length, asked(drive)], [1, 1, []]);

  // Two presses at once: the second waits for the lock; one that cannot get it does nothing.
  const locked = templateReport({ tdocs: TDOCS, drive: makeSharedMinutesDrive({ reportId: REPORT_ID }) });
  locked.build();
  const lockLog = [];
  locked.s.LockService = { getDocumentLock: () => ({ tryLock: (ms) => { lockLog.push('tryLock ' + ms); return true; }, releaseLock: () => { lockLog.push('release'); } }) };
  locked.exec(() => locked.s.createSharedMinutes());
  locked.exec(() => locked.s.createSharedMinutes());
  check('the button runs under the document lock of the report, and gives it back', lockLog, ['tryLock 30000', 'release', 'tryLock 30000', 'release']);
  check('pressed twice: one document, one line', [locked.drive.files.filter((f) => f.mimeType === DOC_MIME).length, locked.drive.count('create'), locked.drive.count('share'), lines(locked.body).length], [1, 1, 1, 1]);
  const busy = templateReport({ tdocs: TDOCS, drive: makeSharedMinutesDrive({ reportId: REPORT_ID }) });
  busy.build();
  busy.s.LockService = { getDocumentLock: () => ({ tryLock: () => false, releaseLock: () => { throw new Error('a lock that was not taken is not released'); } }) };
  const waited = JSON.parse(JSON.stringify(busy.exec(() => busy.s.createSharedMinutes())));
  check('while another action holds the lock: nothing is done, and that is said', [waited.ok, waited.message, asked(busy.drive), busy.docProps.getProperty(KEY)], [false, 'Nothing was done: another action on this report is running right now. Try again in a minute.', [], null]);
  const thrown = templateReport({ tdocs: TDOCS, drive: makeSharedMinutesDrive({ reportId: REPORT_ID }) });
  thrown.s.LockService = { getDocumentLock: () => { throw new Error('synthetic lock failure'); } };
  check('the button never throws: a failure is an answer, with the section as it is', (({ ok, message, view }) => [ok, message, view.createLabel])(JSON.parse(JSON.stringify(thrown.s.createSharedMinutes()))),
    [false, 'The action stopped (synthetic lock failure). Press the button again: it goes on from what is there.', 'Create Shared Minutes']);

  // Sessions: the real build writes the Online information, and the link below it.
  const sessions = JSON.stringify({ v: 1, nextId: 2, sessions: [{ id: 's1', label: 'Kick-off', date: '2026-09-22', start: '15:00', end: '18:00' }] });
  const withSessions = templateReport({ tdocs: TDOCS, drive: makeSharedMinutesDrive({ reportId: REPORT_ID }), props: { ADHOC_SESSIONS: sessions } });
  withSessions.s.Utilities.formatDate = (date) => date.toISOString().slice(0, 19);
  withSessions.build();
  withSessions.exec(() => withSessions.s.createSharedMinutes());
  const shown = texts(withSessions.body);
  check('an ad-hoc report with sessions (real build): the line is below the last line of the Online information, above the Session administration',
    [shown.indexOf('# Online information') !== -1, shown[shown.indexOf(LINE) - 1].indexOf('Docs Folder: ') === 0, shown[shown.indexOf(LINE) + 1], lines(withSessions.body).length], [true, true, '# Session administration', 1]);
  withSessions.build();
  const rebuilt = texts(withSessions.body);
  check('and a rebuild puts it there again, once', [rebuilt[rebuilt.indexOf(LINE) - 1].indexOf('Docs Folder: ') === 0, rebuilt[rebuilt.indexOf(LINE) + 1], lines(withSessions.body).length, mutations(withSessions.drive).length], [true, '# Session administration', 1, 2]);
}

// ------------------------------------------------------------------ 10
console.log('10. A copied report');
{
  const original = world();
  original.create();
  const first = original.minutes()[0];
  const copiedBody = texts(original.body).map((t) => (t.indexOf('# ') === 0 ? [t.slice(2), /^# \d+ /.test(t) ? H2 : H3] : [t]));

  // A copy: the text of the report, no properties, a script project of its own.
  const copy = world({ reportId: COPY_ID, props: { MEETING_ID: null, MEETING_TYPE: null, MEETING_NAME: null, MEETING_DATE: null, REPORT_SUFFIX: null }, paragraphs: copiedBody, drive: original.drive.asProject('copy-project', COPY_ID) });
  copy.body._children[2].editAsText().setLinkUrl(LABEL.length + 1, LINE.length - 1, urlOf(first.id));
  copy.s.DocumentApp.getActiveDocument = () => ({ getId: () => COPY_ID, getBody: () => copy.body });
  const view = JSON.parse(JSON.stringify(copy.s.currentSharedMinutesView_()));
  check('a copy of a built report starts with no Shared Minutes: "Not created", no Open, no Verify', [view.lines.map((l) => l.text), view.openUrl, view.canVerify, view.canForget], [['Status: Not created', 'Resolve the meeting and save the configuration first.'], '', false, false]);
  check('the link it shows is the text of the report it was copied from, and Configure Meeting says so', view.warning,
    'This report shows a link to shared minutes that is not the one recorded for it (a copied report shows the link of the report it was copied from). It is replaced when the Shared Minutes of this report are created or verified; ' +
    'Build Report from Scratch does not write it again.');
  check('looking at it asks Drive nothing', asked(copy.drive), []);
  check('a build of the copy writes no link (and a real build clears the copied text)', [copy.s.finishSharedMinutesRebuild_(makeFakeDocumentBody(copy.s)), asked(copy.drive)], ['', []]);
  check('Verify and Forget have nothing to work on in the copy: the document of the original is not asked about', [copy.verify().ok, copy.forget().ok, asked(copy.drive)], [false, false, []]);

  // The copy is set up for its own meeting, and creates its own document.
  copy.docProps.setProperty('MEETING_ID', '85916');
  copy.docProps.setProperty('MEETING_NAME', 'Synthetic ad-hoc');
  const made = copy.create();
  const second = copy.minutes().filter((f) => f.id !== first.id)[0];
  check('Create in the copy -- even for the same meeting -- does not find the document of the original: it creates its own',
    [made.ok, made.created, made.adopted, copy.minutes().length, copy.record().fileId === second.id, copy.record().reportDocumentId, second.appProperties.sa4ReportId], [true, true, false, 2, true, COPY_ID, COPY_ID]);
  check('the stale line of the copy is replaced by the link of its own document: one line', lines(copy.body), [[LINE, null, urlOf(second.id)]]);
  check('the document of the original is as it was: its marker, its permissions, its record', [first.appProperties.sa4ReportId, first.permissions.length, first.trashed, original.record().fileId, lines(original.body)],
    [REPORT_ID, 2, false, first.id, [[LINE, null, urlOf(first.id)]]]);
  check('each report still finds only its own document', [original.create().adopted, original.record().fileId, copy.create().adopted, copy.record().fileId, original.minutes().length], [false, first.id, false, second.id, 2]);

  // Should the property travel with a copy after all: it names the original, and is ignored.
  const carried = world({ reportId: COPY_ID, props: { SHARED_DOCUMENT: JSON.stringify(original.record()) }, drive: original.drive.asProject('copy-project-2', COPY_ID) });
  carried.s.DocumentApp.getActiveDocument = () => ({ getId: () => COPY_ID, getBody: () => carried.body });
  check('a SHARED_DOCUMENT value that names another document is ignored: not shown, not linked by a build, not verified',
    [carried.view().lines[0].text, carried.view().openUrl, carried.s.finishSharedMinutesRebuild_(carried.body), lines(carried.body), carried.verify().ok, asked(carried.drive)], ['Status: Not created', '', '', [], false, []]);
  const own = carried.create();
  check('and Create makes the document of the copy; the one of the original is not shared again, renamed or re-marked',
    [own.ok, own.created, carried.record().reportDocumentId, carried.record().fileId !== first.id, carried.drive.calls.filter((c) => c[0] === 'share').map((c) => c[1]).indexOf(first.id), first.permissions.length], [true, true, COPY_ID, true, -1, 2]);
  const same = world({ props: { SHARED_DOCUMENT: JSON.stringify(original.record()) }, drive: original.drive.asProject('copy-project-3', REPORT_ID) });
  check('a record that names a file which is not marked for this report (seen from another script project) is not shared and not linked by Create', [same.create().ok, mutations(same.drive), lines(same.body)], [false, [], []]);
}

// ------------------------------------------------------------------ 11
console.log('11. Configure Meeting');
{
  const CONFIGURED = { MEETING_ID: '85916', MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: 'Audio',
    FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
  const drive = makeSharedMinutesDrive({ reportId: REPORT_ID });
  const r = configurableReport({ props: CONFIGURED, drive: drive });
  const body = r.s.DocumentApp.getActiveDocument().getBody();
  ADHOC_BODY.forEach(([text, heading]) => { const p = body.appendParagraph(text); if (heading) p.setHeading(heading); });
  const page = r.open();
  check('a template report: the dialog has the Shared Minutes section, after the options and before Advanced',
    [/<h3>Shared Minutes<\/h3>/.test(page.html), page.html.indexOf('<h3>3. Options</h3>') < page.html.indexOf('<h3>Shared Minutes</h3>'), page.html.indexOf('<h3>Shared Minutes</h3>') < page.html.indexOf('<details class="advanced">')], [true, true, true]);
  check('opening the dialog asks Drive nothing and stores nothing', [asked(drive), r.docProps.getProperty(KEY), r.docProps.getProperty(PENDING)], [[], null, null]);
  check('it shows "Status: Not created" and the button "Create Shared Minutes" -- no Open, no Verify', [/Status: Not created/.test(page.fields.sharedMinutesBox.innerHTML), /<button type='button' onclick='sharedMinutesCreate\(\)'>Create Shared Minutes<\/button>/.test(page.fields.sharedMinutesBox.innerHTML),
    /Open Shared Minutes|Verify/.test(page.fields.sharedMinutesBox.innerHTML)], [true, true, false]);
  check('the section says when something is created or shared, and when not', /Nothing is created or shared by saving this dialog, by an update or by a build\./.test(page.html), true);

  page.save();
  check('Save Configuration sends nothing about the shared minutes, calls only the save, and Drive is not asked', [Object.keys(page.sent[0]).filter((k) => /shared|minutes/i.test(k)), page.called, asked(drive), r.docProps.getProperty(KEY)], [[], ['saveConfigurationSettings'], [], null]);

  const second = r.open();
  second.press('sharedMinutesCreate');
  const doc = drive.files.filter((f) => f.mimeType === DOC_MIME)[0];
  check('the button calls createSharedMinutes() and nothing else', second.called, ['createSharedMinutes']);
  check('the document is created, shared and linked', [doc.name, drive.anyoneCanEdit(doc.id), JSON.parse(r.docProps.getProperty(KEY)).fileId, lines(body)], [NAME, true, doc.id, [[LINE, null, urlOf(doc.id)]]]);
  const box = second.fields.sharedMinutesBox.innerHTML;
  check('the section then shows the name, "Anyone with the link can edit ✓", Open Shared Minutes and Verify -- and no Create button',
    [box.indexOf("<div style='font-weight:bold'>" + NAME + '</div>') !== -1, /Anyone with the link can edit ✓ \(verified \d{4}-\d{2}-\d{2}\)/.test(box), box.indexOf("<a href='" + urlOf(doc.id) + "' target='_blank'>Open Shared Minutes</a>") !== -1,
      /onclick='sharedMinutesVerify\(\)'>Verify<\/button>/.test(box), /sharedMinutesCreate/.test(box)], [true, true, true, true, false]);
  check('with what was done', /The Shared Minutes document ".*" was created\. Anyone with the link can edit it\. The link is in the report\./.test(box), true);
  second.press('sharedMinutesVerify');
  check('Verify calls verifySharedMinutes(), and the section says "Verified"', [second.called, /Verified: ".*" exists, and anyone with the link can edit it\./.test(second.fields.sharedMinutesBox.innerHTML)], [['createSharedMinutes', 'verifySharedMinutes'], true]);

  const third = r.open();
  check('opened again: the section shows the document straight away, and Drive was not asked for it', [/Anyone with the link can edit ✓/.test(third.fields.sharedMinutesBox.innerHTML), /Create Shared Minutes<\/button>/.test(third.fields.sharedMinutesBox.innerHTML)], [true, false]);
  drive.calls.length = 0;
  third.save();
  check('Save with a Shared Minutes document: still nothing about it is sent, and Drive is not asked', [Object.keys(third.sent[0]).filter((k) => /shared|minutes/i.test(k)), asked(drive), lines(body).length], [[], [], 1]);
  third.press('sharedMinutesForget');
  check('Forget calls forgetSharedMinutes(): the record and the line are gone, the button is back, Drive was not asked',
    [third.called.slice(-1), r.docProps.getProperty(KEY), lines(body), /Create Shared Minutes<\/button>/.test(third.fields.sharedMinutesBox.innerHTML), asked(drive), drive.anyoneCanEdit(doc.id)], [['forgetSharedMinutes'], null, [], true, [], true]);

  const failing = configurableReport({ props: CONFIGURED, drive: (() => { const d = makeSharedMinutesDrive({ reportId: REPORT_ID }); d.failShare = true; return d; })() });
  const failed = failing.open().press('sharedMinutesCreate');
  check('sharing refused: the section says SHARING FAILED, shows "NOT shared", and the button is "Finish sharing"',
    [/SHARING FAILED\./.test(failed.fields.sharedMinutesBox.innerHTML), /NOT shared: "anyone with the link can edit" is not confirmed/.test(failed.fields.sharedMinutesBox.innerHTML), />Finish sharing<\/button>/.test(failed.fields.sharedMinutesBox.innerHTML),
      /Anyone with the link can edit ✓/.test(failed.fields.sharedMinutesBox.innerHTML)], [true, true, true, false]);

  const unset = configurableReport({ props: {}, drive: makeSharedMinutesDrive({ reportId: REPORT_ID }) });
  check('a report without a meeting: the button is there and off', /onclick='sharedMinutesCreate\(\)' disabled>Create Shared Minutes<\/button>/.test(unset.open().fields.sharedMinutesBox.innerHTML), true);

  const master = configurableReport({ props: CONFIGURED, drive: makeSharedMinutesDrive({ reportId: TEMPLATE_ID }) });
  master.s.DocumentApp.getActiveDocument = () => ({ getId: () => TEMPLATE_ID, getBody: () => makeFakeDocumentBody(master.s) });
  const refused = JSON.parse(JSON.stringify(master.s.createSharedMinutes()));
  check('in the template itself the button does nothing', [refused.ok, /This is the SA4 Report Template itself/.test(refused.message), asked(master.drive)], [false, true, []]);

  // Not a template report: CENTRAL and Legacy.
  const central = configurableReport({ props: CONFIGURED, release: false });
  const plain = central.open();
  check('without Release.js (CENTRAL, Legacy) the dialog has no such section, no such script and no such button', [/sharedMinutes|Shared Minutes|shared minutes/.test(plain.html), typeof plain.fields.sharedMinutesView], [false, 'undefined']);
  plain.save();
  check('and its Save is as it was', [Object.keys(plain.sent[0]).filter((k) => /shared|minutes/i.test(k)), plain.called], [[], ['saveConfigurationSettings']]);
}

// ------------------------------------------------------------------ 12
console.log('12. Where the code is, and what did not change');
{
  const names = functionNames(CREATOR);
  const section = CREATOR.slice(CREATOR.indexOf('// Shared Minutes: a separate Google Doc'));
  const mine = functionNames(section);
  check('the feature is one section at the end of ReportCreator.js, with these functions', mine,
    ['sharedMinutesUrl_', 'sharedMinutesName_', 'sharedMinutesHtml_', 'sharedMinutesAppProperties_', 'sharedMinutesQuery_', 'isSharedMinutesFile_', 'hasAnyoneWriterPermission_', 'serializeSharedMinutes_', 'readSharedMinutes_', 'sharedMinutesDrive_',
      'currentSharedMinutesWith_', 'writeSharedMinutesLinkWith_', 'createSharedMinutesWith_', 'verifySharedMinutesWith_', 'forgetSharedMinutesWith_', 'sharedMinutesTextHolder_', 'findSharedMinutesLine_', 'sharedMinutesLineLink_',
      'findSharedMinutesInsertIndex_', 'renderSharedMinutesLink_', 'removeSharedMinutesLine_', 'finishSharedMinutesRebuild_', 'sharedMinutesView_', 'currentSharedMinutesView_', 'sharedMinutesDialogParts_', 'liveSharedMinutesDeps_',
      'runSharedMinutesAction_', 'createSharedMinutes', 'verifySharedMinutes', 'forgetSharedMinutes']);
  const others = names.filter((n) => mine.indexOf(n) === -1);
  check('no other function of ReportCreator.js mentions it', others.filter((n) => /sharedMinutes|SharedMinutes|SHARED_MINUTES|SHARED_DOCUMENT/.test(functionSource(CREATOR, n))), []);
  const uses = (src, pattern) => functionNames(src).filter((n) => pattern.test(functionSource(src, n)));
  check('a permission is created in ONE place, and so is a document for other people: the Drive seam of this feature', [uses(CREATOR, /Permissions\.create/), uses(CREATOR, /Drive\.Permissions/), uses(CODE, /Drive\.Permissions|Permissions\.create|addEditor|setSharing|addViewer/)], [['sharedMinutesDrive_'], ['sharedMinutesDrive_'], []]);
  check('the seam is called to create or share by createSharedMinutesWith_() only', [uses(CREATOR, /drive\.create\(/).filter((n) => mine.indexOf(n) !== -1), uses(CREATOR, /shareWithAnyoneAsWriter\(/)], [['createSharedMinutesWith_'], ['createSharedMinutesWith_']]);
  check('createSharedMinutesWith_() is called by the button only, and the button by nothing but the dialog', [uses(CREATOR, /[^n ]createSharedMinutesWith_\b|\(createSharedMinutesWith_\)/), uses(CREATOR, /\bcreateSharedMinutes\(\)/).filter((n) => n !== 'sharedMinutesDialogParts_' && n !== 'createSharedMinutes'), uses(CODE, /createSharedMinutes|verifySharedMinutes|forgetSharedMinutes/)],
    [['createSharedMinutes'], [], []]);
  check('no menu has an item for it', [/createSharedMinutes|verifySharedMinutes|forgetSharedMinutes/.test(functionSource(CREATOR, 'buildTemplateReportMenu_') + functionSource(CREATOR, 'buildTemplateMasterMenu_') + functionSource(CREATOR, 'addTemplateReportMenuHead_') + functionSource(CODE, 'onOpen'))], [false]);
  check('Code.js reaches the feature in two functions, and only in the template runtime', [uses(CODE, /sharedMinutes|SharedMinutes/), (CODE.match(/templateRuntimeRelease_\(\) \? sharedMinutesDialogParts_\(\) : null/g) || []).length,
    (CODE.match(/if \(templateRuntimeRelease_\(\) && typeof finishSharedMinutesRebuild_ === 'function'\) reallocationRestoreNote \+= finishSharedMinutesRebuild_\(body\);/g) || []).length,
    (CODE.match(/sharedMinutesDialogParts_\(\) :|finishSharedMinutesRebuild_\(body\)/g) || []).length, (CODE.match(/sharedMinutesUi\b/g) || []).length],
    [['configureMeetingSettings', 'buildSkeletonWithTdocTables'], 1, 1, 2, 5]);
  // A build of Code.js with a release marker and without ReportCreator.js (as several suites run it) does not fail for it.
  const alone = require('./helpers/load-code.js').loadCode({ documentProperties: PROPS });
  alone.sandbox.SA4_RELEASE_ = RELEASE;
  check('the build asks whether the function exists before it calls it', [!!alone.sandbox.templateRuntimeRelease_(), typeof alone.sandbox.finishSharedMinutesRebuild_], [true, 'undefined']);
  check('the update and the trigger do not mention it', ['continuousUpdate', 'continuousUpdateCore_', 'updateReportNow', 'saveConfigurationSettings', 'persistConfigurationSettings_', 'onOpen'].filter((n) => /sharedMinutes|SHARED_DOCUMENT/i.test(functionSource(CODE, n) || functionSource(CREATOR, n) || '')), []);
  check('SHARED_DOCUMENT is not a key the dialog saves, and not one that is copied to CENTRAL', [/SHARED_DOCUMENT/.test(CODE.slice(CODE.indexOf('*/')))], [false]);

  // Without the lines of this feature, the two functions of Code.js are the released ones.
  let released = null;
  try { released = execFileSync('git', ['show', '03ec9fc92c9227adfe3342c26ef4dc6b568162c7:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, ''); } catch (e) { released = null; }
  if (!released) {
    console.log('  note: the commit of T-2026.10.10 is not available in this checkout; the comparison with it is skipped.');
  } else {
    const dialog = functionSource(CODE, 'configureMeetingSettings')
      .replace('  // A template report: the Shared Minutes section and its buttons (sharedMinutesDialogParts_(), ReportCreator.js). Opening the\n' +
        '  // dialog asks Drive nothing; Save Configuration sends nothing of it. Everywhere else the dialog is as it was.\n' +
        '  const sharedMinutesUi = templateRuntimeRelease_() ? sharedMinutesDialogParts_() : null;\n', '')
      .replace("    </div>${sharedMinutesUi ? '\\n\\n    ' + sharedMinutesUi.sectionHtml : ''}\n", '    </div>\n')
      .replace('  // The script of the Shared Minutes section goes to the browser with it: the dialog script below stays one static text.\n' +
        "  const readinessEvaluatorSource = evaluateMeetingReadiness_.toString() + (sharedMinutesUi ? '\\n\\n      ' + sharedMinutesUi.script : '');\n", '  const readinessEvaluatorSource = evaluateMeetingReadiness_.toString();\n');
    check('the dialog script itself is the released text: the script of the section travels with the readiness evaluator, the one thing that is put into it',
      [(functionSource(CODE, 'configureMeetingSettings').match(/\$\{readinessEvaluatorSource\}/g) || []).length, /\$\{sharedMinutesUi[^}]*script/.test(functionSource(CODE, 'configureMeetingSettings'))], [1, false]);
    const build = functionSource(CODE, 'buildSkeletonWithTdocTables')
      .replace("  // Shared minutes (template reports): the link to them, from what is stored ('' and nothing done without a\n" +
        '  // verified record of this report). Nothing is created or shared here, and Drive is not asked.\n' +
        "  if (templateRuntimeRelease_() && typeof finishSharedMinutesRebuild_ === 'function') reallocationRestoreNote += finishSharedMinutesRebuild_(body);\n", '');
    check('configureMeetingSettings(): without its three additions it is the function of T-2026.10.10', dialog === functionSource(released, 'configureMeetingSettings'), true);
    check('buildSkeletonWithTdocTables(): without its one addition it is the function of T-2026.10.10', build === functionSource(released, 'buildSkeletonWithTdocTables'), true);
    const all = functionNames(released);
    check('no other function of Code.js changed, none was added, none removed', [all.filter((n) => functionSource(CODE, n) !== functionSource(released, n)), functionNames(CODE).filter((n) => all.indexOf(n) === -1), all.filter((n) => functionNames(CODE).indexOf(n) === -1)],
      [['configureMeetingSettings', 'buildSkeletonWithTdocTables'], [], []]);
  }
  check('Code.js is version 2.22.0, and its first changelog entry is that of the shared minutes', [(CODE.match(/^ \* Version: (\d+\.\d+\.\d+) \((\d{4}-\d{2}-\d{2})\)/m) || []).slice(1), (CODE.match(/^ \* (\d+\.\d+\.\d+) \(\d{4}-\d{2}-\d{2}\)$/gm) || [])[0],
    /^ \* 2\.22\.0 \(2026-10-07\)\n \*   - Added \(template reports\): Shared Minutes\./m.test(CODE)], [['2.22.0', '2026-10-07'], ' * 2.22.0 (2026-10-07)', true]);

  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'appsscript.json'), 'utf8'));
  check('the manifest asks for the scopes it asked for: no new one', manifest.oauthScopes, ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/script.external_request', 'https://www.googleapis.com/auth/script.scriptapp', 'https://www.googleapis.com/auth/script.container.ui']);
  check('and has the Drive service the feature uses, in version 3', manifest.dependencies.enabledAdvancedServices.filter((x) => x.userSymbol === 'Drive').map((x) => x.version), ['v3']);
  check('ReportCreator.js is not in the push of CENTRAL or Legacy: the filter names its files, and this is not one', fs.readFileSync(path.join(ROOT, '.claspignore'), 'utf8').replace(/\r/g, '').replace(/^﻿/, '').split('\n').filter(Boolean),
    ['**/**', '!appsscript.json', '!Code.js', '!HyperLink.js']);
  const bare = require('./helpers/load-code.js').loadCode({ documentProperties: PROPS });
  check('Code.js alone (CENTRAL, Legacy) has none of the functions', ['createSharedMinutes', 'verifySharedMinutes', 'forgetSharedMinutes', 'sharedMinutesDrive_', 'finishSharedMinutesRebuild_', 'sharedMinutesDialogParts_'].filter((n) => typeof bare.sandbox[n] !== 'undefined'), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll shared-minutes checks passed.');
process.exitCode = failures ? 1 : 0;
