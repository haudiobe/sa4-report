/**
 * INTEGRATION TEST ONLY -- issue #28. The jobs of the runner in Issue28Check.js (pure; tested by
 * tests/issue28-jobs.test.js), the rules a result is checked against, and the plan of each copy.
 * run-issue28-validation.mjs pushes a job with a stage, reads its result and stops (and cleans up) on a failure.
 *
 * Steps (a fixed list; Issue28Check.js runs nothing else): probe | props | snap | compare {from} | update |
 *   configureMeeting {meetingId, expectType, reportType, overrides} | saveSessions {rows} | setProps {props} |
 *   candidates | expectRevision {decision} | resolveRevision {outcome, reason} | closeReportCheck | meetingControl |
 *   sessionRulePreview | sessionRuleSet {rule} | carryPreview | carrySet {on} | uninstall
 */

// Rows typed by hand in a TDoc table: an update or a code change must never change them.
const MANUAL_ROWS = /^(discussion|decision|comments?|minutes|notes)$/i;
// Rows a candidate update may not touch at all on the settled copy (any field: text, links, formatting): what was
// typed, and the decisions.
const PROTECTED_ROWS = /^(discussion|decision|comments?|minutes|notes|status|disposition)$/i;

// The copies' meetings, from the 3GPP Portal (looked up 2026-11-25), never a meeting in progress:
// main: SA4#138 Calgary (portal 60780, 2026-11-16..20), family 6G Media -- the meeting of the 6G Media report it copies.
const MAIN_MEETING = { meetingId: 60780, expectType: 'main', reportType: '6G', overrides: { meetingFolder: 'TSGS4_138_Calgary', meetingNumber: '138' } };
// ahg: "3GPPSA4-e (AH) on FS_6G_MED" (portal 86178, 2026-09-22..10-28, ended), family 6G Media. The meeting has two
// agendas, so Discover cannot choose and an update is refused ("Discover or select the agenda document", seen
// 2026-11-25): S4aP260149 is taken -- the latest agenda, for the final call of 26-28 Oct, with the whole schedule,
// the one current when the report was last edited (2026-11-01). S4aP260098 was the agenda of the 22 Sep call.
const AHG_MEETING = { meetingId: 86178, expectType: 'adhoc', reportType: '6G', overrides: { agendaTdoc: 'S4aP260149' } };
// Its sessions as the meeting's own agendas give them (S4aP260098, S4aP260149: "Date | Start Time (CET/CEST) | End
// Time"): the conference call of 22 Sep 2026 and the three-day call of 26-28 Oct 2026, 15:00-18:00 each day. The
// offline meetings of 22 Sep and 5 Oct (S4aP260115) were not sessions of the call. Not taken from the document's tabs.
// The AHG settings as Configure Meeting saved them from the Portal on 2026-11-25 14:14 UTC (job ahg-configure-ahg-
// 20261125T141005), saved here through the report's own Save without the Portal: the Portal's TDoc list for 86178
// served another meeting's list that day, and Meeting Control does not read the list.
const AHG_CONFIG = {
  meetingId: '86178', meetingType: 'adhoc', meetingName: 'SA4-e (AH) on FS_6G_MED', meetingDate: 'September 22, 2026',
  ftpBase: 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/',
  tdocUrl: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178', agendaTdoc: 'S4aP260149',
  agendaCsvUrl: '', revisionsUrl: '', reportType: '6G', meetingFolder: '', meetingNumber: '', mailingList: '',
  mailingListMode: 'derived', agendaSourceDocId: '', showPreview: false
};
const SETTLED_REFERENCE = { main2: 's5' };
const AHG_SESSIONS = [
  { id: '', label: 'Call 1', date: '2026-09-22', start: '15:00', end: '18:00' },
  { id: '', label: 'Call 2', date: '2026-10-26', endDate: '2026-10-28', start: '15:00', end: '18:00' }
];

const snapUpdateTwice = () => [
  { do: 'snap' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
  { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }, { do: 'compare', from: 'settled' }
];
const ok = (...steps) => steps.map((s) => ({ step: s, rule: 'ok' }));

/** The jobs: steps and the checks of their result (index into steps). */
const JOBS = {
  // What the copy is: tabs, TDoc tables, first headings and text, its settings. Reads only.
  inventory: () => ({ steps: [{ do: 'probe' }, { do: 'props' }], expect: ok(0, 1) }),
  // Two snapshots of the unchanged document: no difference (the fingerprint is deterministic).
  determinism: () => ({ steps: [{ do: 'snap' }, { do: 'snap' }, { do: 'compare', from: 'previous' }], expect: ok(0, 1).concat([{ step: 2, rule: 'identical' }]) }),
  // The copy's meeting, as Configure Meeting saves it (Resolve from the Portal, then Save), in the copy only.
  'configure-main': () => ({ steps: [Object.assign({ do: 'configureMeeting' }, MAIN_MEETING), { do: 'markStatusDropdowns' }, { do: 'props' }], expect: ok(0, 1, 2) }),
  'configure-ahg': () => ({ steps: [Object.assign({ do: 'configureMeeting' }, AHG_MEETING), { do: 'saveSessions', rows: AHG_SESSIONS }, { do: 'props' }], expect: ok(0, 1, 2) }),
  // The release the copy runs, updated until an update changes nothing; that snapshot is the reference of the stages
  // (a completed report is not a fixed point of its own release: on 2026-11-26 the first T-2026.10.33 update of the
  // clean main copy added a TDoc registered after the meeting, dropped the meeting-room links the copy has no settings
  // for, and re-rendered the e-mail rows; the second changed nothing). The first update may not change typed text.
  settle: () => ({ steps: [{ do: 'snap' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous', all: true },
    { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }, { do: 'markSettled' }],
  expect: ok(0, 1, 2, 4, 5, 7).concat([{ step: 3, rule: 'manualTextKept' }, { step: 6, rule: 'identical' }]) }),
  // The clean main copy, settled by T-2026.10.33 on 2026-11-26 (run 1: s4 after its first update, s5 after a second
  // update that changed nothing; s6-s8 the same). A further T-2026.10.33 update of it ran into the 6-minute limit, so
  // the reference is taken without one: the copy must still be identical to s5, then it is marked.
  'settled-check': () => ({ steps: [{ do: 'snap' }, { do: 'compare', from: SETTLED_REFERENCE.main2 }, { do: 'markSettled' }],
    expect: ok(0, 2).concat([{ step: 1, rule: 'identical' }]) }),
  // The record of the baseline finding (reads only): the untouched clean copy (s1) against the snapshot before the
  // first T-2026.10.33 update (s3: the same), that update (s3 -> s4), the second (s4 -> s5: nothing), and the whole
  // way from the untouched copy to the settled one (s1 -> s5), every line. A finding about T-2026.10.33, reported on
  // its own; not taken as acceptable.
  'baseline-diff': () => ({ steps: [{ do: 'compare', from: 's1', to: 's3' }, { do: 'compare', from: 's3', to: 's4' }, { do: 'compare', from: 's4', to: 's5' },
    { do: 'compare', from: 's1', to: 's5', all: true }],
  expect: ok(1, 3).concat([{ step: 0, rule: 'identical' }, { step: 2, rule: 'identical' }]) }),
  // The rollback to T-2026.10.33: a stage whose updates are run again once when they end at the time limit, as the
  // report's own trigger would (T-2026.10.33 has no time budget). The candidates keep the strict stage.
  rollback: () => ({ steps: snapUpdateTwice().map((s) => (s.do === 'update' ? Object.assign({ retryOnLimit: 1 }, s) : s)),
    expect: ok(0, 1, 4).concat([{ step: 3, rule: 'protectedKept' }, { step: 6, rule: 'identical' }, { step: 7, rule: 'protectedKept' }]) }),
  // A stage (A, A+B, A+B+C), from the settled copy: update twice. The second update changes nothing; no update and
  // no stage changes a manual row or removes text, a TDoc table or a tab.
  stage: () => ({ steps: snapUpdateTwice(),
    expect: ok(0, 1, 4).concat([{ step: 3, rule: 'protectedKept' }, { step: 6, rule: 'identical' }, { step: 7, rule: 'protectedKept' }]) }),
  // Phase A: Heading 2 class lines on, then off again: only the Disposition Summary changes.
  headings: () => ({ steps: [
    { do: 'snap' }, { do: 'setProps', props: { DISPOSITION_CLASS_HEADINGS: 'heading2' } }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
    { do: 'setProps', props: { DISPOSITION_CLASS_HEADINGS: null } }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }],
  expect: ok(1, 2, 5, 6).concat([{ step: 4, rule: 'manualKept' }, { step: 8, rule: 'manualKept' }, { step: 4, rule: 'onlySummary' }, { step: 8, rule: 'onlySummary' }]) }),
  // Phase B: "revision expected" for a TDoc whose successor is never registered blocks Close Report; closed with an
  // outcome and a reason it is information, not blocking; the updates keep everything typed.
  'phase-b': () => ({ steps: [
    { do: 'candidates' }, { do: 'expectRevision', decision: 'noted' }, { do: 'snap' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
    { do: 'closeReportCheck' }, { do: 'resolveRevision', outcome: 'not-submitted', reason: 'Integration test (issue #28): the successor was never registered.' },
    { do: 'closeReportCheck' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }],
  expect: ok(0, 1, 3, 7, 9).concat([{ step: 5, rule: 'manualKept' }, { step: 5, rule: 'linksKept' }, { step: 6, rule: 'closeBlocks' }, { step: 8, rule: 'closeInformation' }, { step: 11, rule: 'manualKept' }, { step: 11, rule: 'linksKept' }]) }),
  // Phase C on the AHG copy as it is: Meeting Control's time on the real document (281 s before the fix of 2026-11-25).
  // No Portal: the settings saved as they were resolved (AHG_CONFIG), the historical sessions, the Status mark only as
  // the copy's own Status entries are; then Meeting Control twice (each the first step of its execution).
  'mc-benchmark': () => ({ steps: [{ do: 'configure', config: Object.assign({}, AHG_CONFIG) }, { do: 'saveSessions', rows: AHG_SESSIONS.map((s) => Object.assign({}, s)) },
    { do: 'markStatusDropdowns' }, { do: 'props' }, { do: 'meetingControl' }, { do: 'meetingControl' }], expect: ok(0, 1, 2, 3, 4, 5) }),
  // Read-only: the header rows of the tables that are not TDoc tables (why Meeting Control finds no registration).
  'table-headers': () => ({ steps: [{ do: 'tableHeaders' }], expect: ok(0) }),
  // Phase C, main meeting: Meeting Control's data (plenary agenda and schedule sources) is read without errors.
  'meeting-control': () => ({ steps: [{ do: 'meetingControl' }], expect: ok(0) }),
  // Phase C, ad-hoc: Meeting Control; the reservation rule (only if its preview, with the clock check, passes on the live
  // list); carry-forward (only after its preview); both back off; every update keeps what was typed.
  'phase-c': () => ({ steps: [
    { do: 'meetingControl' }, { do: 'snap' },
    { do: 'sessionRulePreview' }, { do: 'sessionRuleSet', rule: 'reservation-6h' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
    { do: 'carryPreview' }, { do: 'carrySet', on: true }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
    { do: 'carrySet', on: false }, { do: 'sessionRuleSet', rule: 'upload' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }],
  expect: ok(0, 2, 3, 4, 7, 8, 9, 12, 13, 14).concat([{ step: 6, rule: 'manualKept' }, { step: 11, rule: 'manualKept' }, { step: 16, rule: 'manualKept' }]) }),
  // One update and the settings, to see why an update does not run (its error is in the result).
  diagnose: () => ({ steps: [{ do: 'props' }, { do: 'update' }], expect: ok(0, 1) }),
  // The end: every trigger, document and script property the tests left removed (the snapshot index kept), proven.
  uninstall: () => ({ steps: [{ do: 'uninstall' }], expect: [{ step: 0, rule: 'clean' }] })
};

/** A job for a copy: id unique per push. */
function makeJob(target, name, stamp) {
  if (!JOBS[name]) throw new Error('unknown job ' + name + ' (' + Object.keys(JOBS).join(', ') + ')');
  const j = JOBS[name]();
  return { id: target + '-' + name + '-' + stamp, target: target, name: name, steps: j.steps, expect: j.expect };
}

/** Pure: the failures of a result against the job's rules ([] when it passes). */
function evaluate(job, result) {
  const failures = [];
  if (!result || !Array.isArray(result.steps)) return ['no result'];
  if (result.stoppedAt) failures.push('stopped after step ' + result.stoppedAt + ' (a step that changes something failed)');
  if (result.steps.length !== job.steps.length && !result.stoppedAt) failures.push('the result has ' + result.steps.length + ' steps, the job ' + job.steps.length);
  job.expect.forEach((e) => {
    const r = result.steps[e.step];
    const at = 'step ' + (e.step + 1) + ' (' + job.steps[e.step].do + ')';
    if (!r) { if (!result.stoppedAt) failures.push(at + ': no result'); return; }
    if (!r.ok) { failures.push(at + ': ' + (r.error || 'failed')); return; }
    if (e.rule === 'identical' && !r.identical) failures.push(at + ': expected no difference, got ' + r.lines + ' line(s): ' + (r.first || []).slice(0, 5).join(' | '));
    if (e.rule === 'manualKept') {
      const rows = (r.counters && r.counters.rows) || {};
      const manual = Object.keys(rows).filter((k) => MANUAL_ROWS.test(k.split(' / ')[0]));
      if (manual.length) failures.push(at + ': manual rows changed: ' + manual.map((k) => k + ' x' + rows[k]).join(', '));
      const other = r.counters && r.counters.other;
      if (other && other.removed) failures.push(at + ': ' + other.removed + ' paragraph(s) outside the TDoc tables removed');
      if (r.counters && r.counters.tdocsRemoved) failures.push(at + ': ' + r.counters.tdocsRemoved + ' TDoc table(s) removed');
      if (r.counters && r.counters.tabs && (r.counters.tabs.removed || r.counters.tabs.added)) failures.push(at + ': tabs added or removed');
    }
    if (e.rule === 'protectedKept' || e.rule === 'linksKept') {
      // protectedKept: no field of a typed or decision row changed, no TDoc table, tab or paragraph outside the tables
      // removed, no link lost. linksKept: only the last (the number of links of the document does not go down).
      const c = r.counters || {};
      if (e.rule === 'protectedKept') {
        const touched = Object.keys(c.rows || {}).filter((k) => PROTECTED_ROWS.test(k.split(' / ')[0]));
        if (touched.length) failures.push(at + ': protected rows changed: ' + touched.map((k) => k + ' x' + c.rows[k]).join(', '));
        if (c.other && c.other.removed) failures.push(at + ': ' + c.other.removed + ' paragraph(s) outside the TDoc tables removed');
        if (c.tdocsRemoved) failures.push(at + ': ' + c.tdocsRemoved + ' TDoc table(s) removed');
        if (c.tabs && (c.tabs.removed || c.tabs.added)) failures.push(at + ': tabs added or removed');
      }
      const links = /^(\d+) -> (\d+)$/.exec(String((c.counts || {}).links || ''));
      if (links && Number(links[2]) < Number(links[1])) failures.push(at + ': links lost (' + links[1] + ' -> ' + links[2] + ')');
    }
    if (e.rule === 'manualTextKept') {
      // Typed text and links of the manual rows kept, no TDoc table, tab or paragraph outside the tables removed; a
      // change of paragraph settings alone (the release's own formatting) is allowed here and shown in the result.
      const fields = (r.counters && r.counters.rows) || {};
      const typed = Object.keys(fields).filter((k) => MANUAL_ROWS.test(k.split(' / ')[0]) && !/ \/ formatting$/.test(k));
      if (typed.length) failures.push(at + ': manual rows changed: ' + typed.map((k) => k + ' x' + fields[k]).join(', '));
      const other = r.counters && r.counters.other;
      if (other && other.removed) failures.push(at + ': ' + other.removed + ' paragraph(s) outside the TDoc tables removed');
      if (r.counters && r.counters.tdocsRemoved) failures.push(at + ': ' + r.counters.tdocsRemoved + ' TDoc table(s) removed');
      if (r.counters && r.counters.tabs && (r.counters.tabs.removed || r.counters.tabs.added)) failures.push(at + ': tabs added or removed');
    }
    if (e.rule === 'onlySummary' && r.counters) {
      const c = r.counters;
      const outside = Object.keys(c.rows || {}).length + c.tdocsAdded + c.tdocsRemoved + (c.other ? c.other.changed + c.other.added + c.other.removed : 0);
      if (outside) failures.push(at + ': changes outside the Disposition Summary (' + outside + ')');
    }
    if (e.rule === 'closeBlocks' && !r.tdocBlocking) failures.push(at + ': the unresolved expected revision of ' + r.tdoc + ' does not block Close Report');
    if (e.rule === 'closeInformation' && (r.tdocBlocking || !r.tdocInformation)) failures.push(at + ': the closed record of ' + r.tdoc + ' is ' + (r.tdocBlocking ? 'still blocking' : 'not listed as information'));
    if (e.rule === 'clean' && !r.clean) failures.push(at + ': cleanup incomplete');
  });
  return failures;
}

/** Pure: a result from the description chunks of its Drive files ("IT28 result <id> <i>/<n>"); null while incomplete. */
function assemble(id, files) {
  const parts = {};
  let n = 0;
  files.forEach((f) => {
    const m = new RegExp('^IT28 result ' + id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ' (\\d+)/(\\d+)\\.txt$').exec(f.name || '');
    if (!m) return;
    parts[Number(m[1])] = f.description || '';
    n = Number(m[2]);
  });
  if (!n) return null;
  for (let i = 1; i <= n; i++) if (parts[i] === undefined) return null;
  let text = '';
  for (let i = 1; i <= n; i++) text += parts[i];
  return JSON.parse(text);
}

/** The agreed sequence for each copy: the stage pushed with each job. After the last job the original script goes back. */
const PLAN = {
  // The clean copy of 2026-11-25 (the first main copy was changed by the killed updates and is kept as evidence).
  main2: [['base', 'inventory'], ['base', 'determinism'], ['base', 'configure-main'], ['base', 'settled-check'], ['base', 'baseline-diff'],
    ['A', 'stage'], ['A', 'headings'], ['B', 'stage'], ['B', 'phase-b'], ['C', 'stage'], ['C', 'meeting-control'],
    ['t33', 'rollback'], ['t33', 'uninstall']],
  main: [['base', 'inventory'], ['base', 'determinism'], ['base', 'configure-main'], ['base', 'stage'],
    ['A', 'stage'], ['A', 'headings'], ['B', 'stage'], ['B', 'phase-b'], ['C', 'stage'], ['C', 'meeting-control'],
    ['t33', 'stage'], ['t33', 'uninstall']],
  ahg: [['t33', 'inventory'], ['t33', 'determinism'], ['t33', 'configure-ahg'], ['t33', 'settle'],
    ['A', 'stage'], ['B', 'stage'], ['C', 'stage'], ['C', 'phase-c'], ['t33', 'stage'], ['t33', 'uninstall']]
};

module.exports = { JOBS, MANUAL_ROWS, PROTECTED_ROWS, MAIN_MEETING, AHG_MEETING, AHG_CONFIG, AHG_SESSIONS, SETTLED_REFERENCE, PLAN, makeJob, evaluate, assemble };
