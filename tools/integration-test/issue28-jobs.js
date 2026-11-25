/**
 * INTEGRATION TEST ONLY -- issue #28. The jobs of the runner in Issue28Check.js (pure; tested by
 * tests/issue28-jobs.test.js) and the rules a result is checked against. run-issue28-validation.mjs pushes a job with
 * a stage and reads its result.
 *
 * Steps: probe | props | snap | compare {from: previous|baseline|<label>} | update | configure {config} |
 *        setProps {props} | call {fn, args} | uninstall
 */

// Rows typed by hand in a TDoc table: an update or a code change must never change them.
const MANUAL_ROWS = /^(discussion|decision|comments?|minutes|notes)$/i;

// The SA4#138 main meeting of the copy of the 6G Media report (a completed meeting; the copy only).
const MAIN_CONFIG = { meetingType: 'main', meetingNumber: '138', meetingFolder: 'TSGS4_138_Calgary', reportType: '6G' };

const snapUpdateTwice = () => [
  { do: 'snap' }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
  { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }, { do: 'compare', from: 'baseline' }
];

/** The jobs: steps and the checks of their result (index into steps). */
const JOBS = {
  // What the copy is: tabs, TDoc tables, first headings and text, its settings. Reads only.
  inventory: () => ({ steps: [{ do: 'probe' }, { do: 'props' }], expect: [{ step: 0, rule: 'ok' }, { step: 1, rule: 'ok' }] }),
  // Two snapshots of the unchanged document: no difference (the fingerprint is deterministic).
  determinism: () => ({ steps: [{ do: 'snap' }, { do: 'snap' }, { do: 'compare', from: 'previous' }], expect: [{ step: 2, rule: 'identical' }] }),
  // The main copy's meeting, saved in its document properties (as Configure Meeting > Save does).
  'configure-main': () => ({ steps: [{ do: 'configure', config: MAIN_CONFIG }, { do: 'props' }], expect: [{ step: 0, rule: 'ok' }, { step: 1, rule: 'ok' }] }),
  // A stage (T-2026.10.33, A, B, C, the rollback): update twice. The second update changes nothing (idempotent);
  // no update and no stage changes the manual rows or removes text outside the TDoc tables.
  stage: () => ({ steps: snapUpdateTwice(),
    expect: [{ step: 1, rule: 'ok' }, { step: 4, rule: 'ok' }, { step: 3, rule: 'manualKept' }, { step: 6, rule: 'identical' }, { step: 7, rule: 'manualKept' }] }),
  // Phase A: Heading 2 class lines switched on, then off again; the summary is rewritten in place, nothing else changes.
  headings: () => ({ steps: [
    { do: 'snap' }, { do: 'setProps', props: { DISPOSITION_CLASS_HEADINGS: 'heading2' } }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' },
    { do: 'setProps', props: { DISPOSITION_CLASS_HEADINGS: null } }, { do: 'update' }, { do: 'snap' }, { do: 'compare', from: 'previous' }],
  expect: [{ step: 2, rule: 'ok' }, { step: 6, rule: 'ok' }, { step: 4, rule: 'manualKept' }, { step: 8, rule: 'manualKept' }, { step: 4, rule: 'onlySummary' }, { step: 8, rule: 'onlySummary' }] }),
  // The end: the runner removes its trigger (before the original script goes back).
  uninstall: () => ({ steps: [{ do: 'uninstall' }], expect: [{ step: 0, rule: 'ok' }] })
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
  if (result.steps.length !== job.steps.length) failures.push('the result has ' + result.steps.length + ' steps, the job ' + job.steps.length);
  job.expect.forEach((e) => {
    const r = result.steps[e.step];
    const at = 'step ' + (e.step + 1) + ' (' + job.steps[e.step].do + ')';
    if (!r) { failures.push(at + ': no result'); return; }
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
    if (e.rule === 'onlySummary' && r.counters) {
      const c = r.counters;
      const outside = Object.keys(c.rows || {}).length + c.tdocsAdded + c.tdocsRemoved + (c.other ? c.other.changed + c.other.added + c.other.removed : 0);
      if (outside) failures.push(at + ': changes outside the Disposition Summary (' + outside + ')');
    }
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

/** The sequence for a copy: the stage pushed with each job. */
const PLAN = {
  main: [['base', 'inventory'], ['base', 'determinism'], ['base', 'configure-main'], ['base', 'stage'], ['A', 'stage'], ['A', 'headings'], ['B', 'stage'], ['C', 'stage'], ['t33', 'stage'], ['t33', 'uninstall']],
  ahg: [['t33', 'inventory'], ['t33', 'determinism']]
};

module.exports = { JOBS, MANUAL_ROWS, MAIN_CONFIG, PLAN, makeJob, evaluate, assemble };
