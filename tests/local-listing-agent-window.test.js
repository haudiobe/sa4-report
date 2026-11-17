/**
 * Operating window of the local listing agent (docs/LOCAL_LISTING_AGENT.md, "Operating window and
 * overnight staleness"): scans only 07:30-19:30 Calgary time, daily; a run outside it is a normal
 * pause (exit 0, no file written). Calgary is at UTC-06:00 (ABT) all year from 1 November 2026
 * (docs/MEETING_TIME_ZONES.md), so the default window is a fixed offset, independent of the laptop's
 * Windows time-zone rules (which may still send 'Mountain Standard Time' to UTC-7 in winter).
 *
 *   1. installer: daily trigger, no 30-day expiry, a running scan is not stopped at the window end,
 *      a fixed offset is registered fixed in UTC; example configuration
 *   2. window boundaries: the default (ABT), every season, around 1 November 2026; the Windows-rules
 *      mode ("timeZone") with daylight saving; the configuration installed before this fix
 *   3. a real run outside the window: exit 0, "paused" logged, output folder untouched
 *   4. a run inside the window (or with -IgnoreWindow) passes the gate; a broken window is an error
 *
 * Parts 2-4 run the PowerShell agent and need Windows (powershell.exe); elsewhere they are skipped.
 * Nothing is fetched: every run in 3-4 stops before any network access. All values are synthetic.
 * Run: node tests/local-listing-agent-window.test.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`); else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const DIR = path.join(__dirname, '..', 'tools', 'local-listing-agent');
const AGENT = path.join(DIR, 'sa4-local-listing-agent.ps1');
const AGENT_TEXT = fs.readFileSync(AGENT, 'utf8').replace(/\r/g, '');
const INSTALLER = fs.readFileSync(path.join(DIR, 'install-scheduled-task.ps1'), 'utf8').replace(/\r/g, '');
const EXAMPLE = JSON.parse(fs.readFileSync(path.join(DIR, 'agent.config.example.json'), 'utf8'));

console.log('1. installer and example configuration');
check('daily trigger', /New-ScheduledTaskTrigger -Daily/.test(INSTALLER), true);
check('no repetition lasting days (the old 30-day expiry)', /RepetitionDuration \(New-TimeSpan -Days/.test(INSTALLER), false);
check('no EndBoundary set', /EndBoundary/.test(INSTALLER), false);
check('a running scan is not stopped at the window end', /Repetition\.StopAtDurationEnd = \$false/.test(INSTALLER), true);
check('fixed offset: registered fixed in UTC at <start>, for exactly (end - start)', /if \(\$window\.fixed\) \{[^}]*BaseUtcOffset[^}]*\$duration = \$end - \$start[^}]*StartBoundary = \$firstFixed\.ToString\("yyyy-MM-dd'T'HH:mm:sszzz"\)/.test(INSTALLER), true);
check('Windows zone, same as the laptop: written without a UTC offset (follows daylight saving)', /StartBoundary = \$first\.ToString\("yyyy-MM-dd'T'HH:mm:ss"\)/.test(INSTALLER), true);
check('Windows zone, other than the laptop: one hour wider on each side', /AddHours\(-1\)/.test(INSTALLER) && /New-TimeSpan -Hours 2/.test(INSTALLER), true);
check('registration replaces the task in place (-Force)', /Register-ScheduledTask [^\n]*-Force/.test(INSTALLER), true);
const resolverOf = (text) => text.slice(text.indexOf('function Resolve-OperatingWindow'), text.indexOf('\n}\n', text.indexOf('function Resolve-OperatingWindow')));
check('agent and installer resolve the window with the same code', resolverOf(AGENT_TEXT) === resolverOf(INSTALLER) && resolverOf(AGENT_TEXT).length > 200, true);
check('example window: 07:30-19:30 at UTC-06:00 (ABT)', EXAMPLE.operatingWindow, { utcOffset: '-06:00', label: 'ABT', start: '07:30', end: '19:30' });

if (process.platform !== 'win32') {
  console.log('\n(parts 2-4 skipped: they need Windows PowerShell)');
  process.exit(failures ? 1 : 0);
}

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'sa4-window-'));
const outDir = path.join(tmp, 'out');
fs.mkdirSync(outDir);
function config(name, extra) {
  const p = path.join(tmp, name);
  fs.writeFileSync(p, JSON.stringify(Object.assign({ baseUrl: 'http://192.0.2.1/ftp/SA/SA4/Inbox/Drafts/', outDir: outDir, groups: ['MBS'], timeoutSeconds: 1 }, extra)));
  return p;
}
const CFG = config('agent.config.json', { operatingWindow: EXAMPLE.operatingWindow });
const DEFAULT = config('no-window.json', {});
const WINDOWS_MST = config('windows-mst.json', { operatingWindow: { timeZone: 'Mountain Standard Time', start: '07:30', end: '19:30' } });
function run(cfg, args) {
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', AGENT, '-ConfigPath', cfg].concat(args),
    { encoding: 'utf8', env: Object.assign({}, process.env, { LOCALAPPDATA: tmp }), timeout: 60000 });
  return { code: r.status, out: (r.stdout || '') + (r.stderr || '') };
}
const verdict = (cfg, utc) => { const r = run(cfg, ['-CheckWindow', '-AtUtc', utc]); return (r.out.match(/: (inside|outside) the operating window/) || [])[1] + ' ' + r.code; };

console.log('\n2a. window boundaries, default and example configuration: ABT (UTC-6), Calgary time in brackets');
const abt = [
  ['autumn, start  (07:29 ABT)', '2026-11-17T13:29', 'outside'],
  ['autumn, start  (07:30 ABT)', '2026-11-17T13:30', 'inside'],
  ['autumn, end    (19:29 ABT)', '2026-11-18T01:29', 'inside'],
  ['autumn, end    (19:30 ABT)', '2026-11-18T01:30', 'outside'],
  ['night          (23:59 ABT)', '2026-11-18T05:59', 'outside'],
  ['night          (00:00 ABT)', '2026-11-18T06:00', 'outside'],
  ['1 Nov 2026, the day of Alberta\'s change, start (07:30 ABT)', '2026-11-01T13:30', 'inside'],
  ['1 Nov 2026, end (19:30 ABT)', '2026-11-02T01:30', 'outside'],
  ['winter, start  (07:29 ABT)', '2027-01-15T13:29', 'outside'],
  ['winter, start  (07:30 ABT)', '2027-01-15T13:30', 'inside'],
  ['14 Mar 2027 (no change any more), start (07:30 ABT)', '2027-03-14T13:30', 'inside'],
  ['summer, end    (19:29 ABT)', '2027-07-02T01:29', 'inside'],
  ['summer, end    (19:30 ABT)', '2027-07-02T01:30', 'outside'],
];
for (const [name, utc, expected] of abt) check(`${name} = ${utc}Z: ${expected}`, [verdict(CFG, utc), verdict(DEFAULT, utc)], [expected + ' 0', expected + ' 0']);
check('said in ABT', run(CFG, ['-CheckWindow', '-AtUtc', '2026-11-17T13:30']).out.trim(), '2026-11-17T13:30Z = 2026-11-17 07:30 ABT (UTC-06:00): inside the operating window 07:30-19:30');

console.log('\n2b. Windows-rules mode ("timeZone" without "utcOffset"), with daylight saving');
// With this laptop's Windows rules for 'Mountain Standard Time'. Where they predate Alberta's change, they still go to
// UTC-7 in winter -- the reason the default is a fixed offset. Only checked where Windows' rules are the old ones.
const oldRules = verdict(WINDOWS_MST, '2026-11-17T13:30') === 'outside 0';
if (oldRules) {
  check('Windows rules predating Alberta\'s change: 07:30 ABT is 06:30 there, outside (the configuration installed before this fix)', verdict(WINDOWS_MST, '2026-11-17T13:30'), 'outside 0');
  check('... and its window ends at 20:30 ABT: 20:29 ABT inside', verdict(WINDOWS_MST, '2026-11-18T02:29'), 'inside 0');
}
const dst = [
  ['summer (07:29 MDT)', '2027-07-01T13:29', 'outside'],
  ['summer (07:30 MDT)', '2027-07-01T13:30', 'inside'],
  ['summer (19:29 MDT)', '2027-07-02T01:29', 'inside'],
  ['summer (19:30 MDT)', '2027-07-02T01:30', 'outside'],
];
for (const [name, utc, expected] of dst) check(`${name} = ${utc}Z: ${expected}`, verdict(WINDOWS_MST, utc), expected + ' 0');

console.log('\n3. a run outside the window is a pause, not an error');
const night = run(CFG, ['-AtUtc', '2026-11-18T10:00']);
check('exit code 0', night.code, 0);
check('"paused" logged, resumes at 07:30', /paused: 2026-11-18 04:00 ABT \(UTC-06:00\) is outside the operating window 07:30-19:30; scanning resumes at 07:30/.test(night.out), true);
check('output folder untouched (no listing, no status file)', fs.readdirSync(outDir), []);
check('the pause is in the agent log', /paused:/.test(fs.readFileSync(path.join(tmp, 'SA4LocalListingAgent', 'agent.log'), 'utf8')), true);

console.log('\n4. inside the window the scan proceeds; a broken window is an error');
const MISSING = config('missing-out.json', { outDir: path.join(tmp, 'does-not-exist'), operatingWindow: EXAMPLE.operatingWindow });
// Past the gate, a run stops at the missing output folder (exit 2) -- or, when the real scheduled agent of this laptop is
// scanning at that moment, at the machine-wide one-run-at-a-time lock (exit 0, "skipped"). Both come after the window
// check; a run stopped by the window would say "paused".
const pastGate = (r) => !/paused:/.test(r.out) && ((r.code === 2 && /scan not run: output folder/.test(r.out)) || (r.code === 0 && /another scan is still running; skipped/.test(r.out)));
const day = run(MISSING, ['-AtUtc', '2026-11-17T18:00']);
check('inside the window: passes the gate (stops at the missing output folder, or at the lock of a live agent)', pastGate(day), true);
const forced = run(MISSING, ['-AtUtc', '2026-11-18T10:00', '-IgnoreWindow']);
check('-IgnoreWindow at night: passes the gate (as above)', pastGate(forced), true);
const badZone = run(config('bad-zone.json', { operatingWindow: { timeZone: 'Nowhere/Calgary' } }), ['-AtUtc', '2026-11-17T18:00']);
check('unknown Windows time zone: exit 2, "operatingWindow is invalid"', [badZone.code, /operatingWindow is invalid/.test(badZone.out)], [2, true]);
const badOffset = run(config('bad-offset.json', { operatingWindow: { utcOffset: '-6' } }), ['-AtUtc', '2026-11-17T18:00']);
check('malformed utcOffset: exit 2', [badOffset.code, /utcOffset must look like -06:00/.test(badOffset.out)], [2, true]);
const badOrder = run(config('bad-order.json', { operatingWindow: { start: '19:30', end: '07:30' } }), ['-AtUtc', '2026-11-17T18:00']);
check('start after end: exit 2', [badOrder.code, /must be before end/.test(badOrder.out)], [2, true]);
check('nothing written to the output folder by any run', fs.readdirSync(outDir), []);

fs.rmSync(tmp, { recursive: true, force: true });
console.log(failures ? `\n${failures} local listing agent window check(s) FAILED.` : '\nAll local listing agent window checks passed.');
process.exit(failures ? 1 : 0);
