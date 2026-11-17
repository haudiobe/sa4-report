/**
 * Operating window of the local listing agent (docs/LOCAL_LISTING_AGENT.md, "Operating window and
 * overnight staleness"): scans only 07:30-19:30 Calgary time, daily; a run outside it is a normal
 * pause (exit 0, no file written).
 *
 *   1. installer: daily trigger, no 30-day expiry, a running scan is not stopped at the window end
 *   2. window boundaries in winter (MST) and summer (MDT), and on both daylight-saving change days
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
const INSTALLER = fs.readFileSync(path.join(DIR, 'install-scheduled-task.ps1'), 'utf8');
const EXAMPLE = JSON.parse(fs.readFileSync(path.join(DIR, 'agent.config.example.json'), 'utf8'));

console.log('1. installer and example configuration');
check('daily trigger', /New-ScheduledTaskTrigger -Daily/.test(INSTALLER), true);
check('no repetition lasting days (the old 30-day expiry)', /RepetitionDuration \(New-TimeSpan -Days/.test(INSTALLER), false);
check('no EndBoundary set', /EndBoundary/.test(INSTALLER), false);
check('a running scan is not stopped at the window end', /Repetition\.StopAtDurationEnd = \$false/.test(INSTALLER), true);
check('same-zone trigger is written without a UTC offset (follows daylight saving)', /StartBoundary = \$first\.ToString\("yyyy-MM-dd'T'HH:mm:ss"\)/.test(INSTALLER), true);
check('other-zone trigger is one hour wider on each side', /AddHours\(-1\)/.test(INSTALLER) && /New-TimeSpan -Hours 2/.test(INSTALLER), true);
check('registration replaces the task in place (-Force)', /Register-ScheduledTask [^\n]*-Force/.test(INSTALLER), true);
check('example window', EXAMPLE.operatingWindow, { timeZone: 'Mountain Standard Time', start: '07:30', end: '19:30' });

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
function run(cfg, args) {
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', AGENT, '-ConfigPath', cfg].concat(args),
    { encoding: 'utf8', env: Object.assign({}, process.env, { LOCALAPPDATA: tmp }), timeout: 60000 });
  return { code: r.status, out: (r.stdout || '') + (r.stderr || '') };
}
const verdict = (utc) => { const r = run(CFG, ['-CheckWindow', '-AtUtc', utc]); return (r.out.match(/: (inside|outside) the operating window/) || [])[1] + ' ' + r.code; };

console.log('\n2. window boundaries (Calgary local time in brackets)');
const cases = [
  ['winter, start  (07:29 MST)', '2026-11-17T14:29', 'outside'],
  ['winter, start  (07:30 MST)', '2026-11-17T14:30', 'inside'],
  ['winter, end    (19:29 MST)', '2026-11-18T02:29', 'inside'],
  ['winter, end    (19:30 MST)', '2026-11-18T02:30', 'outside'],
  ['winter, night  (23:59 MST)', '2026-11-18T06:59', 'outside'],
  ['winter, night  (00:00 MST)', '2026-11-18T07:00', 'outside'],
  ['summer, start  (07:29 MDT)', '2027-07-01T13:29', 'outside'],
  ['summer, start  (07:30 MDT)', '2027-07-01T13:30', 'inside'],
  ['summer, end    (19:29 MDT)', '2027-07-02T01:29', 'inside'],
  ['summer, end    (19:30 MDT)', '2027-07-02T01:30', 'outside'],
  // DST ends Sunday 2026-11-01 02:00 MDT -> 01:00 MST
  ['day before DST end, start (07:30 MDT)', '2026-10-31T13:30', 'inside'],
  ['day before DST end, 06:30 MST-equivalent (06:30 MDT)', '2026-10-31T12:30', 'outside'],
  ['DST end day, start (07:29 MST)', '2026-11-01T14:29', 'outside'],
  ['DST end day, start (07:30 MST)', '2026-11-01T14:30', 'inside'],
  ['DST end day, end   (19:29 MST)', '2026-11-02T02:29', 'inside'],
  ['DST end day, end   (19:30 MST)', '2026-11-02T02:30', 'outside'],
  // DST starts Sunday 2027-03-14 02:00 MST -> 03:00 MDT
  ['day before DST start, start (07:30 MST)', '2027-03-13T14:30', 'inside'],
  ['DST start day, start (07:29 MDT)', '2027-03-14T13:29', 'outside'],
  ['DST start day, start (07:30 MDT)', '2027-03-14T13:30', 'inside'],
  ['DST start day, end   (19:29 MDT)', '2027-03-15T01:29', 'inside'],
  ['DST start day, end   (19:30 MDT)', '2027-03-15T01:30', 'outside'],
];
for (const [name, utc, expected] of cases) check(`${name} = ${utc}Z: ${expected}`, verdict(utc), expected + ' 0');

console.log('\n3. a run outside the window is a pause, not an error');
const night = run(CFG, ['-AtUtc', '2026-11-18T10:00']);
check('exit code 0', night.code, 0);
check('"paused" logged, resumes at 07:30', /paused: 2026-11-18 03:00 Mountain Standard Time is outside the operating window 07:30-19:30; scanning resumes at 07:30/.test(night.out), true);
check('output folder untouched (no listing, no status file)', fs.readdirSync(outDir), []);
check('the pause is in the agent log', /paused:/.test(fs.readFileSync(path.join(tmp, 'SA4LocalListingAgent', 'agent.log'), 'utf8')), true);

console.log('\n4. inside the window the scan proceeds; a broken window is an error');
const MISSING = config('missing-out.json', { outDir: path.join(tmp, 'does-not-exist'), operatingWindow: EXAMPLE.operatingWindow });
const day = run(MISSING, ['-AtUtc', '2026-11-17T18:00']);
check('inside the window: passes the gate (stops at the missing output folder, exit 2)', [day.code, /scan not run: output folder/.test(day.out)], [2, true]);
const forced = run(MISSING, ['-AtUtc', '2026-11-18T10:00', '-IgnoreWindow']);
check('-IgnoreWindow at night: passes the gate (exit 2 as above)', [forced.code, /scan not run: output folder/.test(forced.out)], [2, true]);
const noWindow = config('no-window.json', {});
check('no operatingWindow in the config: the same default window', verdict.call(null, '2026-11-17T14:29') === 'outside 0' && run(noWindow, ['-CheckWindow', '-AtUtc', '2026-11-17T14:29']).out.includes('outside the operating window 07:30-19:30'), true);
const badZone = run(config('bad-zone.json', { operatingWindow: { timeZone: 'Nowhere/Calgary' } }), ['-AtUtc', '2026-11-17T18:00']);
check('unknown time zone: exit 2, "operatingWindow is invalid"', [badZone.code, /operatingWindow is invalid/.test(badZone.out)], [2, true]);
const badOrder = run(config('bad-order.json', { operatingWindow: { start: '19:30', end: '07:30' } }), ['-AtUtc', '2026-11-17T18:00']);
check('start after end: exit 2', [badOrder.code, /start must be before end/.test(badOrder.out)], [2, true]);
check('nothing written to the output folder by any run', fs.readdirSync(outDir), []);

fs.rmSync(tmp, { recursive: true, force: true });
console.log(failures ? `\n${failures} local listing agent window check(s) FAILED.` : '\nAll local listing agent window checks passed.');
process.exit(failures ? 1 : 0);
