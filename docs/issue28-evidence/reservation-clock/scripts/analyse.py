"""Analyse reservation/uploaded clocks.
usage: python -I analyse.py <data dir> <head_138.tsv>
"""
import csv, sys, os, collections, datetime as dt
from zoneinfo import ZoneInfo

D, HEADF = sys.argv[1:3]
UTC = dt.timezone.utc
PARIS = ZoneInfo('Europe/Paris')
FIXED = lambda h: dt.timezone(dt.timedelta(hours=h))
MEET = {
    'sa4_138': ('Calgary (UTC-6 fixed; Alberta ABT from 2026-11-01, MDT before)', ZoneInfo('America/Edmonton')),
    'TSGS4_136_Montreal': ('Montreal (EDT UTC-4)', ZoneInfo('America/Toronto')),
    'TSGS4_135_India': ('India (IST UTC+5:30)', ZoneInfo('Asia/Kolkata')),
    'TSGS4_137-e': ('e-meeting (no location)', None),
}
# Note: tzdata may still model Alberta with DST (fall back 1 Nov 2026). Alberta adopted
# permanent UTC-6; for H3 we use a fixed UTC-6 for SA4#138 to follow the law.
MEET['sa4_138'] = (MEET['sa4_138'][0], FIXED(-6))

def P(s):
    return dt.datetime.fromisoformat(s) if s else None

def to_utc(naive, tz):
    return naive.replace(tzinfo=tz).astimezone(UTC).replace(tzinfo=None)

def pct(xs, q):
    xs = sorted(xs); return xs[min(len(xs) - 1, int(q * len(xs)))]

out = []
def pr(*a):
    s = ' '.join(str(x) for x in a); print(s); out.append(s)

# ---- (a) Uploaded vs HEAD Last-Modified (SA4#138 sample) and vs listing (all)
pr('## (a) Uploaded minus file time')
lm = {}
for line in open(HEADF):
    t, v, _ = line.rstrip('\n').split('\t')
    if v:
        lm[t] = dt.datetime.strptime(v, '%a, %d %b %Y %H:%M:%S GMT')
rows138 = {r['tdoc']: r for r in csv.DictReader(open(os.path.join(D, 'sa4_138.csv')))}
d = [(P(rows138[t]['uploaded']) - lm[t]).total_seconds() for t in lm if rows138.get(t, {}).get('uploaded')]
pr(f'SA4#138 HEAD sample n={len(d)}: min={min(d):.0f}s median={pct(d,.5):.0f}s max={max(d):.0f}s;',
   'within 0..60 s:', sum(0 <= x < 60 for x in d), '; within +-2 h of 1h/2h/-6h:',
   {h: sum(abs(x - h*3600) < 120 for x in d) for h in (1, 2, -6)})
for name in MEET:
    rows = list(csv.DictReader(open(os.path.join(D, name + '.csv'))))
    d = [(P(r['uploaded']) - P(r['listing_time'])).total_seconds() for r in rows if r['uploaded'] and r['listing_time']]
    b = collections.Counter(round(x / 3600) for x in d)
    pr(f'{name}: listing n={len(d)} uploaded-listing in [0,60)s: {sum(0 <= x < 60 for x in d)}; by nearest hour: {dict(sorted(b.items()))}')

# ---- (b) Reservation vs Uploaded under hypotheses
pr('\n## (b) upload - reservation (minutes) under each hypothesis')
pr('| meeting | hypothesis | n | inversions (res after upload) | min gap | 1st pct | gaps 0-10 min | gaps 10-60 min | gaps 60-130 min |')
pr('|---|---|---|---|---|---|---|---|---|')
small = []
for name, (lab, tz) in MEET.items():
    rows = [r for r in csv.DictReader(open(os.path.join(D, name + '.csv'))) if r['uploaded'] and r['reservation']]
    hyps = [('H1 UTC', UTC), ('H2 Europe/Paris', PARIS)]
    if tz: hyps.append(('H3 ' + lab, tz))
    for hn, z in hyps:
        g = [(P(r['uploaded']) - to_utc(P(r['reservation']), z)).total_seconds() / 60 for r in rows]
        pr(f'| {name} | {hn} | {len(g)} | {sum(x < 0 for x in g)} | {min(g):.1f} | {pct(g,.01):.1f} | '
           f'{sum(0 <= x < 10 for x in g)} | {sum(10 <= x < 60 for x in g)} | {sum(60 <= x < 130 for x in g)} |')
    for r in rows:
        g = (P(r['uploaded']) - P(r['reservation'])).total_seconds() / 60
        if g < 180:
            small.append((name, r['tdoc'], r['reservation'], r['uploaded'], round(g, 1)))
pr('\nRaw (H1) gaps < 180 min histogram, 10-min bins, all meetings:')
h = collections.Counter(int(x[4] // 10) * 10 for x in small)
for k in sorted(h): pr(f'  {k:>5}..{k+10:<4} min: {h[k]}')
with open(os.path.join(D, 'small_gaps.csv'), 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['meeting', 'tdoc', 'reservation', 'uploaded', 'gap_min_raw']); w.writerows(sorted(small, key=lambda x: x[4]))

# ---- (d) DST split for SA4#138 (and all meetings: CET vs CEST periods)
pr('\n## (d) raw gaps (H1 reading) split by Paris DST state at reservation')
for name in MEET:
    rows = [r for r in csv.DictReader(open(os.path.join(D, name + '.csv'))) if r['uploaded'] and r['reservation']]
    for lab, cond in (('CEST (UTC+2)', True), ('CET (UTC+1)', False)):
        g = [(P(r['uploaded']) - P(r['reservation'])).total_seconds() / 60 for r in rows
             if (P(r['reservation']).replace(tzinfo=UTC).astimezone(PARIS).utcoffset() == dt.timedelta(hours=2)) == cond]
        if g:
            pr(f'{name} {lab}: n={len(g)} min={min(g):.1f} n<60min={sum(x<60 for x in g)} n<120min={sum(x<120 for x in g)} n<10min={sum(x<10 for x in g)}')

# ---- (c) hour of day of reservations during SA4#138 meeting week
pr('\n## (c) SA4#138 reservations 2026-11-16..20 by raw hour-of-day (as written)')
rows = list(csv.DictReader(open(os.path.join(D, 'sa4_138.csv'))))
hc = collections.Counter(P(r['reservation']).hour for r in rows if r['reservation'] and '2026-11-16' <= r['reservation'][:10] <= '2026-11-20')
uc = collections.Counter(P(r['uploaded']).hour for r in rows if r['uploaded'] and '2026-11-16' <= r['uploaded'][:10] <= '2026-11-20')
pr('| raw hour | reservations | uploads | = Calgary hour if UTC | = Calgary hour if Paris |')
pr('|---|---|---|---|---|')
for H in range(24):
    pr(f'| {H:02d} | {hc[H]} | {uc[H]} | {(H-6)%24:02d} | {(H-7)%24:02d} |')
open(os.path.join(D, 'analysis_output.md'), 'w').write('\n'.join(out) + '\n')
