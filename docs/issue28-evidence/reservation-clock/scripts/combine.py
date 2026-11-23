"""Combine per-meeting CSVs + HEAD sample into one per-TDoc CSV. usage: combine.py data_dir head.tsv out.csv"""
import csv, sys, os, datetime as dt
D, H, O = sys.argv[1:4]
lm = {}
for line in open(H):
    t, v, _ = line.rstrip('\n').split('\t')
    if v: lm[t] = dt.datetime.strptime(v, '%a, %d %b %Y %H:%M:%S GMT')
with open(O, 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['meeting', 'tdoc', 'status', 'reservation_raw', 'uploaded_raw', 'ftp_listing_utc_minute', 'http_last_modified_utc',
                'upload_minus_reservation_min_if_H1', 'uploaded_minus_lastmod_s'])
    for name in ('sa4_138', 'TSGS4_137-e', 'TSGS4_136_Montreal', 'TSGS4_135_India'):
        for r in csv.DictReader(open(os.path.join(D, name + '.csv'))):
            u = dt.datetime.fromisoformat(r['uploaded']) if r['uploaded'] else None
            s = dt.datetime.fromisoformat(r['reservation']) if r['reservation'] else None
            L = lm.get(r['tdoc']) if name == 'sa4_138' else None
            w.writerow([name, r['tdoc'], r['status'], r['reservation'], r['uploaded'], r['listing_time'],
                        L.isoformat(sep=' ') if L else '',
                        round((u - s).total_seconds() / 60, 2) if u and s else '',
                        round((u - L).total_seconds()) if u and L else ''])
