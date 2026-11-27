"""Extract TDoc, Reservation date, Uploaded from a TDoc list xlsx and file times
from a saved 3GPP FTP HTML directory listing. Writes a CSV.
usage: python -I extract.py tdoclist.xlsx listing.html out.csv
"""
import csv, re, sys, datetime as dt
import openpyxl

xlsx, listing, out = sys.argv[1:4]

def parse_res(v):
    if v is None or v == '':
        return None
    if isinstance(v, dt.datetime):
        return v
    s = str(v).strip()
    for f in ('%d/%m/%Y %H:%M:%S', '%m/%d/%Y %I:%M:%S %p'):
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    raise ValueError(s)

def parse_up(v):
    if v is None or v == '':
        return None
    if isinstance(v, dt.datetime):
        return v
    s = str(v).strip()
    for f in ('%d/%m/%Y %H:%M:%S', '%m/%d/%Y %I:%M:%S %p', '%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S'):
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    raise ValueError(s)

txt = open(listing, encoding='utf8', errors='replace').read()
lst = {}
for m in re.finditer(r'class="file" href="[^"]*/([^/"]+)\.zip">.*?<td[^>]*>\s*(\d{4}/\d{2}/\d{2} \d{1,2}:\d{2})', txt, re.S):
    lst[m.group(1)] = dt.datetime.strptime(m.group(2), '%Y/%m/%d %H:%M')

wb = openpyxl.load_workbook(xlsx, read_only=True)
ws = wb['TDoc_List']
rows = ws.iter_rows(values_only=True)
hdr = list(next(rows))
iT, iR, iU, iS = hdr.index('TDoc'), hdr.index('Reservation date'), hdr.index('Uploaded'), hdr.index('TDoc Status')
n = 0
with open(out, 'w', newline='') as fo:
    w = csv.writer(fo)
    w.writerow(['tdoc', 'status', 'reservation', 'uploaded', 'listing_time'])
    for r in rows:
        if not r or not r[iT]:
            continue
        res, up = parse_res(r[iR]), parse_up(r[iU])
        L = lst.get(r[iT])
        w.writerow([r[iT], r[iS], res.isoformat(sep=' ') if res else '',
                    up.isoformat(sep=' ', timespec='seconds') if up else '',
                    L.isoformat(sep=' ') if L else ''])
        n += 1
print(n, 'rows;', len(lst), 'listing entries', file=sys.stderr)
