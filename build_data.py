#!/usr/bin/env python3
"""
Membangun data.json (baseline + prakiraan hujan Sulbar berbasis ENSO/IOD) untuk dashboard GitHub Pages.

Struktur repo (root): index.html, build_data.py, Data_Curah_Hujan_-_Mamuju.xlsx, .github/workflows/update.yml
Lokal : pip install numpy openpyxl requests && python build_data.py
Sumber: ENSO = anomali Nino 3.4 mingguan CPC (rata-rata bulanan, min. MIN_WEEKS minggu)
        IOD  = DMI bulanan NOAA PSL (riwayat latih) + DMI mingguan BoM (mengisi bulan yang belum rilis, dikalibrasi ke skala PSL)
Data hujan: file xlsx di repo (atau Google Sheets bila secret SHEET_XLSX_URL diisi).
Cadangan: bila unduhan gagal (mis. BoM memblokir), dipakai berkas terakhir di cache/. Anda dapat menimpa cache/iod_1.txt
          secara manual dengan isi https://www.bom.gov.au/clim_data/IDCK000072/iod_1.txt
Model : regresi ridge anomali terstandar ~ Nino3.4(t-lead) + DMI(t-lead), pool +-1 bulan, validasi leave-one-year-out.
"""
import datetime as dt, json, math, os, pathlib, re, warnings
import numpy as np, requests
from openpyxl import load_workbook

warnings.simplefilter('ignore')
XLSX = 'Data_Curah_Hujan_-_Mamuju.xlsx'
Y0, Y1, LEADS, MIN_WEEKS = 1999, 2024, 6, 2
ENSO_URL = 'https://www.cpc.ncep.noaa.gov/data/indices/wksst9120.for'
DMI_URLS = ['https://psl.noaa.gov/gcos_wgsp/Timeseries/Data/dmi.had.long.data',
            'https://psl.noaa.gov/data/timeseries/month/data/dmi.had.long.data']
BOM_URL = 'https://www.bom.gov.au/clim_data/IDCK000072/iod_1.txt'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
N, T0 = Y1 - Y0 + 1, Y0 * 12
CACHE = pathlib.Path('cache'); CACHE.mkdir(exist_ok=True)
MON = {m: i for i, m in enumerate('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}


# ---------- Unduh & parsing indeks ----------
def fetch(urls, name):
    err = ''
    for u in urls:
        try:
            r = requests.get(u, timeout=60, headers={'User-Agent': UA, 'Referer': 'https://www.bom.gov.au/climate/enso/'})
            r.raise_for_status()
            (CACHE / name).write_text(r.text, encoding='utf-8')
            return r.text, 'unduh langsung OK'
        except Exception as e:
            err = f'{type(e).__name__}: {str(e)[:90]}'
    f = CACHE / name
    if f.exists():
        return f.read_text(encoding='utf-8'), f'cache lokal ({err})'
    return None, f'gagal: {err}'


def monthly(acc):
    return {t: float(np.mean(v)) for t, v in acc.items() if len(v) >= MIN_WEEKS}


def parse_weekly(txt):  # 03JAN1990  23.4-0.4  25.1-0.3  26.6 0.0  28.6 0.3 -> SSTA Nino3.4
    acc = {}
    for l in txt.splitlines():
        d = re.match(r'^\s*(\d{1,2})([A-Z]{3})(\d{4})(.*)$', l)
        if not d or d[2] not in MON:
            continue
        v = [float(x) for x in re.findall(r'-?\d+\.\d', d[4])]
        x = v[5] if len(v) >= 8 else v[2] if len(v) >= 4 else None  # 8 angka = SST+SSTA; 4 angka = SSTA saja
        if x is not None:
            acc.setdefault(int(d[3]) * 12 + MON[d[2]], []).append(x)
    return monthly(acc)


def parse_psl(txt):  # baris: tahun + 12 nilai bulanan; missing <= -90
    o = {}
    for l in txt.splitlines():
        p = l.split()
        if len(p) >= 13 and re.fullmatch(r'\d{4}', p[0]):
            for m, x in enumerate(p[1:13]):
                try:
                    v = float(x)
                except ValueError:
                    continue
                if v > -90:
                    o[int(p[0]) * 12 + m] = v
    return o


def parse_bom(txt):  # baris: yyyymmdd mulai, yyyymmdd akhir, nilai; bulan = titik tengah minggu
    acc = {}
    for l in txt.splitlines():
        p = l.strip().split(',')
        if len(p) < 3 or not re.fullmatch(r'\d{8}', p[0]):
            continue
        try:
            v = float(p[2])
        except ValueError:
            continue
        d = dt.datetime.strptime(p[0], '%Y%m%d') + dt.timedelta(days=3)
        acc.setdefault(d.year * 12 + d.month - 1, []).append(v)
    return monthly(acc)


def merge_bom(dmi, bom):  # isi bulan kosong PSL dengan BoM terkalibrasi (regresi pada bulan tumpang tindih)
    both = [t for t in bom if t in dmi]
    a, b = 0.0, 1.0
    if len(both) >= 24:
        sl, ic = np.polyfit([bom[t] for t in both], [dmi[t] for t in both], 1)
        if 0.3 < sl < 3:
            a, b = ic, sl
    for t, v in bom.items():
        if t not in dmi:
            dmi[t] = a + b * v


def get_indices():
    txt, m1 = fetch([ENSO_URL], 'wksst9120.for')
    oni = parse_weekly(txt) if txt else {}
    if len(oni) < 20:
        raise SystemExit('Data Nino 3.4 tidak terbaca: ' + m1)
    txt, m2 = fetch(DMI_URLS, 'dmi_psl.data')
    dmi = parse_psl(txt) if txt else {}
    if len(dmi) < 100:
        raise SystemExit('Data DMI PSL tidak terbaca: ' + m2)
    txt, msg = fetch([BOM_URL], 'iod_1.txt')
    ok = False
    if txt:
        b = parse_bom(txt)
        if b:
            merge_bom(dmi, b); ok = True
        else:
            msg = 'format tidak terbaca'
    return oni, dmi, ok, msg


# ---------- Data hujan ----------
def num(v):
    try:
        return float(str(v).replace(',', '.'))
    except ValueError:
        return np.nan


def read_stations():
    out = []
    for ws in load_workbook(XLSX, data_only=True).worksheets:
        v = [list(r) for r in ws.iter_rows(values_only=True)]
        meta, r1, r2 = {}, -1, -1
        for i, r in enumerate(v):
            a = str(r[0]) if r and r[0] is not None else ''
            if a == 'Nama Stasiun': meta['name'] = r[1]
            elif a == 'Lokasi': meta['loc'] = r[1]
            elif a.startswith('Longitude'):
                p = str(r[1]).split('/'); meta['lon'], meta['lat'] = num(p[0]), num(p[1])
            elif a.startswith('CURAH HUJAN'): r1 = i + 2
            elif a.startswith('JUMLAH HARI HUJAN'): r2 = i + 2
        if r1 < 0 or r2 < 0:
            continue

        def grab(s):
            o = np.full(N * 12, np.nan)
            for r in v[s:]:
                y = num(r[0])
                if np.isnan(y):
                    break
                if Y0 <= y <= Y1:
                    vals = [num(x) for x in r[1:13]]
                    vals += [np.nan] * (12 - len(vals))
                    k = (int(y) - Y0) * 12
                    o[k:k + 12] = vals
            return o

        ch, hh = grab(r1), grab(r2)
        ch[(ch == 0) & np.isnan(hh)] = np.nan  # nol tanpa data hari hujan (mis. Batetanga Sep-2015) = data hilang
        out.append(dict(id=ws.title, name=meta.get('name') or ws.title, loc=meta.get('loc'),
                        lat=meta.get('lat'), lon=meta.get('lon'), ch=ch, hh=hh))
    return out


# ---------- Statistik & model ----------
def stat(v):
    v = v[~np.isnan(v)]
    if len(v) < 5:
        return dict(mean=0, sd=1, p33=0, p67=0, med=0, min=0, max=0)
    p33, med, p67 = np.percentile(v, [100 / 3, 50, 200 / 3])
    return dict(mean=v.mean(), sd=v.std(ddof=1) or 1.0, p33=p33, p67=p67, med=med, min=v.min(), max=v.max())


def clim(s):
    return [stat(s[m::12]) for m in range(12)]


def corr(x, y):
    if len(x) < 5:
        return 0.0
    c = np.corrcoef(x, y)[0, 1]
    return float(c) if np.isfinite(c) else 0.0


def cor(s, idx):  # korelasi bersamaan CH vs indeks per bulan
    r = []
    for m in range(12):
        pr = [(s[i * 12 + m], idx[(Y0 + i) * 12 + m]) for i in range(N)
              if not np.isnan(s[i * 12 + m]) and (Y0 + i) * 12 + m in idx]
        r.append(corr([p[0] for p in pr], [p[1] for p in pr]))
    return r


def rows(s, cl, m, h, I):
    R = []
    for y in range(N):
        for d in (-1, 0, 1):
            k = m + d; t = (Y0 + y) * 12 + k; i = t - T0
            if i < 0 or i >= len(s) or np.isnan(s[i]):
                continue
            a, b = I['oni'].get(t - h), I['dmi'].get(t - h)
            if a is None or b is None:
                continue
            c = cl[k % 12]
            R.append((y, (s[i] - c['mean']) / c['sd'], a, b))
    return np.array(R)


def fit(R):
    n = len(R)
    if n < 20:
        return None
    _, z, a, b = R.T
    ma, mb, mz = a.mean(), b.mean(), z.mean()
    sa, sb = a.std() or 1.0, b.std() or 1.0
    A, B, Z = (a - ma) / sa, (b - mb) / sb, z - mz
    aa, ab, bb, az, bz = A @ A, A @ B, B @ B, A @ Z, B @ Z
    l = 0.05 * n
    D = (aa + l) * (bb + l) - ab * ab
    return dict(ma=ma, sa=sa, mb=mb, sb=sb, mz=mz, b1=((bb + l) * az - ab * bz) / D, b2=((aa + l) * bz - ab * az) / D)


def pred(M, a, b):
    return M['mz'] + M['b1'] * (a - M['ma']) / M['sa'] + M['b2'] * (b - M['mb']) / M['sb']


def model(s, cl, m, h, I):
    R = rows(s, cl, m, h, I)
    M = fit(R)
    if M is None:
        return None
    sse = sst = 0.0; P, Zs = [], []
    for y in range(N):
        te, Mc = R[R[:, 0] == y], fit(R[R[:, 0] != y])
        if Mc is None or len(te) == 0:
            continue
        p = pred(Mc, te[:, 2], te[:, 3])
        sse += ((te[:, 1] - p) ** 2).sum(); sst += (te[:, 1] ** 2).sum()
        P += list(p); Zs += list(te[:, 1])
    if not P or sst == 0:
        return None
    M.update(rmse=math.sqrt(sse / len(P)), r=corr(P, Zs), skill=1 - sse / sst, n=len(R))
    return M


def build(s, I, t0):
    cc, hc = clim(s['ch']), clim(s['hh'])
    fc = []
    for h in range(1, LEADS + 1):
        t = t0 + h; m = t % 12
        fc.append(dict(y=t // 12, m=m, h=h,
                       ch={**cc[m], 'mod': model(s['ch'], cc, m, h, I)},
                       hh={**hc[m], 'mod': model(s['hh'], hc, m, h, I)}))
    return dict(id=s['id'], name=s['name'], loc=s.get('loc'), lat=s.get('lat'), lon=s.get('lon'),
                clim=dict(ch=cc, hh=hc), cor=dict(oni=cor(s['ch'], I['oni']), dmi=cor(s['ch'], I['dmi'])), fc=fc)


def clean(o):  # numpy -> JSON, NaN -> null
    if isinstance(o, dict): return {k: clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)): return [clean(v) for v in o]
    if isinstance(o, (float, np.floating)): return round(float(o), 4) if math.isfinite(o) else None
    if isinstance(o, np.integer): return int(o)
    return o


def main():
    u = os.environ.get('SHEET_XLSX_URL')  # opsional: ambil data hujan dari Google Sheets (link export xlsx)
    if u:
        try:
            r = requests.get(u, timeout=120); r.raise_for_status()
            pathlib.Path(XLSX).write_bytes(r.content)
            print('Data hujan diambil dari Google Sheets')
        except Exception as e:
            print('Gagal ambil Google Sheets, memakai file xlsx di repo:', e)
    oni, dmi, bom_ok, msg = get_indices()
    I = dict(oni=oni, dmi=dmi)
    t0 = max(oni)                              # bulan terakhir data Nino 3.4
    td = max(t for t in dmi if t <= t0)        # DMI terakhir yang tersedia
    st = read_stations()
    reg = lambda k: np.nanmean(np.vstack([s[k] for s in st]), axis=0)
    allst = [dict(id='WAJO', name=f'WAJO (rata-rata {len(st)} pos)', loc='Rata-rata seluruh pos',
                  ch=reg('ch'), hh=reg('hh'))] + st
    hist = [dict(y=t // 12, m=t % 12, oni=oni.get(t), dmi=dmi.get(t)) for t in range(t0 - 35, t0 + 1)]
    out = dict(stations=[build(s, I, t0) for s in allst],
               idx=dict(y=t0 // 12, m=t0 % 12, oni=oni[t0], dmi=dmi[td], dy=td // 12, dm=td % 12,
                        bom=bom_ok, bomMsg=msg, hist=hist),
               meta=dict(y0=Y0, y1=Y1, n=len(st)),
               updated=dt.datetime.now(dt.timezone.utc).strftime('%Y-%m-%d %H:%M UTC'))
    pathlib.Path('data.json').write_text(json.dumps(clean(out), ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print('data.json ditulis | Nino3.4 terakhir', t0 // 12, t0 % 12 + 1, '| DMI terakhir', td // 12, td % 12 + 1, '| BoM:', msg)


if __name__ == '__main__':
    main()
