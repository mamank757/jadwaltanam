#!/usr/bin/env python3
"""
Membangun data.json untuk dashboard GitHub Pages.
File Excel hanya berisi data curah hujan per kecamatan.
Indeks ONI (Niño 3.4 CPC) dan DMI (NOAA/BoM) ditarik otomatis dari web.
"""
import datetime as dt, glob, json, math, os, pathlib, re, warnings
import numpy as np, requests
from openpyxl import load_workbook
try:
    from scipy.stats import t as tdist
except ImportError:
    class tdist:
        @staticmethod
        def _betacf(a, b, x):
            qab, qap, qam = a + b, a + 1, a - 1
            c, d = 1.0, 1 - qab * x / qap
            d = 1 / (d if abs(d) > 1e-300 else 1e-300); h = d
            for m in range(1, 300):
                m2 = 2 * m
                aa = m * (b - m) * x / ((qam + m2) * (a + m2))
                d = 1 + aa * d; d = 1 / (d if abs(d) > 1e-300 else 1e-300)
                c = 1 + aa / c; c = c if abs(c) > 1e-300 else 1e-300
                h *= d * c
                aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
                d = 1 + aa * d; d = 1 / (d if abs(d) > 1e-300 else 1e-300)
                c = 1 + aa / c; c = c if abs(c) > 1e-300 else 1e-300
                de = d * c; h *= de
                if abs(de - 1) < 1e-14:
                    break
            return h

        @staticmethod
        def _ibeta(a, b, x):
            if x <= 0: return 0.0
            if x >= 1: return 1.0
            bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x))
            if x < (a + 1) / (a + b + 2):
                return bt * tdist._betacf(a, b, x) / a
            return 1 - bt * tdist._betacf(b, a, 1 - x) / b

        @staticmethod
        def sf(x, df):
            tail = 0.5 * tdist._ibeta(df / 2, 0.5, df / (df + x * x))
            return tail if x >= 0 else 1 - tail

        @staticmethod
        def ppf(q, df):
            if q == 0.5: return 0.0
            if q < 0.5: return -tdist.ppf(1 - q, df)
            lo, hi = 0.0, 1e3
            for _ in range(200):
                mid = (lo + hi) / 2
                if 1 - tdist.sf(mid, df) < q: lo = mid
                else: hi = mid
            return (lo + hi) / 2

warnings.simplefilter('ignore')
LEADS = 6
ENSO_URL = 'https://www.cpc.ncep.noaa.gov/data/indices/wksst9120.for'
MIN_WEEKS = 2
DMI_URLS = ['https://psl.noaa.gov/gcos_wgsp/Timeseries/Data/dmi.had.long.data',
            'https://psl.noaa.gov/data/timeseries/month/data/dmi.had.long.data']
BOM_URL = 'https://www.bom.gov.au/clim_data/IDCK000072/iod_1.txt'
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/124.0 Safari/537.36')
CACHE = pathlib.Path('cache'); CACHE.mkdir(exist_ok=True)
MON = {m: i for i, m in enumerate('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}
LAGS = range(7)

def tinv(alpha, df):
    return float(tdist.ppf(1 - alpha / 2, df))

def tdist2(x, df):
    return float(2 * tdist.sf(abs(x), df))

def corr(x, y):
    if len(x) < 3 or np.std(x) == 0 or np.std(y) == 0:
        return 0.0
    return float(np.corrcoef(x, y)[0, 1])

def p_corr(r, df):
    return tdist2(abs(r) * math.sqrt(df / (1 - r * r)), df) if abs(r) < 1 else 0.0

def bh_pass(P, alpha):
    f = np.asarray(P, float).ravel(); m = f.size
    rank = np.array([(f < p).sum() + 1 for p in f])
    ok = f <= rank / m * alpha
    pc = f[ok].max() if ok.any() else 0.0
    return (f <= pc).reshape(np.shape(P))

def find_xlsx():
    env = os.environ.get('XLSX')
    if env and os.path.exists(env):
        return env
    for f in sorted(glob.glob('Data_Curah_Hujan_*.xlsx')):
        try:
            # Validasi cukup pastikan file bisa dibuka dan ada sheet stasiunnya
            wb = load_workbook(f, read_only=True)
            if len(wb.sheetnames) > 0:
                return f
        except Exception:
            pass
    raise SystemExit('File Excel curah hujan tidak ditemukan di folder kerja.')

def num(v):
    try:
        return float(str(v).replace(',', '.'))
    except (ValueError, TypeError):
        return np.nan

def read_stations(wb):
    out = []
    for ws in wb.worksheets:
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
        if r1 < 0:
            continue

        def grab(s):
            if s < 0:
                return {}
            o = {}
            for r in v[s:]:
                y = num(r[0])
                if np.isnan(y):
                    break
                o[int(y)] = np.array([num(x) for x in (list(r[1:13]) + [None] * 12)[:12]])
            return o
        out.append(dict(id=ws.title, name=meta.get('name') or ws.title, loc=meta.get('loc'),
                        lat=meta.get('lat'), lon=meta.get('lon'), ch=grab(r1), hh=grab(r2)))
    return out

# --- Fetch & Parse Indeks Online (ONI & DMI) ---
def fetch(urls, name):
    err = ''
    for u in urls:
        try:
            r = requests.get(u, timeout=60, headers={'User-Agent': UA, 'Referer': 'https://www.bom.gov.au/climate/enso/'})
            r.raise_for_status()
            (CACHE / name).write_text(r.text, encoding='utf-8')
            return r.text, 'unduh langsung OK'
        except Exception as e:
            err = type(e).__name__
    f = CACHE / name
    if f.exists():
        return f.read_text(encoding='utf-8'), f'cache lokal ({err})'
    return None, f'gagal: {err}'

def monthly(acc):
    return {t: float(np.mean(v)) for t, v in acc.items() if len(v) >= MIN_WEEKS}

def parse_weekly(txt):
    acc = {}
    for l in txt.splitlines():
        d = re.match(r'^\s*(\d{1,2})([A-Z]{3})(\d{4})(.*)$', l)
        if not d or d[2] not in MON:
            continue
        v = [float(x) for x in re.findall(r'-?\d+\.\d', d[4])]
        x = v[5] if len(v) >= 8 else v[2] if len(v) >= 4 else None
        if x is not None:
            acc.setdefault(int(d[3]) * 12 + MON[d[2]], []).append(x)
    return monthly(acc)

def oni_from_monthly(nino):
    return {t: float(np.mean([nino[k] for k in (t - 1, t, t + 1)])) for t in nino if all(k in nino for k in (t - 1, t, t + 1))}

def parse_psl(txt):
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

def parse_bom(txt):
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
    return {t: float(np.mean(v)) for t, v in acc.items() if len(v) >= 2}

def smooth_dmi(raw, on, strict=False):
    if not on:
        return dict(raw)
    o = {}
    for t in range(min(raw) if raw else 0, (max(raw) if raw else 0) + 1):
        v = [raw[k] for k in (t - 1, t, t + 1) if k in raw]
        if (strict and len(v) == 3) or (not strict and v and t in raw):
            o[t] = float(np.mean(v))
    return o

def get_online_indices():
    # Ambil data online untuk riwayat bulanan ONI & DMI
    txt, _ = fetch([ENSO_URL], 'wksst9120.for')
    w = parse_weekly(txt) if txt else {}
    oni = oni_from_monthly(w) if w else {}

    txt2, _ = fetch(DMI_URLS, 'dmi_psl.data')
    dmi = parse_psl(txt2) if txt2 else {}
    return oni, smooth_dmi(dmi, 1, strict=False)

# --- Model OLS & Series ---
def lagvec(idx, yrs, m, L):
    return np.array([idx.get(y * 12 + m - L, np.nan) for y in yrs])

def ols(y, x):
    n = len(y); ym, xm = y.mean(), x.mean()
    sxx = ((x - xm) ** 2).sum(); sy = ((y - ym) ** 2).sum()
    b = ((x - xm) * (y - ym)).sum() / sxx if sxx > 0 else 0.0
    a = ym - b * xm
    r = corr(x, y); r2 = r * r
    s = math.sqrt(max(0, (sy - b * b * sxx)) / (n - 2)) if n > 2 else 0.0
    h = 1 / n + (x - xm) ** 2 / sxx if sxx > 0 else np.full(n, 1 / n)
    e = y - a - b * x
    rm = math.sqrt((((e / (1 - h)) ** 2).sum()) / n)
    rc = math.sqrt(sy / n) * n / (n - 1) if n > 1 else 1.0
    se = s / math.sqrt(sxx) if sxx > 0 else np.inf
    return dict(n=n, mean=ym, xbar=xm, sxx=sxx, b=b, a=a, r=r, r2=r2, s=s, p=tdist2(b / se, n - 2) if se and se != np.inf else 1.0,
                rmse=rm, rmseC=rc, msess=1 - (rm / rc) ** 2 if rc > 0 else 0.0, sd=float(np.std(y, ddof=1)))

def combined(y, x, d, ao):
    n = len(y); ym, xm, dm = y.mean(), x.mean(), d.mean()
    Soo = ((x - xm) ** 2).sum(); Sdd = ((d - dm) ** 2).sum(); Sod = ((x - xm) * (d - dm)).sum()
    ryo, ryd, rod = corr(y, x), corr(y, d), corr(x, d)
    sy_, sx_, sd_ = np.std(y, ddof=1), np.std(x, ddof=1), np.std(d, ddof=1)
    denom = (1 - rod ** 2)
    b1 = (ryo - ryd * rod) / denom * sy_ / sx_ if denom != 0 else 0.0
    b2 = (ryd - ryo * rod) / denom * sy_ / sd_ if denom != 0 else 0.0
    a = ym - b1 * xm - b2 * dm
    R2 = (ryo ** 2 + ryd ** 2 - 2 * ryo * ryd * rod) / denom if denom != 0 else 0.0
    SST = ((y - ym) ** 2).sum()
    s = math.sqrt(max(0, (1 - R2)) * SST / (n - 3)) if n > 3 else 0.0
    se2 = s / math.sqrt(Sdd * (1 - rod ** 2)) if Sdd * (1 - rod ** 2) > 0 else np.inf
    pb2 = tdist2(b2 / se2, n - 3) if se2 != np.inf else 1.0
    det = Soo * Sdd - Sod ** 2
    xc, dc = x - xm, d - dm
    h = 1 / n + (Sdd * xc ** 2 - 2 * Sod * xc * dc + Soo * dc ** 2) / det if det != 0 else np.full(n, 1 / n)
    e = y - a - b1 * x - b2 * d
    rm = math.sqrt((((e / (1 - h)) ** 2).sum()) / n)
    return dict(b1=b1, b2=b2, a=a, r2=R2, s=s, pb2=pb2, rmse=rm, soo=Soo, sdd=Sdd, sod=Sod, xo=xm, xd=dm, rod=rod)

def model_series(Y, yrs, oni, dmi):
    al = 0.05
    minN = 20
    minMsess = 0.0
    R = np.zeros((12, 7)); PV = np.ones((12, 7)); NN = [0] * 12
    for m in range(12):
        for L in LAGS:
            x = lagvec(oni, yrs, m, L); y = Y[:, m]; k = ~np.isnan(x) & ~np.isnan(y)
            NN[m] = int(k.sum())
            R[m, L] = corr(x[k], y[k]); PV[m, L] = p_corr(R[m, L], NN[m] - 2) if NN[m] > 2 else 1.0
    FD = bh_pass(PV, al)
    sel = [int(np.argmin(PV[m])) for m in range(12)]
    ev = []
    for m in range(12):
        ev.append('Kuat' if FD[m, sel[m]] else 'Indikatif' if PV[m, sel[m]] < al else 'Lemah / tidak ada')
    
    PP = np.ones((12, 7)); RP = np.zeros((12, 7))
    for m in range(12):
        x0 = lagvec(oni, yrs, m, sel[m]); y = Y[:, m]
        valid_y = ~np.isnan(y)
        ryo = corr(x0[valid_y], y[valid_y]) if valid_y.any() else 0.0
        for L in LAGS:
            d = lagvec(dmi, yrs, m, L); k = ~np.isnan(d) & ~np.isnan(y) & ~np.isnan(x0)
            if k.sum() > 3:
                ryd = corr(d[k], y[k]); rod = corr(x0[k], d[k])
                denom = math.sqrt((1 - ryo ** 2) * (1 - rod ** 2))
                RP[m, L] = (ryd - ryo * rod) / denom if denom > 0 else 0.0
                PP[m, L] = p_corr(RP[m, L], int(k.sum()) - 3)
    FDd = bh_pass(PP, al)
    out = []
    for m in range(12):
        y = Y[:, m]; x = lagvec(oni, yrs, m, sel[m]); k = ~np.isnan(x) & ~np.isnan(y)
        o = ols(y[k], x[k]); n = o['n']
        layak = n >= minN and o['msess'] > minMsess and (ev[m] == 'Kuat' or ev[m] == 'Indikatif')
        lD = int(np.argmin(PP[m])); d = lagvec(dmi, yrs, m, lD)
        k2 = k & ~np.isnan(d)
        c = combined(y[k2], x[k2], d[k2], o) if k2.sum() > 3 else dict(b1=0, b2=0, a=o['mean'], r2=0, s=o['s'], pb2=1, rmse=o['rmse'], soo=1, sdd=1, sod=0, xo=o['xbar'], xd=0, rod=0)
        evD = 'Kuat' if FDd[m, lD] else 'Indikatif' if PP[m, lD] < al else 'Lemah / tidak ada'
        c['msess'] = 1 - (c['rmse'] / o['rmseC']) ** 2 if o['rmseC'] > 0 else 0.0
        c['dm'] = c['msess'] - o['msess']
        add = (c['pb2'] < al and (evD == 'Kuat' or evD == 'Indikatif') and c['dm'] > 0.02 and c['msess'] > minMsess and n >= minN)
        basis = 'ONI+DMI' if add else 'ONI saja' if layak else 'Klimatologi'
        out.append(dict(m=m, lag=sel[m], lagD=lD, ev=ev[m], evD=evD, n=n, p=o['p'], ok=layak, add=add, basis=basis,
                        rD=float(RP[m, lD]), pD=float(PP[m, lD]),
                        mean=o['mean'], sd=o['sd'], xbar=o['xbar'], sxx=o['sxx'], b=o['b'], a=o['a'], r=o['r'], r2=o['r2'],
                        s=o['s'], pct=o['b'] / o['mean'] * 100 if o['mean'] > 0 else 0, rmse=o['rmse'], rmseC=o['rmseC'], msess=o['msess'],
                        c=dict(xo=c['xo'], xd=c['xd'], soo=c['soo'], sdd=c['sdd'], sod=c['sod'], b1=c['b1'], b2=c['b2'],
                               a=c['a'], s=c['s'], r2=c['r2'], pb2=c['pb2'], rmse=c['rmse'], msess=c['msess'], dm=c['dm'], rod=c['rod']),
                        t1=tinv(al, max(1, n - 1)), t2=tinv(al, max(1, n - 2)), t3=tinv(al, max(1, n - 3))))
    return out

def clim_stat(v):
    v = v[~np.isnan(v)]
    if len(v) < 5:
        return dict(mean=0, sd=1, p33=0, p67=0, med=0, min=0, max=0)
    p33, med, p67 = np.percentile(v, [100 / 3, 50, 200 / 3])
    return dict(mean=v.mean(), sd=float(v.std(ddof=1)) or 1.0, p33=p33, p67=p67, med=med, min=v.min(), max=v.max())

def clim(M):
    return [clim_stat(M[:, m]) for m in range(12)]

def clean(o):
    if isinstance(o, dict): return {k: clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)): return [clean(v) for v in o]
    if isinstance(o, (bool, np.bool_)): return bool(o)
    if isinstance(o, (float, np.floating)): return round(float(o), 6) if math.isfinite(o) else None
    if isinstance(o, np.integer): return int(o)
    return o

def main():
    u = os.environ.get('SHEET_XLSX_URL')
    if u:
        try:
            r = requests.get(u, timeout=120); r.raise_for_status()
            pathlib.Path('Data_Curah_Hujan_Wajo.xlsx').write_bytes(r.content)
            print('Excel diambil dari URL')
        except Exception as e:
            print('Gagal ambil dari URL, memakai file lokal:', e)
    
    xl = find_xlsx()
    print('Memakai file Excel:', xl)
    wb = load_workbook(xl, data_only=True)
    st = read_stations(wb)
    
    yrs = sorted(set.intersection(*[set(s['ch']) for s in st]))
    yrs = [y for y in yrs if all(not np.isnan(s['ch'][y]).all() for s in st)]
    Y0, Y1 = yrs[0], yrs[-1]
    
    mat = lambda s, k: np.array([s[k].get(y, np.full(12, np.nan)) for y in yrs])
    for s in st:
        s['CH'] = mat(s, 'ch'); s['HH'] = mat(s, 'hh')
    
    allst = [dict(id='WILAYAH', name=f'WILAYAH (rata-rata {len(st)} kecamatan)', loc='Rata-rata aritmetik seluruh kecamatan',
                  lat=None, lon=None, CH=np.mean([s['CH'] for s in st], axis=0),
                  HH=np.nanmean(np.stack([s['HH'] for s in st]), axis=0))] + st

    oni_c, dmi_c = get_online_indices()
    t_last = max(oni_c) if oni_c else 2026 * 12
    now = dt.datetime.now(dt.timezone.utc); t_now = now.year * 12 + now.month - 1
    hist = [dict(y=t // 12, m=t % 12, oni=oni_c.get(t), dmi=dmi_c.get(t)) for t in range(t_last - 35, t_last + 1)]
    td = max([t for t in dmi_c if t <= t_now] or [max(dmi_c) if dmi_c else t_now])

    stations = []
    for s in allst:
        mods = model_series(s['CH'], yrs, oni_c, dmi_c)
        stations.append(dict(
            id=s['id'], name=s['name'], loc=s.get('loc'), lat=s.get('lat'), lon=s.get('lon'),
            clim=dict(ch=clim(s['CH']), hh=clim(s['HH'])),
            mon=mods,
            fc=[dict(y=(t_now + h - 1) // 12, m=(t_now + h - 1) % 12, h=h) for h in range(1, LEADS + 1)]))

    n = len(yrs)
    out = dict(stations=stations,
               idx=dict(y=t_last // 12, m=t_last % 12, oni=oni_c.get(t_last, 0.0), dmi=dmi_c.get(td, 0.0), dy=td // 12, dm=td % 12,
                        bom=True, bomMsg="Indeks ONI & DMI diambil otomatis dari web NOAA/CPC", live=True, hist=hist),
               meta=dict(y0=Y0, y1=Y1, n=len(st), nYears=n, alpha=0.05, minN=20, minMsess=0.0,
                         useInd=1, smooth=1, dMsess=0.02,
                         rcrit=math.sqrt(tinv(0.05, n - 2) ** 2 / (tinv(0.05, n - 2) ** 2 + n - 2)),
                         source=os.path.basename(xl)),
               updated=now.strftime('%Y-%m-%d %H:%M UTC'))
    
    pathlib.Path('data.json').write_text(json.dumps(clean(out), ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print('data.json berhasil ditulis!')

if __name__ == '__main__':
    main()
