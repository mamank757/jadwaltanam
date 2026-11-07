#!/usr/bin/env python3
"""
Membangun data.json untuk dashboard GitHub Pages dengan LOGIKA YANG SAMA DENGAN FILE EXCEL
"Data Curah Hujan Wajo Model ENSO IOD per Kecamatan.xlsx" (sheet Uji_Korelasi, Uji_FDR, Model_Stasiun,
Uji_DMI, FDR_DMI, Model_ONI_DMI, Skenario, Skenario_DMI).

Struktur repo (root): index.html, build_data.py, Data_Curah_Hujan_Wajo_Model_ENSO_IOD*.xlsx, .github/workflows/update.yml
Lokal : pip install numpy openpyxl requests   (scipy opsional) && python build_data.py

Alur (identik dengan Excel):
  1. Hujan bulanan 14 kecamatan (sheet kecamatan, baris 'CURAH HUJAN BULANAN'); WILAYAH = rata-rata aritmetik 14 kecamatan.
  2. ONI (sheet 'oni') dan DMI (sheet 'dmi'); DMI dihaluskan rata-rata 3 bulan terpusat bila Parameter!B10 = 1.
  3. Korelasi Pearson hujan bulan m vs ONI(m - lag), lag 0-6; uji t dua sisi; lag terpilih = p terkecil.
     FDR Benjamini-Hochberg per seri (84 uji). Bukti: Kuat (lolos FDR) / Indikatif (p < alfa) / Lemah.
  4. Model OLS  R = a + b*ONI(m - lag), validasi leave-one-year-out (LOYO) eksak via leverage, MSESS = 1 - MSE_model/MSE_klim.
     Layak bila n >= B7, MSESS > B8, dan bukti Kuat (atau Indikatif bila B9 = 1).
  5. DMI: korelasi parsial (kontrol ONI) lag 0-6, FDR per seri; model gabungan ONI+DMI (OLS dua prediktor, LOYO eksak).
     DMI 'menambah' bila p b2 < alfa, bukti DMI Kuat/Indikatif, dMSESS > B11, MSESS gabungan > B8, n >= B7.
     Model dipakai: ONI+DMI -> ONI saja -> Klimatologi.
  6. Prediksi = model pada ONI/DMI skenario; selang prediksi (1-alfa) pakai t(1-alfa/2, df) (df = n-3 / n-2 / n-1), batas bawah >= 0.
Semua parameter (alfa, n min, MSESS min, dst.) dibaca dari sheet 'Parameter' Excel.

Nilai ONI/DMI terkini (kartu 'Observasi terakhir' & grafik riwayat): Niño 3.4 mingguan CPC (wksst9120.for) -> rerata bulanan -> rerata 3 bulan terpusat (= ONI); DMI NOAA PSL + BoM (sumber sama seperti kode lama);
bila unduhan gagal dipakai cache/ lalu nilai terakhir di sheet oni/dmi Excel. Pelatihan model SELALU memakai sheet oni/dmi Excel.
"""
import datetime as dt, glob, json, math, os, pathlib, re, warnings
import numpy as np, requests
from openpyxl import load_workbook
try:
    from scipy.stats import t as tdist
except ImportError:  # tanpa scipy: distribusi t dihitung sendiri (beta tak-lengkap terregularisasi)
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
        def sf(x, df):  # P(T > x)
            tail = 0.5 * tdist._ibeta(df / 2, 0.5, df / (df + x * x))
            return tail if x >= 0 else 1 - tail

        @staticmethod
        def ppf(q, df):  # kuantil, bisection
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
ENSO_URL = 'https://www.cpc.ncep.noaa.gov/data/indices/wksst9120.for'   # sumber lama: Niño 3.4 mingguan CPC (tetap)
MIN_WEEKS = 2
DMI_URLS = ['https://psl.noaa.gov/gcos_wgsp/Timeseries/Data/dmi.had.long.data',
            'https://psl.noaa.gov/data/timeseries/month/data/dmi.had.long.data']
BOM_URL = 'https://www.bom.gov.au/clim_data/IDCK000072/iod_1.txt'
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/124.0 Safari/537.36')
CACHE = pathlib.Path('cache'); CACHE.mkdir(exist_ok=True)
SEAS = {s: i for i, s in enumerate('DJF JFM FMA MAM AMJ MJJ JJA JAS ASO SON OND NDJ'.split())}
MON = {m: i for i, m in enumerate('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}
LAGS = range(7)


# ============ Utilitas statistik (setara fungsi Excel) ============
def tinv(alpha, df):            # TINV(alpha, df) dua sisi
    return float(tdist.ppf(1 - alpha / 2, df))


def tdist2(x, df):              # TDIST(|x|, df, 2)
    return float(2 * tdist.sf(abs(x), df))


def corr(x, y):
    if len(x) < 3 or np.std(x) == 0 or np.std(y) == 0:
        return 0.0
    return float(np.corrcoef(x, y)[0, 1])


def p_corr(r, df):
    return tdist2(abs(r) * math.sqrt(df / (1 - r * r)), df) if abs(r) < 1 else 0.0


def bh_pass(P, alpha):
    """Benjamini-Hochberg seperti sheet Uji_FDR: peringkat = jumlah p lebih kecil + 1; p kritis = p terbesar yang <= i/m*alfa."""
    f = np.asarray(P, float).ravel(); m = f.size
    rank = np.array([(f < p).sum() + 1 for p in f])
    ok = f <= rank / m * alpha
    pc = f[ok].max() if ok.any() else 0.0
    return (f <= pc).reshape(np.shape(P))


# ============ Baca Excel ============
def find_xlsx():
    env = os.environ.get('XLSX')
    if env and os.path.exists(env):
        return env
    for f in sorted(glob.glob('Data_Curah_Hujan*.xlsx')):
        try:
            if {'oni', 'dmi', 'Parameter'} <= set(load_workbook(f, read_only=True).sheetnames):
                return f
        except Exception:
            pass
    raise SystemExit('File Excel model (dengan sheet oni, dmi, Parameter) tidak ditemukan di folder kerja.')


def num(v):
    try:
        return float(str(v).replace(',', '.'))
    except (ValueError, TypeError):
        return np.nan


def read_params(wb):
    ws = wb['Parameter']
    g = lambda r, d: (num(ws.cell(r, 2).value) if not np.isnan(num(ws.cell(r, 2).value)) else d)
    return dict(alpha=g(4, .05), minN=g(7, 20), minMsess=g(8, 0.), useInd=int(g(9, 1)),
                smooth=int(g(10, 1)), dMsess=g(11, .02))


def read_index_sheet(ws):
    out = {}
    for r in ws.iter_rows(min_row=2, max_col=13, values_only=True):
        y = num(r[0])
        if np.isnan(y):
            continue
        for m in range(12):
            v = num(r[1 + m])
            if not np.isnan(v):
                out[int(y) * 12 + m] = v
    return out


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


# ============ Model (replika Excel) ============
def lagvec(idx, yrs, m, L):
    return np.array([idx.get(y * 12 + m - L, np.nan) for y in yrs])


def ols(y, x):
    n = len(y); ym, xm = y.mean(), x.mean()
    sxx = ((x - xm) ** 2).sum(); sy = ((y - ym) ** 2).sum()
    b = ((x - xm) * (y - ym)).sum() / sxx if sxx > 0 else 0.0
    a = ym - b * xm
    r = corr(x, y); r2 = r * r
    s = math.sqrt(max(0, (sy - b * b * sxx)) / (n - 2))
    h = 1 / n + (x - xm) ** 2 / sxx if sxx > 0 else np.full(n, 1 / n)
    e = y - a - b * x
    rm = math.sqrt((((e / (1 - h)) ** 2).sum()) / n)
    rc = math.sqrt(sy / n) * n / (n - 1)
    se = s / math.sqrt(sxx) if sxx > 0 else np.inf
    return dict(n=n, mean=ym, xbar=xm, sxx=sxx, b=b, a=a, r=r, r2=r2, s=s, p=tdist2(b / se, n - 2) if se else 1.0,
                rmse=rm, rmseC=rc, msess=1 - (rm / rc) ** 2, sd=float(np.std(y, ddof=1)))


def combined(y, x, d, ao):
    """OLS dua prediktor (rumus tertutup dari korelasi) + LOYO eksak, seperti sheet Model_ONI_DMI."""
    n = len(y); ym, xm, dm = y.mean(), x.mean(), d.mean()
    Soo = ((x - xm) ** 2).sum(); Sdd = ((d - dm) ** 2).sum(); Sod = ((x - xm) * (d - dm)).sum()
    ryo, ryd, rod = corr(y, x), corr(y, d), corr(x, d)
    sy_, sx_, sd_ = np.std(y, ddof=1), np.std(x, ddof=1), np.std(d, ddof=1)
    b1 = (ryo - ryd * rod) / (1 - rod ** 2) * sy_ / sx_
    b2 = (ryd - ryo * rod) / (1 - rod ** 2) * sy_ / sd_
    a = ym - b1 * xm - b2 * dm
    R2 = (ryo ** 2 + ryd ** 2 - 2 * ryo * ryd * rod) / (1 - rod ** 2)
    SST = ((y - ym) ** 2).sum()
    s = math.sqrt(max(0, (1 - R2)) * SST / (n - 3))
    se2 = s / math.sqrt(Sdd * (1 - rod ** 2))
    pb2 = tdist2(b2 / se2, n - 3)
    det = Soo * Sdd - Sod ** 2
    xc, dc = x - xm, d - dm
    h = 1 / n + (Sdd * xc ** 2 - 2 * Sod * xc * dc + Soo * dc ** 2) / det
    e = y - a - b1 * x - b2 * d
    rm = math.sqrt((((e / (1 - h)) ** 2).sum()) / n)
    return dict(b1=b1, b2=b2, a=a, r2=R2, s=s, pb2=pb2, rmse=rm, soo=Soo, sdd=Sdd, sod=Sod, xo=xm, xd=dm, rod=rod)


def model_series(Y, yrs, oni, dmi, P):
    """Y: array (N,12) hujan bulanan satu seri. Kembalikan daftar 12 dict (per bulan)."""
    al = P['alpha']
    # --- ONI: korelasi semua lag + FDR per seri ---
    R = np.zeros((12, 7)); PV = np.ones((12, 7)); NN = [0] * 12
    for m in range(12):
        for L in LAGS:
            x = lagvec(oni, yrs, m, L); y = Y[:, m]; k = ~np.isnan(x) & ~np.isnan(y)
            NN[m] = int(k.sum())
            R[m, L] = corr(x[k], y[k]); PV[m, L] = p_corr(R[m, L], NN[m] - 2)
    FD = bh_pass(PV, al)
    sel = [int(np.argmin(PV[m])) for m in range(12)]
    ev = []
    for m in range(12):
        ev.append('Kuat' if FD[m, sel[m]] else 'Indikatif' if PV[m, sel[m]] < al else 'Lemah / tidak ada')
    # --- DMI: korelasi parsial (kontrol ONI pada lag terpilih) + FDR per seri ---
    PP = np.ones((12, 7)); RP = np.zeros((12, 7))
    for m in range(12):
        x0 = lagvec(oni, yrs, m, sel[m]); y = Y[:, m]
        ryo = corr(x0[~np.isnan(y)], y[~np.isnan(y)])
        for L in LAGS:
            d = lagvec(dmi, yrs, m, L); k = ~np.isnan(d) & ~np.isnan(y) & ~np.isnan(x0)
            ryd = corr(d[k], y[k]); rod = corr(x0[k], d[k])
            RP[m, L] = (ryd - ryo * rod) / math.sqrt((1 - ryo ** 2) * (1 - rod ** 2))
            PP[m, L] = p_corr(RP[m, L], int(k.sum()) - 3)
    FDd = bh_pass(PP, al)
    out = []
    for m in range(12):
        y = Y[:, m]; x = lagvec(oni, yrs, m, sel[m]); k = ~np.isnan(x) & ~np.isnan(y)
        o = ols(y[k], x[k]); n = o['n']
        layak = n >= P['minN'] and o['msess'] > P['minMsess'] and (ev[m] == 'Kuat' or (ev[m] == 'Indikatif' and P['useInd'] == 1))
        lD = int(np.argmin(PP[m])); d = lagvec(dmi, yrs, m, lD)
        k2 = k & ~np.isnan(d)
        c = combined(y[k2], x[k2], d[k2], o)
        evD = 'Kuat' if FDd[m, lD] else 'Indikatif' if PP[m, lD] < al else 'Lemah / tidak ada'
        c['msess'] = 1 - (c['rmse'] / o['rmseC']) ** 2
        c['dm'] = c['msess'] - o['msess']
        add = (c['pb2'] < al and (evD == 'Kuat' or (evD == 'Indikatif' and P['useInd'] == 1)) and c['dm'] > P['dMsess']
               and c['msess'] > P['minMsess'] and n >= P['minN'])
        basis = 'ONI+DMI' if add else 'ONI saja' if layak else 'Klimatologi'
        out.append(dict(m=m, lag=sel[m], lagD=lD, ev=ev[m], evD=evD, n=n, p=o['p'], ok=layak, add=add, basis=basis,
                        rD=float(RP[m, lD]), pD=float(PP[m, lD]),
                        mean=o['mean'], sd=o['sd'], xbar=o['xbar'], sxx=o['sxx'], b=o['b'], a=o['a'], r=o['r'], r2=o['r2'],
                        s=o['s'], pct=o['b'] / o['mean'] * 100, rmse=o['rmse'], rmseC=o['rmseC'], msess=o['msess'],
                        c=dict(xo=c['xo'], xd=c['xd'], soo=c['soo'], sdd=c['sdd'], sod=c['sod'], b1=c['b1'], b2=c['b2'],
                               a=c['a'], s=c['s'], r2=c['r2'], pb2=c['pb2'], rmse=c['rmse'], msess=c['msess'], dm=c['dm'], rod=c['rod']),
                        t1=tinv(al, n - 1), t2=tinv(al, n - 2), t3=tinv(al, n - 3)))
    return out


# ============ Klimatologi (deskriptif) ============
def clim_stat(v):
    v = v[~np.isnan(v)]
    if len(v) < 5:
        return dict(mean=0, sd=1, p33=0, p67=0, med=0, min=0, max=0)
    p33, med, p67 = np.percentile(v, [100 / 3, 50, 200 / 3])
    return dict(mean=v.mean(), sd=float(v.std(ddof=1)) or 1.0, p33=p33, p67=p67, med=med, min=v.min(), max=v.max())


def clim(M):  # M: (N,12)
    return [clim_stat(M[:, m]) for m in range(12)]


# ============ Unduh indeks terkini (best effort) ============
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


def parse_weekly(txt):  # 03JAN1990  23.4-0.4  25.1-0.3  26.6 0.0  28.6 0.3 -> SSTA Nino3.4 (sama seperti kode lama)
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


def oni_from_monthly(nino):
    """ONI = rerata 3 bulan terpusat anomali Niño 3.4 (definisi ONI; sama dengan skala data latih Excel)."""
    return {t: float(np.mean([nino[k] for k in (t - 1, t, t + 1)])) for t in nino if all(k in nino for k in (t - 1, t, t + 1))}


def parse_psl(txt):  # tahun + 12 nilai; missing <= -90
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


def parse_bom(txt):  # yyyymmdd mulai, yyyymmdd akhir, nilai (mingguan) -> rata-rata bulanan (min. 2 minggu)
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
    return {t: float(np.mean(v)) for t, v in acc.items() if len(v) >= 2}


def merge_bom(dmi, bom):  # isi bulan kosong PSL dengan BoM terkalibrasi pada bulan tumpang tindih
    both = [t for t in bom if t in dmi]
    a, b = 0.0, 1.0
    if len(both) >= 24:
        sl, ic = np.polyfit([bom[t] for t in both], [dmi[t] for t in both], 1)
        if 0.3 < sl < 3:
            a, b = ic, sl
    for t, v in bom.items():
        if t not in dmi:
            dmi[t] = a + b * v


def smooth_dmi(raw, on, strict=False):
    """Rata-rata berjalan 3 bulan terpusat (m-1, m, m+1) bila on=1, seperti sheet DMI_Lag.
    strict=True: wajib ketiga bulan ada (untuk nilai terkini); False: AVERAGE Excel (abaikan sel kosong)."""
    if not on:
        return dict(raw)
    o = {}
    for t in range(min(raw) if raw else 0, (max(raw) if raw else 0) + 1):
        v = [raw[k] for k in (t - 1, t, t + 1) if k in raw]
        if (strict and len(v) == 3) or (not strict and v and t in raw):
            o[t] = float(np.mean(v))
    return o


def current_indices(oni_x, dmi_x, P):
    log = []
    oni = dict(oni_x); dmi = dict(dmi_x); live = False
    txt, m1 = fetch([ENSO_URL], 'wksst9120.for')
    w = parse_weekly(txt) if txt else {}
    if len(w) >= 20:
        o = oni_from_monthly(w)
        ov = [t for t in o if t in oni_x]
        if ov:
            log.append(f'Nino3.4 vs ONI Excel: selisih rata2 {np.mean([abs(o[t] - oni_x[t]) for t in ov]):.2f} ({len(ov)} bln)')
        for t, v in o.items():       # bulan yang sudah ada di sheet oni Excel tetap memakai nilai Excel
            oni.setdefault(t, v)
        live = 'langsung' in m1; log.append('Nino3.4 mingguan CPC: ' + m1)
    else:
        log.append('Nino3.4: ' + m1 + ' → nilai terakhir sheet oni Excel')
    txt, m2 = fetch(DMI_URLS, 'dmi_psl.data')
    d = parse_psl(txt) if txt else {}
    if len(d) > 100:
        dmi.update(d); log.append('DMI PSL: ' + m2)
    else:
        log.append('DMI PSL: ' + m2 + ' → nilai terakhir sheet dmi Excel')
    txt, m3 = fetch([BOM_URL], 'iod_1.txt')
    bom_ok = False
    if txt:
        b = parse_bom(txt)
        if b:
            merge_bom(dmi, b); bom_ok = True
    log.append('BoM: ' + m3)
    return oni, smooth_dmi(dmi, P['smooth'], strict=True), bom_ok, '; '.join(log), live


# ============ Utama ============
def clean(o):
    if isinstance(o, dict): return {k: clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)): return [clean(v) for v in o]
    if isinstance(o, (bool, np.bool_)): return bool(o)
    if isinstance(o, (float, np.floating)): return round(float(o), 6) if math.isfinite(o) else None
    if isinstance(o, np.integer): return int(o)
    return o


def main():
    u = os.environ.get('SHEET_XLSX_URL')  # opsional: ambil Excel dari Google Sheets (link export xlsx)
    if u:
        try:
            r = requests.get(u, timeout=120); r.raise_for_status()
            pathlib.Path(os.environ.get('XLSX') or 'Data_Curah_Hujan_-_Wajo.xlsx').write_bytes(r.content)
            print('Excel diambil dari Google Sheets')
        except Exception as e:
            print('Gagal ambil Google Sheets, memakai file xlsx di repo:', e)
    xl = find_xlsx(); print('Memakai', xl)
    wb = load_workbook(xl, data_only=True)
    P = read_params(wb)
    oni_x = read_index_sheet(wb['oni']); dmi_x_raw = read_index_sheet(wb['dmi'])
    dmi_x = smooth_dmi(dmi_x_raw, P['smooth'])            # untuk pelatihan = persis DMI_Lag Excel
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

    oni_c, dmi_c, bom_ok, msg, live = current_indices(oni_x, dmi_x_raw, P)
    t_last = max(oni_c)
    now = dt.datetime.now(dt.timezone.utc); t_now = now.year * 12 + now.month - 1
    hist = [dict(y=t // 12, m=t % 12, oni=oni_c.get(t), dmi=dmi_c.get(t)) for t in range(t_last - 35, t_last + 1)]
    td = max([t for t in dmi_c if t <= t_now] or [max(dmi_c)])

    stations = []
    for s in allst:
        mods = model_series(s['CH'], yrs, oni_x, dmi_x, P)
        # korelasi ringkas per bulan untuk grafik: r ONI (lag terpilih) & r parsial DMI (lag terpilih)
        stations.append(dict(
            id=s['id'], name=s['name'], loc=s.get('loc'), lat=s.get('lat'), lon=s.get('lon'),
            clim=dict(ch=clim(s['CH']), hh=clim(s['HH'])),
            mon=mods,
            fc=[dict(y=(t_now + h - 1) // 12, m=(t_now + h - 1) % 12, h=h) for h in range(1, LEADS + 1)]))

    n = len(yrs)
    out = dict(stations=stations,
               idx=dict(y=t_last // 12, m=t_last % 12, oni=oni_c[t_last], dmi=dmi_c.get(td), dy=td // 12, dm=td % 12,
                        bom=bom_ok, bomMsg=msg, live=live, hist=hist),
               meta=dict(y0=Y0, y1=Y1, n=len(st), nYears=n, alpha=P['alpha'], minN=P['minN'], minMsess=P['minMsess'],
                         useInd=P['useInd'], smooth=P['smooth'], dMsess=P['dMsess'],
                         rcrit=math.sqrt(tinv(P['alpha'], n - 2) ** 2 / (tinv(P['alpha'], n - 2) ** 2 + n - 2)),
                         source=os.path.basename(xl)),
               updated=now.strftime('%Y-%m-%d %H:%M UTC'))
    pathlib.Path('data.json').write_text(json.dumps(clean(out), ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    nm = sum(1 for s in stations for m in s['mon'] if m['basis'] != 'Klimatologi')
    print(f'data.json ditulis | hujan {Y0}-{Y1} ({n} th) | ONI terakhir {t_last // 12}-{t_last % 12 + 1:02d} '
          f'| DMI terakhir {td // 12}-{td % 12 + 1:02d} | seri-bulan memakai model: {nm}/{12 * len(stations)} | {msg}')


if __name__ == '__main__':
    main()
