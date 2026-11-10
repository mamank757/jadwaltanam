import io
import requests
import pandas as pd

URL = "https://www.bom.gov.au/clim_data/IDCK000072/iod_1.txt"
CSV_OUTPUT = "iod_bom_dmi_format.csv"
EXCEL_OUTPUT = "iod_bom_dmi_format.xlsx"
MIN_WEEKS_PER_MONTH = 3  # bulan dengan data < 3 minggu dianggap belum lengkap


def download_text(url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
    }
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response.text


def parse_weekly(text):
    """Format file BOM: start_date,end_date,value (dipisah koma, tanpa spasi)."""
    df = pd.read_csv(
        io.StringIO(text),
        header=None,
        names=["Start", "End", "DMI"],
        comment="#",
        skip_blank_lines=True,
        dtype=str,
    )
    df["Start"] = pd.to_datetime(df["Start"].str.strip(), format="%Y%m%d", errors="coerce")
    df["End"] = pd.to_datetime(df["End"].str.strip(), format="%Y%m%d", errors="coerce")
    df["DMI"] = pd.to_numeric(df["DMI"], errors="coerce")
    df = df.dropna(subset=["Start", "End", "DMI"])

    if df.empty:
        raise Exception(
            "Tidak ada data yang berhasil diekstrak dari file teks BOM!\n"
            f"Cuplikan isi file: {text[:200]!r}"
        )

    # Tanggal acuan tiap minggu = titik tengah (hari ke-4) agar minggu yang
    # melintasi pergantian bulan masuk ke bulan yang paling banyak harinya.
    df["Mid"] = df["Start"] + (df["End"] - df["Start"]) / 2
    return df


def to_monthly(df):
    period = df["Mid"].dt.to_period("M")
    monthly = df.groupby(period)["DMI"].agg(
        Monthly_DMI="mean", Weeks_Count="count"
    )

    # Pastikan tidak ada bulan yang hilang agar rolling 3 bulan tetap benar
    full_index = pd.period_range(monthly.index.min(), monthly.index.max(), freq="M")
    monthly = monthly.reindex(full_index)
    monthly["Weeks_Count"] = monthly["Weeks_Count"].fillna(0).astype(int)

    # Bulan yang belum lengkap (mis. bulan berjalan) tidak dipakai
    monthly.loc[monthly["Weeks_Count"] < MIN_WEEKS_PER_MONTH, "Monthly_DMI"] = float("nan")

    monthly["DMI_3Mo"] = (
        monthly["Monthly_DMI"].rolling(window=3, center=True, min_periods=3).mean()
    )

    monthly = monthly.reset_index(names="Period")
    monthly["Date"] = monthly["Period"].dt.to_timestamp(how="end").dt.normalize()
    monthly["Year"] = monthly["Period"].dt.year
    monthly["Month"] = monthly["Period"].dt.month
    return monthly[
        ["Date", "Year", "Month", "Weeks_Count", "Monthly_DMI", "DMI_3Mo"]
    ]


def main():
    print(f"Mengunduh data dari: {URL}")
    text = download_text(URL)
    print("Berhasil mendownload. Memproses baris teks...")

    weekly = parse_weekly(text)
    print(f"Total baris mingguan valid: {len(weekly)}")

    monthly = to_monthly(weekly)
    monthly.to_csv(CSV_OUTPUT, index=False)
    monthly.to_excel(EXCEL_OUTPUT, index=False, engine="openpyxl")
    print(f"[OK] Sukses! CSV & Excel diperbarui dengan {len(monthly)} baris data bulanan.")


if __name__ == "__main__":
    main()
