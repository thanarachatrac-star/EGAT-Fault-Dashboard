from pathlib import Path
import io
import gzip
from datetime import datetime
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse

import pandas as pd
import requests

EXCEL_SHAREPOINT_URL = (
    "https://egatucc-my.sharepoint.com/:x:/g/personal/599320_egat_co_th/"
    "IQCa3BYjVfj4RqpmHfP8_VEEAWZ5i0luFrsoelLzPAT46Nw?e=yyLIJx"
)
SHEET_NAME = "Logbook"
HEADER_ROW = 1

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "data"
OUT_FILE = OUT_DIR / "fault_snapshot.csv.gz"
META_FILE = OUT_DIR / "fault_snapshot.meta.txt"

def make_download_url(url: str) -> str:
    p = urlparse(url)
    q = dict(parse_qsl(p.query, keep_blank_values=True))
    q["download"] = "1"
    return urlunparse(p._replace(query=urlencode(q)))

print("Downloading Excel from SharePoint...")
r = requests.get(
    make_download_url(EXCEL_SHAREPOINT_URL),
    timeout=60,
    allow_redirects=True,
    headers={"User-Agent": "Mozilla/5.0 EGAT-Fault-Dashboard-Snapshot/1.0"},
)
r.raise_for_status()

data = r.content
ct = (r.headers.get("content-type") or "").lower()
head = data[:500].lstrip().lower()
if "text/html" in ct or head.startswith(b"<!doctype html") or head.startswith(b"<html"):
    raise RuntimeError("SharePoint returned a sign-in page instead of the Excel file.")

print("Parsing Logbook sheet...")
df = pd.read_excel(
    io.BytesIO(data),
    sheet_name=SHEET_NAME,
    header=HEADER_ROW,
    engine="openpyxl",
)

if "วันที่" in df.columns:
    df["วันที่"] = pd.to_datetime(df["วันที่"], errors="coerce")

# Build canonical CSV bytes first. This lets us compare the actual data
# instead of gzip timestamps or metadata timestamps.
csv_text = df.to_csv(index=False, lineterminator="\n")
csv_bytes = csv_text.encode("utf-8-sig")

OUT_DIR.mkdir(parents=True, exist_ok=True)

old_csv_bytes = None
if OUT_FILE.exists():
    try:
        with gzip.open(OUT_FILE, "rb") as f:
            old_csv_bytes = f.read()
    except Exception:
        old_csv_bytes = None

if old_csv_bytes == csv_bytes:
    print("No data changes detected. Snapshot left unchanged.")
    print(f"Rows: {len(df):,}")
else:
    # Write deterministic gzip (mtime=0) so identical data produces identical bytes.
    with OUT_FILE.open("wb") as raw:
        with gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as gz:
            gz.write(csv_bytes)

    META_FILE.write_text(
        datetime.now().isoformat(timespec="seconds"),
        encoding="utf-8",
    )

    print(f"Snapshot updated: {OUT_FILE}")
    print(f"Rows: {len(df):,}")
    print(f"Size: {OUT_FILE.stat().st_size/1024:.1f} KB")
