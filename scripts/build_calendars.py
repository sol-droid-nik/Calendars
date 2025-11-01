# -*- coding: utf-8 -*-
import re, uuid, html
from pathlib import Path
from datetime import datetime, timedelta, time, timezone

import pandas as pd
import pytz

EXCEL_PATH = Path("data") / "tyovuorot_2025.xlsx"
OUT_DIR = Path("public") / "calendars"
OUT_DIR.mkdir(parents=True, exist_ok=True)

TZ = pytz.timezone("Europe/Helsinki")
WEEKDAYS_FI = ["MA","TI","KE","TO","PE","LA","SU"]

def parse_header_date(h: str, year=2025):
    if not isinstance(h, str): return None
    m = re.match(r"([A-ZÅÄÖ]{2})\s+(\d{1,2})[.](\d{1,2})", h.strip())
    if not m: return None
    d, mth = int(m.group(2)), int(m.group(3))
    try: return datetime(year, mth, d)
    except: return None

def extract_times(txt: str):
    if not isinstance(txt, str): return (None, None)
    t = txt.lower().replace("–","-").replace("—","-")
    m = re.search(r"klo\s*([0-9]{1,2}(?::[0-9]{2}|[.][0-9]{2})?|[0-9]{1,2})\s*-\s*([0-9]{1,2}(?::[0-9]{2}|[.][0-9]{2})?)", t)
    def norm(x):
        if not x: return None
        x=x.replace(".",":")
        if ":" not in x: x=f"{int(x):02d}:00"
        else:
            hh,mm=x.split(":"); x=f"{int(hh):02d}:{int(mm):02d}"
        return x
    if m: return (norm(m.group(1)), norm(m.group(2)))
    m2 = re.search(r"klo\s*([0-9]{1,2}(?::[0-9]{2}|[.][0-9]{2})?)", t)
    if m2: return (norm(m2.group(1)), None)
    return (None, None)

def to_time(s):
    if not isinstance(s, str) or not s: return None
    hh,mm = map(int, s.split(":")); return time(hh,mm)

def esc_ics(s: str) -> str:
    return s.replace("\\","\\\\").replace(";","\\;").replace(",","\\,").replace("\n","\\n") if s else ""

def slug_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]+", "_", name.strip()).strip("_") or "person"

def to_utc_str(local_dt: datetime) -> str:
    return local_dt.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

def read_long_from_excel(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    rows = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        df.columns = [str(c) for c in df.columns]
        df = df.rename(columns={df.columns[0]: "Name"})
        df["Name"] = df["Name"].ffill()
        day_cols = [c for c in df.columns if any(c.startswith(w+" ") for w in WEEKDAYS_FI)]
        if not day_cols: continue
        long_df = df.melt(id_vars=["Name"], value_vars=day_cols, var_name="DayHeader", value_name="Shift")
        long_df = long_df[long_df["Shift"].notna() & (long_df["Shift"].astype(str).str.strip()!="")]
        long_df["Date"] = long_df["DayHeader"].apply(parse_header_date)
        times = long_df["Shift"].apply(extract_times)
        long_df["Start"] = times.apply(lambda t: t[0])
        long_df["End"] = times.apply(lambda t: t[1])
        rows.append(long_df[["Date","Name","Shift","Start","End"]])
    out = pd.concat(rows, ignore_index=True)
    out = out.dropna(subset=["Date"]).copy()
    out["Name"] = out["Name"].str.strip()
    return out.sort_values(["Name","Date"]).reset_index(drop=True)

def build_ics_for_person(name: str, df_person: pd.DataFrame):
    default_start = time(7,0); default_hours = 8
    lines = ["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Work Shifts//Auto ICS//FI","METHOD:PUBLISH"]
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    for _, r in df_person.sort_values("Date").iterrows():
        d = r["Date"].date()
        start = to_time(r["Start"]) if isinstance(r["Start"], str) else None
        end   = to_time(r["End"])   if isinstance(r["End"], str)   else None
        if start and not end: end = (datetime.combine(d, start) + timedelta(hours=default_hours)).time()
        if not start and not end:
            start = default_start
            end   = (datetime.combine(d, start) + timedelta(hours=default_hours)).time()
        if not start and end: start = (datetime.combine(d, end) - timedelta(hours=default_hours)).time()
        dtstart_utc = to_utc_str(TZ.localize(datetime.combine(d, start)))
        dtend_utc   = to_utc_str(TZ.localize(datetime.combine(d, end)))
        summary = esc_ics(str(r["Shift"]).strip())
        lines += ["BEGIN:VEVENT",f"UID:{uuid.uuid4().hex}@workshifts","SEQUENCE:0",f"DTSTAMP:{dtstamp}",f"DTSTART:{dtstart_utc}",f"DTEND:{dtend_utc}",f"SUMMARY:{summary}","END:VEVENT"]
    lines.append("END:VCALENDAR")
    (OUT_DIR / f"{slug_name(name)}.ics").write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")

def main():
    df = read_long_from_excel(EXCEL_PATH)
    today = datetime.now(TZ).date()
    d_w = today.isoweekday()
    days_to_mon = (8 - d_w) % 7 or 7
    start = today + timedelta(days=days_to_mon)
    end   = start + timedelta(days=8*7 - 1)
    df = df[(df["Date"].dt.date >= start) & (df["Date"].dt.date <= end)].copy()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for person, grp in df.groupby("Name"):
        build_ics_for_person(person, grp)
    links = []
    for p in sorted(OUT_DIR.glob("*.ics")):
        links.append(f'<li><a href="calendars/{p.name}">{html.escape(p.stem)}</a></li>')
    idx = "<h1>Work shifts calendars</h1><ul>" + "\n".join(links) + "</ul>"
    Path("public/index.html").write_text(idx, encoding="utf-8")

if __name__ == "__main__":
    main()
