# -*- coding: utf-8 -*-
import re
import uuid
import html
from pathlib import Path
from datetime import datetime, timedelta, time, timezone, date

import pandas as pd
import pytz


TZ = pytz.timezone("Europe/Helsinki")
WEEKDAYS_FI = ["MA", "TI", "KE", "TO", "PE", "LA", "SU"]
EXCLUDE_CALENDARS = {"vanha caddy", "uusi caddy"}

DATA_DIR = Path("data")
OUT_DIR = Path("public") / "calendars"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def pick_excel_file(data_dir: Path) -> Path:
    """Берём самый свежий .xlsx из data/ (по mtime)."""
    cands = list(data_dir.glob("*.xlsx"))
    if not cands:
        raise FileNotFoundError("В папке data нет .xlsx. Загрузите файл расписания в data/.")
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0]


EXCEL_PATH = pick_excel_file(DATA_DIR)
print("Using Excel:", EXCEL_PATH)


m = re.search(r"(20\d{2})", EXCEL_PATH.name)
BASE_YEAR = int(m.group(1)) if m else datetime.now(TZ).year
print("BASE_YEAR:", BASE_YEAR)


def extract_year_from_text(s: str):
    m = re.search(r"\b(20\d{2})\b", str(s))
    return int(m.group(1)) if m else None


def parse_header_date(header: str, year_hint=None, week_hint=None):
    if not isinstance(header, str):
        return None

    m = re.match(
        r"([A-ZÅÄÖ]{2})\s+(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?",
        header.strip()
    )
    if not m:
        return None

    day = int(m.group(2))
    month = int(m.group(3))

    # 1) Если год указан прямо в заголовке (1.1.2026) — он главный
    if m.group(4):
        y = int(m.group(4))
        year = 2000 + y if y < 100 else y

    # 2) Иначе берём год из вкладки (vko 1 2026)
    elif year_hint:
        year = year_hint

        # ✅ ISO-ловушка: vko 1 2026 может включать дни декабря 2025
        if week_hint == 1 and month == 12:
            year = year_hint - 1

    # 3) Иначе запасной вариант
    else:
        now = datetime.now(TZ)
        year = now.year
        if month <= 3 and now.month >= 11:
            year += 1
        elif month >= 11 and now.month <= 3:
            year -= 1

    try:
        return datetime(year, month, day)
    except:
        return None
        

def extract_times(txt: str):
    """Возвращает ('HH:MM','HH:MM') или (start,None). Понимает 'klo 7-15' и т.п."""
    if not isinstance(txt, str):
        return (None, None)
    t = txt.lower().replace("–", "-").replace("—", "-")
    m = re.search(
        r"klo\s*([0-9]{1,2}(?::[0-9]{2}|[.][0-9]{2})?|[0-9]{1,2})\s*-\s*([0-9]{1,2}(?::[0-9]{2}|[.][0-9]{2})?)",
        t,
    )

    def norm(x):
        if not x:
            return None
        x = x.replace(".", ":")
        if ":" not in x:
            x = f"{int(x):02d}:00"
        else:
            hh, mm = x.split(":")
            x = f"{int(hh):02d}:{int(mm):02d}"
        return x

    if m:
        return (norm(m.group(1)), norm(m.group(2)))

    m2 = re.search(r"klo\s*([0-9]{1,2}(?::[0-9]{2}|[.][0-9]{2})?)", t)
    if m2:
        return (norm(m2.group(1)), None)

    return (None, None)


def to_time(s):
    if not isinstance(s, str) or not s:
        return None
    hh, mm = map(int, s.split(":"))
    return time(hh, mm)


def esc_ics(s: str) -> str:
    return (
        s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")
        if s
        else ""
    )


def slug_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]+", "_", name.strip()).strip("_") or "person"


def to_utc_str(local_dt: datetime) -> str:
    return local_dt.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def stable_uid(name: str, dt_local: datetime, summary: str) -> str:
    """Стабильный UID, чтобы обновления не плодили дубликаты."""
    base = f"{name}|{dt_local.strftime('%Y-%m-%dT%H:%M')}|{summary}"
    return uuid.uuid5(uuid.NAMESPACE_URL, base).hex + "@workshifts"

def extract_week_from_sheet(sheet_name: str):
    m = re.search(r"\bvko\s*(\d{1,2})\b", str(sheet_name), flags=re.IGNORECASE)
    return int(m.group(1)) if m else None
    

def read_long_from_excel(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)

    base_year = extract_year_from_text(path.name) or datetime.now(TZ).year
    rows = []

    for sheet in xls.sheet_names:
        m_year = re.search(r"\b(20\d{2})\b", str(sheet))
        sheet_year = int(m_year.group(1)) if m_year else None
        sheet_week = extract_week_from_sheet(sheet)

        # ✅ главное: если в имени вкладки нет года — используем год из имени файла
        year_for_sheet = sheet_year or base_year

        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        df.columns = [str(c) for c in df.columns]
        df = df.rename(columns={df.columns[0]: "Name"})
        df["Name"] = df["Name"].ffill()

        def norm_ws(s: str) -> str:
            return s.replace("\u00A0", " ").strip()

        day_cols = [
            c for c in df.columns
            if any(norm_ws(c).upper().startswith(w + " ") for w in WEEKDAYS_FI)
        ]
        if not day_cols:
            continue

        long_df = df.melt(
            id_vars=["Name"], value_vars=day_cols, var_name="DayHeader", value_name="Shift"
        )
        long_df = long_df[
            long_df["Shift"].notna() & (long_df["Shift"].astype(str).str.strip() != "")
        ]

        long_df["Date"] = long_df["DayHeader"].apply(
            lambda h: parse_header_date(h, year_hint=year_for_sheet, week_hint=sheet_week)
        )

        times = long_df["Shift"].apply(extract_times)
        long_df["Start"] = times.apply(lambda t: t[0])
        long_df["End"] = times.apply(lambda t: t[1])

        rows.append(long_df[["Date", "Name", "Shift", "Start", "End"]])

    if not rows:
        return pd.DataFrame(columns=["Date", "Name", "Shift", "Start", "End"])

    out = pd.concat(rows, ignore_index=True)
    out = out.dropna(subset=["Date"]).copy()
    out["Name"] = out["Name"].astype(str).str.strip()
    return out.sort_values(["Name", "Date"]).reset_index(drop=True)
    

def build_ics_for_person(name: str, df_person: pd.DataFrame):
    default_start = time(7, 0)
    default_hours = 8

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Work Shifts//Auto ICS//FI",
        "METHOD:PUBLISH",
    ]

    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    for _, r in df_person.sort_values("Date").iterrows():
        d = r["Date"].date()

        start = to_time(r["Start"]) if isinstance(r["Start"], str) else None
        end = to_time(r["End"]) if isinstance(r["End"], str) else None

        if start and not end:
            end = (datetime.combine(d, start) + timedelta(hours=default_hours)).time()
        if not start and not end:
            start = default_start
            end = (datetime.combine(d, start) + timedelta(hours=default_hours)).time()
        if not start and end:
            start = (datetime.combine(d, end) - timedelta(hours=default_hours)).time()

        local_start = TZ.localize(datetime.combine(d, start))
        local_end = TZ.localize(datetime.combine(d, end))

        dtstart_utc = to_utc_str(local_start)
        dtend_utc = to_utc_str(local_end)

        summary = esc_ics(str(r["Shift"]).strip())
        uid = stable_uid(name, local_start, summary)

        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            "SEQUENCE:0",
            f"DTSTAMP:{dtstamp}",
            f"DTSTART:{dtstart_utc}",
            f"DTEND:{dtend_utc}",
            f"SUMMARY:{summary}",
            "END:VEVENT",
        ]

    lines.append("END:VCALENDAR")
    (OUT_DIR / f"{slug_name(name)}.ics").write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")


def main():
    df = read_long_from_excel(EXCEL_PATH)

    # --- окно, чтобы календарь не разрастался ---
    today = datetime.now(TZ).date()
    window_start = today - timedelta(days=14)
    window_end = today + timedelta(days=120)

    if not df.empty:
        df = df[(df["Date"].dt.date >= window_start) & (df["Date"].dt.date <= window_end)].copy()

    # Генерируем .ics + собираем отображаемые имена
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    slug_to_display = {}  # file_slug -> красивое имя из Excel

    for person, grp in df.groupby("Name"):
        person_clean = str(person).strip()
        if person_clean.lower() in EXCLUDE_CALENDARS:
            continue
        if "caddy" in person_clean.lower():
            continue

        build_ics_for_person(person_clean, grp)

        slug_to_display[slug_name(person_clean)] = person_clean

    # Собираем красивую главную страницу из slug_to_display (а не из имён файлов)
    cards = []
    for slug, display_name in sorted(slug_to_display.items(), key=lambda x: x[1].lower()):
        fname = f"{slug}.ics"
        cards.append(f"""
        <div class="person">
          <div class="name">{html.escape(display_name)}</div>
          <div class="btns">
            <a class="apple" data-file="{fname}" href="calendars/{fname}"> Apple</a>
            <a class="google" data-file="{fname}" href="calendars/{fname}">Google</a>
            <a class="raw" href="calendars/{fname}" download>.ics</a>
          </div>
        </div>
        """)

    html_page = f"""<!DOCTYPE html>
<html lang="fi"><head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Työvuorot kalenterit</title>
  <style>
    body {{ font-family: system-ui, -apple-system, Arial, sans-serif; background:#fafafa; margin:0; }}
    .wrap {{ max-width: 760px; margin: 32px auto; padding: 0 16px; }}
    h1 {{ text-align:center; margin: 0 0 16px; }}
    p.note {{ text-align:center; color:#666; margin: 0 0 24px; }}
    .person {{ background:#fff; margin:10px 0; padding:12px 14px; border-radius:10px;
              box-shadow:0 1px 3px rgba(0,0,0,.08); display:flex; justify-content:space-between;
              align-items:center; gap:10px; }}
    .name {{ font-weight:600; }}
    .btns a {{ display:inline-block; padding:8px 10px; border-radius:8px; text-decoration:none;
               border:1px solid #ddd; margin-left:6px; }}
  </style>
</head><body>
<div class="wrap">
  <h1>📅 Työvuorot</h1>
  <p class="note">Valitse oma nimi ja lisää kalenteri.</p>
  {''.join(cards) if cards else '<p class="note">Ei kalentereita löytynyt.</p>'}
</div>
<script>
  const base = location.origin + location.pathname.replace(/\\/?$/, '/') + 'calendars/';
  document.querySelectorAll('.apple').forEach(a => {{
    const u = base + a.dataset.file;
    a.href = 'webcal://' + u.replace(/^https?:\\/\\//, '');
  }});
  document.querySelectorAll('.google').forEach(a => {{
    const u = base + a.dataset.file;
    a.href = 'https://calendar.google.com/calendar/u/0/r?cid=' + encodeURIComponent(u);
  }});
</script>
</body></html>"""
    Path("public/index.html").write_text(html_page, encoding="utf-8")


if __name__ == "__main__":
    main()
