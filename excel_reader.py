from __future__ import annotations

from openpyxl import load_workbook
from datetime import datetime, date, timedelta
from pathlib import Path
import csv

# =====================
# CONFIG
# =====================
# Chemin du fichier Excel
BASE_DIR = Path(__file__).resolve().parent
XLSX_PATH = BASE_DIR / "ROW 2026 FR113 NextGen Schedule 4.1_.xlsx"
SHEET_NAME = None
ROW_MONTH = 1
ROW_DAY = 4          # jour du mois / date
ROW_WEEKDAY = 3      # Mon/Tue/Wed/Fri...
ROW_TARGET = 37      # codes

ASTREINTE_CODES = {"-14", "-X"}
BRIDGE_CODE = "H"
CONGES_CODES = {"F", "V"}

SAMEDI_TRAVAIL_CODE = 8.5
SAMEDI_TRAVAIL_WEEKDAY = "sat"

EXPORT_CSV = True
# Les noms finaux seront suffixés avec JJMMAAHHMM
CSV_OUT_BASE = "exportcsv"

EXPORT_ICS = True
# --- Export calendriers ICS ---
EXPORT_ICS_PER_CATEGORY = True   # 1 fichier ICS par catégorie
EXPORT_ICS_ALL_IN_ONE = False    # + 1 fichier global (tout) si True
ICS_OUT_BASE = "exportcalendrier"
ICS_CALNAME = "Astreintes"
ICS_TIMEZONE = "Europe/Paris"
ICS_EVENT_TITLE_ASTREINTE = "Astreinte"
ICS_EVENT_TITLE_CONGES = "Congés"
ICS_EVENT_TITLE_SAMEDI_TRAV = "Samedi travaillé"
ICS_EVENT_TITLE_REPOS_HEBDO = "Repos hebdomadaire"

INCLUDE_SINGLE_DAY_CONGES = True
RUN_TESTS = False

# =====================
# HELPERS
# =====================
FR_DAYS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]


def fmt(d: date) -> str:
    return f"{FR_DAYS[d.weekday()]} {d.strftime('%d/%m/%y')}"


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def norm(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return str(v).replace('.0', '')
    return str(v).strip()


def to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '.'))
    except Exception:
        return None


def is_weekday_sat(v) -> bool:
    return v is not None and str(v).strip().lower().startswith(SAMEDI_TRAVAIL_WEEKDAY)



def is_samedi_trav_code(v) -> bool:
    f = to_float(v)
    return f is not None and abs(f - SAMEDI_TRAVAIL_CODE) < 1e-9


def is_weekend_label(v) -> bool:
    """True si la ligne 3 indique Sat/Sun (peu importe la casse)."""
    if v is None:
        return False
    s = str(v).strip().lower()
    return s.startswith("sat") or s.startswith("sun")


def is_repos_hebdo_code(v) -> bool:
    """Repos hebdo si code = 'X' (exact)."""
    if v is None:
        return False
    return str(v).strip().upper() == "X"


def group_consecutive_days(days: set[date]):
    if not days:
        return []
    ds = sorted(days)
    blocks = []
    start = prev = ds[0]
    for d in ds[1:]:
        if (d - prev).days == 1:
            prev = d
        else:
            blocks.append((start, prev))
            start = prev = d
    blocks.append((start, prev))
    return blocks


def filter_conges_blocks(blocks, include_single):
    return blocks if include_single else [b for b in blocks if b[0] != b[1]]


def ics_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def block_uid(prefix: str, start: date, end: date) -> str:
    return f"{prefix}-{start.isoformat()}_{end.isoformat()}@local"

# =====================
# READ EXCEL
# =====================
wb = load_workbook(XLSX_PATH, data_only=True)
ws = wb[SHEET_NAME] if SHEET_NAME else wb.active

max_col = ws.max_column
date_code = {}
unknown_cols = []
samedi_trav_days: set[date] = set()
repos_hebdo_days: set[date] = set()

for col in range(1, max_col + 1):
    code_raw = ws.cell(row=ROW_TARGET, column=col).value
    code = norm(code_raw)
    if not code:
        continue

    d = to_date(ws.cell(row=ROW_DAY, column=col).value)
    if d is None:
        day = ws.cell(row=ROW_DAY, column=col).value
        month = ws.cell(row=ROW_MONTH, column=col).value
        if isinstance(day, (int, float)) and isinstance(month, (int, float)):
            try:
                d = date(2026, int(month), int(day))
            except Exception:
                pass

    if d is None:
        unknown_cols.append(col)
        continue

    wd = ws.cell(row=ROW_WEEKDAY, column=col).value
    if is_weekday_sat(wd) and is_samedi_trav_code(code_raw):
        # Samedi travaillé : 8,5 sur un 'Sat' -> événement sur le samedi (même date)
        samedi_trav_days.add(d)

    # Repos hebdomadaire : si 'X' et que le libellé jour n'est pas Sat/Sun
    if is_repos_hebdo_code(code_raw) and not is_weekend_label(wd):
        repos_hebdo_days.add(d)

    priority = {"-14": 3, "-X": 3, "H": 2, "F": 1, "V": 1}
    prev = date_code.get(d)
    if prev is None or priority.get(code, 0) > priority.get(prev, 0):
        date_code[d] = code

# =====================
# COMPUTE SETS
# =====================
sorted_days = sorted(date_code.items())
astreinte_days = set()
conges_days = set()

for d, c in sorted_days:
    if c in ASTREINTE_CODES:
        astreinte_days.add(d)
    if c in CONGES_CODES:
        conges_days.add(d)

# H pont
seg = []
segments = []
prev_d = None
for d, c in sorted_days:
    if prev_d is None or (d - prev_d).days == 1:
        seg.append((d, c))
    else:
        segments.append(seg)
        seg = [(d, c)]
    prev_d = d
if seg:
    segments.append(seg)

for seg in segments:
    i = 0
    while i < len(seg):
        if seg[i][1] in ASTREINTE_CODES:
            j = i + 1
            while j < len(seg) and seg[j][1] == BRIDGE_CODE:
                j += 1
            if j < len(seg) and seg[j][1] in ASTREINTE_CODES:
                for k in range(i + 1, j):
                    astreinte_days.add(seg[k][0])
            i = j
        else:
            i += 1

# =====================
# BLOCKS
# =====================
astreinte_blocks = group_consecutive_days(astreinte_days)
conges_blocks = filter_conges_blocks(group_consecutive_days(conges_days), INCLUDE_SINGLE_DAY_CONGES)
samedi_trav_blocks = group_consecutive_days(samedi_trav_days)
repos_hebdo_blocks = group_consecutive_days(repos_hebdo_days)

# =====================
# DISPLAY
# =====================
print("\n=== ASTREINTES ===")
print(f"Jours : {len(astreinte_days)} | Blocs : {len(astreinte_blocks)}")
for s, e in astreinte_blocks:
    print(f"Astreinte : {fmt(s)}" if s == e else f"Astreinte : {fmt(s)} -> {fmt(e)}")

print("\n=== CONGÉS ===")
print(f"Jours : {len(conges_days)} | Blocs : {len(conges_blocks)}")
for s, e in conges_blocks:
    print(f"Congés : {fmt(s)}" if s == e else f"Congés : {fmt(s)} -> {fmt(e)}")

print("\n=== SAMEDI TRAVAILLÉ ===")
print(f"Jours : {len(samedi_trav_days)} | Blocs : {len(samedi_trav_blocks)}")
for s, e in samedi_trav_blocks:
    print(f"Samedi travaillé : {fmt(s)}" if s == e else f"Samedi travaillé : {fmt(s)} -> {fmt(e)}")

print("\n=== REPOS HEBDOMADAIRE ===")
print(f"Jours : {len(repos_hebdo_days)} | Blocs : {len(repos_hebdo_blocks)}")
for s, e in repos_hebdo_blocks:
    print(f"Repos hebdo : {fmt(s)}" if s == e else f"Repos hebdo : {fmt(s)} -> {fmt(e)}")

if unknown_cols:
    print("\n[WARN] Colonnes sans date exploitable :", unknown_cols)

# =====================
# EXPORT CSV
# =====================
# Horodatage JJMMAAHHMM
timestamp = datetime.now().strftime("%d%m%y%H%M")
CSV_OUT = f"{CSV_OUT_BASE}_{timestamp}.csv"
ICS_OUT = f"{ICS_OUT_BASE}_{timestamp}.ics"  # (base, mais l'export ICS utilise des suffixes)


def _block_days(s: date, e: date) -> int:
    return (e - s).days + 1


def _all_dates_union(*sets_: set[date]) -> list[date]:
    u: set[date] = set()
    for st in sets_:
        u |= st
    return sorted(u)


def _write_clean_csv(path: Path) -> None:
    """CSV en 3 sections : SUMMARY / BLOCKS / DAYS (séparées par une ligne vide).

    - Délimiteur ';' (pratique pour Excel FR)
    - SUMMARY : totaux par catégorie + total jours uniques
    - BLOCKS  : blocs consécutifs avec nb_jours
    - DAYS    : ligne par date (0/1 par catégorie)
    """

    totals = {
        "astreinte_jours": len(astreinte_days),
        "astreinte_blocs": len(astreinte_blocks),
        "conges_jours": len(conges_days),
        "conges_blocs": len(conges_blocks),
        "samedi_trav_jours": len(samedi_trav_days),
        "samedi_trav_blocs": len(samedi_trav_blocks),
        "repos_hebdo_jours": len(repos_hebdo_days),
        "repos_hebdo_blocs": len(repos_hebdo_blocks),
    }
    totals["total_jours_uniques"] = len(_all_dates_union(astreinte_days, conges_days, samedi_trav_days, repos_hebdo_days))

    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")

        # --- SUMMARY ---
        w.writerow(["# SUMMARY"])
        w.writerow(["champ", "valeur"])
        w.writerow(["export_timestamp", timestamp])
        w.writerow(["xlsx_path", XLSX_PATH])
        for k, v in totals.items():
            w.writerow([k, v])
        if unknown_cols:
            w.writerow(["warn_unknown_cols_count", len(unknown_cols)])

        w.writerow([])

        # --- BLOCKS ---
        w.writerow(["# BLOCKS"])
        w.writerow(["categorie", "debut", "fin", "nb_jours"])

        def write_blocks(cat: str, blocks: list[tuple[date, date]]):
            for s, e in blocks:
                w.writerow([cat, fmt(s), fmt(e), _block_days(s, e)])

        write_blocks("astreinte", astreinte_blocks)
        write_blocks("conges", conges_blocks)
        write_blocks("samedi_travaille", samedi_trav_blocks)
        write_blocks("repos_hebdo", repos_hebdo_blocks)

        w.writerow([])

        # --- DAYS ---
        w.writerow(["# DAYS"])
        w.writerow(["date", "astreinte", "conges", "samedi_travaille", "repos_hebdo"])
        for d in _all_dates_union(astreinte_days, conges_days, samedi_trav_days, repos_hebdo_days):
            w.writerow([
                fmt(d),
                1 if d in astreinte_days else 0,
                1 if d in conges_days else 0,
                1 if d in samedi_trav_days else 0,
                1 if d in repos_hebdo_days else 0,
            ])


if EXPORT_CSV:
    csv_path = Path(CSV_OUT)
    _write_clean_csv(csv_path)
    print("\nCSV exporté :", csv_path.resolve())

# =====================
# EXPORT ICS
# =====================
# NOTE: si EXPORT_ICS_PER_CATEGORY=True, on génère 1 fichier par catégorie.
#       si EXPORT_ICS_ALL_IN_ONE=True, on génère aussi un fichier "tout".
#       si EXPORT_ICS_PER_CATEGORY=False, on génère uniquement le fichier "tout".


def write_ics(path: Path, calname: str, events: list[tuple[str, str, date, date]]) -> None:
    """Ecrit un fichier ICS (événements journée entière)."""
    lines: list[str] = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//excel_reader//FR",
        "CALSCALE:GREGORIAN",
        f"X-WR-CALNAME:{ics_escape(calname)}",
        f"X-WR-TIMEZONE:{ics_escape(ICS_TIMEZONE)}",
    ]

    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    for title, prefix, s, e in events:
        uid = ics_escape(block_uid(prefix, s, e))
        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{dtstamp}",
            f"DTSTART;VALUE=DATE:{s.strftime('%Y%m%d')}",
            f"DTEND;VALUE=DATE:{(e + timedelta(days=1)).strftime('%Y%m%d')}",
            f"SUMMARY:{ics_escape(title)}",
            "END:VEVENT",
        ])

    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")


if EXPORT_ICS:
    # Nom de fichier : exportcalendrier_<suffix>_JJMMAAHHMM.ics
    def ics_path_for(suffix: str) -> Path:
        return Path(f"{ICS_OUT_BASE}_{suffix}_{timestamp}.ics")

    events_astreintes = [(ICS_EVENT_TITLE_ASTREINTE, "astreinte", s, e) for s, e in astreinte_blocks]
    events_conges = [(ICS_EVENT_TITLE_CONGES, "conges", s, e) for s, e in conges_blocks]
    events_samedi = [(ICS_EVENT_TITLE_SAMEDI_TRAV, "samedi_travaille", s, e) for s, e in samedi_trav_blocks]
    events_repos = [(ICS_EVENT_TITLE_REPOS_HEBDO, "repos_hebdo", s, e) for s, e in repos_hebdo_blocks]

    events_all = events_astreintes + events_conges + events_samedi + events_repos

    if EXPORT_ICS_PER_CATEGORY:
        write_ics(ics_path_for("astreintes"), "Astreintes", events_astreintes)
        write_ics(ics_path_for("conges"), "Congés", events_conges)
        write_ics(ics_path_for("samedi_travaille"), "Samedi travaillé", events_samedi)
        write_ics(ics_path_for("repos_hebdo"), "Repos hebdomadaire", events_repos)

    if EXPORT_ICS_ALL_IN_ONE or not EXPORT_ICS_PER_CATEGORY:
        write_ics(ics_path_for("tout"), ICS_CALNAME, events_all)
