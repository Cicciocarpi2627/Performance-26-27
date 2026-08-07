"""
Aggiorna il file strength-dashboard.html sostituendo l'array `const D=[...]`
con i dati provenienti dal foglio "carichi giornalieri" del file Excel Forza.
Uso:
    python genera_strength.py dati_forza.xlsx strength-dashboard.html
"""
import sys
import re
import json
import pandas as pd

SHEET_NAME = "carichi giornalieri"

EXERCISES = [
    "Deadlift", "Romanian Deadlift (RDL)", "Rack Pull", "Sumo Deadlift", "TrapBar Deadlift",
    "Single Leg Deadlift (SLDL)", "Single Leg Romanian Deadlift", "Goog Morning", "Hip Thrust",
    "American Swing", "Russian Swing", "Goblet Squat", "Back Squat", "Front Squat", "Box Squat",
    "Overhead Squat", "Zercher Squat", "Pin Squat", "Kang Squat", "Nordic Hamstring",
    "Reverse Nordic Curl", "GHD Glute e Hip Raise", "Calf Raises", "Bulgarian Squat",
    "Front e Reverse Lunge", "Walking Lunge", "Lateral Lunge", "Crossover Lunge",
    "Rear Foot Elevated Split Squat", "Front Foot Elevated Split Squat", "Step Up", "Side Step Up",
    "Pistol Squat", "SL Hip Thrust", "Pin Split Squat", "Leg Press", "Leg Extension", "Leg Curl",
    "Adductor Machine", "Abductor Machine", "Clean", "Snatch", "Pwr Clean", "Pwr Snatch",
    "Hang Clean", "Hang Snatch", "High Pull", "Push Press", "Push Jerk", "Thruster", "Swing",
    "Wall/Slam Ball",
]

COLS = [
    "PLAYER", "DATA", "AM/PM", "WEEK", "MD DIST", "SESSION N.", "PREV MD", "NEXT MD",
    "SESSION TYPE", "DIFF", "totale",
] + EXERCISES

TEXT_COLS = {"PLAYER", "SESSION TYPE", "AM/PM"}
INT_COLS = {"WEEK", "MD DIST", "SESSION N.", "PREV MD", "NEXT MD", "DIFF"}
EXCEL_ERRORS = {"#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "#NAME?"}


def round_value(col, val):
    if pd.isna(val):
        return None
    if isinstance(val, str) and val.strip().upper() in EXCEL_ERRORS:
        return None
    if col == "DATA":
        return val.strftime("%Y-%m-%d") if hasattr(val, "strftime") else str(val)
    if col in TEXT_COLS:
        return val
    if col in INT_COLS:
        try:
            return int(val)
        except (ValueError, TypeError):
            return None
    if isinstance(val, float):
        return round(val, 1)
    if isinstance(val, str):
        try:
            return round(float(val), 1)
        except (ValueError, TypeError):
            return None
    return val


def build_records(xlsx_path):
    df = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME)
    df = df.dropna(axis=1, how="all")
    print("Colonne trovate:", list(df.columns))  # debug temporaneo

    error_cells = []
    for col in COLS:
        if col not in df.columns:
            continue
        mask = df[col].apply(lambda v: isinstance(v, str) and v.strip().upper() in EXCEL_ERRORS)
        for idx in df[mask].index:
            player = df.at[idx, "PLAYER"] if "PLAYER" in df.columns else "?"
            date = df.at[idx, "DATA"] if "DATA" in df.columns else "?"
            error_cells.append(f"  - {date} | {player} | colonna '{col}' = {df.at[idx, col]!r}")
    if error_cells:
        print(f"ATTENZIONE: trovati {len(error_cells)} valori con errore Excel (verranno trattati come vuoti):")
        for line in error_cells:
            print(line)

    records = []
    for _, row in df.iterrows():
        rec = {}
        for col in COLS:
            v = round_value(col, row.get(col))
            if v is not None:
                rec[col] = v
        records.append(rec)
    return records


def update_html(html_path, records, output_path=None):
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
    new_array = "const D=" + json.dumps(records, ensure_ascii=False, separators=(",", ":")) + ";"
    pattern = re.compile(r"const D=\[.*?\];", re.DOTALL)
    if not pattern.search(html):
        raise RuntimeError("Non ho trovato 'const D=[...]' nel file HTML.")
    new_html = pattern.sub(new_array, html, count=1)
    out = output_path or html_path
    with open(out, "w", encoding="utf-8") as f:
        f.write(new_html)
    print(f"Aggiornati {len(records)} record. File scritto in: {out}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python genera_strength.py dati_forza.xlsx strength-dashboard.html")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    html_path = sys.argv[2]
    recs = build_records(xlsx_path)
    update_html(html_path, recs)
