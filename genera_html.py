"""
Aggiorna il file index.html (dashboard GPS) sostituendo l'array `const D=[...]`
con i dati provenienti dal file Excel "Dati Giornalieri per pBI".
Uso:
    python genera_html.py dati.xlsx index.html
"""
import sys
import re
import json
import pandas as pd

COLS = [
    "DATA", "AM/PM", "WEEK", "NEXT MD", "SESSION TYPE", "PLAYER", "DIFF",
    "TQR", "sRPE", "MIN", "TL",
    "FC rip", "FCmed", "FCmax", "tFC>60", "tFC>70", "tFC>85", "tFC>92",
    "D/MIN", "D", "D>16", "D>20", "D>25", "D>30",
    "DA>2", "DA>3", "DD<-2", "DD<-3",
    "nD>16", "nD>20", "nD>25", "nD>30",
    "nDA>2", "nDA>3", "nDD<-2", "nDD<-3",
    "TLGPSL", "TLGPSM", "TLGPSH", "TLGPST",
    "Vmax", "Amax", "Dmax",
]
INT_COLS = {
    "WEEK", "NEXT MD", "DIFF", "TQR", "sRPE",
    "nD>16", "nD>20", "nD>25", "nD>30",
    "nDA>2", "nDA>3", "nDD<-2", "nDD<-3",
}
EXCEL_ERRORS = {"#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "#NAME?"}
def round_value(col, val):
    if pd.isna(val):
        return None
    if isinstance(val, str) and val.strip().upper() in EXCEL_ERRORS:
        return None
    if col == "DATA":
        return val.strftime("%Y-%m-%d") if hasattr(val, "strftime") else str(val)
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
    df = pd.read_excel(xlsx_path)
    df = df.dropna(axis=1, how="all")
    print("Colonne trovate:", list(df.columns))  # debug temporaneo

    # Segnala eventuali errori Excel (#REF!, #VALUE!, ecc.) trovati nel file,
    # cosi' si sa subito quale riga correggere alla fonte invece di scoprirlo da un crash.
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
        print("Uso: python genera_html.py dati.xlsx index.html")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    html_path = sys.argv[2]
    recs = build_records(xlsx_path)
    update_html(html_path, recs)
