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
    "DATA", "WEEK", "NEXT MD", "SESSION TYPE", "PLAYER", "DIFF",
    "MIN", "D", "D/MIN", "D>16", "D>20", "D>25", "D>30",
    "DA>2", "DA>3", "DD<-2", "DD<-3",
]

INT_COLS = {"WEEK", "NEXT MD", "DIFF"}


def round_value(col, val):
    if pd.isna(val):
        return None
    if col == "DATA":
        return val.strftime("%Y-%m-%d") if hasattr(val, "strftime") else str(val)
    if col in INT_COLS:
        return int(val)
    if isinstance(val, float):
        return round(val, 1)
    return val


def build_records(xlsx_path):
    df = pd.read_excel(xlsx_path)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(subset=["PLAYER", "D"])
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
