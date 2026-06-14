import pandas as pd

df = pd.read_excel("dati.xlsx")

html_table = df.to_html(index=False, classes="tabella", border=0)

template = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <title>Dati</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 2rem; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        tr:nth-child(even) {{ background-color: #fafafa; }}
    </style>
</head>
<body>
    <h1>Dati aggiornati</h1>
    {html_table}
</body>
</html>"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(template)
