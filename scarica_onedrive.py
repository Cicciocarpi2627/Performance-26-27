"""
Scarica un file Excel condiviso su OneDrive tramite link pubblico
("Chiunque abbia il collegamento puo' visualizzare").

Uso:
    python scarica_onedrive.py "<LINK_CONDIVISIONE>" output.xlsx

Funziona sia con OneDrive personale (1drv.ms / onedrive.live.com)
sia con OneDrive aziendale / SharePoint (...sharepoint.com).
"""
import sys
import base64
import urllib.request
import urllib.error

USER_AGENT = "Mozilla/5.0 (compatible; dashboard-updater/1.0)"


def share_api_url(link):
    """OneDrive personale: converte il link di condivisione nell'URL API di download diretto."""
    b64 = base64.urlsafe_b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"


def direct_url(link):
    """Restituisce l'URL da cui scaricare direttamente il file."""
    link = link.strip()
    if "sharepoint.com" in link:
        if "download=1" in link:
            return link
        return link + ("&" if "?" in link else "?") + "download=1"
    return share_api_url(link)


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.read()


def save_xlsx(data, out_path):
    """Salva il file solo se e' davvero un Excel (.xlsx = archivio zip che inizia con 'PK')."""
    if data[:2] == b"PK":
        with open(out_path, "wb") as f:
            f.write(data)
        print(f"OK: scaricato {len(data):,} byte in {out_path}")
        return
    if data[:4] == b"\xd0\xcf\x11\xe0":
        sys.exit("ERRORE: il file e' nel vecchio formato .xls. Salvalo come .xlsx e riprova.")
    sys.exit(
        "ERRORE: il link non ha restituito un file Excel (probabilmente una pagina di login). "
        "Controlla che il collegamento sia impostato su 'Chiunque abbia il collegamento' "
        "e che il file non sia stato cancellato e ricaricato (in quel caso il link cambia)."
    )


def main():
    if len(sys.argv) != 3 or not sys.argv[1].strip():
        sys.exit("Uso: python scarica_onedrive.py \"<LINK>\" output.xlsx  (link vuoto: secret mancante su GitHub?)")
    link, out_path = sys.argv[1], sys.argv[2]
    url = direct_url(link)
    try:
        data = fetch(url)
    except urllib.error.HTTPError as e:
        sys.exit(f"ERRORE: OneDrive ha risposto {e.code} {e.reason}. Controlla il link di condivisione.")
    except urllib.error.URLError as e:
        sys.exit(f"ERRORE di rete: {e.reason}")
    save_xlsx(data, out_path)


if __name__ == "__main__":
    main()
