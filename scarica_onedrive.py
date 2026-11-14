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
import urllib.parse
import http.cookiejar

USER_AGENT = "Mozilla/5.0 (compatible; dashboard-updater/1.0)"


def share_api_url(link):
    """OneDrive personale: converte il link di condivisione nell'URL API di download diretto."""
    b64 = base64.urlsafe_b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"


_jar = http.cookiejar.CookieJar()
_opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(_jar))


def resolve(link):
    """Segue i reindirizzamenti dei link brevi (1drv.ms) e restituisce l'URL finale."""
    try:
        req = urllib.request.Request(link, headers={"User-Agent": USER_AGENT})
        with _opener.open(req, timeout=60) as resp:
            return resp.geturl()
    except Exception:
        return link


def candidate_urls(link):
    """Elenco di URL di download da provare, dal piu' probabile al meno."""
    link = link.strip()
    final = resolve(link) if ("1drv.ms" in link or "sharepoint.com" in link) else link
    urls = []
    for base in (final, link):
        if "sharepoint.com" in base:
            u = base if "download=1" in base else base + ("&" if "?" in base else "?") + "download=1"
            urls.append(u)
            host = urllib.parse.urlsplit(base).netloc
            b64 = base64.urlsafe_b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
            urls.append(f"https://{host}/_api/v2.0/shares/u!{b64}/driveitem/content")
    urls.append(share_api_url(link))
    for base in (final, link):
        if "onedrive.live.com" in base:
            q = urllib.parse.parse_qs(urllib.parse.urlsplit(base).query)
            resid = (q.get("resid") or q.get("id") or [""])[0]
            auth = (q.get("authkey") or [""])[0]
            if resid:
                d = "https://onedrive.live.com/download?resid=" + urllib.parse.quote(resid)
                if auth:
                    d += "&authkey=" + urllib.parse.quote(auth)
                cid = (q.get("cid") or [""])[0]
                if cid:
                    d += "&cid=" + urllib.parse.quote(cid)
                urls.insert(0, d)
            urls.append(base.replace("/redir?", "/download?").replace("/edit?", "/download?").replace("/view.aspx?", "/download?").replace("/edit.aspx?", "/download?"))
    out = []
    for u in urls:
        if u not in out:
            out.append(u)
    return out


LAST_INFO = {}


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with _opener.open(req, timeout=120) as resp:
        LAST_INFO["final_host"] = urllib.parse.urlsplit(resp.geturl()).netloc
        LAST_INFO["ctype"] = resp.headers.get("Content-Type", "")
        return resp.read()


def describe_html(data):
    """Info sicure (niente link) su una pagina HTML ricevuta al posto del file."""
    import re
    txt = data[:200000].decode("utf-8", "ignore")
    m = re.search(r"<title[^>]*>(.*?)</title>", txt, re.S | re.I)
    title = (m.group(1).strip() if m else "")[:100]
    flags = [k for k in ("login", "signin", "sign in", "accedi", "microsoftonline", "password", "verification", "codice") if k in txt.lower()]
    return f"host finale={LAST_INFO.get('final_host')}, titolo='{title}', parole chiave={flags}"


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
    last = "nessun tentativo"
    for i, url in enumerate(candidate_urls(link), 1):
        host = urllib.parse.urlsplit(url).netloc
        try:
            data = fetch(url)
        except urllib.error.HTTPError as e:
            last = f"{e.code} {e.reason}"
            print(f"Tentativo {i} ({host}): {last}")
            continue
        except urllib.error.URLError as e:
            last = f"rete: {e.reason}"
            print(f"Tentativo {i} ({host}): {last}")
            continue
        if data[:2] == b"PK" or data[:4] == b"\xd0\xcf\x11\xe0":
            save_xlsx(data, out_path)
            return
        last = "risposta non Excel (pagina HTML)"
        print(f"Tentativo {i} ({host}): {last} -> {describe_html(data)}")
    sys.exit(f"ERRORE: tutti i tentativi falliti (ultimo: {last}). "
             "Controlla che il link sia 'Chiunque abbia il collegamento puo' visualizzare'.")


if __name__ == "__main__":
    main()
