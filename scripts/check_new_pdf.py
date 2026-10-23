#!/usr/bin/env python3
"""
check_new_pdf.py
Detecteert een nieuwe VLIF-codelijst op de Vlaanderen-pagina,
downloadt de PDF en start de verwerking (PDF -> Excel -> validatie -> mail).
"""

import os
import re
import json
import hashlib
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Configuratie
# ---------------------------------------------------------------------------
URL = "https://lv.vlaanderen.be/subsidies/vlif-steun/vlif-investeringssteun-voor-land-en-tuinbouwers"
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
PDF_DIR = BASE_DIR / "pdf"
EXCEL_DIR = BASE_DIR / "excel"
STATE_FILE = DATA_DIR / "last_version.txt"
PDF_STATE_FILE = DATA_DIR / "last_pdf_hash.txt"

DATA_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)
EXCEL_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Regexes
# ---------------------------------------------------------------------------
# Vangt "VLIF4 Codelijst blok ..." in alle varianten (spaties, streepjes, enz.)
LINK_TEXT_PATTERN = re.compile(r"vlif4.*codelijst.*blok", re.IGNORECASE)

# Accepteert zowel "03-2026" als "4 2026" (spatie of streepje tussen blok en jaar)
VERSION_PATTERN = re.compile(
    r"blok\s*([0-9]{1,2})\s*[-\s]\s*([0-9]{4})",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def normalize_version(raw: str) -> str:
    """Normaliseer '4-2026' of '4 2026' naar '4-2026'."""
    parts = re.split(r"[-\s]+", raw.strip())
    if len(parts) == 2:
        return f"{parts[0]}-{parts[1]}"
    return raw.strip()

def fetch_html(url: str) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        )
    }
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    return resp.text

def find_codelijst_link(html: str, base_url: str):
    """
    Zoek de eerste link met tekst die op de VLIF-codelijst lijkt.
    Retourneert (href, link_text, version) of (None, None, None).
    """
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        if LINK_TEXT_PATTERN.search(text):
            href = urljoin(base_url, a["href"])
            match = VERSION_PATTERN.search(text) or VERSION_PATTERN.search(href)
            version = None
            if match:
                version = normalize_version(f"{match.group(1)}-{match.group(2)}")
            return href, text, version
    return None, None, None

def download_pdf(url: str, dest: Path) -> bytes:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        )
    }
    resp = requests.get(url, headers=headers, timeout=120)
    resp.raise_for_status()
    if not resp.content.startswith(b"%PDF"):
        raise ValueError(f"Geen geldige PDF op {url}")
    dest.write_bytes(resp.content)
    return resp.content

def pdf_sha256(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()

def read_state(path: Path, default: str = "") -> str:
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    return default

def write_state(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8")

# ---------------------------------------------------------------------------
# Hoofdlogica
# ---------------------------------------------------------------------------
def main():
    print(f"[{datetime.now().isoformat()}] Ophalen: {URL}")
    html = fetch_html(URL)

    href, link_text, version = find_codelijst_link(html, URL)
    if not href:
        print("Geen VLIF-codelijstlink gevonden op de pagina.")
        # Optioneel: alle links loggen voor debug
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            if ".pdf" in a["href"].lower():
                print("  PDF-link:", a.get_text(strip=True), "->", a["href"])
        return

    print(f"Gevonden link: '{link_text}'")
    print(f"  href: {href}")
    print(f"  versie: {version}")

    if version is None:
        print("WAARSCHUWING: geen versie in linktekst/href gevonden. Ga verder met hash-check.")

    last_version = read_state(STATE_FILE)
    print(f"Laatst verwerkte versie: {last_version!r}")

    # Bepaal of dit een nieuwe versie is
    is_new_version = version is not None and version != last_version
    print(f"Nieuwe versie? {is_new_version}")

    # Download PDF (altijd, om hash te kunnen vergelijken)
    safe_name = f"vlif_codelijst_{version or 'onbekend'}.pdf".replace(" ", "_")
    pdf_path = PDF_DIR / safe_name
    content = download_pdf(href, pdf_path)
    new_hash = pdf_sha256(content)
    last_hash = read_state(PDF_STATE_FILE)

    print(f"PDF-hash: {new_hash[:16]}... (vorige: {last_hash[:16] if last_hash else 'geen'})")

    is_new_hash = new_hash != last_hash

    if not is_new_version and not is_new_hash:
        print("Geen nieuwe versie en geen gewijzigde PDF. Stoppen.")
        return

    # State bijwerken
    if version:
        write_state(STATE_FILE, version)
    write_state(PDF_STATE_FILE, new_hash)

    print(f"Nieuwe PDF opgeslagen: {pdf_path}")

    # -----------------------------------------------------------------------
    # Start de rest van de pipeline
    # -----------------------------------------------------------------------
    # Pas dit aan naar jouw bestaande modules/functies:
    #
    # from pdf_to_excel import convert_pdf_to_excel
    # from validate_excel import validate_excel
    # from send_email import send_email_with_attachment
    #
    # excel_path = EXCEL_DIR / f"VLIF_codelijst_{version or 'nieuw'}.xlsx"
    # convert_pdf_to_excel(pdf_path, excel_path)
    # validate_excel(excel_path)
    # send_email_with_attachment(
    #     subject=f"Nieuwe VLIF-codelijst {version or ''}",
    #     body=f"Bijgevoegd de nieuwe codelijst ({version or 'onbekende versie'}).",
    #     attachment=excel_path,
    # )
    #
    # Plaats hier de aanroepen zoals ze in jouw project heten.
    # -----------------------------------------------------------------------
    print("Pipeline gereed (vervolgstappen nog niet geconfigureerd in dit voorbeeld).")

if __name__ == "__main__":
    main()
