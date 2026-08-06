#!/usr/bin/env python3
"""
check_new_pdf.py

Controleert de VLIF-webpagina op een nieuwe versie van de
"VLIF4 Codelijst blok ..." PDF, en downloadt deze naar input/ als
er een nieuwe versie beschikbaar is.

1. Haalt de HTML op van de bronpagina.
2. Zoekt de <a>-tag waarvan de zichtbare linktekst begint met
   "VLIF4 Codelijst blok" (case-insensitive).
3. Haalt uit de linktekst een versie-string, bv. "03-2026" uit
   "VLIF4 Codelijst blok 03-2026 .pdf".
4. Vergelijkt die versie met data/last_version.txt.
5. Bij een nieuwe versie: PDF downloaden naar input/, en
   data/last_version.txt bijwerken.
6. Zet GitHub Actions outputs "new_version", "pdf_path", "version".
"""

import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

SOURCE_URL = (
    "https://lv.vlaanderen.be/subsidies/vlif-steun/"
    "vlif-investeringssteun-voor-land-en-tuinbouwers"
)

LINK_TEXT_PATTERN = re.compile(r"vlif4\s*codelijst\s*blok", re.IGNORECASE)
VERSION_PATTERN = re.compile(r"blok\s*([0-9]{1,2}\s*-\s*[0-9]{4})", re.IGNORECASE)

REPO_ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = REPO_ROOT / "input"
DATA_DIR = REPO_ROOT / "data"
LAST_VERSION_FILE = DATA_DIR / "last_version.txt"


def normalize_version(raw: str) -> str:
    return re.sub(r"\s*-\s*", "-", raw.strip())


def find_codelijst_link(html: str, base_url: str):
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        if LINK_TEXT_PATTERN.search(text):
            href = urljoin(base_url, a["href"])
            match = VERSION_PATTERN.search(text) or VERSION_PATTERN.search(href)
            version = normalize_version(match.group(1)) if match else None
            return href, text, version
    return None, None, None


def write_github_output(name: str, value: str) -> None:
    gh_output = os.environ.get("GITHUB_OUTPUT")
    if gh_output:
        with open(gh_output, "a", encoding="utf-8") as f:
            f.write(f"{name}={value}\n")
    else:
        print(f"[output] {name}={value}")


def main() -> int:
    print(f"Bronpagina ophalen: {SOURCE_URL}")
    resp = requests.get(SOURCE_URL, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()

    href, link_text, version = find_codelijst_link(resp.text, SOURCE_URL)

    if not href:
        print("FOUT: geen link gevonden die begint met 'VLIF4 Codelijst blok'.", file=sys.stderr)
        write_github_output("new_version", "false")
        return 1

    if not version:
        print(f"FOUT: kon geen versie afleiden uit '{link_text}'.", file=sys.stderr)
        write_github_output("new_version", "false")
        return 1

    print(f"Gevonden link: {link_text}")
    print(f"URL: {href}")
    print(f"Versie: {version}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    last_version = LAST_VERSION_FILE.read_text(encoding="utf-8").strip() \
        if LAST_VERSION_FILE.exists() else None

    if version == last_version:
        print(f"Geen nieuwe versie (huidige versie '{version}' al verwerkt).")
        write_github_output("new_version", "false")
        write_github_output("previous_version", last_version or "")
        return 0

    print(f"Nieuwe versie gedetecteerd: '{last_version}' -> '{version}'.")

    oude_pdfs = list(INPUT_DIR.glob("*.pdf"))
    if oude_pdfs:
        print(f"Oude PDF('s) verwijderen uit {INPUT_DIR}:")
        for oude_pdf in oude_pdfs:
            print(f"  - {oude_pdf.name}")
            oude_pdf.unlink()

    print("Nieuwe PDF downloaden...")
    pdf_resp = requests.get(href, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    pdf_resp.raise_for_status()

    filename = href.split("/")[-1].split("?")[0].replace("%20", " ")
    if not filename.lower().endswith(".pdf"):
        filename = f"VLIF4 Codelijst blok {version}.pdf"

    pdf_path = INPUT_DIR / filename
    pdf_path.write_bytes(pdf_resp.content)
    print(f"PDF opgeslagen als: {pdf_path}")

    LAST_VERSION_FILE.write_text(version, encoding="utf-8")

    write_github_output("new_version", "true")
    write_github_output("pdf_path", str(pdf_path))
    write_github_output("version", version)
    write_github_output("previous_version", last_version or "")
    return 0


if __name__ == "__main__":
    sys.exit(main())
