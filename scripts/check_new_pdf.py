#!/usr/bin/env python3
"""
check_new_pdf.py

Trigger = gewijzigde bestandsnaam van de VLIF-codelijst.
Geen versie-parsing als voorwaarde meer: als de naam anders is,
is het een nieuw bestand en start de pipeline.

State: data/last_filename.txt  (de laatst verwerkte bestandsnaam)
"""

import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, unquote

import requests
from playwright.sync_api import sync_playwright

SOURCE_URL = (
    "https://lv.vlaanderen.be/subsidies/vlif-steun/"
    "vlif-investeringssteun-voor-land-en-tuinbouwers"
)

LINK_TEXT_PATTERN = re.compile(r"vlif4\s*codelijst\s*blok", re.IGNORECASE)

# Alleen voor de mail-output; blokkeert niets.
VERSION_PATTERN = re.compile(
    r"blok\s*([0-9]{1,2})\s*[-\s]\s*([0-9]{4})", re.IGNORECASE
)

REPO_ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = REPO_ROOT / "input"
DATA_DIR = REPO_ROOT / "data"
LAST_NAME_FILE = DATA_DIR / "last_filename.txt"


def find_codelijst_link(links, base_url):
    for link in links:
        if LINK_TEXT_PATTERN.search(link["text"]):
            return urljoin(base_url, link["href"]), link["text"]
    return None, None


def pdf_filename(href: str) -> str:
    name = href.split("/")[-1].split("?")[0]
    return unquote(name).strip()


def extract_version(text: str) -> str:
    m = VERSION_PATTERN.search(text)
    return f"{m.group(1)}-{m.group(2)}" if m else ""


def write_github_output(name, value):
    gh_output = os.environ.get("GITHUB_OUTPUT")
    if gh_output:
        with open(gh_output, "a", encoding="utf-8") as f:
            f.write(f"{name}={value}\n")
    else:
        print(f"[output] {name}={value}")


def main() -> int:
    print(f"Bronpagina renderen: {SOURCE_URL}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(SOURCE_URL, wait_until="networkidle", timeout=120_000)
        try:
            page.wait_for_selector("text=VLIF4 Codelijst", timeout=30_000)
        except Exception:
            print("Waarschuwing: 'VLIF4 Codelijst' niet binnen 30s gevonden.")
        links = page.eval_on_selector_all(
            "a[href]",
            """els => els.map(a => ({
                text: (a.innerText || '').trim(),
                href: a.href
            }))""",
        )
        browser.close()

    href, link_text = find_codelijst_link(links, SOURCE_URL)
    if not href:
        print("FOUT: geen VLIF4-codelijstlink gevonden.", file=sys.stderr)
        for l in links:
            if ".pdf" in l["href"].lower():
                print(f"  '{l['text']}' -> {l['href']}")
        write_github_output("new_version", "false")
        return 1

    new_name = pdf_filename(href)
    new_version = extract_version(link_text) or extract_version(new_name)

    print(f"Linktekst    : {link_text}")
    print(f"Bestandsnaam : {new_name}")
    print(f"Versie (info): {new_version or 'onbekend'}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    last_name = LAST_NAME_FILE.read_text(encoding="utf-8").strip() \
        if LAST_NAME_FILE.exists() else ""

    print(f"Laatste naam : {last_name or '(geen)'}")

    # === TRIGGER: bestandsnaam verschilt -> nieuw bestand ===
    if new_name == last_name:
        print("Bestandsnaam ongewijzigd. Geen nieuwe versie.")
        write_github_output("new_version", "false")
        write_github_output("previous_version", last_name)
        return 0

    print(f"NIEUW BESTAND: '{last_name}' -> '{new_name}'")

    for old in INPUT_DIR.glob("*.pdf"):
        print(f"  Verwijderen: {old.name}")
        old.unlink()

    print("Downloaden...")
    r = requests.get(href, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    if not r.content.startswith(b"%PDF"):
        print("FOUT: gedownloade inhoud is geen PDF.", file=sys.stderr)
        write_github_output("new_version", "false")
        return 1

    pdf_path = INPUT_DIR / new_name
    pdf_path.write_bytes(r.content)
    print(f"Opgeslagen   : {pdf_path}")

    LAST_NAME_FILE.write_text(new_name, encoding="utf-8")

    write_github_output("new_version", "true")
    write_github_output("pdf_path", str(pdf_path))
    write_github_output("version", new_version)
    write_github_output("previous_version", last_name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
