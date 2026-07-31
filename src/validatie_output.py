"""
validatie_output.py

Derde stap in de VLIF4-pipeline (na layout_engine.py en build_excel.py).
Controleert het gegenereerde Excel-bestand op:

  1. Deterministische fouten (snel, 100% betrouwbaar):
     - rijtelling: aantal records in de brontekst (via parse_pdf() uit
       layout_engine.py) vs. aantal rijen in Excel
     - anomale cellen: cellen met newlines / abnormale lengte (teken van
       samengevoegde rijen)
     - kolomtype-consistentie: een numerieke kolom die plots tekst bevat
       (optioneel, alleen als --numeric-columns is opgegeven)

  2. Semantische controle via een lokaal Phi-model (llama-cpp-python):
     - voor elke verdachte rij wordt het model gevraagd of de rij
       overeenkomt met precies één record uit de brontekst.

Resultaat: validation_report.xlsx met tabbladen "Summary" en "Issues".
Exit code 0 = geen problemen, 1 = problemen gevonden (zodat de GitHub
Actions workflow hierop kan reageren, bv. in de e-mail vermelden).

Verwacht als input:
  --excel       pad naar het gegenereerde Excel-bestand (build_excel.py output)
  --source-pdf  pad naar de originele PDF
  --model       pad naar een lokaal Phi GGUF-model (bv. Phi-3.5-mini-instruct-Q4_K_M.gguf)
  --numeric-columns  kommagescheiden kolomindices die numeriek horen te zijn (optioneel)
  --report-out  pad voor het validatierapport (default validation_report.xlsx)
"""

import argparse
import json
import sys
from pathlib import Path

import fitz  # PyMuPDF, al gebruikt in layout_engine.py
from openpyxl import load_workbook, Workbook

# layout_engine.py staat in de root, validatie_output.py in src/ —
# dit voegt de root toe aan het importpad zodat parse_pdf() gevonden wordt.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from layout_engine import parse_pdf


# ---------------------------------------------------------------------------
# Laag 1: deterministische checks
# ---------------------------------------------------------------------------

def extract_expected_record_count(pdf_path: str) -> int:
    """
    Gebruikt dezelfde parse_pdf() als layout_engine.py, zodat de rijtelling
    exact aansluit bij wat de layout-engine als geldig record beschouwt
    (enkel records met een geldige code, zoals ook gefilterd in _save()).
    """
    records = parse_pdf(pdf_path)
    return len([r for r in records if r.get("code", "") != ""])


def find_anomalous_cells(ws, max_len: int = 200):
    """
    Zoekt cellen die wijzen op samengevoegde rijen:
      - bevat een newline-teken (meerdere records in één cel geplakt)
      - abnormaal lange tekst t.o.v. de rest van de kolom
    Retourneert lijst van (rij, kolom, waarde, reden).
    """
    issues = []
    col_lengths: dict[int, list[int]] = {}

    for row in ws.iter_rows(min_row=2):  # rij 1 = header
        for cell in row:
            if cell.value is None:
                continue
            val = str(cell.value)
            col_lengths.setdefault(cell.column, []).append(len(val))
            if "\n" in val:
                issues.append((cell.row, cell.column, val, "bevat newline (mogelijk samengevoegde rij)"))

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value is None:
                continue
            lengths = col_lengths.get(cell.column, [])
            if not lengths:
                continue
            lengths_sorted = sorted(lengths)
            median = lengths_sorted[len(lengths_sorted) // 2]
            val_len = len(str(cell.value))
            if median > 0 and val_len > max(max_len, median * 3):
                issues.append((cell.row, cell.column, str(cell.value)[:80] + "...",
                                f"abnormale lengte ({val_len} tekens, mediaan kolom = {median})"))

    return issues


def
