"""
validate_output.py

Derde stap in de VLIF4-pipeline (na layout_engine.py en build_excel.py).
Controleert het gegenereerde Excel-bestand op:

  1. Deterministische fouten (snel, 100% betrouwbaar):
     - rijtelling: aantal records in de brontekst vs. aantal rijen in Excel
     - anomale cellen: cellen met newlines / abnormale lengte (teken van
       samengevoegde rijen)
     - kolomtype-consistentie: een numerieke kolom die plots tekst bevat

  2. Semantische controle via een lokaal Phi-model (llama-cpp-python):
     - voor elke verdachte rij (of een steekproef) wordt het model gevraagd
       of de rij overeenkomt met precies één record uit de brontekst.

Resultaat: validation_report.xlsx met tabbladen "Summary" en "Issues".
Exit code 0 = geen problemen, 1 = problemen gevonden (zodat de GitHub
Actions workflow hierop kan reageren, bv. in de e-mail vermelden).

Verwacht als input:
  --excel       pad naar het gegenereerde Excel-bestand (build_excel.py output)
  --source-pdf  pad naar de originele PDF
  --model       pad naar een lokaal Phi GGUF-model (bv. Phi-3.5-mini-instruct-Q4_K_M.gguf)
  --sample-rate fractie van de verdachte rijen die door het Phi-model
                gecontroleerd wordt (default 1.0 = alle verdachte rijen)
"""

import argparse
import json
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF, al gebruikt in layout_engine.py
from openpyxl import load_workbook, Workbook


# ---------------------------------------------------------------------------
# Laag 1: deterministische checks
# ---------------------------------------------------------------------------

def extract_expected_record_count(pdf_path: str) -> int:
    """
    Telt het verwachte aantal records in de brontekst.

    LET OP: dit is een placeholder-heuristiek (telt regels die op een
    subsidiecode-patroon lijken, bv. een reeks cijfers/letters aan het
    begin van een regel). Vervang dit door de echte detectielogica uit
    layout_engine.py zodra je dat script hierheen kopieert of importeert,
    zodat de rijtelling exact overeenkomt met wat de layout-engine als
    "record" beschouwt.
    """
    doc = fitz.open(pdf_path)
    count = 0
    code_pattern = re.compile(r"^\s*[A-Z0-9]{2,10}[-/]?\d*\s")
    for page in doc:
        for line in page.get_text().splitlines():
            if code_pattern.match(line):
                count += 1
    doc.close()
    return count


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

    # kolommen met een gemiddelde lengte > 3x de mediaan zijn verdacht
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


def check_column_type_consistency(ws, numeric_columns: list[int]):
    """
    Controleert of kolommen die numeriek horen te zijn, ook effectief
    numerieke waarden bevatten. Pas `numeric_columns` aan naar de
    kolomindices die in jouw build_excel.py als bedrag/getal gelden.
    """
    issues = []
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in numeric_columns and cell.value is not None:
                val = str(cell.value).replace(",", ".").replace("€", "").strip()
                try:
                    float(val)
                except ValueError:
                    issues.append((cell.row, cell.column, cell.value,
                                    "verwacht numeriek, maar bevat tekst (mogelijk kolomverschuiving)"))
    return issues


# ---------------------------------------------------------------------------
# Laag 2: Phi-model semantische check
# ---------------------------------------------------------------------------

def load_phi_model(model_path: str):
    from llama_cpp import Llama
    return Llama(model_path=model_path, n_ctx=4096, n_threads=4, verbose=False)


PROMPT_TEMPLATE = """Je bent een data-validatie-assistent. Hieronder staat een fragment
brontekst uit een PDF, en een rij uit een Excel-bestand die daaruit werd
geëxtraheerd.

BRONTEKST:
{source_text}

EXCEL-RIJ:
{row_data}

Vraag: komt deze Excel-rij overeen met PRECIES ÉÉN record uit de brontekst,
of lijken er meerdere records samengevoegd, of kolommen verschoven?

Antwoord ALLEEN met geldige JSON, geen andere tekst:
{{"issue": true/false, "reason": "korte uitleg in het Nederlands"}}
"""


def phi_check_row(llm, source_text: str, row_data: dict) -> dict:
    prompt = PROMPT_TEMPLATE.format(source_text=source_text[:1500], row_data=json.dumps(row_data, ensure_ascii=False))
    output = llm(prompt, max_tokens=200, temperature=0.0, stop=["\n\n"])
    text = output["choices"][0]["text"].strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"issue": True, "reason": f"kon modeloutput niet parsen: {text[:200]}"}


# ---------------------------------------------------------------------------
# Rapportage
# ---------------------------------------------------------------------------

def write_report(path: str, summary: dict, issues: list[tuple]):
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.append(["Check", "Resultaat"])
    for key, val in summary.items():
        ws_summary.append([key, val])

    ws_issues = wb.create_sheet("Issues")
    ws_issues.append(["Rij", "Kolom", "Waarde", "Reden"])
    for row, col, val, reason in issues:
        ws_issues.append([row, col, val, reason])

    wb.save(path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", required=True)
    parser.add_argument("--source-pdf", required=True)
    parser.add_argument("--model", required=False, help="pad naar lokaal Phi GGUF-model")
    parser.add_argument("--numeric-columns", default="", help="kommagescheiden kolomindices, bv. 3,5,7")
    parser.add_argument("--report-out", default="validation_report.xlsx")
    args = parser.parse_args()

    wb = load_workbook(args.excel)
    all_issues = []
    summary = {}

    # Laag 1
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        anomalies = find_anomalous_cells(ws)
        numeric_cols = [int(x) for x in args.numeric_columns.split(",") if x.strip()]
        type_issues = check_column_type_consistency(ws, numeric_cols) if numeric_cols else []
        all_issues.extend([(r, c, v, f"[{sheet_name}] {reason}") for r, c, v, reason in anomalies])
        all_issues.extend([(r, c, v, f"[{sheet_name}] {reason}") for r, c, v, reason in type_issues])

    expected = extract_expected_record_count(args.source_pdf)
    actual = sum(ws.max_row - 1 for ws in wb.worksheets)  # -1 voor header per tab
    summary["Verwacht aantal records (PDF)"] = expected
    summary["Aantal rijen in Excel (alle tabs)"] = actual
    summary["Verschil"] = actual - expected
    if actual != expected:
        all_issues.append((0, 0, "", f"Rijtelling klopt niet: {actual} vs verwacht {expected}"))

    # Laag 2 (optioneel, alleen als een modelpad is meegegeven)
    if args.model and all_issues:
        llm = load_phi_model(args.model)
        doc = fitz.open(args.source_pdf)
        full_text = "\n".join(page.get_text() for page in doc)
        doc.close()
        for i, (row, col, val, reason) in enumerate(all_issues):
            if row == 0:
                continue
            result = phi_check_row(llm, full_text, {"rij": row, "kolom": col, "waarde": val})
            all_issues[i] = (row, col, val, f"{reason} | Phi-check: {result.get('reason', '')}")

    summary["Totaal aantal gevonden issues"] = len(all_issues)
    write_report(args.report_out, summary, all_issues)

    print(f"Validatie klaar. {len(all_issues)} issue(s) gevonden. Rapport: {args.report_out}")
    sys.exit(1 if all_issues else 0)


if __name__ == "__main__":
    main()
