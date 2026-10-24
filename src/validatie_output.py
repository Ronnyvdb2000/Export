"""
validatie_output.py

Derde stap in de VLIF4-pipeline (na layout_engine.py en build_excel.py).
Controleert het gegenereerde Excel-bestand op:

  1. Deterministische fouten (snel, 100% betrouwbaar):
     - rijtelling: aantal records in de brontekst (via parse_pdf() uit
       layout_engine.py) vs. aantal rijen in Excel — met correctie voor
       de titelblok-structuur (4 titelregels + 1 headerrij per tabblad)
       en met het Inhoudsopgave-tabblad uitgesloten
     - anomale cellen: cellen met newlines / abnormale lengte (teken van
       samengevoegde rijen)
     - kolomtype-consistentie: een numerieke kolom die plots tekst bevat
       (optioneel, alleen als --numeric-columns is opgegeven)

  2. Semantische controle via een lokaal Phi-model (llama-cpp-python):
     - voor elke verdachte rij wordt gezocht naar het relevante fragment
       in de brontekst (rond de plek waar de celwaarde letterlijk
       voorkomt), en dat fragment — niet zomaar het begin van de PDF —
       wordt aan het model gegeven om te beoordelen.

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
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF, al gebruikt in layout_engine.py
from openpyxl import load_workbook, Workbook

# layout_engine.py staat in de root, validatie_output.py in src/ —
# dit voegt de root toe aan het importpad zodat parse_pdf() gevonden wordt.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from layout_engine import parse_pdf

# Structuur van elk data-tabblad in build_excel.py: 4 titelregels + 1 headerrij
# vóór de eerste datarij (zie _title_rows() + header-append in _data_sheet()).
NON_DATA_ROWS_PER_SHEET = 5
TOC_SHEET_NAME = "Inhoudsopgave"

# NIEUW: een subfilter-label (kolom 1) is kort ("2.12.10 Zaai/ planttechniek", max. ~40
# tekens in blok 03-2026). Veel langer wijst op vervuiling met tekst uit de omgeving.
SUBFILTER_MAX_LEN = 80


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
    
    # NIEUW: layout_engine.py geeft nu een pandas DataFrame terug
    if hasattr(records, "columns"):
        col = "code (%)" if "code (%)" in records.columns else "code"
        if col in records.columns:
            return len(records[records[col].astype(str).str.strip() != ""])
        return len(records)
        
    # Fallback voor de oude dict-lijst (voor het geval dat)
    return len([r for r in records if isinstance(r, dict) and r.get("code", "") != ""])


def count_actual_rows(wb) -> int:
    """
    Telt de echte datarijen over alle data-tabbladen, met correctie voor:
      - het Inhoudsopgave-tabblad (bevat geen PDF-records, wordt overgeslagen)
      - de titelblok-structuur (4 titelregels + 1 headerrij per tabblad,
        niet slechts 1 header zoals eerder aangenomen)
    """
    total = 0
    for ws in wb.worksheets:
        if ws.title == TOC_SHEET_NAME:
            continue
        total += max(ws.max_row - NON_DATA_ROWS_PER_SHEET, 0)
    return total


def _norm(text: str) -> str:
    """Witruimte (spaties, tabs, regeleinden) samenvoegen tot één spatie."""
    return re.sub(r"\s+", " ", str(text)).strip()


def is_long_text_confirmed(text: str, norm_source: str,
                           ruimte: float = 0.25, anker: int = 80) -> bool:
    """
    NIEUW. Controleert of een lange celtekst als één passage in de PDF voorkomt.

    Een lange cel is op zich geen fout: sommige notities in de codelijst zijn
    echt 1700+ tekens (bv. "Niet aan te vragen bij afbraak van een gebouw ...",
    18x aanwezig). Verdacht is een lange cel die NIET als één passage in de
    bron staat, want dat is wat er gebeurt bij samengevoegde rijen of tekst
    die uit de omgeving is meegeplakt.

    Bevestigd als:
      - de volledige tekst (witruimte genormaliseerd) letterlijk in de PDF staat, of
      - de eerste en de laatste `anker` tekens allebei in de PDF staan en het
        stuk PDF ertussen hooguit `ruimte` (25%) langer is dan de cel. Dat vangt
        paginakoppen/-voeten op die midden in een lange notitie in de PDF-tekst staan.
    """
    t = _norm(text)
    if not t or not norm_source:
        return False
    if t in norm_source:
        return True
    if len(t) < 2 * anker:
        return False
    start, einde = t[:anker], t[-anker:]
    i = norm_source.find(start)
    while i != -1:
        j = norm_source.find(einde, i)
        if j != -1 and (j + anker - i) <= len(t) * (1 + ruimte):
            return True
        i = norm_source.find(start, i + 1)
    return False


def find_anomalous_cells(ws, max_len: int = 200, source_text=None):
    """
    Zoekt cellen die wijzen op samengevoegde rijen:
      - bevat een newline-teken (meerdere records in één cel geplakt)
      - abnormaal lange tekst t.o.v. de rest van de kolom
      - subfilter-label (kolom 1) ongewoon lang (NIEUW)
    Met `source_text` (genormaliseerde PDF-tekst) worden lange cellen die
    letterlijk als één passage in de PDF staan niet meer gemeld (NIEUW);
    zonder `source_text` is het gedrag zoals voorheen.
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
            # NIEUW: subfilter-label (kolom 1, datarijen) mag niet lang zijn. Rij < 6 zijn
            # de titelrijen/kolomkop van build_excel.py en worden hier overgeslagen.
            if (cell.column == 1 and cell.row > NON_DATA_ROWS_PER_SHEET
                    and val_len > SUBFILTER_MAX_LEN):
                issues.append((cell.row, cell.column, str(cell.value)[:80] + "...",
                                f"subfilter-label ongewoon lang ({val_len} tekens, "
                                f"verwacht hooguit {SUBFILTER_MAX_LEN})"))
                continue
            if median > 0 and val_len > max(max_len, median * 3):
                # NIEUW: lange cel die als één passage in de PDF staat is geen fout.
                # Oorspronkelijk werd elke lange cel gemeld (61 meldingen op blok 03-2026,
                # waarvan 60 gewone lange notities).
                if source_text and is_long_text_confirmed(str(cell.value), source_text):
                    continue
                issues.append((cell.row, cell.column, str(cell.value)[:80] + "...",
                                f"abnormale lengte ({val_len} tekens, mediaan kolom = {median})"))

    return issues


def check_column_type_consistency(ws, numeric_columns: list[int]):
    """
    Controleert of kolommen die numeriek horen te zijn, ook effectief
    numerieke waarden bevatten.
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


def find_relevant_context(norm_source: str, cell_value: str, window: int = 900) -> str:
    """
    NIEUW. Zoekt in de genormaliseerde PDF-tekst (witruimte samengevoegd, zie _norm)
    het fragment waar het begin van de celwaarde voorkomt, en geeft een venster rond
    die plek terug. De waarde wordt eerst genormaliseerd en een afsluitende "..."
    wordt verwijderd.

    Geeft een LEGE string terug als niets gevonden wordt. De aanroeper slaat de
    Phi-check dan over, in plaats van het model de titelpagina te laten beoordelen.
    """
    needle = _norm(cell_value)
    if needle.endswith("..."):
        needle = needle[:-3].rstrip()
    if not needle:
        return ""

    idx = -1
    for lengte in (120, 60, 40):
        idx = norm_source.find(needle[:lengte])
        if idx != -1:
            break
    if idx == -1:
        return ""

    start = max(0, idx - window // 2)
    end = min(len(norm_source), idx + window // 2)
    return norm_source[start:end]


PROMPT_TEMPLATE = (
    "Je bent een data-validatie-assistent. Hieronder staat een fragment "
    "brontekst uit een PDF, en een rij uit een Excel-bestand die daaruit werd "
    "geëxtraheerd.\n\n"
    "BRONTEKST:\n{source_text}\n\n"
    "EXCEL-RIJ:\n{row_data}\n\n"
    "Vraag: komt deze Excel-rij overeen met PRECIES EEN record uit de brontekst, "
    "of lijken er meerdere records samengevoegd, of kolommen verschoven?\n\n"
    "Antwoord ALLEEN met geldige JSON, geen andere tekst:\n"
    '{{"issue": true/false, "reason": "uitleg in het Nederlands, maximaal 25 woorden"}}\n'
)


def phi_check_row(llm, source_text: str, row_data: dict) -> dict:
    prompt = PROMPT_TEMPLATE.format(source_text=source_text[:1500], row_data=json.dumps(row_data, ensure_ascii=False))
    output = llm(prompt, max_tokens=300, temperature=0.0, stop=["\n\n"])
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

    # NIEUW: PDF-tekst één keer inlezen (laag 1 en laag 2 gebruiken dezelfde tekst)
    doc = fitz.open(args.source_pdf)
    full_text = "\n".join(page.get_text() for page in doc)
    doc.close()
    norm_source = _norm(full_text)

    # Laag 1
    for sheet_name in wb.sheetnames:
        if sheet_name == TOC_SHEET_NAME:
            continue
        ws = wb[sheet_name]
        anomalies = find_anomalous_cells(ws, source_text=norm_source)
        numeric_cols = [int(x) for x in args.numeric_columns.split(",") if x.strip()]
        type_issues = check_column_type_consistency(ws, numeric_cols) if numeric_cols else []
        all_issues.extend([(r, c, v, f"[{sheet_name}] {reason}") for r, c, v, reason in anomalies])
        all_issues.extend([(r, c, v, f"[{sheet_name}] {reason}") for r, c, v, reason in type_issues])

    expected = extract_expected_record_count(args.source_pdf)
    actual = count_actual_rows(wb)
    summary["Verwacht aantal records (PDF)"] = expected
    summary["Aantal rijen in Excel (gecorrigeerd voor titelblok + TOC)"] = actual
    summary["Verschil"] = actual - expected
    if actual != expected:
        all_issues.append((0, 0, "", f"Rijtelling klopt niet: {actual} vs verwacht {expected}"))

    # Laag 2 (optioneel, alleen als een modelpad is meegegeven)
    if args.model and all_issues:
        llm = load_phi_model(args.model)
        for i, (row, col, val, reason) in enumerate(all_issues):
            if row == 0:
                continue
            context = find_relevant_context(norm_source, val)
            if not context:
                all_issues[i] = (row, col, val,
                                 f"{reason} | Phi-check overgeslagen: fragment niet gevonden in de PDF")
                continue
            result = phi_check_row(llm, context, {"rij": row, "kolom": col, "waarde": val})
            all_issues[i] = (row, col, val, f"{reason} | Phi-check: {result.get('reason', '')}")

    summary["Totaal aantal gevonden issues"] = len(all_issues)
    write_report(args.report_out, summary, all_issues)

    print(f"Validatie klaar. {len(all_issues)} issue(s) gevonden. Rapport: {args.report_out}")
    sys.exit(1 if all_issues else 0)


if __name__ == "__main__":
    main()
