import argparse
import re
from pathlib import Path
import fitz  # PyMuPDF
import pandas as pd


def extract_pdf_data(pdf_path):
    print(f"📖 PDF Analyseren met PyMuPDF: {pdf_path}")
    doc = fitz.open(pdf_path)
    all_rows = []

    # Regex om te kijken of een regel begint met een geldige code (bijv. 2.1, 2.12, 3.1.2, 4.1)
    code_pattern = re.compile(r"^(\d+\.\d+(?:\.\d+)?|\d+)\s")

    current_subfilter = ""

    for page_num, page in enumerate(doc, start=1):
        text_instances = page.get_text("blocks")
        # Sorteer blokken van boven naar beneden, links naar rechts
        text_instances.sort(key=lambda x: (x[1], x[0]))

        for block in text_instances:
            line = block[4].strip()
            if not line:
                continue

            # Splits de bloktekst op in afzonderlijke regels indien aanwezig
            lines = [l.strip() for l in line.split("\n") if l.strip()]

            for l in lines:
                # 1. Detecteer subfilters (hoofdstuktitels zonder code aan het begin)
                if (
                    ("HOOFDSTUK" in l.upper() or "STAL" in l.upper() or "AERi" in l)
                    if not code_pattern.match(l)
                    else False
                ):
                    current_subfilter = l
                    continue

                # 2. Match data rijen die beginnen met een code
                match = code_pattern.match(l)
                if match:
                    code = match.group(1)
                    rest_of_text = l[len(code) :].strip()

                    # Probeer eenheid/kenmerken te splitsen aan het einde van de regel (bijv. % of EUR)
                    # Dit is een basis fallback-split; openpyxl / pandas regelt de rest in build_excel
                    parts = rest_of_text.split("  ")
                    parts = [p.strip() for p in parts if p.strip()]

                    omschrijving = parts[0] if len(parts) > 0 else rest_of_text
                    extra_info = parts[1] if len(parts) > 1 else ""
                    kenmerk = parts[2] if len(parts) > 2 else ""
                    eenheid = parts[3] if len(parts) > 3 else ""

                    all_rows.append(
                        {
                            "subfilter": current_subfilter,
                            "Code": code,
                            "Omschrijving": omschrijving,
                            "Kolom_5": extra_info,  # extra_info mapping
                            "Kolom_4": kenmerk,  # kenmerk mapping
                            "Kolom_3": eenheid,  # eenheid mapping
                        }
                    )

    print(f"✨ PDF succesvol uitgelezen. {len(all_rows)} datarijen gevonden.")
    return pd.DataFrame(all_rows)


def build_excel(input_path, output_path):
    # Extraheer data uit de PDF
    df = extract_pdf_data(input_path)

    # Als de dataframe leeg is, maak dan een lege template met de juiste kolommen
    if df.empty:
        df = pd.DataFrame(
            columns=["subfilter", "Code", "Omschrijving", "Kolom_5", "Kolom_4", "Kolom_3"]
        )

    # Schrijf weg naar de rauwe Excel-bestandsindeling
    df.to_excel(output_path, index=False)
    print(f"💾 Ruwe data opgeslagen naar: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Layout Engine voor VLIF PDF naar Excel")
    parser.add_argument("--input", required=True, help="Pad naar het input PDF-bestand")
    parser.add_argument("--output", required=True, help="Pad naar het output XLSX-bestand")
    args = parser.parse_args()

    build_excel(Path(args.input), Path(args.output))


if __name__ == "__main__":
    main()
