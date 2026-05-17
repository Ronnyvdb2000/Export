import argparse
import os
import re
import fitz  # PyMuPDF
import pandas as pd


def extract_vlif_table(pdf_path):
    print(f"📄 Analyseren van PDF: {pdf_path}")
    doc = fitz.open(pdf_path)
    all_records = []

    # VLIF tabellen hebben vaak kolommen zoals: Code, Omschrijving, Selectie/Voorwaarden, Punten, etc.
    # We zoeken flexibel naar regels die beginnen met een code (bijv. 3 of 4 cijfers, of letters/cijfers)
    code_pattern = re.compile(
        r"^\d{3,4}/?.*"
    )  # Matches bijv. '1041', '2101/1', etc.

    for page_num, page in enumerate(doc, start=1):
        # Haal de tekst op als gestructureerde woorden met coördinaten
        words = page.get_text("words")
        if not words:
            continue

        # Groepeer woorden per regel op basis van hun Y-coördinaat (met een marge van 3 pixels)
        lines_dict = {}
        for w in words:
            x0, y0, x1, y1, text, block_no, line_no, word_no = w
            # Rond de Y-coördinaat af om kleine hoogteverschillen op te vangen
            y_key = round(y0 / 3) * 3
            if y_key not in lines_dict:
                lines_dict[y_key] = []
            lines_dict[y_key].append(w)

        # Sorteer de regels van boven naar beneden
        for y_key in sorted(lines_dict.keys()):
            # Sorteer woorden binnen de regel van links naar rechts
            sorted_words = sorted(lines_dict[y_key], key=lambda x: x[0])

            # Combineer woorden tot kolommen op basis van hun X-afstand
            line_text_elements = []
            current_element = sorted_words[0][4]
            last_x1 = sorted_words[0][2]

            for w in sorted_words[1:]:
                x0, _, x1, _, text = w[0], w[1], w[2], w[3], w[4]
                # Als de afstand tot het vorige woord groter is dan 15 pixels, zien we het als een nieuwe kolom
                if x0 - last_x1 > 15:
                    line_text_elements.append(current_element)
                    current_element = text
                else:
                    # Anders hoort het bij hetzelfde tekstblok (bijv. een zin met spaties)
                    current_element += " " + text
                last_x1 = x1
            line_text_elements.append(current_element)

            # Filterlogica: Een geldige VLIF-regel begint meestal met een specifieke Code
            if line_text_elements and code_pattern.match(line_text_elements[0]):
                # Zorg dat we lege kolommen opvullen tot een vaste structuur (bijv. max 5 kolommen)
                while len(line_text_elements) < 4:
                    line_text_elements.append("")
                all_records.append(line_text_elements[:5])

    print(f"    records gevonden met nieuwe engine: {len(all_records)}")
    return all_records


def main():
    parser = argparse.ArgumentParser(
        description="Flexibele PDF naar RAW XLSX VLIF Converter"
    )
    parser.add_argument("--input-pdf", required=True, help="Pad naar de input PDF")
    parser.add_argument(
        "--output-xlsx", required=True, help="Pad naar de output RAW XLSX"
    )
    args = parser.parse_args()

    records = extract_vlif_table(args.input_pdf)

    if not records:
        print("⚠️ Geen records gevonden. We maken een lege nood-tabel aan.")
        # Maak een dummy record aan zodat de workflow in ieder geval niet crasht en je kunt debuggen
        df = pd.DataFrame(
            [["GEEN_DATA", "Tabelstructuur niet herkend", "", ""]],
            columns=["Code", "Omschrijving", "Info_1", "Info_2"],
        )
    else:
        # Dynamische kolommen toewijzen op basis van het aantal gevonden elementen
        col_names = [f"Kolom_{i+1}" for i in range(len(records[0]))]
        # Geef de eerste twee kolommen logische namen die vaak bij VLIF horen
        col_names[0] = "Code"
        if len(col_names) > 1:
            col_names[1] = "Omschrijving"

        df = pd.DataFrame(records, columns=col_names)

    # Zorg dat de output map bestaat
    os.makedirs(os.path.dirname(args.output_xlsx), exist_ok=True)

    # Sla op naar Excel
    df.to_excel(args.output_xlsx, index=False)
    print(f"✅ Bestand succesvol opgeslagen: {args.output_xlsx}")


if __name__ == "__main__":
    main()
