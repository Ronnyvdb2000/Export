import argparse
import os
import re
import fitz  # PyMuPDF
import pandas as pd


def extract_vlif_table(pdf_path):
    print(f"📄 Analyseren van PDF: {pdf_path}")
    doc = fitz.open(pdf_path)
    all_records = []

    # Matcht codes zoals 1041, 2101/1, 4.1, etc.
    code_pattern = re.compile(r"^\d{1,4}/?.*")
    current_subfilter = "Algemeen"

    for page_num, page in enumerate(doc, start=1):
        words = page.get_text("words")
        if not words:
            continue

        # 1. Groepeer woorden per regel op basis van afgeronde Y-coördinaat (marge van 3pt)
        lines_dict = {}
        for w in words:
            x0, y0, x1, y1, text, block_no, line_no, word_no = w
            # Sla paginanummers of headers over
            if y0 < 50 or y0 > 800:
                continue
            
            y_key = round(y0 / 3) * 3
            if y_key not in lines_dict:
                lines_dict[y_key] = []
            lines_dict[y_key].append(w)

        # 2. Verwerk regels van boven naar beneden
        for y_key in sorted(lines_dict.keys()):
            sorted_words = sorted(lines_dict[y_key], key=lambda x: x[0])
            
            # Bouw de tekstvlakken op basis van de X-afstand
            line_text_elements = []
            current_element = sorted_words[0][4]
            last_x1 = sorted_words[0][2]
            start_x = sorted_words[0][0]

            # We houden ook de start-X positie van elk tekstblok bij
            element_positions = [(start_x, current_element)]

            for w in sorted_words[1:]:
                x0, _, x1, _, text = w[0], w[1], w[2], w[3], w[4]
                # Als het gat groter is dan 20pt, is het een nieuwe kolom
                if x0 - last_x1 > 20:
                    element_positions.append((x0, text))
                else:
                    # Voeg toe aan lopende tekstblok
                    element_positions[-1] = (element_positions[-1][0], element_positions[-1][1] + " " + text)
                last_x1 = x1

            # Sorteer elementen naar de 5 doelkolommen op basis van absolute X-positie (vrijblijvende inschatting)
            # Kolom 1 (Code): X < 100
            # Kolom 2 (Omschrijving): 100 <= X < 350
            # Kolom 3 (Eenheid): 350 <= X < 450
            # Kolom 4 (Bedrag/Kenmerk): 450 <= X < 550
            # Kolom 5 (Extra info): X >= 550
            
            row_data = ["", "", "", "", ""]
            first_text = element_positions[0][1].strip()

            # Controleer of dit een subfilter/tussentitel is (Geen code, wel letters, langwerpig)
            if not code_pattern.match(first_text) and len(first_text) > 4 and len(element_positions) == 1:
                # Sla opmerkingen zoals "Geldig vanaf" of headers over
                if "vlif" not in first_text.lower() and "pagina" not in first_text.lower():
                    current_subfilter = first_text
                    continue

            # Als het een reguliere datarij is (begint met een code)
            if code_pattern.match(first_text):
                for start_x, text in element_positions:
                    text = text.strip()
                    if start_x < 90:
                        row_data[0] = text  # Code
                    elif start_x >= 90 and start_x < 320:
                        row_data[1] = text  # Omschrijving
                    elif start_x >= 320 and start_x < 440:
                        row_data[2] = text  # Kolom 3 (Eenheid)
                    elif start_x >= 440 and start_x < 540:
                        row_data[3] = text  # Kolom 4 (Bedrag)
                    else:
                        row_data[4] = text  # Kolom 5 (Extra)

                # Voeg het actieve subfilter toe als onzichtbare 6e kolom voor de Excel builder
                row_data.append(current_subfilter)
                all_records.append(row_data)

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
        df = pd.DataFrame(
            [["GEEN_DATA", "Tabelstructuur niet herkend", "", "", "", "Algemeen"]],
            columns=["Code", "Omschrijving", "Kolom_3", "Kolom_4", "Kolom_5", "subfilter"],
        )
    else:
        df = pd.DataFrame(
            records, 
            columns=["Code", "Omschrijving", "Kolom_3", "Kolom_4", "Kolom_5", "subfilter"]
        )

    os.makedirs(os.path.dirname(args.output_xlsx), exist_ok=True)
    df.to_excel(args.output_xlsx, index=False)
    print(f"✅ Bestand succesvol opgeslagen: {args.output_xlsx}")


if __name__ == "__main__":
    main()
