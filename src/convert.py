import os
import argparse
import pdfplumber
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


START_PAGE = 9
START_TEXT = "2. Productieve investeringen"


def clean_row(row):
    cleaned = []

    for cell in row:
        if cell is None:
            cleaned.append("")
        else:
            value = str(cell).replace("\n", " ").strip()
            cleaned.append(value)

    return cleaned

def extract_tables(pdf_path):
    all_rows = []
    inhoud = []

    started = False

    with pdfplumber.open(pdf_path) as pdf:

        total_pages = len(pdf.pages)


        for i, page in enumerate(pdf.pages, start=1):

            print(f"📄 Verwerken pagina {i}")

            text = page.extract_text() or ""

            inhoud.append({
                "Pagina": i,
                "Tekst": text[:150]
            })

            if i < START_PAGE:
                continue

            if not started:
                if START_TEXT in text:
                    started = True
                else:
                    continue

            tables = page.extract_tables()

            found_table = False

            if tables:
                for table in tables:
                    if not table:
                        continue

                    for row in table:
                        cleaned = clean_row(row)

                        if any(cell != "" for cell in cleaned):
                            all_rows.append(cleaned)
                            found_table = True

            # fallback indien geen tabellen gevonden
            if not found_table:
                words = page.extract_words()

                for word in words:
                    txt = word.get("text", "").strip()

                    if txt:
                        all_rows.append([txt])

    if not all_rows:
        raise Exception(
            f"❌ Geen data gevonden in {os.path.basename(pdf_path)}"
        )

    return inhoud, all_rows


def save_excel(output_path, inhoud, rows):

    wb = Workbook()

    # TAB 1 = inhoudstafel
    ws1 = wb.active
    ws1.title = "Inhoudstafel"

    inhoud_df = pd.DataFrame(inhoud)

    for r in dataframe_to_rows(inhoud_df, index=False, header=True):
        ws1.append(r)

    # TAB 2 = alle tabellen
    ws2 = wb.create_sheet(title="Tabellen")

    for row in rows:
        ws2.append(row)

    wb.save(output_path)


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output-dir", required=True)

    args = parser.parse_args()

    pdf_files = [
        f for f in os.listdir(args.input_dir)
        if f.lower().endswith(".pdf")
    ]

    print(f"📁 Gevonden PDF bestanden: {pdf_files}")

    if not pdf_files:
        raise Exception("❌ Geen PDF bestanden gevonden")

    for pdf_file in pdf_files:

        pdf_path = os.path.join(args.input_dir, pdf_file)

        print(f"\n🚀 Start conversie: {pdf_file}")

        inhoud, rows = extract_tables(pdf_path)

        output_name = os.path.splitext(pdf_file)[0] + ".xlsx"

        output_path = os.path.join(args.output_dir, output_name)

        save_excel(output_path, inhoud, rows)

        print(f"✅ Excel opgeslagen: {output_path}")


if __name__ == "__main__":
    main()
