import pandas as pd
import pdfplumber
import re
from pathlib import Path
import argparse

START_MARKER = "2. Productieve investeringen"

IGNORE_PATTERNS = [
    "VLAANDEREN.BE",
    "AGENTSCHAP",
    "LANDBOUW & ZEEVISSERIJ",
    "Productieve investeringssteun",
]

HEADER_PATTERNS = [
    "SUBFILTER",
    "OMSCHRIJVING",
    "KENMERK",
    "EENHEID",
]

def clean_line(line):
    line = line.strip()

    # verwijder dubbele spaties
    line = re.sub(r"\s+", " ", line)

    return line


def should_ignore(line):

    if not line:
        return True

    # paginanummers
    if re.match(r"^\d+$", line):
        return True

    # vaste headers
    for pattern in IGNORE_PATTERNS:
        if pattern.lower() in line.lower():
            return True

    return False


def looks_like_header(line):

    matches = 0

    for pattern in HEADER_PATTERNS:
        if pattern.lower() in line.lower():
            matches += 1

    return matches >= 2


def process_pdf(pdf_path):

    toc_rows = []
    data_rows = []

    found_data_start = False
    header_added = False

    with pdfplumber.open(pdf_path) as pdf:

        for page_num, page in enumerate(pdf.pages, start=1):

            print(f"📄 Verwerken pagina {page_num}")

            text = page.extract_text()

            if not text:
                print(f"⚠️ Geen tekst op pagina {page_num}")
                continue

            lines = text.split("\n")

            for raw_line in lines:

                line = clean_line(raw_line)

                if should_ignore(line):
                    continue

                # detecteer start data
                if START_MARKER.lower() in line.lower():
                    found_data_start = True
                    print(f"✅ Start data gevonden op pagina {page_num}")
                    continue

                # inhoudstafel
                if not found_data_start:
                    toc_rows.append([line])
                    continue

                # headers slechts 1x
                if looks_like_header(line):

                    if header_added:
                        continue

                    header_added = True

                # multiline cleanup
                line = line.replace("\n", " ")

                data_rows.append([line])

    toc_df = pd.DataFrame(toc_rows, columns=["INHOUD"])

    data_df = pd.DataFrame(data_rows, columns=["DATA"])

    return toc_df, data_df


def export_excel(toc_df, data_df, output_path):

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        toc_df.to_excel(
            writer,
            sheet_name="INHOUDSTAFEL",
            index=False
        )

        data_df.to_excel(
            writer,
            sheet_name="DATA",
            index=False
        )


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument("--input-dir", default="input")
    parser.add_argument("--output-dir", default="output")

    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)

    output_dir.mkdir(exist_ok=True)

    pdf_files = list(input_dir.glob("*.pdf"))

    print(f"📁 Gevonden PDF bestanden: {[f.name for f in pdf_files]}")

    if not pdf_files:
        raise Exception("❌ Geen PDF bestanden gevonden")

    for pdf_path in pdf_files:

        print(f"\n🚀 Start conversie: {pdf_path.name}")

        toc_df, data_df = process_pdf(pdf_path)

        if data_df.empty:
            raise Exception(
                f"❌ Geen data gevonden in {pdf_path.name}"
            )

        output_file = output_dir / f"{pdf_path.stem}.xlsx"

        export_excel(
            toc_df,
            data_df,
            output_file
        )

        print(f"✅ Excel opgeslagen: {output_file}")

        print(f"📊 TOC rijen: {len(toc_df)}")
        print(f"📊 DATA rijen: {len(data_df)}")


if __name__ == "__main__":
    main()
