import argparse
from pathlib import Path

import pandas as pd
import pdfplumber


def clean_cell(value):
    if value is None:
        return ""

    return str(value).strip()


def extract_tables(pdf_path):
    all_rows = []
    inhoud = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            print(f"📄 Verwerken pagina {page_number}")

            try:
                tables = page.extract_tables()

                if not tables:
                    continue

                for table in tables:
                    if not table:
                        continue

                    for row in table:
                        if not row:
                            continue

                        cleaned_row = [clean_cell(cell) for cell in row]

                        # Skip lege rijen
                        if not any(cleaned_row):
                            continue

                        all_rows.append(cleaned_row)
                        inhoud.append(" | ".join(cleaned_row))

            except Exception as e:
                print(f"⚠️ Fout op pagina {page_number}: {e}")
                continue

    # GEEN crash meer
    if len(all_rows) == 0:
        print(f"⚠️ Geen data gevonden in {pdf_path.name}")
        return "", []

    return "\n".join(inhoud), all_rows


def save_excel(rows, output_path):
    max_cols = max(len(r) for r in rows)

    normalized_rows = []

    for row in rows:
        normalized_rows.append(
            row + [""] * (max_cols - len(row))
        )

    df = pd.DataFrame(normalized_rows)

    df.to_excel(output_path, index=False)

    print(f"✅ Excel opgeslagen: {output_path}")


def process_pdf(pdf_path, output_dir):
    print(f"\n🚀 Start conversie: {pdf_path.name}")

    inhoud, rows = extract_tables(pdf_path)

    # Geen data → gewoon overslaan
    if not rows:
        print(f"⏭️ Bestand overgeslagen: {pdf_path.name}")
        return

    output_file = output_dir / f"{pdf_path.stem}.xlsx"

    save_excel(rows, output_file)


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--input-dir",
        required=True,
        help="Map met PDF bestanden"
    )

    parser.add_argument(
        "--output-dir",
        required=True,
        help="Map voor output bestanden"
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = list(input_dir.glob("*.pdf"))

    print(f"📁 Gevonden PDF bestanden: {[p.name for p in pdf_files]}")

    if not pdf_files:
        print("⚠️ Geen PDF bestanden gevonden")
        return

    for pdf_path in pdf_files:
        try:
            process_pdf(pdf_path, output_dir)

        except Exception as e:
            print(f"❌ Fout bij verwerken van {pdf_path.name}: {e}")

    print("\n🎉 Conversie voltooid")


if __name__ == "__main__":
    main()
