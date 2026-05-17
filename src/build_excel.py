import argparse
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd


def verwerk_en_bouw_excel(input_path, output_path):
    print(f"📖 Inlezen van ruwe data: {input_path}")

    # Lees de ruwe Excel in die door layout_engine.py is gegenereerd
    try:
        df = pd.read_excel(input_path, dtype=str)
    except Exception as e:
        print(f"❌ Fout bij het openen van de Excel-file: {e}")
        sys.exit(1)

    # Kolommapping fixen (van generieke Kolom_X naar wat build_excel verwacht)
    rename_dict = {}
    if "Kolom_3" in df.columns:
        rename_dict["Kolom_3"] = "eenheid"
    if "Kolom_4" in df.columns:
        rename_dict["Kolom_4"] = "bedrag"

    if rename_dict:
        print(f"🔄 Kolommen hernoemen: {rename_dict}")
        df = df.rename(columns=rename_dict)

    # Zorg dat de cruciale kolommen aanwezig zijn, zelfs als ze leeg zijn
    for col in ["subfilter", "Code", "Omschrijving", "eenheid", "bedrag"]:
        if col not in df.columns:
            df[col] = ""

    # Eventuele NaN waarden omzetten naar lege strings
    df = df.fillna("")

    print(f"📊 Excel opmaken en structureren...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VLIF Codelijst"

    # Hulplijnen aanzetten
    ws.views.sheetView[0].showGridLines = True

    # Stijlen definiëren
    font_title = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="Segoe UI", size=11, bold=True, color="2C3E50")
    font_data = Font(name="Segoe UI", size=10, color="333333")
    font_code = Font(name="Segoe UI", size=10, bold=True, color="2C3E50")

    fill_title = PatternFill(
        start_color="2C3E50", end_color="2C3E50", fill_type="solid"
    )
    fill_header = PatternFill(
        start_color="ECF0F1", end_color="ECF0F1", fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin", color="BDC3C7"),
        right=Side(style="thin", color="BDC3C7"),
        top=Side(style="thin", color="BDC3C7"),
        bottom=Side(style="thin", color="BDC3C7"),
    )

    # Schrijf Headers
    headers = ["Subfilter", "Code", "Omschrijving", "Eenheid", "Bedrag"]
    ws.append(headers)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = font_title
        cell.fill = fill_title
        cell.alignment = Alignment(
            horizontal="center" if col_num in [1, 2, 4, 5] else "left",
            vertical="center",
        )

    # Data schrijven en opmaken
    for idx, row in df.iterrows():
        ws.append(
            [
                row["subfilter"],
                row["Code"],
                row["Omschrijving"],
                row["eenheid"],
                row["bedrag"],
            ]
        )
        current_row = ws.max_row

        # Pas stijlen toe per cel in de rij
        for col_num in range(1, 6):
            cell = ws.cell(row=current_row, column=col_num)
            cell.font = font_code if col_num == 2 else font_data
            cell.border = thin_border

            # Uitlijning
            if col_num in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_num == 5:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    # Kolombreedtes automatisch aanpassen
    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                # Voorkom gigantische kolommen bij lange omschrijvingen
                val_len = len(str(cell.value))
                if col_letter == "C" and val_len > 70:
                    max_len = 70
                else:
                    max_len = max(max_len, val_len)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Opslaan
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"✅ Definitieve Excel succesvol gegenereerd: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Bouw opgemaakte Excel vanaf ruwe data."
    )
    parser.add_argument("--input", required=True, help="Pad naar het input XLSX-bestand")
    parser.add_argument(
        "--output", required=True, help="Pad naar het definitieve XLSX-bestand"
    )
    args = parser.parse_args()

    verwerk_en_bouw_excel(args.input, args.output)


if __name__ == "__main__":
    main()
