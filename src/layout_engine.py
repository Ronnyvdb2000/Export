import argparse
import re
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================
# CONFIGURATIE & METADATA
# =========================================================
TITLE_LINE1 = "VLIF CODELIJST"
TITLE_LINE2 = "Geldig voor blokperiode 20/04/2026 – 30/06/2026"
TITLE_LINE3 = "Versie 20/04/2026"

TABS = [
    # BLOK 2
    ("Rundveestal", "2.1"),
    ("Varkensstal", "2.2"),
    ("Pluimveestal", "2.3"),
    ("Paardenstal", "2.4"),
    ("Schapenstal", "2.5"),
    ("Geitenstal", "2.6"),
    ("Hertenstal", "2.7"),
    ("Konijnenstal", "2.8"),
    ("Serre", "2.9"),
    ("Loods", "2.10"),
    ("Onroerend buiten", "2.11"),
    ("Machines en Varia", "2.12"),
    # BLOK 3
    ("AERi Rundvee", "3.1"),
    ("AERi Varkens", "3.2"),
    ("AERi Pluimvee", "3.3"),
    ("AERi Onroerend", "3.4"),
    ("Roerende AEA technieken", "3.5"),
    # BLOK 4
    ("Niet-productieve investeringen", "4"),
]

HEADERS = [
    "Subfilter",
    "Code (%)",
    "Omschrijving",
    "Extra info",
    "Kenmerk",
    "Eenheid",
    "Offerte",
]

CLR_TITLE1 = "1F4E79"
CLR_TITLE2 = "2E75B6"
CLR_TITLE3 = "4472C4"
CLR_HEADER = "1F4E79"
CLR_SECTION = "D6E4F0"
CLR_WHITE = "FFFFFF"


def thin_border():
    s = Side(style="thin", color="B8CCE4")
    return Border(top=s, bottom=s, left=s, right=s)


def fix_data(df):
    df = df.copy().fillna("")

    # Mocht 'offerte' nog niet bestaan, maak hem aan
    if "offerte" not in df.columns:
        df["offerte"] = ""

    # Offerte logica herstellen op basis van eenheid tekst
    mask = df["eenheid"].str.upper().str.contains("OFFERTE", na=False)
    df.loc[mask, "offerte"] = "OFFERTE"
    df.loc[mask, "eenheid"] = (
        df.loc[mask, "eenheid"]
        .str.replace(r"\s*/?\s*OFFERTE\s*/?", "", regex=True, flags=re.IGNORECASE)
        .str.strip(" /")
    )

    # Subfilter en prefix bepaling voor tabblad splitsing
    df["subfilter"] = df["subfilter"].astype(str).str.strip()
    df["prefix"] = df["Code"].astype(str).str.extract(r"^(\d+(?:\.\d+)?)")
    return df


def write_title_rows(ws, title):
    ncols = len(HEADERS)

    # Rij 1: Hoofdtitel
    ws.append([TITLE_LINE1])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1)
    c.font = Font(bold=True, color="FFFFFF", size=16, name="Arial")
    c.fill = PatternFill("solid", fgColor=CLR_TITLE1)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    # Rij 2: Geldigheidsperiode
    ws.append([TITLE_LINE2])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(2, 1)
    c.font = Font(color="FFFFFF", size=11, name="Arial")
    c.fill = PatternFill("solid", fgColor=CLR_TITLE2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # Rij 3: Versiebeheer en Tabblad info
    ws.append([f"{TITLE_LINE3}   |   {title}"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    c = ws.cell(3, 1)
    c.font = Font(bold=True, color="FFFFFF", size=10, name="Arial")
    c.fill = PatternFill("solid", fgColor=CLR_TITLE3)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 16

    ws.append([])  # Lege bufferrij 4


def write_sheet(ws, df, title):
    write_title_rows(ws, title)
    border = thin_border()

    # Headers schrijven (Rij 5)
    ws.append(HEADERS)
    ws.row_dimensions[5].height = 20

    for col in range(1, len(HEADERS) + 1):
        c = ws.cell(5, col)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.fill = PatternFill("solid", fgColor=CLR_HEADER)
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Kolombreedtes instellen
    widths = [28, 10, 42, 52, 30, 12, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Data wegschrijven (Vanaf rij 6)
    row = 6
    prev_sub = None

    for _, rec in df.iterrows():
        sub = rec.get("subfilter", "").strip()

        if sub and sub != prev_sub:
            ws.cell(row, 1, sub)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(HEADERS))
            c = ws.cell(row, 1)
            c.font = Font(bold=True, name="Arial", size=10, color="1F4E79")
            c.fill = PatternFill("solid", fgColor=CLR_SECTION)
            c.border = border
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[row].height = 20
            row += 1
            prev_sub = sub

        values = [
            rec.get("subfilter", ""),
            rec.get("Code", ""),
            rec.get("Omschrijving", ""),
            rec.get("extra_info", ""),
            rec.get("kenmerk", ""),
            rec.get("eenheid", ""),
            rec.get("offerte", ""),
        ]

        for ci, val in enumerate(values, start=1):
            c = ws.cell(row, ci, val)
            c.font = Font(name="Arial", size=9)
            c.fill = PatternFill("solid", fgColor=CLR_WHITE)
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=True)

        row += 1

    ws.freeze_panes = "A6"
    ws.views.sheetView[0].showGridLines = True


def write_toc(ws, tabs):
    write_title_rows(ws, "Inhoudsopgave")
    ws.append(["#", "Tabblad"])
    border = thin_border()

    for col in range(1, 3):
        c = ws.cell(5, col)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.fill = PatternFill("solid", fgColor=CLR_HEADER)
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    row = 6
    for i, tab in enumerate(tabs, start=1):
        ws.cell(row, 1, i).border = border
        ws.cell(row, 1).alignment = Alignment(horizontal="center")

        cell = ws.cell(row, 2, tab)
        cell.border = border
        cell.hyperlink = f"#'{tab}'!A1"
        cell.font = Font(color="1F4E79", underline="single", bold=True, name="Arial", size=10)
        row += 1

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40
    ws.freeze_panes = "A6"
    ws.views.sheetView[0].showGridLines = True


def build_excel(input_path, output_path):
    print(f"📖 Inlezen: {input_path}")
    df = pd.read_excel(input_path, dtype=str)

    # Flexibele kolom-mapping om KeyError te voorkomen
    rename_dict = {}
    if "Kolom_3" in df.columns:
        rename_dict["Kolom_3"] = "eenheid"
    if "Kolom_4" in df.columns:
        rename_dict["Kolom_4"] = "kenmerk"
    if "Kolom_5" in df.columns:
        rename_dict["Kolom_5"] = "extra_info"

    if rename_dict:
        df = df.rename(columns=rename_dict)

    # Garandeer aanwezigheid van kolommen voor fix_data logica
    for col in ["subfilter", "eenheid", "kenmerk", "extra_info", "offerte"]:
        if col not in df.columns:
            df[col] = ""

    df = fix_data(df)

    wb = Workbook()
    wb.remove(wb.active)  # Verwijder standaard leeg tabblad

    toc = wb.create_sheet("Inhoudsopgave", 0)
    created_tabs = []

    for tab_name, prefix in TABS:
        if prefix == "4":
            tab_df = df[df["prefix"].str.startswith("4", na=False)].copy()
        else:
            tab_df = df[df["prefix"] == prefix].copy()

        if tab_df.empty:
            print(f"⚠️ Geen data gevonden voor [{prefix}] ({tab_name})")
            continue

        ws = wb.create_sheet(tab_name)
        write_sheet(ws, tab_df, tab_name)
        created_tabs.append(tab_name)
        print(f"✅ {tab_name:30s} {len(tab_df):5d} rijen verwerkt")

    write_toc(toc, created_tabs)
    wb.save(output_path)
    print(f"\n🎉 Succesvol opgeslagen: {output_path}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    build_excel(Path(args.input), Path(args.output))


if __name__ == "__main__":
    main()
