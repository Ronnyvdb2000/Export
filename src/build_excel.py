"""
build_excel.py – Ruwe Excel → gestructureerd werkboek met tabbladen

Fixes tov vorige versie:
- OFFERTE-bug: verplaatst van eenheid → offerte kolom
- Subfilter nr/naam correct via lookup (niet meer leeg)
- NPI volledig (prefix 3.5)
- 3 titelregels bovenaan elk tabblad + inhoudsopgave
- Tabs: AERi gesplitst in AERi (algemeen 3.x), NPI apart
"""

import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Titeltekst uit pagina 1 van de PDF
TITLE_LINE1 = "VLIF CODELIJST"
TITLE_LINE2 = "Geldig voor blokperiode 20/04/2026 – 30/06/2026"
TITLE_LINE3 = "Versie 20/04/2026"

# (tabblad_naam, hoofdstuk_prefix, volledige_titel)
TABS = [
    ("Rundveestal",       "2.1",  "2.1 Rundveestal"),
    ("Varkensstal",       "2.2",  "2.2 Varkensstal"),
    ("Pluimveestal",      "2.3",  "2.3 Pluimveestal"),
    ("Paardenstal",       "2.4",  "2.4 Paardenstal"),
    ("Schapenstal",       "2.5",  "2.5 Schapenstal"),
    ("Geitenstal",        "2.6",  "2.6 Geitenstal"),
    ("Hertenstal",        "2.7",  "2.7 Hertenstal"),
    ("Konijnenstal",      "2.8",  "2.8 Konijnenstal"),
    ("Serre",             "2.9",  "2.9 Serre"),
    ("Loods",             "2.10", "2.10 Loods"),
    ("Onroerend buiten",  "2.11", "2.11 Onroerend buiten"),
    ("Machines en Varia", "2.12", "2.12 Machines en Varia"),
    ("AERi RV",           "3.1",  "3.1 AERi Rundveestal"),
    ("AERi V",            "3.2",  "3.2 AERi Varkensstal"),
    ("AERi Pl",           "3.3",  "3.3 AERi Pluimveestal"),
    ("AERi OG",           "3.4",  "3.4 AERi Onroerend buiten"),
    ("NPI OG",            "3.5",  "3.5 NPI Onroerend buiten"),
]

HEADERS = ["Subfilter nr", "Subfilter naam", "Code (%)", "Omschrijving",
           "Extra info", "Kenmerk", "Eenheid", "Offerte"]

CLR_DARK   = "1F4E79"
CLR_LIGHT  = "D6E4F0"
CLR_ALT    = "EBF0F7"
CLR_WHITE  = "FFFFFF"
CLR_TOC_BG = "2E75B6"
CLR_TOC_LT = "DEEAF1"
CLR_TITLE1 = "1F4E79"
CLR_TITLE2 = "2E75B6"
CLR_TITLE3 = "4472C4"


def _border():
    t = Side(style="thin", color="B8CCE4")
    return Border(top=t, bottom=t, left=t, right=t)


def _fix_data(df):
    df = df.copy().fillna("")

    # OFFERTE in eenheid → offerte kolom
    mask_eh = df["eenheid"].str.upper().str.contains("OFFERTE", na=False)
    df.loc[mask_eh, "offerte"] = "OFFERTE"
    df.loc[mask_eh, "eenheid"] = df.loc[mask_eh, "eenheid"].str.replace(
        r"\s*/?\s*OFFERTE\s*/?", "", regex=True, flags=re.IGNORECASE
    ).str.strip(" /")

    # OFFERTE in kenmerk → offerte kolom
    mask_km = df["kenmerk"].str.upper().str.contains("OFFERTE", na=False)
    df.loc[mask_km, "offerte"] = "OFFERTE"
    df.loc[mask_km, "kenmerk"] = df.loc[mask_km, "kenmerk"].str.replace(
        r"\s*/?\s*OFFERTE\s*/?", "", regex=True, flags=re.IGNORECASE
    ).str.strip(" /")

    for col in ("hoofdstuk", "subfilter_nr", "subfilter_naam"):
        df[col] = df[col].replace("", pd.NA).ffill().fillna("")

    return df


def _write_title_rows(ws, tab_title):
    """Schrijft 3 titelrijen bovenaan elk tabblad."""
    ncols = len(HEADERS)

    # Rij 1: VLIF CODELIJST
    ws.append([TITLE_LINE1])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1)
    c.font = Font(bold=True, name="Arial", size=16, color=CLR_WHITE)
    c.fill = PatternFill("solid", fgColor=CLR_TITLE1)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    # Rij 2: blokperiode
    ws.append([TITLE_LINE2])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(2, 1)
    c.font = Font(bold=False, name="Arial", size=11, color=CLR_WHITE)
    c.fill = PatternFill("solid", fgColor=CLR_TITLE2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # Rij 3: versie + sectietitel
    ws.append([f"{TITLE_LINE3}   |   {tab_title}"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    c = ws.cell(3, 1)
    c.font = Font(bold=True, name="Arial", size=10, color=CLR_WHITE)
    c.fill = PatternFill("solid", fgColor=CLR_TITLE3)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 16

    # Lege scheidingrij
    ws.append([])
    ws.row_dimensions[4].height = 6


def _write_data_sheet(ws, tab_df, tab_title):
    _write_title_rows(ws, tab_title)

    hdr_fill  = PatternFill("solid", fgColor=CLR_DARK)
    hdr_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    sec_fill  = PatternFill("solid", fgColor=CLR_LIGHT)
    sec_font  = Font(bold=True, name="Arial", size=10, color="1F4E79")
    data_font = Font(name="Arial", size=9)
    brd       = _border()
    ncols     = len(HEADERS)
    wrap_cols = {4, 5, 6}   # omschrijving, extra_info, kenmerk

    # Header op rij 5
    ws.append(HEADERS)
    for c in range(1, ncols + 1):
        cell = ws.cell(5, c)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = brd
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[5].height = 26

    # Kolombreedtes
    for i, w in enumerate([14, 28, 8, 42, 52, 30, 12, 9], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row_idx = 6
    alt = False
    prev_sf = None

    for _, rec in tab_df.iterrows():
        sf_nr   = rec.get("subfilter_nr",   "")
        sf_naam = rec.get("subfilter_naam",  "")
        sf_key  = f"{sf_nr}|{sf_naam}"

        # Subfilter-sectieregel
        if sf_key != prev_sf and (sf_nr or sf_naam):
            label = f"{sf_nr}   {sf_naam}".strip()
            ws.cell(row_idx, 1, label)
            ws.merge_cells(start_row=row_idx, start_column=1,
                           end_row=row_idx, end_column=ncols)
            cell = ws.cell(row_idx, 1)
            cell.font      = sec_font
            cell.fill      = sec_fill
            cell.border    = brd
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[row_idx].height = 18
            row_idx += 1
            prev_sf = sf_key
            alt = False

        # Data-rij
        bg   = CLR_ALT if alt else CLR_WHITE
        fill = PatternFill("solid", fgColor=bg)
        values = [
            sf_nr, sf_naam,
            rec.get("code", ""),
            rec.get("omschrijving", ""),
            rec.get("extra_info", ""),
            rec.get("kenmerk", ""),
            rec.get("eenheid", ""),
            rec.get("offerte", ""),
        ]
        for ci, val in enumerate(values, start=1):
            cell = ws.cell(row_idx, ci, val or "")
            cell.font      = data_font
            cell.fill      = fill
            cell.border    = brd
            cell.alignment = Alignment(vertical="top", wrap_text=(ci in wrap_cols))
        ws.row_dimensions[row_idx].height = 15
        row_idx += 1
        alt = not alt

    ws.freeze_panes = "A6"
    ws.sheet_view.showGridLines = False


def _write_toc(ws, tabs_present):
    ws.title = "Inhoudsopgave"
    ws.sheet_view.showGridLines = False
    ncols = 3

    # 3 titelrijen
    ws.append([TITLE_LINE1])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1)
    c.font = Font(bold=True, name="Arial", size=16, color=CLR_WHITE)
    c.fill = PatternFill("solid", fgColor=CLR_TITLE1)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    ws.append([TITLE_LINE2])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(2, 1)
    c.font = Font(bold=False, name="Arial", size=11, color=CLR_WHITE)
    c.fill = PatternFill("solid", fgColor=CLR_TITLE2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    ws.append([f"{TITLE_LINE3}   |   Inhoudsopgave"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    c = ws.cell(3, 1)
    c.font = Font(bold=True, name="Arial", size=10, color=CLR_WHITE)
    c.fill = PatternFill("solid", fgColor=CLR_TITLE3)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 16

    ws.append([])
    ws.row_dimensions[4].height = 6

    # Header
    ws.append(["#", "Tabblad", "Volledige titel"])
    hdr_fill = PatternFill("solid", fgColor=CLR_TOC_BG)
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    brd = _border()
    for c in range(1, 4):
        cell = ws.cell(5, c)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = brd
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 22

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 50

    lnk_font = Font(name="Arial", size=10, color="1F4E79", underline="single", bold=True)
    txt_font  = Font(name="Arial", size=10)

    for i, (tab_name, prefix, full_title) in enumerate(tabs_present, start=1):
        row = 5 + i
        bg = CLR_TOC_LT if i % 2 == 0 else CLR_WHITE
        fill = PatternFill("solid", fgColor=bg)

        n_cell = ws.cell(row, 1, i)
        n_cell.font = txt_font; n_cell.fill = fill; n_cell.border = brd
        n_cell.alignment = Alignment(horizontal="center", vertical="center")

        lnk_cell = ws.cell(row, 2, tab_name)
        lnk_cell.font = lnk_font; lnk_cell.fill = fill; lnk_cell.border = brd
        lnk_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lnk_cell.hyperlink = f"#{tab_name}!A1"

        t_cell = ws.cell(row, 3, full_title)
        t_cell.font = txt_font; t_cell.fill = fill; t_cell.border = brd
        t_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 18

    ws.freeze_panes = "A6"


def build(input_xlsx, output_xlsx):
    print(f"📖 Inlezen: {input_xlsx}")
    df = pd.read_excel(input_xlsx, dtype=str)
    df = _fix_data(df)

    wb = Workbook()
    wb.remove(wb.active)
    toc_ws = wb.create_sheet("Inhoudsopgave", 0)

    tabs_present = []
    for tab_name, prefix, full_title in TABS:
        tab_df = df[df["hoofdstuk"] == prefix].copy()
        if tab_df.empty:
            print(f"   ⚠️  Geen data voor {prefix} ({tab_name})")
            continue
        ws = wb.create_sheet(tab_name)
        _write_data_sheet(ws, tab_df, full_title)
        tabs_present.append((tab_name, prefix, full_title))
        print(f"   ✅ {tab_name:22s}  {len(tab_df):4d} rijen")

    _write_toc(toc_ws, tabs_present)
    wb.save(output_xlsx)
    print(f"\n🎉 Opgeslagen: {output_xlsx}")


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--input",  required=True)
    p.add_argument("--output", required=True)
    a = p.parse_args()
    build(a.input, a.output)
