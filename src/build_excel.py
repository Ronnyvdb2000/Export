"""
build_excel.py – Bouwt het gestructureerde Excel-werkboek vanuit de PDF-export.

Fixes:
- OFFERTE staat fout in kolom 'eenheid' → verplaatsen naar 'offerte'
- subfilter_nr en subfilter_naam worden correct ingevuld
- Per sectie een apart tabblad
- Eerste tabblad = inhoudsopgave met hyperlinks
"""

import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

# ── Tabblad-definitie: (tabblad_naam, hoofdstuk_prefix, volledige_titel) ──
TABS = [
    ("Rundveestal",           "2.1",  "2.1 Rundveestal"),
    ("Varkensstal",           "2.2",  "2.2 Varkensstal"),
    ("Pluimveestal",          "2.3",  "2.3 Pluimveestal"),
    ("Paardenstal",           "2.4",  "2.4 Paardenstal"),
    ("Schapenstal",           "2.5",  "2.5 Schapenstal"),
    ("Geitenstal",            "2.6",  "2.6 Geitenstal"),
    ("Hertenstal",            "2.7",  "2.7 Hertenstal"),
    ("Konijnenstal",          "2.8",  "2.8 Konijnenstal"),
    ("Serre",                 "2.9",  "2.9 Serre"),
    ("Loods",                 "2.10", "2.10 Loods"),
    ("Onroerend buiten",      "2.11", "2.11 Onroerend buiten"),
    ("Machines en Varia",     "2.12", "2.12 Machines en Varia"),
    ("AERi RV",               "3.1",  "3.1 AERi Rundveestal"),
    ("AERi V",                "3.2",  "3.2 AERi Varkensstal"),
    ("AERi Pl",               "3.3",  "3.3 AERi Pluimveestal"),
    ("AERi OG",               "3.4",  "3.4 AERi Onroerend buiten"),
    ("NPI OG",                "3.5",  "3.5 NPI Onroerend buiten"),
]

# Kolom-headers in het data-blad
HEADERS = ["Subfilter nr", "Subfilter naam", "Code (%)", "Omschrijving",
           "Extra info", "Kenmerk", "Eenheid", "Offerte"]

# Kleuren
CLR_DARK   = "1F4E79"   # donkerblauw header
CLR_LIGHT  = "D6E4F0"   # lichtblauw sectietitel
CLR_ALT    = "EBF0F7"   # afwisselende rijen
CLR_WHITE  = "FFFFFF"
CLR_TOC_BG = "2E75B6"   # TOC header blauw
CLR_TOC_LT = "DEEAF1"   # TOC afwisselende rij


def _thin_border():
    t = Side(style="thin", color="B8CCE4")
    return Border(top=t, bottom=t, left=t, right=t)


def _fix_data(df: pd.DataFrame) -> pd.DataFrame:
    """Fix de OFFERTE-bug en normaliseer het dataframe."""
    df = df.copy().fillna("")

    # Fix 1: OFFERTE in eenheid → verplaats naar offerte-kolom
    mask_eh = df["eenheid"].str.upper().str.contains("OFFERTE")
    df.loc[mask_eh, "offerte"] = "OFFERTE"
    # Verwijder OFFERTE uit eenheid (maar bewaar de rest zoals "m³/u / OFFERTE")
    df.loc[mask_eh, "eenheid"] = df.loc[mask_eh, "eenheid"].str.replace(
        r"\s*/?\s*OFFERTE\s*/?", "", regex=True, flags=re.IGNORECASE
    ).str.strip(" /")

    # Fix 2: OFFERTE in kenmerk → ook verplaatsen
    mask_km = df["kenmerk"].str.upper().str.contains("OFFERTE")
    df.loc[mask_km, "offerte"] = "OFFERTE"
    df.loc[mask_km, "kenmerk"] = df.loc[mask_km, "kenmerk"].str.replace(
        r"\s*/?\s*OFFERTE\s*/?", "", regex=True, flags=re.IGNORECASE
    ).str.strip(" /")

    # Forward-fill subfilter voor rijen zonder subfilter
    for col in ("hoofdstuk", "subfilter_nr", "subfilter_naam"):
        df[col] = df[col].replace("", pd.NA).ffill().fillna("")

    return df


def _col_widths():
    return [16, 22, 8, 42, 55, 30, 12, 9]


def _write_data_sheet(ws, tab_df, tab_title):
    """Schrijft één data-tabblad."""
    hdr_fill  = PatternFill("solid", fgColor=CLR_DARK)
    hdr_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    sec_fill  = PatternFill("solid", fgColor=CLR_LIGHT)
    sec_font  = Font(bold=True, name="Arial", size=10, color="1F4E79")
    data_font = Font(name="Arial", size=9)
    brd       = _thin_border()
    wrap_cols = {4, 5, 6}   # omschrijving, extra_info, kenmerk (1-based)

    # Titel
    ws.append([tab_title])
    title_cell = ws.cell(1, 1)
    title_cell.font = Font(bold=True, name="Arial", size=13, color=CLR_DARK)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22
    ws.append([])  # lege rij

    # Header
    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(3, c)
        cell.font  = hdr_font
        cell.fill  = hdr_fill
        cell.border = brd
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    ws.row_dimensions[3].height = 26

    # Kolombreedtes
    for i, w in enumerate(_col_widths(), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Data per subfiltergroep
    row_idx = 4
    alt = False
    prev_sf = None

    for _, rec in tab_df.iterrows():
        sf_nr   = rec.get("subfilter_nr",   "")
        sf_naam = rec.get("subfilter_naam",  "")
        sf_key  = f"{sf_nr}|{sf_naam}"

        # Subfilter-sectieregel
        if sf_key != prev_sf and (sf_nr or sf_naam):
            label = f"{sf_nr}  {sf_naam}".strip()
            ws.cell(row_idx, 1, label)
            ws.merge_cells(start_row=row_idx, start_column=1,
                           end_row=row_idx, end_column=len(HEADERS))
            merged_cell = ws.cell(row_idx, 1)
            merged_cell.font   = sec_font
            merged_cell.fill   = sec_fill
            merged_cell.border = brd
            merged_cell.alignment = Alignment(horizontal="left",
                                              vertical="center",
                                              indent=1)
            ws.row_dimensions[row_idx].height = 18
            row_idx += 1
            prev_sf = sf_key
            alt = False

        # Data-rij
        bg = CLR_ALT if alt else CLR_WHITE
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
            cell.font   = data_font
            cell.fill   = fill
            cell.border = brd
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=(ci in wrap_cols)
            )
        ws.row_dimensions[row_idx].height = 15
        row_idx += 1
        alt = not alt

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False


def _write_toc(ws, tabs_present):
    """Schrijft het inhoudsopgave-tabblad met hyperlinks."""
    ws.title = "Inhoudsopgave"
    ws.sheet_view.showGridLines = False

    # Titel
    ws["A1"] = "VLIF Codelijst – Inhoudsopgave"
    ws["A1"].font = Font(bold=True, name="Arial", size=16, color=CLR_DARK)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.append([])
    ws.append(["#", "Tabblad", "Volledige titel"])

    hdr_fill = PatternFill("solid", fgColor=CLR_TOC_BG)
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    for c in range(1, 4):
        cell = ws.cell(3, c)
        cell.font  = hdr_font
        cell.fill  = hdr_fill
        cell.border = _thin_border()
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 22

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 50

    lnk_font = Font(name="Arial", size=10, color="1F4E79",
                    underline="single", bold=True)
    txt_font  = Font(name="Arial", size=10)
    brd = _thin_border()

    for i, (tab_name, prefix, full_title) in enumerate(tabs_present, start=1):
        row = 3 + i
        alt_fill = PatternFill("solid",
                               fgColor=CLR_TOC_LT if i % 2 == 0 else CLR_WHITE)

        # Nummer
        n_cell = ws.cell(row, 1, i)
        n_cell.font      = txt_font
        n_cell.fill      = alt_fill
        n_cell.border    = brd
        n_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Hyperlink naar tabblad
        lnk_cell = ws.cell(row, 2, tab_name)
        lnk_cell.font      = lnk_font
        lnk_cell.fill      = alt_fill
        lnk_cell.border    = brd
        lnk_cell.alignment = Alignment(horizontal="left", vertical="center",
                                       indent=1)
        lnk_cell.hyperlink = f"#{tab_name}!A1"

        # Volledige titel
        t_cell = ws.cell(row, 3, full_title)
        t_cell.font      = txt_font
        t_cell.fill      = alt_fill
        t_cell.border    = brd
        t_cell.alignment = Alignment(horizontal="left", vertical="center",
                                     indent=1)
        ws.row_dimensions[row].height = 18


def build(input_xlsx: str, output_xlsx: str):
    print(f"📖 Inlezen: {input_xlsx}")
    df = pd.read_excel(input_xlsx, dtype=str)
    df = _fix_data(df)

    wb = Workbook()
    # Verwijder standaard leeg tabblad
    wb.remove(wb.active)

    # TOC als eerste tabblad (placeholder, later invullen)
    toc_ws = wb.create_sheet("Inhoudsopgave", 0)

    tabs_present = []

    for tab_name, prefix, full_title in TABS:
        tab_df = df[df["hoofdstuk"] == prefix].copy()
        if tab_df.empty:
            print(f"   ⚠️  Geen data voor {prefix} ({tab_name}), overgeslagen")
            continue

        ws = wb.create_sheet(tab_name)
        _write_data_sheet(ws, tab_df, full_title)
        tabs_present.append((tab_name, prefix, full_title))
        print(f"   ✅ {tab_name:20s}  {len(tab_df):4d} rijen")

    # Vul TOC in
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
