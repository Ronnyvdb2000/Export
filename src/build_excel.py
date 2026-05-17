"""
build_excel.py – Ruwe Excel → finale VLIF Excel

Verbeteringen:
- Subfilter = 1 gecombineerd veld
- Geen alternerende kleuren meer
- Correcte ondersteuning:
    2.x Productieve investeringen
    3.x AERi
    3.5 NPI roerende technieken
    4.x Niet-productieve investeringen
- Apart tabblad voor roerende AEA-technieken
- Inhoudsopgave
"""

import re
import argparse
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# Titel
# ------------------------------------------------------------

TITLE_LINE1 = "VLIF CODELIJST"
TITLE_LINE2 = "Automatisch geëxporteerd uit PDF"
TITLE_LINE3 = "Conversie via PyMuPDF"

# ------------------------------------------------------------
# Tabbladen
# ------------------------------------------------------------

TABS = [

    # --------------------------------------------------------
    # 2.x PRODUCTIEVE INVESTERINGEN
    # --------------------------------------------------------

    (
        "2.1 Rundveestal",
        "2.1",
    ),

    (
        "2.2 Varkensstal",
        "2.2",
    ),

    (
        "2.3 Pluimveestal",
        "2.3",
    ),

    (
        "2.4 Paardenstal",
        "2.4",
    ),

    (
        "2.5 Schapenstal",
        "2.5",
    ),

    (
        "2.6 Geitenstal",
        "2.6",
    ),

    (
        "2.7 Hertenstal",
        "2.7",
    ),

    (
        "2.8 Konijnenstal",
        "2.8",
    ),

    (
        "2.9 Serre",
        "2.9",
    ),

    (
        "2.10 Loods",
        "2.10",
    ),

    (
        "2.11 Onroerend buiten",
        "2.11",
    ),

    (
        "2.12 Machines en varia",
        "2.12",
    ),

    # --------------------------------------------------------
    # 3.x AERi
    # --------------------------------------------------------

    (
        "3.1 AERi Rundvee",
        "3.1",
    ),

    (
        "3.2 AERi Varkens",
        "3.2",
    ),

    (
        "3.3 AERi Pluimvee",
        "3.3",
    ),

    (
        "3.4 AERi Buiten",
        "3.4",
    ),

    # --------------------------------------------------------
    # 3.5 NPI
    # --------------------------------------------------------

    (
        "3.5 Roerende AEA",
        "3.5",
    ),

    # --------------------------------------------------------
    # 4.x NIET-PRODUCTIEF
    # --------------------------------------------------------

    (
        "4 Niet-productief",
        "4",
    ),
]

# ------------------------------------------------------------

HEADERS = [
    "Subfilter",
    "Code (%)",
    "Omschrijving",
    "Extra info",
    "Kenmerk",
    "Eenheid",
    "Offerte",
]

# ------------------------------------------------------------

CLR_DARK = "1F4E79"
CLR_LIGHT = "D9EAF7"
CLR_WHITE = "FFFFFF"

# ------------------------------------------------------------


def border():

    thin = Side(
        style="thin",
        color="B8CCE4"
    )

    return Border(
        top=thin,
        bottom=thin,
        left=thin,
        right=thin,
    )


# ------------------------------------------------------------


def fix_dataframe(df):

    df = df.copy()

    df = df.fillna("")

    # --------------------------------------------------------
    # OFFERTE fout corrigeren
    # --------------------------------------------------------

    mask = df["eenheid"].str.upper().str.contains(
        "OFFERTE",
        na=False
    )

    df.loc[mask, "offerte"] = "OFFERTE"

    df.loc[mask, "eenheid"] = (
        df.loc[mask, "eenheid"]
        .str.replace(
            r"\s*/?\s*OFFERTE\s*/?",
            "",
            regex=True,
            flags=re.IGNORECASE
        )
        .str.strip(" /")
    )

    # --------------------------------------------------------

    mask2 = df["kenmerk"].str.upper().str.contains(
        "OFFERTE",
        na=False
    )

    df.loc[mask2, "offerte"] = "OFFERTE"

    df.loc[mask2, "kenmerk"] = (
        df.loc[mask2, "kenmerk"]
        .str.replace(
            r"\s*/?\s*OFFERTE\s*/?",
            "",
            regex=True,
            flags=re.IGNORECASE
        )
        .str.strip(" /")
    )

    # --------------------------------------------------------

    for col in [
        "hoofdstuk",
        "subfilter",
    ]:

        df[col] = (
            df[col]
            .replace("", pd.NA)
            .ffill()
            .fillna("")
        )

    return df


# ------------------------------------------------------------


def write_title(ws, title):

    ncols = len(HEADERS)

    # --------------------------------------------------------
    # Titel 1
    # --------------------------------------------------------

    ws.append([TITLE_LINE1])

    ws.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=ncols
    )

    c = ws.cell(1, 1)

    c.font = Font(
        bold=True,
        size=16,
        color=CLR_WHITE,
        name="Arial"
    )

    c.fill = PatternFill(
        "solid",
        fgColor=CLR_DARK
    )

    c.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1
    )

    ws.row_dimensions[1].height = 24

    # --------------------------------------------------------
    # Titel 2
    # --------------------------------------------------------

    ws.append([TITLE_LINE2])

    ws.merge_cells(
        start_row=2,
        start_column=1,
        end_row=2,
        end_column=ncols
    )

    c = ws.cell(2, 1)

    c.font = Font(
        size=11,
        color=CLR_WHITE,
        name="Arial"
    )

    c.fill = PatternFill(
        "solid",
        fgColor="4472C4"
    )

    c.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1
    )

    ws.row_dimensions[2].height = 18

    # --------------------------------------------------------
    # Titel 3
    # --------------------------------------------------------

    ws.append([title])

    ws.merge_cells(
        start_row=3,
        start_column=1,
        end_row=3,
        end_column=ncols
    )

    c = ws.cell(3, 1)

    c.font = Font(
        bold=True,
        size=10,
        color=CLR_WHITE,
        name="Arial"
    )

    c.fill = PatternFill(
        "solid",
        fgColor="5B9BD5"
    )

    c.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1
    )

    ws.row_dimensions[3].height = 16

    # lege rij

    ws.append([])


# ------------------------------------------------------------


def write_sheet(ws, df, title):

    write_title(ws, title)

    brd = border()

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    ws.append(HEADERS)

    for col in range(1, len(HEADERS) + 1):

        cell = ws.cell(5, col)

        cell.font = Font(
            bold=True,
            color=CLR_WHITE,
            name="Arial",
            size=10
        )

        cell.fill = PatternFill(
            "solid",
            fgColor=CLR_DARK
        )

        cell.border = brd

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    ws.row_dimensions[5].height = 26

    # --------------------------------------------------------
    # Kolombreedtes
    # --------------------------------------------------------

    widths = [
        35,
        10,
        55,
        55,
        35,
        14,
        10,
    ]

    for i, w in enumerate(widths, start=1):

        ws.column_dimensions[
            get_column_letter(i)
        ].width = w

    # --------------------------------------------------------
    # Data
    # --------------------------------------------------------

    row = 6

    prev_sf = None

    for _, rec in df.iterrows():

        sf = rec.get("subfilter", "")

        # ----------------------------------------------------
        # Nieuwe subfilter-sectie
        # ----------------------------------------------------

        if sf != prev_sf and sf:

            ws.cell(row, 1, sf)

            ws.merge_cells(
                start_row=row,
                start_column=1,
                end_row=row,
                end_column=len(HEADERS)
            )

            c = ws.cell(row, 1)

            c.font = Font(
                bold=True,
                color="1F1F1F",
                name="Arial",
                size=10
            )

            c.fill = PatternFill(
                "solid",
                fgColor=CLR_LIGHT
            )

            c.border = brd

            c.alignment = Alignment(
                horizontal="left",
                vertical="center",
                indent=1
            )

            ws.row_dimensions[row].height = 18

            row += 1

            prev_sf = sf

        # ----------------------------------------------------
        # Dataregel
        # ----------------------------------------------------

        values = [

            sf,

            rec.get("code", ""),

            rec.get("omschrijving", ""),

            rec.get("extra_info", ""),

            rec.get("kenmerk", ""),

            rec.get("eenheid", ""),

            rec.get("offerte", ""),
        ]

        for ci, val in enumerate(values, start=1):

            cell = ws.cell(row, ci, val)

            cell.font = Font(
                name="Arial",
                size=9
            )

            # altijd wit
            cell.fill = PatternFill(
                "solid",
                fgColor=CLR_WHITE
            )

            cell.border = brd

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

        ws.row_dimensions[row].height = 15

        row += 1

    ws.freeze_panes = "A6"

    ws.sheet_view.showGridLines = False


# ------------------------------------------------------------


def write_toc(ws, tabs):

    ws.title = "Inhoudsopgave"

    write_title(ws, "Inhoudsopgave")

    brd = border()

    headers = [
        "#",
        "Tabblad",
    ]

    ws.append(headers)

    for c in range(1, 3):

        cell = ws.cell(5, c)

        cell.font = Font(
            bold=True,
            color=CLR_WHITE,
            name="Arial"
        )

        cell.fill = PatternFill(
            "solid",
            fgColor=CLR_DARK
        )

        cell.border = brd

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40

    row = 6

    for i, name in enumerate(tabs, start=1):

        ws.cell(row, 1, i)

        link = ws.cell(row, 2, name)

        link.hyperlink = f"#{name}!A1"

        link.font = Font(
            color="0563C1",
            underline="single",
            bold=True,
            name="Arial"
        )

        for c in range(1, 3):

            cell = ws.cell(row, c)

            cell.border = brd

            cell.alignment = Alignment(
                vertical="center"
            )

        row += 1

    ws.freeze_panes = "A6"


# ------------------------------------------------------------


def build_excel(input_xlsx, output_xlsx):

    print(f"📖 Inlezen: {input_xlsx}")

    df = pd.read_excel(
        input_xlsx,
        dtype=str
    )

    df = fix_dataframe(df)

    wb = Workbook()

    wb.remove(wb.active)

    toc = wb.create_sheet(
        "Inhoudsopgave",
        0
    )

    created_tabs = []

    # --------------------------------------------------------
    # Alle tabbladen
    # --------------------------------------------------------

    for tab_name, prefix in TABS:

        tab_df = df[
            df["hoofdstuk"]
            .str.startswith(prefix, na=False)
        ].copy()

        if tab_df.empty:

            print(f"⚠️ Geen data voor {tab_name}")

            continue

        ws = wb.create_sheet(tab_name)

        write_sheet(
            ws,
            tab_df,
            tab_name
        )

        created_tabs.append(tab_name)

        print(
            f"✅ {tab_name:30s} "
            f"{len(tab_df):5d} rijen"
        )

    # --------------------------------------------------------

    write_toc(
        toc,
        created_tabs
    )

    wb.save(output_xlsx)

    print(f"\n🎉 Opgeslagen: {output_xlsx}")


# ------------------------------------------------------------


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--input",
        required=True
    )

    parser.add_argument(
        "--output",
        required=True
    )

    args = parser.parse_args()

    build_excel(
        args.input,
        args.output
    )


# ------------------------------------------------------------

if __name__ == "__main__":
    main()
