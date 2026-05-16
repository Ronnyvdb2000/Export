"""
build_excel.py – verwerkt *_raw.xlsx naar finale Excel + SQLite

Deze versie:
✔ Maakt automatisch de output-map aan
✔ Schrijft Excel én SQLite in dezelfde map
✔ Houdt jouw bestaande structuur volledig intact
✔ Formatteert Excel netjes (kolombreedtes, header, zebra)
✔ Forward-fill hiërarchie
"""

import argparse
from pathlib import Path
import pandas as pd
import sqlite3

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ───────────────────────────────────────────────────────────────
#  U T I L S
# ───────────────────────────────────────────────────────────────

def ensure_output_dir(path):
    """Zorgt dat de output-map bestaat."""
    out_dir = Path(path).parent
    out_dir.mkdir(parents=True, exist_ok=True)


# ───────────────────────────────────────────────────────────────
#  E X C E L   O P S L A A N
# ───────────────────────────────────────────────────────────────

def save_excel(df, output_path):
    """Schrijft de finale Excel weg met formatting."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Codelijst")
        ws = writer.sheets["Codelijst"]

        # Kolombreedtes
        widths = [30, 14, 22, 6, 40, 55, 30, 12, 8]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # Header formatting
        hdr_fill = PatternFill("solid", fgColor="1F4E79")
        hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Data formatting
        thin = Side(style="thin", color="D9D9D9")
        brd = Border(bottom=thin)
        dfont = Font(name="Arial", size=9)
        wrap_cols = {5, 6, 7}  # omschrijving, extra_info, kenmerk

        for ri, row in enumerate(ws.iter_rows(min_row=2), start=2):
            fill = PatternFill("solid", fgColor="EBF0F7" if ri % 2 == 0 else "FFFFFF")
            for cell in row:
                cell.font = dfont
                cell.fill = fill
                cell.border = brd
                cell.alignment = Alignment(vertical="top", wrap_text=(cell.column in wrap_cols))

        ws.freeze_panes = "A2"

    print(f"📄 Excel opgeslagen: {output_path}")


# ───────────────────────────────────────────────────────────────
#  S Q L I T E   O P S L A A N
# ───────────────────────────────────────────────────────────────

def save_sqlite(df, output_path):
    """Schrijft dezelfde data weg naar SQLite."""
    db_path = Path(output_path).with_suffix(".db")
    conn = sqlite3.connect(db_path)

    df.to_sql("vlif", conn, if_exists="replace", index=False)

    # Indexen voor snelle queries
    conn.execute("CREATE INDEX IF NOT EXISTS idx_code ON vlif(code);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sf ON vlif(subfilter_nr);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_hfd ON vlif(hoofdstuk);")

    conn.commit()
    conn.close()

    print(f"🗄️ SQLite opgeslagen: {db_path}")


# ───────────────────────────────────────────────────────────────
#  M A I N
# ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Pad naar *_raw.xlsx")
    parser.add_argument("--output", required=True, help="Pad naar finale Excel")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    ensure_output_dir(output_path)

    print(f"📘 Inlezen raw Excel: {input_path}")
    df = pd.read_excel(input_path)

    # Forward-fill hiërarchie
    for col in ("hoofdstuk", "subfilter_nr", "subfilter_naam"):
        df[col] = df[col].replace("", pd.NA).ffill().fillna("")

    # Verwijder rijen zonder code
    df = df[df["code"].notna() & (df["code"] != "")].reset_index(drop=True)

    # Excel + SQLite wegschrijven
    save_excel(df, output_path)
    save_sqlite(df, output_path)

    print("🎉 build_excel voltooid!")


if __name__ == "__main__":
    main()
