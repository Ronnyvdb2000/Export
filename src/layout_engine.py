"""
layout_engine.py – VLIF4 PDF → Excel + SQLite (PyMuPDF)

Deze versie:
✔ Maakt automatisch de output-map aan
✔ Schrijft Excel én SQLite in dezelfde map
✔ Houdt jouw bestaande structuur volledig intact
"""

import re
import argparse
from pathlib import Path

import fitz
import pandas as pd
import sqlite3

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ───────────────────────────────────────────────────────────────
#  K O L O M - G R E N Z E N
# ───────────────────────────────────────────────────────────────

X_NR         = (35,   80)
X_NAAM       = (68,  216)
X_CODE       = (196, 216)
X_OMSCHR     = (216, 371)
X_EXTRA      = (371, 596)
X_KENMERK    = (596, 696)
X_EENHEID    = (696, 771)
X_OFFERTE    = (771, 842)

Y_HEADER_MAX = 60
Y_FOOTER_MIN = 520

CODE_RE          = re.compile(r"^(15|30|40)$")
SUBFILTER_NR_RE  = re.compile(r"^\d+\.\d+\.\d+$")
HOOFD_NR_RE      = re.compile(r"^\d+\.?\d*$")


# ───────────────────────────────────────────────────────────────
#  U T I L S
# ───────────────────────────────────────────────────────────────

def ensure_output_dir(path):
    """Zorgt dat de output-map bestaat."""
    out_dir = Path(path).parent
    out_dir.mkdir(parents=True, exist_ok=True)


def _classify(x0, size):
    big = size >= 14
    if X_NR[0] <= x0 < X_NR[1]:
        return ("nr", big)
    if X_NAAM[0] <= x0 < X_CODE[0]:
        return ("naam", big)
    if X_CODE[0] <= x0 < X_CODE[1]:
        return ("code", False)
    if X_OMSCHR[0] <= x0 < X_EXTRA[0]:
        return ("omschrijving", False)
    if X_EXTRA[0] <= x0 < X_KENMERK[0]:
        return ("extra_info", False)
    if X_KENMERK[0] <= x0 < X_EENHEID[0]:
        return ("kenmerk", False)
    if X_EENHEID[0] <= x0 < X_OFFERTE[0]:
        return ("eenheid", False)
    if X_OFFERTE[0] <= x0:
        return ("offerte", False)
    return ("ignore", False)


def _extract_spans(pdf_path):
    doc = fitz.open(pdf_path)
    spans = []

    for page_no, page in enumerate(doc, start=1):
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = " ".join(span["text"].split())
                    if not text:
                        continue
                    x0 = span["bbox"][0]
                    y0 = span["bbox"][1]
                    size = span["size"]

                    if y0 < Y_HEADER_MAX or y0 > Y_FOOTER_MIN:
                        continue

                    col, big = _classify(x0, size)
                    if col == "ignore":
                        continue

                    spans.append((page_no, round(y0, 1), round(x0, 1), col, big, text))

    spans.sort(key=lambda s: (s[0], s[1], s[2]))
    return spans


def _group_lines(spans):
    lines = []
    cur = None
    cur_page = cur_y = None

    for page, y, x0, col, big, text in spans:
        new = (cur is None) or (page != cur_page) or (abs(y - cur_y) > 3)

        if new:
            if cur:
                lines.append(cur)
            cur = {"page": page, "y": y, "big": big}
            cur_page, cur_y = page, y

        if big:
            cur["big"] = True

        cur[col] = (cur[col] + " " + text) if col in cur else text

    if cur:
        lines.append(cur)

    return lines


def _build_records(lines):
    records = []

    cur_hoofdstuk = ""
    cur_sf_nr = cur_sf_naam = ""
    cur_code = None
    cur_omschr = []
    cur_extra = []
    cur_kenmerk = []
    cur_eenheid = []
    cur_offerte = ""

    def reset_record():
        nonlocal cur_code, cur_omschr, cur_extra, cur_kenmerk, cur_eenheid, cur_offerte
        cur_code = None
        cur_omschr = []
        cur_extra = []
        cur_kenmerk = []
        cur_eenheid = []
        cur_offerte = ""

    def flush():
        if cur_code is None and not cur_omschr:
            return
        records.append({
            "hoofdstuk": cur_hoofdstuk,
            "subfilter_nr": cur_sf_nr,
            "subfilter_naam": cur_sf_naam,
            "code": cur_code or "",
            "omschrijving": " ".join(cur_omschr).strip(),
            "extra_info": " ".join(cur_extra).strip(),
            "kenmerk": " / ".join(cur_kenmerk),
            "eenheid": " / ".join(cur_eenheid),
            "offerte": cur_offerte,
        })
        reset_record()

    for line in lines:
        nr     = line.get("nr", "").strip()
        naam   = line.get("naam", "").strip()
        code   = line.get("code", "").strip()
        omschr = line.get("omschrijving", "").strip()
        extra  = line.get("extra_info", "").strip()
        kenmrk = line.get("kenmerk", "").strip()
        eenhd  = line.get("eenheid", "").strip()
        offrte = line.get("offerte", "").strip()
        big    = line.get("big", False)

        # Hoofdstuk
        if big and nr and HOOFD_NR_RE.match(nr):
            flush()
            cur_hoofdstuk = f"{nr} {naam}".strip()
            cur_sf_nr = cur_sf_naam = ""
            reset_record()
            continue

        # Subfilter
        if nr and SUBFILTER_NR_RE.match(nr):
            flush()
            cur_sf_nr = nr
            cur_sf_naam = naam
            continue

        # Nieuwe code
        if code and CODE_RE.match(code):
            flush()
            cur_code = code
            if omschr: cur_omschr.append(omschr)
            if extra:  cur_extra.append(extra)
            if kenmrk: cur_kenmerk.append(kenmrk)
            if eenhd:  cur_eenheid.append(eenhd)
            if offrte: cur_offerte = offrte
            continue

        # Vervolgregels
        if cur_code is not None:
            if omschr: cur_omschr.append(omschr)
            if extra:  cur_extra.append(extra)
            if kenmrk: cur_kenmerk.append(kenmrk)
            if eenhd:  cur_eenheid.append(eenhd)
            if offrte and not cur_offerte:
                cur_offerte = offrte

    flush()
    return records


# ───────────────────────────────────────────────────────────────
#  E X C E L   +   S Q L I T E
# ───────────────────────────────────────────────────────────────

def save_excel(df, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Codelijst")
        ws = writer.sheets["Codelijst"]

        # Kolombreedtes
        widths = [30, 14, 22, 6, 40, 55, 30, 12, 8]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # Header
        hdr_fill = PatternFill("solid", fgColor="1F4E79")
        hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Data
        thin = Side(style="thin", color="D9D9D9")
        brd = Border(bottom=thin)
        dfont = Font(name="Arial", size=9)
        wrap_cols = {5, 6, 7}

        for ri, row in enumerate(ws.iter_rows(min_row=2), start=2):
            fill = PatternFill("solid", fgColor="EBF0F7" if ri % 2 == 0 else "FFFFFF")
            for cell in row:
                cell.font = dfont
                cell.fill = fill
                cell.border = brd
                cell.alignment = Alignment(vertical="top", wrap_text=(cell.column in wrap_cols))

        ws.freeze_panes = "A2"

    print(f"📄 Excel opgeslagen: {output_path}")


def save_sqlite(df, output_path):
    db_path = Path(output_path).with_suffix(".db")
    conn = sqlite3.connect(db_path)

    df.to_sql("vlif", conn, if_exists="replace", index=False)

    conn.execute("CREATE INDEX IF NOT EXISTS idx_code ON vlif(code);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sf ON vlif(subfilter_nr);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_hfd ON vlif(hoofdstuk);")

    conn.commit()
    conn.close()

    print(f"🗄️ SQLite opgeslagen: {db_path}")


# ───────────────────────────────────────────────────────────────
#  M A I N
# ───────────────────────────────────────────────────────────────

def parse_pdf(pdf_path):
    print(f"📘 Parsing: {pdf_path.name}")
    spans = _extract_spans(pdf_path)
    lines = _group_lines(spans)
    records = _build_records(lines)
    return records


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-pdf", required=True)
    parser.add_argument("--output-xlsx", required=True)
    args = parser.parse_args()

    pdf_path = Path(args.input_pdf)
    output_path = Path(args.output_xlsx)

    ensure_output_dir(output_path)

    records = parse_pdf(pdf_path)
    df = pd.DataFrame(records)

    # hiërarchie forward-fill
    for col in ("hoofdstuk", "subfilter_nr", "subfilter_naam"):
        df[col] = df[col].replace("", pd.NA).ffill().fillna("")

    df = df[df["code"] != ""].reset_index(drop=True)

    save_excel(df, output_path)
    save_sqlite(df, output_path)

    print("🎉 Klaar!")


if __name__ == "__main__":
    main()
