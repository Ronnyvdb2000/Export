"""
layout_engine.py
VLIF4 PDF -> ruwe Excel

Features:
- Ondersteunt blok 2, 3 en 4
- Ondersteunt AERi + NPI + roerende AEA technieken
- Dynamische subfilters
- Robuustere x-ranges
- Werkt met VLIF4 2026 PDF's
"""

import re
import argparse
from pathlib import Path

import fitz
import pandas as pd
from openpyxl.utils import get_column_letter


# =========================================================
# X-RANGES
# =========================================================

X_SF_NR = (20, 80)
X_SF_NAAM = (80, 220)

X_CODE = (220, 260)

X_OMSCHR = (260, 430)
X_EXTRA = (430, 610)
X_KENMERK = (610, 710)
X_EENHEID = (710, 790)
X_OFFERTE = (790, 860)

Y_HEADER_MAX = 55
Y_FOOTER_MIN = 540


# =========================================================
# REGEX
# =========================================================

CODE_RE = re.compile(r"^\d+$")

SUBFILTER_RE = re.compile(r"^\d+\.\d+\.\d+$")

HOOFDSTUK_RE = re.compile(r"^\d+\.\d+$")


SKIP_VALUES = {
    "",
    "SUBFILTER",
    "OMSCHRIJVING",
    "EXTRA_INFO",
    "KENMERK",
    "EENHEID",
    "OFFERTE",
    "%",
}


# =========================================================
# KOLOM DETECTIE
# =========================================================

def detect_column(x0):

    if X_SF_NR[0] <= x0 < X_SF_NR[1]:
        return "sf_nr"

    if X_SF_NAAM[0] <= x0 < X_SF_NAAM[1]:
        return "sf_naam"

    if X_CODE[0] <= x0 < X_CODE[1]:
        return "code"

    if X_OMSCHR[0] <= x0 < X_EXTRA[0]:
        return "omschrijving"

    if X_EXTRA[0] <= x0 < X_KENMERK[0]:
        return "extra_info"

    if X_KENMERK[0] <= x0 < X_EENHEID[0]:
        return "kenmerk"

    if X_EENHEID[0] <= x0 < X_OFFERTE[0]:
        return "eenheid"

    if X_OFFERTE[0] <= x0:
        return "offerte"

    return "ignore"


# =========================================================
# PDF EXTRACT
# =========================================================

def extract_spans(pdf_path):

    doc = fitz.open(pdf_path)

    spans = []

    for page_no, page in enumerate(doc, start=1):

        blocks = page.get_text(
            "dict",
            flags=fitz.TEXT_PRESERVE_WHITESPACE
        )["blocks"]

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

                    if y0 < Y_HEADER_MAX:
                        continue

                    if y0 > Y_FOOTER_MIN:
                        continue

                    col = detect_column(x0)

                    if col == "ignore":
                        continue

                    size = span["size"]

                    spans.append((
                        page_no,
                        round(y0, 1),
                        round(x0, 1),
                        col,
                        size >= 14,
                        text
                    ))

    spans.sort(key=lambda x: (x[0], x[1], x[2]))

    return spans


# =========================================================
# GROUP LINES
# =========================================================

def group_lines(spans):

    lines = []

    cur = None

    for page, y, x0, col, big, text in spans:

        new_line = (
            cur is None
            or page != cur["page"]
            or abs(y - cur["y"]) > 3
        )

        if new_line:

            if cur:
                lines.append(cur)

            cur = {
                "page": page,
                "y": y,
                "big": big
            }

        if big:
            cur["big"] = True

        if col in cur:
            cur[col] += " " + text
        else:
            cur[col] = text

    if cur:
        lines.append(cur)

    return lines


# =========================================================
# PARSE
# =========================================================

def parse_lines(lines):

    records = []

    subfilters = []

    current_hoofdstuk = ""
    current_prefix = ""

    pending_hoofdstuk = ""

    cur_code = None
    cur_page = 0
    cur_y = 0

    cur_omschrijving = []
    cur_extra = []
    cur_kenmerk = []
    cur_eenheid = []
    cur_offerte = ""

    def flush():

        nonlocal cur_code
        nonlocal cur_omschrijving
        nonlocal cur_extra
        nonlocal cur_kenmerk
        nonlocal cur_eenheid
        nonlocal cur_offerte

        if cur_code is None:
            return

        records.append({

            "hoofdstuk": current_hoofdstuk,

            "subfilter": "",

            "code": cur_code,

            "omschrijving":
                " ".join(cur_omschrijving).strip(),

            "extra_info":
                " ".join(cur_extra).strip(),

            "kenmerk":
                " / ".join(cur_kenmerk).strip(),

            "eenheid":
                " / ".join(cur_eenheid).strip(),

            "offerte":
                cur_offerte,

            "_page": cur_page,
            "_y": cur_y,
            "_prefix": current_prefix,
        })

        cur_code = None

        cur_omschrijving = []
        cur_extra = []
        cur_kenmerk = []
        cur_eenheid = []

        cur_offerte = ""

    for line in lines:

        sf_nr = line.get("sf_nr", "").strip()
        sf_naam = line.get("sf_naam", "").strip()

        code = line.get("code", "").strip()

        omschrijving = line.get("omschrijving", "").strip()

        extra_info = line.get("extra_info", "").strip()

        kenmerk = line.get("kenmerk", "").strip()

        eenheid = line.get("eenheid", "").strip()

        offerte = line.get("offerte", "").strip()

        big = line.get("big", False)

        page = line["page"]
        y = line["y"]

        # =====================================================
        # SKIP HEADERS
        # =====================================================

        if (
            sf_nr.upper() in SKIP_VALUES
            or sf_naam.upper() in SKIP_VALUES
            or code.upper() in SKIP_VALUES
        ):
            continue

        # =====================================================
        # HOOFDSTUK
        # =====================================================

        if big:

            # 2.1 + naam op zelfde lijn

            if (
                sf_nr
                and HOOFDSTUK_RE.match(sf_nr)
                and sf_naam
            ):

                flush()

                current_hoofdstuk = f"{sf_nr} {sf_naam}".strip()

                current_prefix = sf_nr

                pending_hoofdstuk = ""

                continue

            # enkel nummer

            if (
                sf_nr
                and HOOFDSTUK_RE.match(sf_nr)
                and not sf_naam
            ):

                flush()

                pending_hoofdstuk = sf_nr

                continue

            # enkel naam

            if sf_naam and not sf_nr:

                flush()

                if pending_hoofdstuk:

                    current_hoofdstuk = (
                        f"{pending_hoofdstuk} {sf_naam}"
                    ).strip()

                    current_prefix = pending_hoofdstuk

                    pending_hoofdstuk = ""

                else:

                    current_hoofdstuk = sf_naam

                    current_prefix = ""

                continue

        # =====================================================
        # SUBFILTER
        # =====================================================

        if sf_nr and SUBFILTER_RE.match(sf_nr):

            prefix = ".".join(sf_nr.split(".")[:2])

            label = f"{sf_nr} {sf_naam}".strip()

            subfilters.append((
                page,
                y,
                prefix,
                label
            ))

            continue

        # =====================================================
        # CODE
        # =====================================================

        if code and CODE_RE.match(code):

            flush()

            cur_code = code

            cur_page = page
            cur_y = y

            cur_omschrijving = (
                [omschrijving]
                if omschrijving else []
            )

            cur_extra = (
                [extra_info]
                if extra_info else []
            )

            cur_kenmerk = (
                [kenmerk]
                if kenmerk else []
            )

            cur_eenheid = (
                [eenheid]
                if eenheid else []
            )

            cur_offerte = offerte

            continue

        # =====================================================
        # CONTINUATION
        # =====================================================

        if cur_code is not None:

            if omschrijving:
                cur_omschrijving.append(omschrijving)

            if extra_info:
                cur_extra.append(extra_info)

            if kenmerk:
                cur_kenmerk.append(kenmerk)

            if eenheid:
                cur_eenheid.append(eenheid)

            if offerte and not cur_offerte:
                cur_offerte = offerte

    flush()

    return records, subfilters


# =========================================================
# ASSIGN SUBFILTERS
# =========================================================

def assign_subfilters(records, subfilters):

    for rec in records:

        page = rec["_page"]
        y = rec["_y"]

        prefix = rec["_prefix"]

        best = ""

        for sf_page, sf_y, sf_prefix, sf_label in subfilters:

            if sf_prefix != prefix:
                continue

            if (sf_page, sf_y) >= (page, y):

                best = sf_label

                break

        rec["subfilter"] = best

        del rec["_page"]
        del rec["_y"]
        del rec["_prefix"]

    return records


# =========================================================
# SAVE EXCEL
# =========================================================

def save_excel(records, output_path):

    df = pd.DataFrame(records)

    df = df[df["code"] != ""]

    df = df.reset_index(drop=True)

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Ruwe_data"
        )

        ws = writer.sheets["Ruwe_data"]

        widths = [
            34,
            40,
            8,
            50,
            60,
            30,
            14,
            10,
        ]

        for i, width in enumerate(widths, start=1):

            ws.column_dimensions[
                get_column_letter(i)
            ].width = width

        ws.freeze_panes = "A2"

    print(f"✅ Saved: {output_path}")


# =========================================================
# MAIN PARSER
# =========================================================

def parse_pdf(pdf_path):

    pdf_path = Path(pdf_path)

    print(f"📄 {pdf_path.name}")

    spans = extract_spans(pdf_path)

    print(f"   spans: {len(spans):,}")

    lines = group_lines(spans)

    print(f"   lines: {len(lines):,}")

    records, subfilters = parse_lines(lines)

    print(f"   records: {len(records):,}")

    print(f"   subfilters: {len(subfilters):,}")

    records = assign_subfilters(records, subfilters)

    return records


# =========================================================
# MAIN
# =========================================================

def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--input-pdf",
        required=True
    )

    parser.add_argument(
        "--output-xlsx",
        required=True
    )

    args = parser.parse_args()

    records = parse_pdf(args.input_pdf)

    if not records:

        print("⚠️ Geen records gevonden")

        return

    save_excel(
        records,
        args.output_xlsx
    )

    print("🎉 Klaar")


if __name__ == "__main__":
    main()
