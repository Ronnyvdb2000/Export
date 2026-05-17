"""
layout_engine.py – VLIF PDF → ruwe Excel

Verbeteringen:
- Subfilter = één gecombineerd veld
- Correcte koppeling records ↔ subfilters
- Ondersteunt:
    2. Productieve investeringen
    3. Ammoniak-emissiereducerende investeringen
    4. Niet-productieve investeringen
- Robuustere hoofdstukdetectie
- Betere verwerking van meerregelige records
"""

import re
import argparse
from pathlib import Path

import fitz
import pandas as pd
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# Kolomgrenzen (x-posities uit PDF)
# ------------------------------------------------------------

X_SF_NR   = (35, 70)
X_SF_NM   = (70, 210)
X_CODE    = (190, 220)
X_OMSCHR  = (220, 375)
X_EXTRA   = (375, 600)
X_KENMERK = (600, 700)
X_EENHEID = (700, 770)
X_OFFERTE = (770, 850)

Y_HEADER_MAX = 55
Y_FOOTER_MIN = 540

# ------------------------------------------------------------
# Regex
# ------------------------------------------------------------

CODE_RE = re.compile(r"^(15|30|40)$")

SUBFILTER_RE = re.compile(r"^\d+\.\d+\.\d+$")

HOOFDSTUK_RE = re.compile(r"^\d+(\.\d+)?$")

# ------------------------------------------------------------

SKIP = {
    "SUBFILTER",
    "OMSCHRIJVING",
    "EXTRA_INFO",
    "KENMERK",
    "EENHEID",
    "OFFERTE",
    "%",
}

# ------------------------------------------------------------


def detect_col(x0):

    if X_SF_NR[0] <= x0 < X_SF_NR[1]:
        return "sf_nr"

    if X_SF_NM[0] <= x0 < X_SF_NM[1]:
        return "sf_nm"

    if X_CODE[0] <= x0 < X_CODE[1]:
        return "code"

    if X_OMSCHR[0] <= x0 < X_EXTRA[0]:
        return "omschrijving"

    if X_EXTRA[0] <= x0 < X_KENMERK[0]:
        return "extra"

    if X_KENMERK[0] <= x0 < X_EENHEID[0]:
        return "kenmerk"

    if X_EENHEID[0] <= x0 < X_OFFERTE[0]:
        return "eenheid"

    if X_OFFERTE[0] <= x0:
        return "offerte"

    return "ignore"


# ------------------------------------------------------------


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
                    size = span["size"]

                    if y0 < Y_HEADER_MAX:
                        continue

                    if y0 > Y_FOOTER_MIN:
                        continue

                    col = detect_col(x0)

                    if col == "ignore":
                        continue

                    spans.append({
                        "page": page_no,
                        "y": round(y0, 1),
                        "x": round(x0, 1),
                        "col": col,
                        "big": size >= 13.5,
                        "text": text,
                    })

    spans.sort(key=lambda s: (s["page"], s["y"], s["x"]))

    return spans


# ------------------------------------------------------------


def group_lines(spans):

    lines = []

    cur = None

    for s in spans:

        new_line = (
            cur is None
            or s["page"] != cur["page"]
            or abs(s["y"] - cur["y"]) > 3
        )

        if new_line:

            if cur:
                lines.append(cur)

            cur = {
                "page": s["page"],
                "y": s["y"],
                "big": s["big"],
            }

        if s["big"]:
            cur["big"] = True

        col = s["col"]

        if col in cur:
            cur[col] += " " + s["text"]
        else:
            cur[col] = s["text"]

    if cur:
        lines.append(cur)

    return lines


# ------------------------------------------------------------


def parse_lines(lines):

    records = []
    subfilters = []

    current_hoofdstuk = ""

    pending_nr = ""

    cur_code = None
    cur_page = 0
    cur_y = 0

    cur_omschrijving = []
    cur_extra = []
    cur_kenmerk = []
    cur_eenheid = []
    cur_offerte = ""

    # --------------------------------------------------------

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
            "omschrijving": " ".join(cur_omschrijving).strip(),
            "extra_info": " ".join(cur_extra).strip(),
            "kenmerk": " / ".join(cur_kenmerk),
            "eenheid": " / ".join(cur_eenheid),
            "offerte": cur_offerte,
            "_page": cur_page,
            "_y": cur_y,
        })

        cur_code = None
        cur_omschrijving = []
        cur_extra = []
        cur_kenmerk = []
        cur_eenheid = []
        cur_offerte = ""

    # --------------------------------------------------------

    for line in lines:

        sf_nr = line.get("sf_nr", "").strip()
        sf_nm = line.get("sf_nm", "").strip()

        code = line.get("code", "").strip()

        omschr = line.get("omschrijving", "").strip()

        extra = line.get("extra", "").strip()

        kenmerk = line.get("kenmerk", "").strip()

        eenheid = line.get("eenheid", "").strip()

        offerte = line.get("offerte", "").strip()

        big = line.get("big", False)

        page = line["page"]
        y = line["y"]

        # ----------------------------------------------------
        # Headers overslaan
        # ----------------------------------------------------

        if (
            sf_nr.upper() in SKIP
            or sf_nm.upper() in SKIP
            or code in SKIP
        ):
            continue

        # ----------------------------------------------------
        # Hoofdstukken
        # ----------------------------------------------------

        if big:

            # bv "2.1"
            if sf_nr and HOOFDSTUK_RE.match(sf_nr) and not sf_nm:

                flush()

                pending_nr = sf_nr

                continue

            # bv "RUNDVEESTAL"
            if sf_nm and not sf_nr:

                flush()

                if pending_nr:
                    current_hoofdstuk = f"{pending_nr} {sf_nm}".strip()
                    pending_nr = ""
                else:
                    current_hoofdstuk = sf_nm

                continue

            # bv "2.1 RUNDVEESTAL"
            if sf_nr and sf_nm:

                flush()

                current_hoofdstuk = f"{sf_nr} {sf_nm}".strip()

                continue

        # ----------------------------------------------------
        # Subfilter
        # ----------------------------------------------------

        if sf_nr and SUBFILTER_RE.match(sf_nr):

            label = f"{sf_nr} {sf_nm}".strip()

            subfilters.append({
                "page": page,
                "y": y,
                "hoofdstuk": current_hoofdstuk,
                "label": label,
            })

            continue

        # ----------------------------------------------------
        # Nieuwe code
        # ----------------------------------------------------

        if code and CODE_RE.match(code):

            flush()

            cur_code = code

            cur_page = page
            cur_y = y

            cur_omschrijving = [omschr] if omschr else []
            cur_extra = [extra] if extra else []
            cur_kenmerk = [kenmerk] if kenmerk else []
            cur_eenheid = [eenheid] if eenheid else []

            cur_offerte = offerte

            continue

        # ----------------------------------------------------
        # Vervolgregel
        # ----------------------------------------------------

        if cur_code is not None:

            if omschr:
                cur_omschrijving.append(omschr)

            if extra:
                cur_extra.append(extra)

            if kenmerk:
                cur_kenmerk.append(kenmerk)

            if eenheid:
                cur_eenheid.append(eenheid)

            if offerte and not cur_offerte:
                cur_offerte = offerte

    flush()

    return records, subfilters


# ------------------------------------------------------------


def assign_subfilters(records, subfilters):

    sf_sorted = sorted(
        subfilters,
        key=lambda s: (s["page"], s["y"])
    )

    for rec in records:

        rp = rec["_page"]
        ry = rec["_y"]

        hoofdstuk = rec["hoofdstuk"]

        for sf in sf_sorted:

            if sf["hoofdstuk"] != hoofdstuk:
                continue

            if (sf["page"], sf["y"]) >= (rp, ry):

                rec["subfilter"] = sf["label"]

                break

        del rec["_page"]
        del rec["_y"]

    return records


# ------------------------------------------------------------


def save_excel(records, output_path):

    df = pd.DataFrame(records)

    df = df[df["code"] != ""]

    cols = [
        "hoofdstuk",
        "subfilter",
        "code",
        "omschrijving",
        "extra_info",
        "kenmerk",
        "eenheid",
        "offerte",
    ]

    df = df[cols]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Ruwe_data"
        )

        ws = writer.sheets["Ruwe_data"]

        widths = [35, 35, 8, 55, 55, 30, 12, 10]

        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        ws.freeze_panes = "A2"

    print(f"✅ Excel opgeslagen: {output_path}")


# ------------------------------------------------------------


def parse_pdf(pdf_path):

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


# ------------------------------------------------------------


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

    records = parse_pdf(Path(args.input_pdf))

    if not records:
        print("⚠️ Geen records gevonden")
        return

    save_excel(records, Path(args.output_xlsx))

    print("🎉 Klaar")


# ------------------------------------------------------------

if __name__ == "__main__":
    main()
