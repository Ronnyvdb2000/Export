"""
layout_engine.py
VLIF4 PDF -> ruwe Excel
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
Y_FOOTER_MIN = 520   # ⭐ FIX: 540 → 520


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

        nonlocal cur_code, cur_omschrijving, cur_extra, cur_kenmerk, cur_eenheid, cur_offerte

        if cur_code is None:
            return

        records.append({
            "hoofdstuk": current_hoofdstuk,
            "subfilter": "",
            "code": cur_code,
