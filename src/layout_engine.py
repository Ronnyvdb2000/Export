"""
layout_engine.py – VLIF4 PDF → Excel (PyMuPDF)

Gemeten kolomgrenzen (liggend A4, 841 × 595 pt):
  x  35 –  80  → nummering  (subfilter of hoofdstuk)
  x  68 – 196  → naam       (hoofdstuk bij size≥14, anders subfilter-naam)
  x 196 – 216  → code       (15 / 30 / 40)
  x 216 – 371  → omschrijving (titel maatregel)
  x 371 – 596  → extra_info
  x 596 – 696  → kenmerk
  x 696 – 771  → eenheid
  x 771 – 841  → offerte

Hoofdstuk vs subfilter:
  - font size ≥ 14 → hoofdstuk (bv. "2.2 VARKENSSTAL")
  - font size 11–13 → subfilter (bv. "2.2.1 Isolatie")
"""

import re
import argparse
from pathlib import Path

import fitz
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# x-grenzen (punten)
X_NR         = (35,   80)
X_NAAM       = (68,  216)
X_CODE       = (196, 216)
X_OMSCHR     = (216, 371)
X_EXTRA      = (371, 596)
X_KENMERK    = (596, 696)
X_EENHEID    = (696, 771)
X_OFFERTE    = (771, 842)

# y-zones voor header/footer
Y_HEADER_MAX = 60
Y_FOOTER_MIN = 520

CODE_RE          = re.compile(r"^(15|30|40)$")
SUBFILTER_NR_RE  = re.compile(r"^\d+\.\d+\.\d+$")   # bv. 2.1.3
HOOFD_NR_RE      = re.compile(r"^\d+\.?\d*$")        # bv. 2 of 2.1
PAGINA_NR_RE     = re.compile(r"^\d{1,3}$")          # paginanummer in TOC


def _classify(x0, size):
    """Geeft (kolom, is_groot) terug; is_groot = font-size ≥ 14."""
    big = size >= 14
    if X_NR[0] <= x0 < X_NR[1]:
        return ("nr", big)
    # naam: hoofdstuk gebruikt x≈68, subfilter x≈75; beide vallen in X_NAAM
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
                    x0   = span["bbox"][0]
                    y0   = span["bbox"][1]
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
    """Groepeert spans op dezelfde (page, y±3) tot regeldict."""
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
    cur_omschr = cur_extra = cur_kenmerk = cur_eenheid = []
    cur_offerte = ""

    def reset_record():
        nonlocal cur_code, cur_omschr, cur_extra, cur_kenmerk, cur_eenheid, cur_offerte
        cur_code = None
        cur_omschr = []; cur_extra = []; cur_kenmerk = []; cur_eenheid = []
        cur_offerte = ""

    def flush():
        if cur_code is None and not cur_omschr:
            return
        records.append({
            "hoofdstuk"     : cur_hoofdstuk,
            "subfilter_nr"  : cur_sf_nr,
            "subfilter_naam": cur_sf_naam,
            "code"          : cur_code or "",
            "omschrijving"  : " ".join(cur_omschr).strip(),
            "extra_info"    : " ".join(cur_extra).strip(),
            "kenmerk"       : " / ".join(cur_kenmerk) if cur_kenmerk else "",
            "eenheid"       : " / ".join(cur_eenheid) if cur_eenheid else "",
            "offerte"       : cur_offerte,
        })
        reset_record()

    for line in lines:
        nr     = line.get("nr",          "").strip()
        naam   = line.get("naam",        "").strip()
        code   = line.get("code",        "").strip()
        omschr = line.get("omschrijving","").strip()
        extra  = line.get("extra_info",  "").strip()
        kenmrk = line.get("kenmerk",     "").strip()
        eenhd  = line.get("eenheid",     "").strip()
        offrte = line.get("offerte",     "").strip()
        big    = line.get("big", False)

        # ── Kolomhoofd-rij overslaan ("SUBFILTER", "%", …) ───────────────
        if naam in ("SUBFILTER", "OMSCHRIJVING", "EXTRA_INFO", "KENMERK",
                    "EENHEID", "OFFERTE", "PRODUCTIEVE INVESTERINGSSTEUN",
                    "AMMONIAK-EMISSIEREDUCERENDE INVESTERINGSSTEUN"):
            continue
        if code == "%":
            continue

        # ── Hoofdstuk (font-size ≥ 14, bv. "2.1 RUNDVEESTAL") ───────────
        if big and nr and HOOFD_NR_RE.match(nr):
            flush()
            cur_hoofdstuk = f"{nr} {naam}".strip()
            cur_sf_nr = cur_sf_naam = ""
            reset_record()
            continue

        # ── Sectietitel zonder nr maar wel big (bv. "2 PRODUCTIEVE…") ───
        if big and not nr and naam:
            flush()
            cur_hoofdstuk = naam
            cur_sf_nr = cur_sf_naam = ""
            reset_record()
            continue

        # ── Subfilter (bv. "2.1.3 Isolatie") ────────────────────────────
        if nr and SUBFILTER_NR_RE.match(nr):
            flush()
            cur_sf_nr   = nr
            cur_sf_naam = naam
            # soms staat code op dezelfde regel
            if code and CODE_RE.match(code):
                cur_code = code
            if omschr:  cur_omschr  = [omschr]
            if extra:   cur_extra   = [extra]
            if kenmrk:  cur_kenmerk = [kenmrk]
            if eenhd:   cur_eenheid = [eenhd]
            if offrte:  cur_offerte = offrte
            continue

        # ── Nieuwe code (15 / 30 / 40) ───────────────────────────────────
        if code and CODE_RE.match(code):
            flush()
            cur_code    = code
            cur_omschr  = [omschr]  if omschr  else []
            cur_extra   = [extra]   if extra   else []
            cur_kenmerk = [kenmrk]  if kenmrk  else []
            cur_eenheid = [eenhd]   if eenhd   else []
            cur_offerte = offrte
            continue

        # ── Vervolgregel van huidig record ───────────────────────────────
        if cur_code is not None:
            if omschr:  cur_omschr.append(omschr)
            if extra:   cur_extra.append(extra)
            if kenmrk:  cur_kenmerk.append(kenmrk)
            if eenhd:   cur_eenheid.append(eenhd)
            if offrte and not cur_offerte:
                cur_offerte = offrte
        elif naam and not nr and not code:
            # Subfilternaam loopt over meerdere regels
            cur_sf_naam = (cur_sf_naam + " " + naam).strip()

    flush()
    return records


def _save_excel(records, output_path):
    df = pd.DataFrame(records, columns=[
        "hoofdstuk","subfilter_nr","subfilter_naam","code",
        "omschrijving","extra_info","kenmerk","eenheid","offerte"
    ])
    # Forward-fill hiërarchie
    for col in ("hoofdstuk","subfilter_nr","subfilter_naam"):
        df[col] = df[col].replace("", pd.NA).ffill().fillna("")

    # Verwijder rijen zonder code
    df = df[df["code"] != ""].reset_index(drop=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Codelijst")
        ws = writer.sheets["Codelijst"]

        # Kolombreedtes A–I
        for i, w in enumerate([30,14,22,6,40,55,30,12,8], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # Header
        hdr_fill = PatternFill("solid", fgColor="1F4E79")
        hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        for cell in ws[1]:
            cell.fill   = hdr_fill
            cell.font   = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)

        # Data
        thin   = Side(style="thin", color="D9D9D9")
        brd    = Border(bottom=thin)
        dfont  = Font(name="Arial", size=9)
        wrap_c = {5, 6, 7}   # omschrijving, extra_info, kenmerk

        for ri, row in enumerate(ws.iter_rows(min_row=2), start=2):
            fill = PatternFill("solid", fgColor="EBF0F7" if ri % 2 == 0 else "FFFFFF")
            for cell in row:
                cell.font   = dfont
                cell.fill   = fill
                cell.border = brd
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(cell.column in wrap_c))
        ws.freeze_panes = "A2"

    print(f"✅ {output_path}  ({len(df)} rijen)")


def parse_pdf(pdf_path):
    pdf_path = Path(pdf_path)
    print(f"📄 {pdf_path.name}")
    spans = _extract_spans(pdf_path)
    print(f"   {len(spans):,} spans")
    lines = _group_lines(spans)
    print(f"   {len(lines):,} regels")
    records = _build_records(lines)
    print(f"   {len(records):,} records")
    return records


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-pdf",   required=True)
    parser.add_argument("--output-xlsx", required=True)
    args = parser.parse_args()
    records = parse_pdf(Path(args.input_pdf))
    if not records:
        print("⚠️  Geen records gevonden")
        return
    _save_excel(records, Path(args.output_xlsx))
    print("🎉 Klaar")


if __name__ == "__main__":
    main()
