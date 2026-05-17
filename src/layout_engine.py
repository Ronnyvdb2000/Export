"""
layout_engine.py – VLIF4 PDF naar ruwe Excel

Opbouw uit de PDF (pagina 9 als voorbeeld):
  y=306  sf_nr=40 "2.1"          big=True   → start hoofdstuk
  y=309  sf_naam  "RUNDVEESTAL"  big=True   → naam van het hoofdstuk
  y=339  code=30  omschr="Ijsbatterij"      → eerste record (nog zonder subfilter)
  y=349  sf_nr="2.1.1" sf_naam="Bewaren en koelen" → subfilter voor bovenstaande code
  y=443  code=30  omschr="Selectiepoort..."
  y=453  sf_nr="2.1.2" sf_naam="Gestuurd koeverkeer"
  ...

Aanpak:
  1. Parse alle lijnen en sla records op met hun (page, y)
  2. Verzamel subfilters apart met (page, y, nr, naam, hoofd_prefix)
  3. Wijs aan elke record het EERSTE subfilter toe dat >= zijn (page, y)
     binnen hetzelfde hoofdstuk (op basis van hoofd_prefix = "2.1" etc.)
"""

import re
import argparse
from pathlib import Path

import fitz
import pandas as pd
from openpyxl.utils import get_column_letter

# x-grenzen (punten)
X_SF_NR   = (35,   65)   # subfilter/hoofdstuk nummer  (bv. "2.1", "2.1.1")
X_SF_NAAM = (65,  196)   # subfilter/hoofdstuk naam    (bv. "RUNDVEESTAL", "Bewaren en koelen")
X_CODE    = (196, 216)
X_OMSCHR  = (216, 371)
X_EXTRA   = (371, 596)
X_KENMERK = (596, 696)
X_EENHEID = (696, 771)
X_OFFERTE = (771, 842)

Y_HEADER_MAX = 60
Y_FOOTER_MIN = 520

CODE_RE         = re.compile(r"^(15|30|40)$")
SUBFILTER_NR_RE = re.compile(r"^\d+\.\d+\.\d+$")   # bv. 2.1.1
HOOFD_NR_RE     = re.compile(r"^\d+\.\d+$")          # bv. 2.1  (niet 2.1.1)

SKIP = {
    "SUBFILTER", "OMSCHRIJVING", "EXTRA_INFO", "KENMERK", "EENHEID",
    "OFFERTE", "%",
}


def _col(x0):
    if X_SF_NR[0]   <= x0 < X_SF_NR[1]:    return "sf_nr"
    if X_SF_NAAM[0] <= x0 < X_SF_NAAM[1]:   return "sf_naam"
    if X_CODE[0]    <= x0 < X_CODE[1]:      return "code"
    if X_OMSCHR[0]  <= x0 < X_EXTRA[0]:     return "omschrijving"
    if X_EXTRA[0]   <= x0 < X_KENMERK[0]:   return "extra_info"
    if X_KENMERK[0] <= x0 < X_EENHEID[0]:   return "kenmerk"
    if X_EENHEID[0] <= x0 < X_OFFERTE[0]:   return "eenheid"
    if X_OFFERTE[0] <= x0:                   return "offerte"
    return "ignore"


def _extract(pdf_path):
    doc = fitz.open(pdf_path)
    spans = []
    for page_no, page in enumerate(doc, start=1):
        for block in page.get_text("dict",
                flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]:
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
                    c = _col(x0)
                    if c == "ignore":
                        continue
                    spans.append((page_no, round(y0, 1), round(x0, 1),
                                  c, size >= 14, text))
    spans.sort(key=lambda s: (s[0], s[1], s[2]))
    return spans


def _group(spans):
    lines = []
    cur = None
    for page, y, x0, col, big, text in spans:
        new = (cur is None or page != cur["page"] or abs(y - cur["y"]) > 3)
        if new:
            if cur:
                lines.append(cur)
            cur = {"page": page, "y": y, "big": big}
        if big:
            cur["big"] = True
        cur[col] = (cur[col] + " " + text) if col in cur else text
    if cur:
        lines.append(cur)
    return lines


def _parse(lines):
    records    = []
    subfilters = []   # (page, y, nr, naam, hoofd_prefix)

    cur_hoofdstuk    = ""
    cur_hoofd_prefix = ""   # bv. "2.1"
    pending_hoofd_nr = ""   # buffer voor hoofdstuk-nummer dat nog geen naam heeft
    cur_code    = None
    cur_omschr  = []
    cur_extra   = []
    cur_kenmerk = []
    cur_eenheid = []
    cur_offerte = ""
    cur_page    = 0
    cur_y       = 0.0

    def flush():
        nonlocal cur_code, cur_omschr, cur_extra
        nonlocal cur_kenmerk, cur_eenheid, cur_offerte
        if cur_code is None and not cur_omschr:
            return
        records.append({
            "hoofdstuk"     : cur_hoofdstuk,
            "hoofd_prefix"  : cur_hoofd_prefix,
            "subfilter_nr"  : "",
            "subfilter_naam": "",
            "code"          : cur_code or "",
            "omschrijving"  : " ".join(cur_omschr).strip(),
            "extra_info"    : " ".join(cur_extra).strip(),
            "kenmerk"       : " / ".join(cur_kenmerk) if cur_kenmerk else "",
            "eenheid"       : " / ".join(cur_eenheid) if cur_eenheid else "",
            "offerte"       : cur_offerte,
            "_page"         : cur_page,
            "_y"            : cur_y,
        })
        cur_code    = None
        cur_omschr  = []
        cur_extra   = []
        cur_kenmerk = []
        cur_eenheid = []
        cur_offerte = ""

    for line in lines:
        sf_nr  = line.get("sf_nr",        "").strip()
        sf_nm  = line.get("sf_naam",      "").strip()
        code   = line.get("code",         "").strip()
        omschr = line.get("omschrijving", "").strip()
        extra  = line.get("extra_info",   "").strip()
        kenmrk = line.get("kenmerk",      "").strip()
        eenhd  = line.get("eenheid",      "").strip()
        offrte = line.get("offerte",      "").strip()
        big    = line.get("big",          False)
        page   = line["page"]
        y      = line["y"]

        # Sla header-rijen over
        if sf_nr.upper() in SKIP or sf_nm.upper() in SKIP or code in SKIP:
            continue

        # ── Grote tekst (hoofdstuk) ───────────────────────────────────────
        if big:
            # Geval 1: nr + naam op zelfde regel (bv. "2 PRODUCTIEVE...")
            if sf_nr and HOOFD_NR_RE.match(sf_nr) and sf_nm:
                flush()
                cur_hoofdstuk    = f"{sf_nr} {sf_nm}".strip()
                cur_hoofd_prefix = sf_nr
                pending_hoofd_nr = ""
                continue

            # Geval 2: alleen nummer (bv. "2.1") → bewaar voor volgende regel
            if sf_nr and HOOFD_NR_RE.match(sf_nr) and not sf_nm:
                flush()
                pending_hoofd_nr = sf_nr
                continue

            # Geval 3: alleen naam (bv. "RUNDVEESTAL") → combineer met pending
            if sf_nm and not sf_nr:
                flush()
                if pending_hoofd_nr:
                    cur_hoofdstuk    = f"{pending_hoofd_nr} {sf_nm}".strip()
                    cur_hoofd_prefix = pending_hoofd_nr
                    pending_hoofd_nr = ""
                else:
                    cur_hoofdstuk    = sf_nm
                    cur_hoofd_prefix = ""
                continue

        # ── Subfilter-label (bv. "2.1.1" + "Bewaren en koelen") ──────────
        if sf_nr and SUBFILTER_NR_RE.match(sf_nr):
            sf_hoofd = ".".join(sf_nr.split(".")[:2])
            subfilters.append((page, y, sf_nr, sf_nm, sf_hoofd))
            continue   # geen flush — vorige code hoort bij dit subfilter

        # ── Code (15/30/40) ───────────────────────────────────────────────
        if code and CODE_RE.match(code):
            flush()
            cur_code    = code
            cur_page    = page
            cur_y       = y
            cur_omschr  = [omschr] if omschr else []
            cur_extra   = [extra]  if extra  else []
            cur_kenmerk = [kenmrk] if kenmrk else []
            cur_eenheid = [eenhd]  if eenhd  else []
            cur_offerte = offrte
            continue

        # ── Vervolgregel ─────────────────────────────────────────────────
        if cur_code is not None:
            if omschr:  cur_omschr.append(omschr)
            if extra:   cur_extra.append(extra)
            if kenmrk:  cur_kenmerk.append(kenmrk)
            if eenhd:   cur_eenheid.append(eenhd)
            if offrte and not cur_offerte:
                cur_offerte = offrte

    flush()
    return records, subfilters


def _assign(records, subfilters):
    """Wijs aan elke record het EERSTE subfilter toe dat NA zijn positie komt."""
    sf_sorted = sorted(subfilters, key=lambda s: (s[0], s[1]))

    for rec in records:
        rp  = rec["_page"]
        ry  = rec["_y"]
        pfx = rec["hoofd_prefix"]

        if not pfx:
            del rec["_page"]; del rec["_y"]; del rec["hoofd_prefix"]
            continue

        for sp, sy, snr, snaam, shoofd in sf_sorted:
            if shoofd != pfx:
                continue
            if (sp, sy) >= (rp, ry):
                rec["subfilter_nr"]   = snr
                rec["subfilter_naam"] = snaam
                break

        del rec["_page"]
        del rec["_y"]
        del rec["hoofd_prefix"]

    return records


def _save(records, output_path):
    df = pd.DataFrame(records, columns=[
        "hoofdstuk", "subfilter_nr", "subfilter_naam", "code",
        "omschrijving", "extra_info", "kenmerk", "eenheid", "offerte"
    ])
    df = df[df["code"] != ""].reset_index(drop=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ruwe_data")
        ws = writer.sheets["Ruwe_data"]
        for i, w in enumerate([28,14,30,6,40,55,30,12,8], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
    print(f"✅ {output_path}  ({len(df)} rijen)")


def parse_pdf(pdf_path):
    pdf_path = Path(pdf_path)
    print(f"📄 {pdf_path.name}")
    spans = _extract(pdf_path)
    print(f"   {len(spans):,} spans")
    lines = _group(spans)
    print(f"   {len(lines):,} regels")
    records, subfilters = _parse(lines)
    print(f"   {len(subfilters)} subfilters, {len(records)} records")
    records = _assign(records, subfilters)
    return records


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-pdf",   required=True)
    parser.add_argument("--output-xlsx", required=True)
    args = parser.parse_args()
    records = parse_pdf(Path(args.input_pdf))
    if not records:
        print("⚠️  Geen records")
        return
    _save(records, Path(args.output_xlsx))
    print("🎉 Klaar")


if __name__ == "__main__":
    main()
