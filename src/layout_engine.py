"""
layout_engine.py – VLIF4 PDF naar ruwe Excel

Exacte kolomgrenzen (gemeten):
  sf_nr    x=39.5     zone 35–62
  sf_naam  x=75.5     zone 62–197
  code     x=200.5    zone 197–213
  omschr   x=223.8    zone 213–370
  extra    x=372.6    zone 370–597
  kenmerk  x=599.5    zone 597–697
  eenheid  x=698.7    zone 697–770
  offerte  x=771+     zone 770+

NPI (pg 175+):
  code+omschr samen  x=229  zone 225–256  → splitsen op CODE_COMBINED_RE
  omschr vervolg     x=254  zone 256–396
  extra              x=397  zone 396–624
  kenmerk            x=624  zone 624–723
  eenheid            x=723  zone 723–784
  offerte            x=784  zone 784+

Subfilter-toewijzing:
  De subfilter-label staat ALTIJD een paar pt ONDER de eerste code.
  Per record: neem het EERSTE subfilter dat VOLGT op zijn (pagina,y)-positie.
  Subfilter = "nr naam" in één veld, komt in kolom A van elke rij.
"""

import re
import argparse
import collections
from pathlib import Path

import fitz
import pandas as pd
from openpyxl.utils import get_column_letter

Y_TOP = 60
Y_BOT = 520

CODE_RE          = re.compile(r"^(15|30|40|50|100)$")
CODE_COMBINED_RE = re.compile(r"^(15|30|40|50|100)\s+(.+)$")
SF_NR_RE         = re.compile(r"^\d+\.\d+\.\d+$")
HOOFD_NR_RE      = re.compile(r"^\d+\.?\d*$")

SKIP = {"SUBFILTER", "OMSCHRIJVING", "EXTRA_INFO", "KENMERK",
        "EENHEID", "OFFERTE", "%"}


def _col(x0, npi=False):
    if  35 <= x0 <  62: return "sf_nr"
    if  62 <= x0 < 197: return "sf_naam"
    if npi:
        if 197 <= x0 < 225: return "sf_naam"
        if 225 <= x0 < 256: return "code"
        if 256 <= x0 < 396: return "omschrijving"
        if 396 <= x0 < 624: return "extra_info"
        if 624 <= x0 < 723: return "kenmerk"
        if 723 <= x0 < 784: return "eenheid"
        if 784 <= x0:        return "offerte"
    else:
        if 197 <= x0 < 213: return "code"
        if 213 <= x0 < 370: return "omschrijving"
        if 370 <= x0 < 597: return "extra_info"
        if 597 <= x0 < 697: return "kenmerk"
        if 697 <= x0 < 770: return "eenheid"
        if 770 <= x0:        return "offerte"
    return "ignore"


def _extract(pdf_path):
    doc = fitz.open(pdf_path)
    spans = []
    npi = False
    for page_no, page in enumerate(doc, start=1):
        if page_no >= 175:
            npi = True
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
                    if y0 < Y_TOP or y0 > Y_BOT:
                        continue
                    c = _col(x0, npi)
                    if c == "ignore":
                        continue
                    spans.append((page_no, round(y0, 1), round(x0, 1),
                                  c, size >= 14, text))
    spans.sort(key=lambda s: (s[0], s[1], s[2]))
    return spans


def _group(spans):
    """
    Groepeert spans op (page, y±4) tot regeldict.
    Spans in dezelfde kolom op dezelfde regel worden samengevoegd.
    Spans in VERSCHILLENDE kolommen worden NIET gemengd.
    """
    lines = []
    cur = None
    for page, y, x0, col, big, text in spans:
        new = (cur is None or page != cur["page"] or abs(y - cur["y"]) > 4)
        if new:
            if cur:
                lines.append(cur)
            cur = {"page": page, "y": y, "big": big}
        if big:
            cur["big"] = True
        # Zelfde kolom: samenvoegen met spatie
        cur[col] = (cur[col] + " " + text) if col in cur else text
    if cur:
        lines.append(cur)
    return lines


def _parse(lines):
    records    = []
    subfilters = []

    cur_hoofd    = ""
    cur_prefix   = ""
    pending_nr   = ""
    cur_code     = None
    cur_omschr   = []
    cur_extra    = []
    cur_kenmerk  = []
    cur_eenheid  = []
    cur_offerte  = ""
    cur_page     = 0
    cur_y        = 0.0

    def flush():
        nonlocal cur_code, cur_omschr, cur_extra
        nonlocal cur_kenmerk, cur_eenheid, cur_offerte
        if cur_code is None and not cur_omschr:
            return
        records.append({
            "hoofdstuk" : cur_hoofd,
            "prefix"    : cur_prefix,
            "subfilter" : "",
            "code"      : cur_code or "",
            "omschrijving": " ".join(cur_omschr).strip(),
            "extra_info": " ".join(cur_extra).strip(),
            "kenmerk"   : " / ".join(cur_kenmerk) if cur_kenmerk else "",
            "eenheid"   : " / ".join(cur_eenheid) if cur_eenheid else "",
            "offerte"   : cur_offerte,
            "_p": cur_page, "_y": cur_y,
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

        if sf_nr.upper() in SKIP or sf_nm.upper() in SKIP or code in SKIP:
            continue
        if kenmrk == "(leeg)": kenmrk = ""
        if eenhd  == "(leeg)": eenhd  = ""

        # NPI: "100 Naam" in code-veld → splitsen
        cm = CODE_COMBINED_RE.match(code)
        if cm:
            extra_omschr = cm.group(2)
            code = cm.group(1)
            omschr = (extra_omschr + " " + omschr).strip() if omschr else extra_omschr

        # ── Grote kop ────────────────────────────────────────────────────
        if big:
            if sf_nr and not HOOFD_NR_RE.match(sf_nr) and not SF_NR_RE.match(sf_nr):
                sf_nm = (sf_nr + " " + sf_nm).strip()
                sf_nr = ""
            if sf_nr and HOOFD_NR_RE.match(sf_nr) and sf_nm:
                flush()
                cur_hoofd  = f"{sf_nr} {sf_nm}".strip()
                cur_prefix = sf_nr
                pending_nr = ""
            elif sf_nr and HOOFD_NR_RE.match(sf_nr):
                flush()
                pending_nr = sf_nr
            elif sf_nm:
                flush()
                if pending_nr:
                    cur_hoofd  = f"{pending_nr} {sf_nm}".strip()
                    cur_prefix = pending_nr
                    pending_nr = ""
                else:
                    cur_hoofd = (cur_hoofd + " " + sf_nm).strip()
            continue

        # ── Subfilter-label ───────────────────────────────────────────────
        if sf_nr and SF_NR_RE.match(sf_nr):
            sf_h = ".".join(sf_nr.split(".")[:2])
            subfilters.append((page, y, f"{sf_nr} {sf_nm}".strip(), sf_h))
            continue

        # ── Code ─────────────────────────────────────────────────────────
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
    sf_sorted = sorted(subfilters, key=lambda s: (s[0], s[1]))
    for rec in records:
        rp, ry, pfx = rec["_p"], rec["_y"], rec["prefix"]
        if pfx:
            for sp, sy, sf_str, shoofd in sf_sorted:
                if shoofd != pfx:
                    continue
                if (sp, sy) >= (rp, ry):
                    rec["subfilter"] = sf_str
                    break
        del rec["_p"], rec["_y"], rec["prefix"]
    return records


def _save(records, output_path):
    df = pd.DataFrame(records, columns=[
        "hoofdstuk", "subfilter", "code",
        "omschrijving", "extra_info", "kenmerk", "eenheid", "offerte"
    ])
    df = df[df["code"] != ""].reset_index(drop=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ruwe_data")
        ws = writer.sheets["Ruwe_data"]
        for i, w in enumerate([30,30,6,42,55,30,12,9], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
    print(f"✅ {output_path}  ({len(df)} rijen)")


def parse_pdf(pdf_path):
    pdf_path = Path(pdf_path)
    print(f"📄 {pdf_path.name}")
    spans   = _extract(pdf_path)
    print(f"   {len(spans):,} spans")
    lines   = _group(spans)
    print(f"   {len(lines):,} regels")
    records, subfilters = _parse(lines)
    print(f"   {len(subfilters)} subfilters  {len(records)} records")
    records = _assign(records, subfilters)
    cnt = collections.Counter(r["hoofdstuk"] for r in records)
    for h, c in sorted(cnt.items()):
        print(f"   {h:55s}  {c:3d} rijen")
    return records


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--input-pdf",   required=True)
    p.add_argument("--output-xlsx", required=True)
    a = p.parse_args()
    recs = parse_pdf(Path(a.input_pdf))
    if not recs:
        print("⚠️  Geen records")
        return
    _save(recs, Path(a.output_xlsx))
    print("🎉 Klaar")


if __name__ == "__main__":
    main()
