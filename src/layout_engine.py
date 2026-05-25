"""
layout_engine.py – VLIF4 PDF naar ruwe Excel

Ondersteunt meerdere blokken/versies automatisch:
- Detecteert NPI-startpagina via kolomhoofd-rij (SUBFILTER % OMSCHRIJVING...)
- Detecteert NPI kolomposities uit die kolomhoofd-rij
- Werkt voor blok 01-2026, blok 02-2026 en toekomstige versies

Kolomindeling blok 2&3 (gemeten):
  sf_nr    x=39.5    zone 35–62
  sf_naam  x=75.5    zone 62–197
  code     x=200.5   zone 197–213
  omschr   x=223.8   zone 213–370
  extra    x=372.6   zone 370–597
  kenmerk  x=599.5   zone 597–697
  eenheid  x=698.7   zone 697–770
  offerte  x=769.7   zone 770+

NPI kolomindeling: automatisch gedetecteerd uit PDF.
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


def _extract_titles(pdf_path):
    """Leest de 3 titelregels van pagina 1: VLIF CODELIJST, blokperiode, versie."""
    doc = fitz.open(pdf_path)
    page = doc[0]
    spans = []
    for block in page.get_text('dict')['blocks']:
        if block.get('type') != 0: continue
        for line in block['lines']:
            for span in line['spans']:
                txt = span['text'].strip()
                if txt:
                    spans.append((round(span['bbox'][1],1), span['size'], txt))
    spans.sort()
    
    title1 = title2 = title3 = ""
    for y, sz, t in spans:
        if sz >= 20 and not title1:
            title1 = t          # bv. "VLIF CODELIJST"
        elif sz >= 16 and not title2 and title1:
            title2 = t          # bv. "Geldig voor blokperiode..."
        elif sz <= 12 and not title3 and title2:
            if t.lower().startswith('versie'):
                title3 = t      # bv. "Versie 20/01/2026"
        if title1 and title2 and title3:
            break
    return title1, title2, title3


def _detect_columns(pdf_path):
    """
    Detecteert kolomgrenzen automatisch uit de kolomhoofd-rij.
    Geeft (normal_bounds, npi_start, npi_bounds) terug.
    """
    doc = fitz.open(pdf_path)
    normal_bounds = None
    npi_start     = None
    npi_bounds    = None
    has_100_seen  = False

    for page_no, page in enumerate(doc, start=1):
        cols = {}
        has_100_on_page = False

        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                for span in line["spans"]:
                    txt = span["text"].strip()
                    x0  = round(span["bbox"][0], 1)
                    y0  = span["bbox"][1]
                    if not (60 < y0 < 520): continue

                    for key in ("SUBFILTER", "EXTRA_INFO", "KENMERK",
                                "EENHEID", "OFFERTE"):
                        if txt == key:
                            cols[key] = x0
                    if txt == "%":
                        cols["%"] = x0
                    if "% OMSCHRIJVING" in txt:
                        cols["%"]          = x0
                        cols["OMSCHRIJVING"] = x0 + 20
                    elif txt == "OMSCHRIJVING":
                        cols["OMSCHRIJVING"] = x0
                    if (txt == "100" or txt.startswith("100 ")) and 60 < y0 < 520:
                        has_100_on_page = True

        has_all = all(k in cols for k in
                      ("SUBFILTER", "EXTRA_INFO", "KENMERK", "EENHEID"))

        if has_all and normal_bounds is None and not has_100_on_page:
            # Normale sectie (blok 2 & 3)
            subf_x  = cols["SUBFILTER"]
            pct_x   = cols.get("%", cols.get("OMSCHRIJVING", 200) - 30)
            omschr_x = cols.get("OMSCHRIJVING", pct_x + 25)
            extra_x  = cols["EXTRA_INFO"]
            kenm_x   = cols["KENMERK"]
            eenh_x   = cols["EENHEID"]
            offerte_x = cols.get("OFFERTE", eenh_x + 70)

            normal_bounds = {
                "sf_nr_max"  : 62,
                "sf_naam_max": 195,  # vaste grens: codes staan op x≈197-202, namen op x≈62-85
                "code_max"   : pct_x + (omschr_x - pct_x) / 2,
                "omschr_max" : omschr_x + (extra_x - omschr_x) / 2,
                "extra_max"  : extra_x  + (kenm_x - extra_x) / 2,
                "kenm_max"   : kenm_x   + (eenh_x - kenm_x) / 2,
                "eenh_max"   : eenh_x   + (offerte_x - eenh_x) / 2,
            }

        if has_all and has_100_on_page and npi_start is None:
            # NPI sectie
            npi_start = page_no
            subf_x   = cols["SUBFILTER"]
            pct_x    = cols.get("%", 196)
            omschr_x = cols.get("OMSCHRIJVING", pct_x + 25)
            extra_x  = cols["EXTRA_INFO"]
            kenm_x   = cols["KENMERK"]
            eenh_x   = cols["EENHEID"]

            npi_bounds = {
                "sf_nr_max"  : 62,
                "sf_naam_max": pct_x - 2,
                "code_max"   : pct_x + (omschr_x - pct_x) / 2,
                "omschr_max" : omschr_x + (extra_x - omschr_x) / 2,
                "kenm_max"   : extra_x  + (kenm_x - extra_x) / 2,
                "eenh_max"   : kenm_x   + (eenh_x - kenm_x) / 2,
                "_no_offerte": True,
            }

        # Stop als beide gevonden, of aan het einde
        if normal_bounds and npi_start:
            break

    return normal_bounds, npi_start, npi_bounds


def _detect_npi(pdf_path):
    """
    Detecteert de NPI-startpagina en kolomgrenzen automatisch
    door te zoeken naar de kolomhoofd-rij met SUBFILTER + % + OMSCHRIJVING.
    Geeft (npi_start_page, npi_col_bounds) terug.
    npi_col_bounds = dict met x-grenzen per kolom.
    """
    doc = fitz.open(pdf_path)
    npi_start = None
    npi_bounds = None

    for page_no, page in enumerate(doc, start=1):
        col_positions = {}
        has_100_on_page = False
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                for span in line["spans"]:
                    txt = span["text"].strip()
                    x0 = round(span["bbox"][0], 1)
                    y0 = span["bbox"][1]
                    if txt in ("SUBFILTER", "%", "OMSCHRIJVING", "EXTRA_INFO",
                               "KENMERK", "EENHEID") and 60 < y0 < 520:
                        col_positions[txt] = x0
                    if (txt == "100" or txt.startswith("100 ")) and 60 < y0 < 520:
                        has_100_on_page = True

        # Als we alle 6 kolomkoppen op één pagina vinden → NPI-pagina
        if all(k in col_positions for k in
               ("SUBFILTER", "%", "OMSCHRIJVING", "EXTRA_INFO", "KENMERK", "EENHEID")):
            # Controleer: is dit de NPI pagina?
            # NPI is de enige sectie met code "100" → check op dezelfde pagina
            pct_x = col_positions["%"]
            if has_100_on_page:  # NPI heeft code 100, andere secties niet
                npi_start = page_no
                # Bouw kolomgrenzen
                subf_x  = col_positions["SUBFILTER"]
                pct_x   = col_positions["%"]
                omschr_x = col_positions["OMSCHRIJVING"]
                extra_x  = col_positions["EXTRA_INFO"]
                kenm_x   = col_positions["KENMERK"]
                eenh_x   = col_positions["EENHEID"]

                # Grenzen = midden tussen twee kolommen
                npi_bounds = {
                    # sf_nr: alleen het nummer zelf (x ≈ 39-41), max tot x=62
                    "sf_nr_max"  : 62,
                    # sf_naam: namen staan op x ≈ 68-85, tot net voor code-zone
                    "sf_naam_max": pct_x - 2,
                    # code: pct_x tot net voor omschr_x ("100 Naam" op x ≈ 198)
                    # code_max: halverwege pct_x en omschr_x
                    # zodat omschrijving (x≈224) in omschrijving-zone valt
                    "code_max"   : pct_x + (omschr_x - pct_x) / 2,
                    # omschrijving: vanaf omschr_x tot halverwege extra
                    "omschr_max" : omschr_x + (extra_x - omschr_x) / 2,
                    # extra_info: tot halverwege extra en kenmerk
                    "kenm_max"   : extra_x  + (kenm_x - extra_x) / 2,
                    # eenheid: van kenmerk tot rechtermarge (geen offerte in NPI)
                    "eenh_max"   : kenm_x + (eenh_x - kenm_x) / 2,
                    # offerte: niet aanwezig in NPI → alles boven eenh_x is eenheid
                    "_no_offerte": True
                }
                break

    return npi_start, npi_bounds


def _col_normal(x0, b):
    """Kolomindeling blok 2 & 3, dynamisch op basis van gedetecteerde grenzen."""
    # sf_nr: alleen echte nummers (x ≈ 36-41), vaste grens op 55
    if x0 < 55:               return "sf_nr"
    if x0 < b["sf_naam_max"]: return "sf_naam"
    if x0 < b["code_max"]:    return "code"
    if x0 < b["omschr_max"]:  return "omschrijving"
    if x0 < b["extra_max"]:   return "extra_info"
    if x0 < b["kenm_max"]:    return "kenmerk"
    if x0 < b["eenh_max"]:    return "eenheid"
    return "offerte"


def _col_npi(x0, bounds):
    """Kolomindeling NPI (automatisch gedetecteerd)."""
    if x0 < bounds["sf_nr_max"]:   return "sf_nr"
    if x0 < bounds["sf_naam_max"]: return "sf_naam"
    if x0 < bounds["code_max"]:    return "code"
    if x0 < bounds["omschr_max"]:  return "omschrijving"
    if x0 < bounds["kenm_max"]:    return "extra_info"
    if x0 < bounds["eenh_max"]:    return "eenheid"
    return "offerte"


def _extract(pdf_path, normal_bounds, npi_start, npi_bounds):
    doc = fitz.open(pdf_path)
    spans = []
    for page_no, page in enumerate(doc, start=1):
        npi = (npi_start is not None and page_no >= npi_start)
        for block in page.get_text("dict",
                flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = " ".join(span["text"].split())
                    if not text: continue
                    x0   = span["bbox"][0]
                    y0   = span["bbox"][1]
                    size = span["size"]
                    if y0 < Y_TOP or y0 > Y_BOT: continue
                    if npi:
                        c = _col_npi(x0, npi_bounds)
                    else:
                        c = _col_normal(x0, normal_bounds)
                    if c == "ignore": continue
                    spans.append((page_no, round(y0, 1), round(x0, 1),
                                  c, size >= 14, text))
    spans.sort(key=lambda s: (s[0], s[1], s[2]))
    return spans


def _group(spans):
    lines = []
    cur = None
    for page, y, x0, col, big, text in spans:
        new = (cur is None or page != cur["page"] or abs(y - cur["y"]) > 4)
        if new:
            if cur: lines.append(cur)
            cur = {"page": page, "y": y, "big": big}
        if big: cur["big"] = True
        cur[col] = (cur[col] + " " + text) if col in cur else text
    if cur: lines.append(cur)
    return lines


def _parse(lines):
    records    = []
    subfilters = []

    cur_hoofd   = ""
    cur_prefix  = ""
    pending_nr  = ""
    last_sf_idx = None
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
        if cur_code is None and not cur_omschr: return
        records.append({
            "hoofdstuk"   : cur_hoofd,
            "prefix"      : cur_prefix,
            "subfilter"   : "",
            "code"        : cur_code or "",
            "omschrijving": " ".join(cur_omschr).strip(),
            "extra_info"  : " ".join(cur_extra).strip(),
            "kenmerk"     : " / ".join(cur_kenmerk) if cur_kenmerk else "",
            "eenheid"     : " / ".join(cur_eenheid) if cur_eenheid else "",
            "offerte"     : cur_offerte,
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

        # NPI: "100 Naam" gecombineerd in code-veld → splitsen
        cm = CODE_COMBINED_RE.match(code)
        if cm:
            omschr = (cm.group(2) + " " + omschr).strip()
            code   = cm.group(1)

        # Code-zone tekst die geen geldige code is → omschrijving-vervolg
        if code and not CODE_RE.match(code) and not CODE_COMBINED_RE.match(code):
            if cur_code is not None:
                cur_omschr.append(code)
            continue

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
                pending_nr  = sf_nr
                cur_hoofd   = ""
                cur_prefix  = ""
            elif sf_nm:
                flush()
                if pending_nr:
                    # Pending nummer + naam = nieuw hoofdstuk
                    cur_hoofd  = f"{pending_nr} {sf_nm}".strip()
                    cur_prefix = pending_nr
                    pending_nr = ""
                elif not cur_prefix:
                    # Naam zonder nummer en zonder vorig hoofdstuk → nieuw
                    cur_hoofd = sf_nm
                else:
                    # Naam loopt door over meerdere regels (bv. "ROERENDE AEA-TECHNIEKEN" + "IN BESTAANDE STALLEN")
                    cur_hoofd = (cur_hoofd + " " + sf_nm).strip()
            continue

        # ── Subfilter-label ───────────────────────────────────────────────
        if sf_nr and SF_NR_RE.match(sf_nr):
            sf_h = ".".join(sf_nr.split(".")[:2])
            subfilters.append((page, y, f"{sf_nr} {sf_nm}".strip(), sf_h))
            last_sf_idx = len(subfilters) - 1
            # Vervolgtekst die samen gegroepeerd werd
            if cur_code is not None:
                if omschr:  cur_omschr.append(omschr)
                if extra:   cur_extra.append(extra)
                if kenmrk:  cur_kenmerk.append(kenmrk)
                if eenhd:   cur_eenheid.append(eenhd)
                if offrte and not cur_offerte: cur_offerte = offrte
            continue
        
        # ── Vervolg subfilter-naam over meerdere regels ───────────────────
        # Bv. "Voeder- en" / "drinkwater-" / "voorziening" op 3 regels
        if (sf_nm and not sf_nr and not code and not big
                and last_sf_idx is not None):
            pg, sy, sf_str, sf_h = subfilters[last_sf_idx]
            # Voeg samen: afbreekstreepje = woord loopt door, anders spatie
            if sf_str.endswith('-'):
                sf_str_clean = sf_str[:-1]
                subfilters[last_sf_idx] = (pg, sy, sf_str_clean + sf_nm, sf_h)
            else:
                subfilters[last_sf_idx] = (pg, sy, sf_str + " " + sf_nm, sf_h)
            # Verwerk ook andere velden uit dezelfde lijn
            if cur_code is not None:
                if omschr:  cur_omschr.append(omschr)
                if extra:   cur_extra.append(extra)
                if kenmrk:  cur_kenmerk.append(kenmrk)
                if eenhd:   cur_eenheid.append(eenhd)
                if offrte and not cur_offerte: cur_offerte = offrte
            continue

        # ── Nieuwe code ───────────────────────────────────────────────────
        if code and CODE_RE.match(code):
            flush()
            last_sf_idx = None  # subfilter-naam is volledig
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
            if offrte and not cur_offerte: cur_offerte = offrte

    flush()
    return records, subfilters


def _assign(records, subfilters):
    sf_sorted = sorted(subfilters, key=lambda s: (s[0], s[1]))
    from collections import defaultdict
    sf_per = defaultdict(list)
    for sp, sy, sf_str, sh in sf_sorted:
        sf_per[sh].append((sp, sy, sf_str))

    for rec in records:
        rp, ry, pfx = rec["_p"], rec["_y"], rec["prefix"]
        if pfx and pfx in sf_per:
            sfs = sf_per[pfx]
            next_sf = None
            prev_sf = None
            for sp, sy, sf_str in sfs:
                if (sp, sy) > (rp, ry):
                    next_sf = (sp, sy, sf_str)
                    break
            for sp, sy, sf_str in reversed(sfs):
                if (sp, sy) <= (rp, ry):
                    prev_sf = (sp, sy, sf_str)
                    break
            same_page_close = (next_sf and next_sf[0] == rp
                               and (next_sf[1] - ry) <= 20)
            if same_page_close:
                rec["subfilter"] = next_sf[2]
            elif prev_sf:
                rec["subfilter"] = prev_sf[2]
            elif next_sf:
                rec["subfilter"] = next_sf[2]
        del rec["_p"], rec["_y"], rec["prefix"]
    return records


def _save(records, output_path, titles=None):
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
        # Sla titels op als metadata tabblad
        if titles:
            t1, t2, t3 = titles
            meta_df = pd.DataFrame({"titel": [t1, t2, t3]})
            meta_df.to_excel(writer, index=False, sheet_name="_meta")
    print(f"✅ {output_path}  ({len(df)} rijen)")


def parse_pdf(pdf_path):
    pdf_path = Path(pdf_path)
    print(f"📄 {pdf_path.name}")
    titles = _extract_titles(pdf_path)
    print(f"   Titel: {titles[0]} | {titles[1]}")
    print("   Kolomgrenzen detecteren...")
    normal_bounds, npi_start, npi_bounds = _detect_columns(pdf_path)
    if normal_bounds:
        print(f"   Kolommen: code_max={normal_bounds['code_max']:.0f} "
              f"omschr_max={normal_bounds['omschr_max']:.0f} "
              f"kenm_max={normal_bounds['kenm_max']:.0f} "
              f"eenh_max={normal_bounds['eenh_max']:.0f}")
    if npi_start:
        print(f"   NPI start pagina: {npi_start}")
    else:
        print("   Geen NPI sectie gevonden")
    spans   = _extract(pdf_path, normal_bounds, npi_start, npi_bounds)
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
    titles = _extract_titles(Path(a.input_pdf))
    _save(recs, Path(a.output_xlsx), titles=titles)
    print("🎉 Klaar")


if __name__ == "__main__":
    main()
