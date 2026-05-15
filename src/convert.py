import re
import argparse
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd


CODE_REGEX = re.compile(r"^(15|30|40)\s*$")
HOOFDSTUK_REGEX = re.compile(r"^\d+\.\d+\s+.+")
SUBCODE_REGEX = re.compile(r"^\d+\.\d+\.\d+\s+.+")
EENHEID_TOKENS = {"m²", "m³", "m³/u", "kW", "MW", "l", "stuk(s)", "dieren", "m"}


def normalize_text(text: str) -> str:
    return " ".join(text.replace("\xa0", " ").split())


def extract_blocks(pdf_path: Path):
    doc = fitz.open(pdf_path)
    all_blocks = []

    for page_index, page in enumerate(doc, start=1):
        print(f"📄 Layout analyse pagina {page_index}")
        blocks = page.get_text("blocks")  # (x0, y0, x1, y1, text, block_no, ...)
        for b in blocks:
            x0, y0, x1, y1, text, *_ = b
            text = normalize_text(text)
            if text.strip():
                all_blocks.append(
                    {
                        "page": page_index,
                        "x0": x0,
                        "y0": y0,
                        "x1": x1,
                        "y1": y1,
                        "text": text,
                    }
                )

    # sorteer globaal: eerst pagina, dan y, dan x
    all_blocks.sort(key=lambda b: (b["page"], b["y0"], b["x0"]))
    return all_blocks


def is_code(text: str) -> bool:
    return bool(CODE_REGEX.match(text.strip()))


def is_hoofdstuk(text: str) -> bool:
    return bool(HOOFDSTUK_REGEX.match(text.strip()))


def is_subcode(text: str) -> bool:
    return bool(SUBCODE_REGEX.match(text.strip()))


def looks_like_eenheid(text: str) -> bool:
    t = text.strip()
    if len(t.split()) <= 3:
        for token in EENHEID_TOKENS:
            if token in t:
                return True
    return False


def parse_pdf(pdf_path: Path):
    blocks = extract_blocks(pdf_path)

    current_hoofdstuk = ""
    current_subcode = ""
    current_code = None

    current_titel = ""
    current_eenheid_label = ""
    current_eenheid = ""
    current_omschrijving_lines = []

    rows = []

    def flush_current():
        nonlocal current_code, current_titel, current_eenheid_label, current_eenheid, current_omschrijving_lines

        if current_code is None and not current_titel and not current_omschrijving_lines:
            return

        omschrijving = "\n".join(current_omschrijving_lines).strip()

        rows.append(
            {
                "hoofdstuk": current_hoofdstuk,
                "subcode": current_subcode,
                "code": current_code,
                "titel": current_titel.strip(),
                "eenheid_label": current_eenheid_label.strip(),
                "eenheid": current_eenheid.strip(),
                "omschrijving": omschrijving,
            }
        )

        current_code = None
        current_titel = ""
        current_eenheid_label = ""
        current_eenheid = ""
        current_omschrijving_lines = []

    i = 0
    n = len(blocks)

    while i < n:
        b = blocks[i]
        text = b["text"]

        # Hoofdstuk
        if is_hoofdstuk(text):
            flush_current()
            current_hoofdstuk = text
            i += 1
            continue

        # Subcode
        if is_subcode(text):
            flush_current()
            current_subcode = text
            i += 1
            continue

        # Code (15/30/40)
        if is_code(text):
            flush_current()
            current_code = text.strip()

            # Volgende blok(ken) horen vaak bij titel + eenheid
            j = i + 1
            # Titel: eerste niet-lege blok na code
            if j < n:
                titel_text = blocks[j]["text"]
                current_titel = titel_text
                j += 1

            # Daarna eventueel eenheid-label + eenheid
            # We kijken een paar blokken vooruit
            eenheid_label_candidate = ""
            eenheid_candidate = ""
            while j < n:
                t = blocks[j]["text"]
                # Stop als we een nieuwe code/subcode/hoofdstuk tegenkomen
                if is_code(t) or is_subcode(t) or is_hoofdstuk(t):
                    break

                if looks_like_eenheid(t):
                    if not eenheid_candidate:
                        eenheid_candidate = t
                    else:
                        # tweede eenheid → voeg toe aan label
                        eenheid_label_candidate += (" " + t if eenheid_label_candidate else t)
                else:
                    # langere tekst → omschrijving
                    current_omschrijving_lines.append(t)

                j += 1

            current_eenheid_label = eenheid_label_candidate
            current_eenheid = eenheid_candidate

            i = j
            continue

        # Overige tekst: als we al in een code‑context zitten → omschrijving
        if current_code is not None:
            current_omschrijving_lines.append(text)
        else:
            # tekst buiten code‑context → negeren of later gebruiken
            pass

        i += 1

    # flush laatste
    flush_current()

    return rows


def save_excel(rows, output_path: Path):
    df = pd.DataFrame(rows)
    df.to_excel(output_path, index=False)
    print(f"✅ Excel opgeslagen: {output_path}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-pdf", required=True, help="Pad naar PDF bestand")
    parser.add_argument("--output-xlsx", required=True, help="Pad naar output Excel")
    args = parser.parse_args()

    pdf_path = Path(args.input_pdf)
    output_path = Path(args.output_xlsx)

    rows = parse_pdf(pdf_path)

    if not rows:
        print("⚠️ Geen gestructureerde data gevonden")
        return

    save_excel(rows, output_path)
    print("🎉 Conversie voltooid")


if __name__ == "__main__":
    main()
