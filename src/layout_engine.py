import fitz  # PyMuPDF
import pandas as pd
import argparse
import re
from pathlib import Path

def extract_table_from_page(page):
    """Haalt tekstblokken op en probeert ze in kolommen te plaatsen."""
    blocks = page.get_text("blocks")
    # Sorteer blokken van boven naar beneden, dan van links naar rechts
    blocks.sort(key=lambda b: (b[1], b[0]))
    
    rows = []
    current_row = []
    
    for b in blocks:
        text = b[4].strip()
        if not text:
            continue
            
        # Bepaal de x-positie (links, midden, rechts) om kolommen te herkennen
        x0 = b[0]
        
        # Als het blok heel erg links staat, is het een nieuwe rij (Subfilter)
        # Als het in het midden staat, is het een omschrijving
        # Als het rechts staat, is het een kenmerk/eenheid
        # We splitsen de tekst in losse regels om de enorme lappen tekst te breken
        lines = text.split('\n')
        
        for line in lines:
            if not line.strip():
                continue
            
            # Splits de regel op basis van dubbele spaties of tabs (dit is een gok, maar werkt vaak goed)
            parts = re.split(r'\s{2,}', line)
            
            # Als we een nieuw subfilter vinden (begint met cijfer.cijfer), start een nieuwe rij
            if re.match(r'^\d+\.\d+', parts[0]):
                if current_row:
                    rows.append(current_row)
                current_row = [parts[0]]
                parts = parts[1:]
            
            # Voeg de rest van de delen toe aan de huidige rij
            current_row.extend(parts)
            
    if current_row:
        rows.append(current_row)
        
    return rows

def parse_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    all_data = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_rows = extract_table_from_page(page)
        all_data.extend(page_rows)
        
    # Maak er een DataFrame van en geef de kolommen een naam
    # Dit is een schatting op basis van de PDF; je moet dit controleren!
    max_cols = max(len(r) for r in all_data) if all_data else 0
    columns = ["Subfilter", "Code (%)", "Omschrijving", "Extra info", "Kenmerk", "Eenheid", "Offerte"]
    
    # Vul rijen aan tot het maximum aantal kolommen
    for r in all_data:
        while len(r) < max_cols:
            r.append("")
            
    df = pd.DataFrame(all_data, columns=columns[:max_cols])
    return df

def main():
    parser = argparse.ArgumentParser(description="Zet VLIF PDF om naar ruwe Excel")
    parser.add_argument("--input-pdf", required=True, help="Pad naar de PDF")
    parser.add_argument("--output-xlsx", required=True, help="Pad naar de output Excel")
    args = parser.parse_args()
    
    pdf_path = Path(args.input_pdf)
    output_path = Path(args.output_xlsx)
    
    print(f"📄 {pdf_path.name}")
    
    # Parse de PDF
    df = parse_pdf(pdf_path)
    
    # Sla op als Excel
    df.to_excel(output_path, index=False)
    print(f"✅ {output_path} ({len(df)} rijen)")

if __name__ == "__main__":
    main()
