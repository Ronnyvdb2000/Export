import pymupdf
import pandas as pd
import argparse
from pathlib import Path

def parse_pdf(pdf_path):
    doc = pymupdf.open(pdf_path)
    all_data = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Gebruik de ingebouwde tabel-herkenning van PyMuPDF
        tabs = page.find_tables()
        
        for tab in tabs:
            data = tab.extract()
            if data:
                # Voeg de tabel toe aan de lijst
                for row in data:
                    all_data.append(row)
                        
    if not all_data:
        # Fallback voor pagina's zonder tabellen (bijv. de inleiding)
        text = ""
        for page_num in range(len(doc)):
            text += doc[page_num].get_text() + "\n"
        return pd.DataFrame({"Text": text.splitlines()})
        
    # Bepaal het maximum aantal kolommen dat we tegenkwamen
    max_cols = max(len(row) for row in all_data) if all_data else 0
    
    # Vul rijen aan tot het maximum aantal kolommen (voorkomt crashes)
    for row in all_data:
        while len(row) < max_cols:
            row.append("")
            
    df = pd.DataFrame(all_data)
    return df

def main():
    parser = argparse.ArgumentParser(description="Zet VLIF PDF om naar ruwe Excel")
    parser.add_argument("--input-pdf", required=True, help="Pad naar de PDF")
    parser.add_argument("--output-xlsx", required=True, help="Pad naar de output Excel")
    args = parser.parse_args()
    
    pdf_path = Path(args.input_pdf)
    output_path = Path(args.output_xlsx)
    
    print(f"📄 {pdf_path.name}")
    df = parse_pdf(pdf_path)
    df.to_excel(output_path, index=False)
    print(f"✅ {output_path} ({len(df)} rijen)")

if __name__ == "__main__":
    main()
