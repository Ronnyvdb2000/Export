import pandas as pd
import re
import os
from pathlib import Path
import argparse

def extract_tables_from_text(text):
    """Haal tabelstructuur uit vrije tekst"""
    lines = text.split('\n')
    tables = []
    current_table = []
    in_table = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Herken tabelrijen (bevatten getallen of eenheden zoals m², l, stuk(s))
        if re.search(r'\d+\s*(m²|m²K/W|l|stuk\(s\)|%)', line) or re.search(r'^\d+\s+\d', line):
            # Split op meerdere spaties (kolomscheiding)
            parts = re.split(r'\s{2,}', line)
            if len(parts) >= 2:
                current_table.append(parts)
                in_table = True
        elif in_table and line and not re.search(r'^\d', line):
            # Einde van tabel
            if current_table:
                tables.append(current_table)
                current_table = []
            in_table = False
    
    if current_table:
        tables.append(current_table)
    
    return tables

def convert_pdf_to_excel(input_path, output_path):
    """Converteer PDF naar Excel met tekst-extractie"""
    try:
        import pdfplumber
    except ImportError:
        print("Installeer pdfplumber: pip install pdfplumber")
        return False
    
    all_data = []
    
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue
                
            tables = extract_tables_from_text(text)
            
            for table_idx, table_rows in enumerate(tables):
                if table_rows:
                    # Maak DataFrame van gevonden tabel
                    max_cols = max(len(row) for row in table_rows)
                    df_data = []
                    for row in table_rows:
                        padded = row + [''] * (max_cols - len(row))
                        df_data.append(padded)
                    
                    df = pd.DataFrame(df_data)
                    all_data.append((f"Pagina{page_num}_Tabel{table_idx+1}", df))
    
    # Exporteer naar Excel
    if all_data:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in all_data:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)
        print(f"✅ {len(all_data)} tabellen geëxtraheerd naar {output_path}")
        return True
    else:
        # Fallback: exporteer ruwe tekst per pagina
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            with pdfplumber.open(input_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    if text:
                        lines = [line for line in text.split('\n') if line.strip()]
                        df = pd.DataFrame(lines, columns=[f"Pagina {page_num}"])
                        df.to_excel(writer, sheet_name=f"Pagina{page_num}", index=False)
        print(f"⚠️ Geen tabellen gevonden, ruwe tekst geëxporteerd naar {output_path}")
        return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input-dir', default='input')
    parser.add_argument('--output-dir', default='output')
    args = parser.parse_args()
    
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(exist_ok=True)
    
    pdf_files = list(input_dir.glob('*.pdf'))
    print(f"📁 Gevonden bestanden: {[f.name for f in pdf_files]}")
    
    success_count = 0
    for pdf_path in pdf_files:
        print(f"📄 Converteren: {pdf_path.name}")
        excel_path = output_dir / f"{pdf_path.stem}.xlsx"
        
        if convert_pdf_to_excel(pdf_path, excel_path):
            success_count += 1
    
    print(f"\n📊 Resultaat: {success_count}/{len(pdf_files)} bestanden geconverteerd")

if __name__ == "__main__":
    main()
