#!/usr/bin/env python3
"""Converter voor PDF, ODS en ODT bestanden naar Excel"""

import pandas as pd
from pathlib import Path
import argparse
import sys

def read_pdf_simple(pdf_path):
    """Lees PDF bestand - eenvoudige extractie"""
    try:
        # Probeer tabula-py eerst (werkt goed voor tabellen)
        import tabula
        dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        if dfs:
            return {f"PDF_Tabel_{i+1}": df for i, df in enumerate(dfs) if not df.empty}
    except:
        pass
    
    try:
        # Fallback naar pdfplumber
        import pdfplumber
        sheets = {}
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                for table_num, table in enumerate(tables, 1):
                    if table and len(table) > 1:
                        df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                        if not df.empty:
                            sheets[f"Pagina_{page_num}_Tabel_{table_num}"] = df
        return sheets
    except:
        return {}

def read_ods_simple(ods_path):
    """Lees ODS bestand"""
    try:
        return pd.read_excel(ods_path, sheet_name=None, engine='odf')
    except:
        return {}

def convert_file_to_excel(input_file, output_file):
    """Converteer elk bestand naar Excel"""
    try:
        print(f"📄 Converteren: {input_file.name}")
        
        sheets = {}
        
        if input_file.suffix.lower() == '.pdf':
            sheets = read_pdf_simple(input_file)
            print(f"   📑 Gevonden: {len(sheets)} tabellen uit PDF")
        elif input_file.suffix.lower() == '.ods':
            sheets = read_ods_simple(input_file)
            print(f"   📑 Gevonden: {len(sheets)} sheets uit ODS")
        else:
            print(f"⚠️ Niet-ondersteund bestandstype: {input_file.suffix}")
            return False
        
        if not sheets:
            print("   ⚠️ Geen data gevonden, maak leeg Excel bestand")
            df_empty = pd.DataFrame({'Melding': ['Geen tabellen gevonden in bestand']})
            sheets = {'Info': df_empty}
        
        # Schrijf naar Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                clean_name = sheet_name.replace('/', '_').replace('\\', '_')[:31]
                # Limiteer DataFrame grootte voor performance
                if len(df) > 10000:
                    df = df.head(10000)
                df.to_excel(writer, sheet_name=clean_name, index=False)
        
        print(f"   ✅ Succes: {len(sheets)} sheets geconverteerd")
        return True
        
    except Exception as e:
        print(f"   ❌ Fout: {e}")
        return False

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input-dir', default='input')
    parser.add_argument('--output-dir', default='output')
    args = parser.parse_args()
    
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    
    # Maak output directory aan
    output_dir.mkdir(exist_ok=True)
    
    # Zoek alle ondersteunde bestanden
    supported_files = []
    for ext in ['*.pdf', '*.ods', '*.odt']:
        supported_files.extend(input_dir.glob(ext))
    
    if not supported_files:
        print(f"❌ Geen .pdf, .ods of .odt bestanden gevonden in {input_dir}")
        print(f"📁 Bestanden in input: {list(input_dir.iterdir())}")
        sys.exit(1)
    
    print(f"📁 Gevonden bestanden: {[f.name for f in supported_files]}")
    
    # Converteer elk bestand
    success_count = 0
    for input_file in supported_files:
        output_file = output_dir / f"{input_file.stem}.xlsx"
        if convert_file_to_excel(input_file, output_file):
            success_count += 1
    
    print(f"\n📊 Resultaat: {success_count}/{len(supported_files)} bestanden geconverteerd")
    
    if success_count == 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
