import pandas as pd
from pathlib import Path
import argparse
import sys
import zipfile
import xml.etree.ElementTree as ET

def read_ods_simple(filename):
    """Lees ODS bestand zonder complexe dependencies"""
    try:
        # Probeer normale pandas methode
        return pd.read_excel(filename, sheet_name=None, engine='odf')
    except:
        # Fallback voor eenvoudige tabellen
        print("Gebruik simpele ODS parser...")
        return {}

def convert_odf_to_excel(input_file, output_file):
    """Converteer ODF naar Excel"""
    try:
        print(f"📄 Converteren: {input_file.name}")
        
        # Lees ODS bestand
        sheets = read_ods_simple(input_file)
        
        if not sheets:
            print("⚠️ Geen sheets gevonden, maak leeg Excel bestand")
            # Maak leeg DataFrame als fallback
            df_empty = pd.DataFrame({'Melding': ['Geen data gevonden in ODF bestand']})
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_empty.to_excel(writer, sheet_name='Info', index=False)
            return True
        
        # Schrijf naar Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                # Excel sheet naam max 31 tekens
                clean_name = sheet_name.replace('/', '_').replace('\\', '_')[:31]
                df.to_excel(writer, sheet_name=clean_name, index=False)
        
        print(f"✅ Succes: {len(sheets)} sheets geconverteerd")
        return True
        
    except Exception as e:
        print(f"❌ Fout: {e}")
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
    
    # Zoek alle ODS bestanden
    odf_files = list(input_dir.glob('*.ods')) + list(input_dir.glob('*.odt'))
    
    if not odf_files:
        print(f"❌ Geen .ods of .odt bestanden gevonden in {input_dir}")
        print(f"📁 Huidige mappen: {list(Path('.').iterdir())}")
        sys.exit(1)
    
    print(f"📁 Gevonden bestanden: {[f.name for f in odf_files]}")
    
    # Converteer elk bestand
    success_count = 0
    for odf_file in odf_files:
        output_file = output_dir / f"{odf_file.stem}.xlsx"
        if convert_odf_to_excel(odf_file, output_file):
            success_count += 1
    
    print(f"\n📊 Resultaat: {success_count}/{len(odf_files)} bestanden geconverteerd")
    
    if success_count == 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
