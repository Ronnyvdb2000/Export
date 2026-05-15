#!/usr/bin/env python3
"""
VLIF ODF naar Excel Converter
Converteert ODF-bestanden (.ods of .odt) naar gestructureerd Excel (.xlsx)
Speciaal ontworpen voor VLIF codelijsten met tabellen
"""

import pandas as pd
import re
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

# ============================================
# 1. ODF BESTANDEN LEZEN
# ============================================

def read_ods(filename):
    """
    Lees ODS bestand (OpenDocument Spreadsheet)
    
    Parameters:
    filename: pad naar .ods bestand
    
    Returns:
    dict met sheetnamen als keys en DataFrames als values
    """
    try:
        # Probeer met pandas (meest eenvoudig)
        ods_file = pd.ExcelFile(filename, engine='odf')
        sheets = {}
        for sheet_name in ods_file.sheet_names:
            df = pd.read_excel(filename, sheet_name=sheet_name, engine='odf')
            sheets[sheet_name] = df
        return sheets
    except Exception as e:
        print(f"Fout bij lezen ODS met pandas: {e}")
        print("Probeer alternatieve methode met manual XML parsing...")
        return read_ods_manual(filename)


def read_ods_manual(filename):
    """
    Lees ODS bestand via handmatige XML parsing
    Werkt wanneer pandas/odf engine faalt
    """
    sheets = {}
    
    with zipfile.ZipFile(filename, 'r') as ods_zip:
        # Lees content.xml (bevat alle tabeldata)
        with ods_zip.open('content.xml') as content_file:
            tree = ET.parse(content_file)
            root = tree.getroot()
            
            # Namespaces in ODS bestanden
            ns = {
                'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
                'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
                'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
            }
            
            # Zoek alle tabellen
            tables = root.findall('.//table:table', ns)
            
            for table_idx, table in enumerate(tables):
                sheet_name = table.get('table:name', f'Sheet{table_idx + 1}')
                rows_data = []
                
                # Zoek alle rijen in de tabel
                rows = table.findall('.//table:table-row', ns)
                
                for row in rows:
                    row_cells = []
                    cells = row.findall('.//table:table-cell', ns)
                    
                    for cell in cells:
                        # Haal de waarde op
                        value_elem = cell.find('.//text:p', ns)
                        if value_elem is not None and value_elem.text:
                            cell_value = value_elem.text
                        else:
                            # Check voor float/int waarden in attributen
                            float_val = cell.get('office:value')
                            if float_val:
                                try:
                                    cell_value = float(float_val)
                                except:
                                    cell_value = float_val
                            else:
                                cell_value = ''
                        
                        row_cells.append(cell_value)
                    
                    # Alleen rijen met data toevoegen
                    if any(cell and str(cell).strip() for cell in row_cells):
                        rows_data.append(row_cells)
                
                if rows_data:
                    # Bepaal maximale kolomlengte
                    max_cols = max(len(row) for row in rows_data)
                    # Vul kortere rijen aan met lege waarden
                    for row in rows_data:
                        while len(row) < max_cols:
                            row.append('')
                    
                    sheets[sheet_name] = pd.DataFrame(rows_data)
    
    return sheets


def read_odt(filename):
    """
    Lees ODT bestand (OpenDocument Text)
    Extraheert tabellen uit tekstdocument
    """
    tables = []
    
    with zipfile.ZipFile(filename, 'r') as odt_zip:
        with odt_zip.open('content.xml') as content_file:
            tree = ET.parse(content_file)
            root = tree.getroot()
            
            ns = {
                'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
                'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
            }
            
            tables_elem = root.findall('.//table:table', ns)
            
            for table_idx, table_elem in enumerate(tables_elem):
                rows_data = []
                rows = table_elem.findall('.//table:table-row', ns)
                
                for row in rows:
                    row_cells = []
                    cells = row.findall('.//table:table-cell', ns)
                    
                    for cell in cells:
                        paras = cell.findall('.//text:p', ns)
                        cell_text = ' '.join([p.text for p in paras if p.text])
                        row_cells.append(cell_text)
                    
                    if any(cell.strip() for cell in row_cells):
                        rows_data.append(row_cells)
                
                if rows_data:
                    max_cols = max(len(row) for row in rows_data)
                    for row in rows_data:
                        while len(row) < max_cols:
                            row.append('')
                    tables.append(pd.DataFrame(rows_data))
    
    return tables


# ============================================
# 2. VLIF-SPECIFIEKE DATAVERWERKING
# ============================================

def clean_vlif_dataframe(df):
    """
    Maak VLIF dataframe schoon en gestructureerd
    """
    if df.empty:
        return df
    
    # Verwijder volledig lege rijen en kolommen
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Reset index
    df = df.reset_index(drop=True)
    
    return df


def extract_investments(df):
    """
    Extraheer investeringsregels uit VLIF tabel
    Herkent patronen zoals:
    - Steunpercentages (30, 40, 15, 50, 75, 100)
    - Omschrijving
    - Kenmerken en eenheden
    """
    investments = []
    
    for idx, row in df.iterrows():
        row_values = [str(val).strip() if pd.notna(val) else '' for val in row.values]
        
        # Zoek naar percentage in eerste kolommen
        percentage = None
        for i, val in enumerate(row_values[:3]):
            if re.match(r'^\d{2,3}$', val):
                percentage = int(val)
                break
        
        # Zoek naar omschrijving (tekst van minstens 5 karakters)
        description = None
        for i, val in enumerate(row_values):
            if len(val) > 5 and not re.match(r'^\d+$', val):
                description = val
                start_idx = i
                break
        
        if not description:
            continue
        
        # Verzamel overige kolommen
        extra_info = None
        kenmerk = None
        eenheid = None
        
        if start_idx + 1 < len(row_values):
            extra_info = row_values[start_idx + 1] if len(row_values[start_idx + 1]) > 3 else None
        if start_idx + 2 < len(row_values):
            kenmerk = row_values[start_idx + 2] if not re.match(r'^\d+$', row_values[start_idx + 2]) else None
        if start_idx + 3 < len(row_values):
            eenheid = row_values[start_idx + 3]
        
        investments.append({
            'steunpercentage': percentage,
            'omschrijving': description,
            'extra_info': extra_info,
            'kenmerk': kenmerk,
            'eenheid': eenheid,
            'brongegevens': ' | '.join(row_values)
        })
    
    return pd.DataFrame(investments)


def classify_investment(row):
    """
    Classificeer investering op basis van omschrijving
    """
    text = str(row['omschrijving']).lower()
    
    # AERI (Ammoniak Emissie Reducerende Investeringen)
    aeri_keywords = ['luchtwasser', 'ammoniak', 'dichte vloer', 'mestrobot', 
                     'koeldeksysteem', 'mestband', 'volière', 'beluchting']
    
    # NPI (Niet-Productieve Investeringen)
    npi_keywords = ['haag', 'houtkant', 'poel', 'boom', 'nestkast', 'biodiversiteit',
                    'erosie', 'landschap', 'waterberging', 'infiltratie']
    
    # Bepaal type
    if any(keyword in text for keyword in aeri_keywords):
        return 'AERI'
    elif any(keyword in text for keyword in npi_keywords):
        return 'NPI'
    else:
        return 'Productief'


def add_metadata(df, source_file):
    """
    Voeg metadata toe aan dataframe
    """
    df['bronbestand'] = source_file
    df['extractiedatum'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df['investeringstype'] = df.apply(classify_investment, axis=1)
    return df


# ============================================
# 3. EXPORT NAAR EXCEL
# ============================================

def export_to_excel(data_dict, output_path):
    """
    Exporteer meerdere DataFrames naar Excel met meerdere sheets
    
    Parameters:
    data_dict: dict met sheetnamen als keys en DataFrames als values
    output_path: pad voor .xlsx bestand
    """
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in data_dict.items():
            if not df.empty:
                # Limiteer sheetnaam naar 31 tekens (Excel limiet)
                safe_name = sheet_name[:31].replace('[', '').replace(']', '').replace(':', '')
                df.to_excel(writer, sheet_name=safe_name, index=False)
                
                # Auto-width kolommen
                worksheet = writer.sheets[safe_name]
                for column in df:
                    column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
                    worksheet.column_dimensions[column].width = min(column_width, 50)
    
    print(f"✅ Excel-bestand opgeslagen: {output_path}")


# ============================================
# 4. HOOFDFUNCTIE
# ============================================

def convert_odf_to_excel(input_path, output_path=None):
    """
    Converteer ODF bestand (.ods of .odt) naar Excel (.xlsx)
    
    Parameters:
    input_path: pad naar .ods of .odt bestand
    output_path: optioneel pad voor .xlsx bestand (standaard:zelfde naam met .xlsx)
    
    Returns:
    Dict met alle geëxtraheerde data
    """
    input_path = Path(input_path)
    
    if output_path is None:
        output_path = input_path.with_suffix('.xlsx')
    
    result = {}
    
    print(f"📄 Verwerken: {input_path.name}")
    
    # Lees ODS of ODT
    if input_path.suffix.lower() == '.ods':
        sheets = read_ods(str(input_path))
        
        for sheet_name, df in sheets.items():
            print(f"  📑 Verwerk sheet: {sheet_name}")
            df_clean = clean_vlif_dataframe(df)
            investments = extract_investments(df_clean)
            investments = add_metadata(investments, input_path.name)
            result[f"{sheet_name}_investeringen"] = investments
            result[f"{sheet_name}_ruw"] = df_clean
    
    elif input_path.suffix.lower() == '.odt':
        tables = read_odt(str(input_path))
        
        for idx, df in enumerate(tables):
            print(f"  📑 Verwerk tabel {idx + 1}")
            df_clean = clean_vlif_dataframe(df)
            investments = extract_investments(df_clean)
            investments = add_metadata(investments, input_path.name)
            result[f"tabel_{idx + 1}_investeringen"] = investments
            result[f"tabel_{idx + 1}_ruw"] = df_clean
    
    else:
        raise ValueError(f"Ongeldig bestandstype: {input_path.suffix}. Gebruik .ods of .odt")
    
    # Exporteer naar Excel
    if result:
        export_to_excel(result, output_path)
        
        # Toon statistieken
        total_inv = sum(len(df) for name, df in result.items() if 'investeringen' in name)
        print(f"\n📊 STATISTIEKEN:")
        print(f"  - Aantal sheets in Excel: {len(result)}")
        print(f"  - Totaal geëxtraheerde investeringen: {total_inv}")
        
        # Per type
        for inv_name, df in result.items():
            if 'investeringen' in name and 'investeringstype' in df.columns:
                print(f"\n  📈 {inv_name}:")
                for inv_type, count in df['investeringstype'].value_counts().items():
                    print(f"     - {inv_type}: {count}")
    
    return result


# ============================================
# 5. COMMAND LINE INTERFACE
# ============================================

if __name__ == "__main__":
    import sys
    
    # Configureer pandas weergave
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    
    # Gebruik command line argument of vraag input
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
    else:
        print("VLIF ODF naar Excel Converter")
        print("=" * 40)
        input_file = input("Voer pad naar .ods of .odt bestand in: ").strip()
        output_file = input("Voer pad voor .xlsx output (optional, enter voor default): ").strip() or None
    
    try:
        result = convert_odf_to_excel(input_file, output_file)
        print("\n✅ Conversie succesvol afgerond!")
    except FileNotFoundError:
        print(f"❌ Bestand niet gevonden: {input_file}")
    except Exception as e:
        print(f"❌ Fout tijdens conversie: {e}")
        print("\n💡 Tip: Controleer of je de benodigde packages hebt geïnstalleerd:")
        print("   pip install pandas openpyxl odfpy")
