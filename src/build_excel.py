# === ZOEK DIT STUK OP IN src/build_excel.py ===
    print(f"📖 Inlezen: {input_path}")
    df = pd.read_excel(input_path)

    # 🛠️ VOEG DEZE FIX HIERONDER TOE:
    # Als de nieuwe layout_engine generieke kolomnamen heeft aangemaakt, 
    # mappen we deze naar de verwachtingen van dit script.
    if "Kolom_3" in df.columns or "Kolom_4" in df.columns:
        rename_dict = {}
        if "Kolom_3" in df.columns: rename_dict["Kolom_3"] = "eenheid"
        if "Kolom_4" in df.columns: rename_dict["Kolom_4"] = "bedrag"
        if "Kolom_5" in df.columns: rename_dict["Kolom_5"] = "extra_info"
        df = df.rename(columns=rename_dict)
    
    # Zorg dat 'eenheid' altijd bestaat om de KeyError te voorkomen, zelfs als die leeg is
    if "eenheid" not in df.columns:
        df["eenheid"] = ""
    if "bedrag" not in df.columns:
        df["bedrag"] = ""

    # === HIER GAAT DE BESTAANDE CODE WEER VERDER ===
    df = fix_data(df)
