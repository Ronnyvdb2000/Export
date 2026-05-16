import sqlite3
from pathlib import Path

def save_sqlite(df, output_path):
    """
    Slaat de dataframe op in een SQLite database.
    De tabelnaam = 'vlif'.
    Bestandsnaam = zelfde naam als Excel, maar .db
    """
    db_path = Path(output_path).with_suffix(".db")
    conn = sqlite3.connect(db_path)

    # DataFrame → SQLite
    df.to_sql("vlif", conn, if_exists="replace", index=False)

    # Indexen voor snelle queries
    conn.execute("CREATE INDEX IF NOT EXISTS idx_vlif_code ON vlif(code);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_vlif_subfilter ON vlif(subfilter_nr);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_vlif_hoofdstuk ON vlif(hoofdstuk);")

    conn.commit()
    conn.close()

    print(f"🗄️  SQLite database opgeslagen: {db_path}")
