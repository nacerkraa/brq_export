"""
export_brq.py

Full daily workflow (XAMPP / MariaDB version):
  1. Clean the .sql export (remove the invalid trailing XML block)
  2. Re-import it into your local MariaDB (XAMPP), replacing the old database
  3. Run the BRQ_Category query for a given date
  4. Save the result as an Excel (.xlsx) file

Setup (run once):
    pip install pymysql pandas openpyxl
    (Start Apache and MySQL from the XAMPP Control Panel before running this)

Usage:
    python export_brq.py                # uses today's date
    python export_brq.py 2026-09-24      # uses a specific date
"""

import sys
import os
import subprocess
import datetime
import pymysql
import pandas as pd

# ---------------------------------------------------------------------------
# 1. EDIT THESE VALUES to match your setup
# ---------------------------------------------------------------------------
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "",                       # XAMPP's default root user has NO password
    "database": "UAB_OULEDHAMLA_COM26",
    "port": 3306,
}

# Path to XAMPP's mysql.exe command-line tool (default XAMPP install location)
MYSQL_EXE_PATH = "C:\\xampp\\mysql\\bin\\mysql.exe"

# Folder where everything lives
WORK_FOLDER = "C:\\BRQ"

# The .sql file you drop into WORK_FOLDER every day (must match this exact name)
SQL_FILE_NAME = "UAB_OULEDHAMLA_COM26.sql"

# ---------------------------------------------------------------------------
SQL_FILE_PATH = os.path.join(WORK_FOLDER, SQL_FILE_NAME)
CLEANED_SQL_PATH = os.path.join(WORK_FOLDER, "_cleaned_import.sql")

QUERY = """
SELECT
    CASE
        WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'CHAIR' THEN 'CHAIR'
        WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'PONTE' THEN 'PONTE'
        WHEN f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN') THEN 'RUMINANTS'
        WHEN f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS' THEN 'DIVERS'
    END AS BRQ_Category,
    e.Date AS Date_mvt,
    e.Libelle_Client AS Client_Name,
    e.Code_FC AS Code,
    d.Produit_Code,
    d.Quantite,
    d.Valeur AS Total_Amount
FROM entete_mvts e
JOIN detail_mvts d
    ON d.Entete_mvts_Code = e.Code
   AND d.Type_mvt = e.Type_mvt
   AND d.Type_bon = e.Type_bon
JOIN produit p ON p.Code = d.Produit_Code
LEFT JOIN produit_famille pf ON pf.Produit_Code = p.Code
LEFT JOIN famille f ON f.Code = pf.Famille_Code
LEFT JOIN produit_sousfammille psf ON psf.Produit_Code = p.Code
LEFT JOIN sous_famille sf ON sf.Code = psf.SousFamille_Code
WHERE
    e.Date = %(brq_date)s
    AND e.Type_mvt = 'S'
    AND e.Type_bon = 'F'
    AND (
        (f.Libelle = 'VOLAILLE' AND sf.Libelle IN ('CHAIR', 'PONTE'))
        OR (f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN'))
        OR (f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS')
    )
ORDER BY
    CASE
        WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'CHAIR' THEN 1
        WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'PONTE' THEN 2
        WHEN f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN') THEN 3
        WHEN f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS' THEN 4
    END,
    e.Libelle_Client, e.Code, d.Produit_Code;
"""


def get_target_date():
    if len(sys.argv) > 1:
        date_str = sys.argv[1]
        try:
            return datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            print(f"Invalid date '{date_str}'. Use format YYYY-MM-DD.")
            sys.exit(1)
    return datetime.date.today()


def clean_sql_file(input_path, output_path):
    """Remove the invalid trailing XML (<?xml ... <PARAMETRAGE>...) from the dump."""
    if not os.path.exists(input_path):
        print(f"Could not find the SQL file at: {input_path}")
        sys.exit(1)

    text = None
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            with open(input_path, "r", encoding=enc) as f:
                text = f.read()
            break
        except UnicodeDecodeError:
            continue

    if text is None:
        print("Could not read the SQL file — unrecognized text encoding.")
        sys.exit(1)

    marker_index = text.find("<?xml")
    if marker_index == -1:
        cleaned = text
        print("No trailing XML block found (file was already clean).")
    else:
        cleaned = text[:marker_index]
        removed_chars = len(text) - len(cleaned)
        print(f"Removed {removed_chars} characters of invalid trailing XML.")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(cleaned)


def import_sql_into_mariadb(cleaned_sql_path):
    """Drop and recreate the database in local MariaDB (XAMPP), then import."""
    if not os.path.exists(MYSQL_EXE_PATH):
        print(f"Could not find mysql.exe at: {MYSQL_EXE_PATH}")
        print("Open XAMPP's install folder and update MYSQL_EXE_PATH in the script.")
        sys.exit(1)

    user = DB_CONFIG["user"]
    password = DB_CONFIG["password"]
    database = DB_CONFIG["database"]

    base_cmd = [MYSQL_EXE_PATH, "-h", DB_CONFIG["host"], "-P", str(DB_CONFIG["port"]), "-u", user]
    if password:
        base_cmd.append(f"-p{password}")

    print(f"Replacing database '{database}'...")
    drop_create_sql = f"DROP DATABASE IF EXISTS `{database}`; CREATE DATABASE `{database}`;"
    result = subprocess.run(base_cmd + ["-e", drop_create_sql], capture_output=True, text=True)
    if result.returncode != 0:
        print("Failed to drop/create the database:")
        print(result.stderr)
        sys.exit(1)

    print("Importing cleaned SQL file (this can take a moment)...")
    with open(cleaned_sql_path, "rb") as f:
        result = subprocess.run(base_cmd + [database], stdin=f, capture_output=True)
    if result.returncode != 0:
        print("Import failed:")
        print(result.stderr.decode(errors="replace"))
        sys.exit(1)

    print("Import complete.")


def main():
    target_date = get_target_date()

    clean_sql_file(SQL_FILE_PATH, CLEANED_SQL_PATH)
    import_sql_into_mariadb(CLEANED_SQL_PATH)

    print(f"Running query for date: {target_date}")
    connection = pymysql.connect(**DB_CONFIG)
    try:
        df = pd.read_sql(QUERY, connection, params={"brq_date": target_date})
    finally:
        connection.close()

    if df.empty:
        print("No rows found for this date. Nothing was exported.")
        return

    filename = os.path.join(WORK_FOLDER, f"BRQ_Export_{target_date}.xlsx")

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BRQ")
        worksheet = writer.sheets["BRQ"]
        for i, col in enumerate(df.columns, start=1):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = max_len

    print(f"Done. Saved {len(df)} rows to: {filename}")


if __name__ == "__main__":
    main()
