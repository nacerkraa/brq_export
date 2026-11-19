"""
Export all products from MariaDB to an Excel file.

Excel columns:
- Nom de produit
- Code de produit
- Prix unitaire (from produit.Prix3)

Install dependencies:
    py -m pip install mysql-connector-python pandas openpyxl

Run:
    py export_products.py
"""

from pathlib import Path

import mysql.connector
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment


# Update these values if your MariaDB configuration is different.
DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "uab_ouledhamla_com26",
}

SQL = """
SELECT
    Libelle AS `Nom de produit`,
    Code AS `Code de produit`,
    Prix3 AS `Prix unitaire`
FROM produit
ORDER BY Libelle, Code;
"""


def main():
    connection = None

    try:
        print("Connecting to MariaDB...")
        connection = mysql.connector.connect(**DB_CONFIG)

        print("Reading products...")
        df = pd.read_sql_query(SQL, connection)

        if df.empty:
            print("No products were found in the produit table.")
            return

        output_file = Path.cwd() / "Liste_Produits.xlsx"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Produits", index=False)
            ws = writer.sheets["Produits"]

            # Format column headers.
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center")

            # Format unit prices.
            for cell in ws["C"][1:]:
                cell.number_format = '#,##0.00;[Red]-#,##0.00'

            # Make the worksheet convenient to use.
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.column_dimensions["A"].width = 40
            ws.column_dimensions["B"].width = 20
            ws.column_dimensions["C"].width = 18

        print(f"Export completed: {output_file}")
        print(f"Products exported: {len(df)}")

    except mysql.connector.Error as error:
        print(f"MariaDB error: {error}")
        print("Check DB_CONFIG and make sure MariaDB is running.")
    except Exception as error:
        print(f"Unexpected error: {error}")
    finally:
        if connection is not None and connection.is_connected():
            connection.close()


if __name__ == "__main__":
    main()
