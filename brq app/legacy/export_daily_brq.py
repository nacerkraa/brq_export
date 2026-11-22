"""
Export daily BRQ quantities from MariaDB to Excel.

Install:
    py -m pip install mysql-connector-python pandas openpyxl

Run:
    py export_daily_brq.py
"""

from datetime import datetime
from pathlib import Path

import mysql.connector
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Edit these settings to match your XAMPP/MariaDB configuration.
DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "uab_ouledhamla_com26",
}

SQL = """
SELECT
    DATE(e.Date) AS Date_mvt,
    SUM(CASE WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'CHAIR'
        THEN d.Quantite ELSE 0 END) AS CHAIR,
    SUM(CASE WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'PONTE'
        THEN d.Quantite ELSE 0 END) AS PONTE,
    SUM(CASE WHEN f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN')
        THEN d.Quantite ELSE 0 END) AS RUMINANTS,
    SUM(CASE WHEN f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS'
        THEN d.Quantite ELSE 0 END) AS DIVERS,
    SUM(CASE
        WHEN (f.Libelle = 'VOLAILLE' AND sf.Libelle IN ('CHAIR', 'PONTE'))
          OR (f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN'))
          OR (f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS')
        THEN d.Quantite ELSE 0
    END) AS TOTAL
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
    e.Date >= %s
    AND e.Date < DATE_ADD(%s, INTERVAL 1 DAY)
    AND e.Type_mvt = 'S'
    AND e.Type_bon IN ('F', 'FA')
GROUP BY DATE(e.Date)
ORDER BY DATE(e.Date);
"""


def ask_date(prompt):
    while True:
        value = input(prompt).strip()
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            print("Use YYYY-MM-DD format, for example 2026-10-01.")


def main():
    print("=== Daily BRQ Excel Export ===")
    start_date = ask_date("Start date (YYYY-MM-DD): ")
    end_date = ask_date("End date   (YYYY-MM-DD): ")

    if end_date < start_date:
        print("Error: end date must not be earlier than start date.")
        return

    connection = None
    try:
        print("\nConnecting to MariaDB...")
        connection = mysql.connector.connect(**DB_CONFIG)

        print("Executing query...")
        df = pd.read_sql_query(SQL, connection, params=(start_date, end_date))

        if df.empty:
            print("No matching movements found for this period.")
            return

        quantity_columns = ["CHAIR", "PONTE", "RUMINANTS", "DIVERS", "TOTAL"]
        for column in quantity_columns:
            df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

        output_file = Path.cwd() / (
            f"BRQ_Daily_{start_date:%Y%m%d}_{end_date:%Y%m%d}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Daily BRQ", index=False)
            ws = writer.sheets["Daily BRQ"]

            last_data_row = len(df) + 1          # header is row 1
            total_row = last_data_row + 1

            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center")

            for row in ws.iter_rows(min_row=2, max_row=last_data_row, max_col=1):
                for cell in row:
                    cell.number_format = "yyyy-mm-dd"

            # Totals row: one SUM formula per quantity column (B..F)
            ws.cell(row=total_row, column=1, value="TOTAL")
            for col in range(2, 2 + len(quantity_columns)):
                letter = get_column_letter(col)
                ws.cell(row=total_row, column=col,
                        value=f"=SUM({letter}2:{letter}{last_data_row})")

            for row in ws.iter_rows(min_row=2, max_row=total_row,
                                    min_col=2, max_col=1 + len(quantity_columns)):
                for cell in row:
                    cell.number_format = '#,##0.00;[Red]-#,##0.00'

            top_border = Border(top=Side(style="medium"))
            for cell in ws[total_row]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="DDEBF7")
                cell.border = top_border

            ws.freeze_panes = "A2"
            # Filter only the data rows, so sorting never moves the totals row
            ws.auto_filter.ref = (
                f"A1:{get_column_letter(1 + len(quantity_columns))}{last_data_row}"
            )

            for column, width in {
                "A": 16, "B": 15, "C": 15,
                "D": 17, "E": 15, "F": 15
            }.items():
                ws.column_dimensions[column].width = width

        print(f"\nExcel file created: {output_file}")
        print(f"Days with movements: {len(df)}")
        print("Totals for the period:")
        for column in quantity_columns:
            print(f"  {column:<10} {df[column].sum():>14,.2f}")

    except mysql.connector.Error as error:
        print(f"MariaDB connection/query error: {error}")
        print("Check DB_CONFIG and make sure MariaDB is running.")
    except Exception as error:
        print(f"Error: {error}")
    finally:
        if connection is not None and connection.is_connected():
            connection.close()


if __name__ == "__main__":
    main()
