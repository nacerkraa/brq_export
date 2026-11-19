"""
Export sales totals accumulated per client to Excel.

Output columns:
Code client | Nom du client | Adresse | Quantité | Total HT | TVA | Timbre | TTC

Install:
    py -m pip install mysql-connector-python openpyxl

Run:
    py export_sales_by_client.py

The script asks for a start and end date in YYYY-MM-DD format.
It exports sales invoices only (Type_mvt='S', Type_bon='F').
BLAF delivery notes are excluded to avoid duplicate sales.
"""

from datetime import datetime
from pathlib import Path

import mysql.connector
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "uab_ouledhamla_com26",
}

SQL = """
SELECT
    e.Code_FC AS `Code client`,
    COALESCE(NULLIF(MAX(c.Raison), ''), MAX(e.Libelle_Client), '') AS `Nom du client`,
    COALESCE(MAX(c.Adresse), '') AS `Adresse`,
    SUM(COALESCE(q.Quantite, 0)) AS `Quantité`,
    SUM(e.Total) AS `Total HT`,
    SUM(e.TVA) AS `TVA`,
    SUM(e.Timbre) AS `Timbre`,
    SUM(e.TTC) AS `TTC`
FROM entete_mvts e
LEFT JOIN client c
    ON c.Client_Code = e.Code_FC
LEFT JOIN (
    SELECT Entete_mvts_Code, Type_mvt, Type_bon, SUM(Quantite) AS Quantite
    FROM detail_mvts
    GROUP BY Entete_mvts_Code, Type_mvt, Type_bon
) q
    ON q.Entete_mvts_Code = e.Code
   AND q.Type_mvt = e.Type_mvt
   AND q.Type_bon = e.Type_bon
WHERE
    e.Date >= %s
    AND e.Date < DATE_ADD(%s, INTERVAL 1 DAY)
    AND e.Type_mvt = 'S'
    AND e.Type_bon IN ('F', 'FA')
GROUP BY e.Code_FC
ORDER BY MAX(c.Raison), e.Code_FC;
"""


def ask_date(prompt):
    while True:
        try:
            return datetime.strptime(input(prompt).strip(), "%Y-%m-%d").date()
        except ValueError:
            print("Date invalide. Format attendu : YYYY-MM-DD (exemple : 2026-09-01).")


def main():
    print("=== Export des ventes cumulées par client ===")
    start_date = ask_date("Date de début (YYYY-MM-DD) : ")
    end_date = ask_date("Date de fin   (YYYY-MM-DD) : ")

    if end_date < start_date:
        print("Erreur : la date de fin doit être postérieure ou égale à la date de début.")
        return

    connection = None
    try:
        print("\\nConnexion à MariaDB...")
        connection = mysql.connector.connect(**DB_CONFIG)

        print("Récupération des ventes...")
        cursor = connection.cursor()
        cursor.execute(SQL, (start_date, end_date))
        rows = cursor.fetchall()
        headers = [column[0] for column in cursor.description]
        cursor.close()

        if not rows:
            print("Aucune vente trouvée pour cette période.")
            return

        output_file = Path.cwd() / f"Ventes_par_client_{start_date}_{end_date}.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Ventes par client"
        worksheet.append(headers)

        for row in rows:
            worksheet.append(list(row))

        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in worksheet.iter_rows(min_row=2):
            row[3].number_format = '#,##0.000'
            for cell in row[4:8]:
                cell.number_format = '#,##0.00;[Red]-#,##0.00'

        widths = [18, 35, 45, 16, 18, 16, 16, 18]
        for index, width in enumerate(widths, start=1):
            worksheet.column_dimensions[get_column_letter(index)].width = width

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 24
        workbook.save(output_file)

        print(f"\\nExport terminé : {output_file}")
        print(f"Nombre de clients : {len(rows)}")

    except mysql.connector.Error as error:
        print(f"Erreur MariaDB : {error}")
        print("Vérifiez DB_CONFIG et assurez-vous que MariaDB est démarré.")
    except Exception as error:
        print(f"Erreur inattendue : {error}")
    finally:
        if connection is not None and connection.is_connected():
            connection.close()


if __name__ == "__main__":
    main()
