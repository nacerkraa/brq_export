
"""
Export sales totals accumulated per client to Excel.

Output columns:
Code client | Nom du client | Adresse | Quantité |
Quantité CHAIR | Quantité PONTE | Quantité RUMINANTS |
Quantité DIVERS | Total HT | TVA | TTC

Install:
    py -m pip install mysql-connector-python openpyxl

Run:
    py export_sales_by_client.py
"""

from datetime import datetime
from pathlib import Path

import mysql.connector
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# --------------------------------------------------
# DATABASE CONFIGURATION
# --------------------------------------------------

DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "uab_ouledhamla_com26",
}


# --------------------------------------------------
# SQL QUERY
# --------------------------------------------------

SQL = """
SELECT
    e.Code_FC AS `Code client`,

    COALESCE(
        NULLIF(MAX(c.Raison), ''),
        MAX(e.Libelle_Client),
        ''
    ) AS `Nom du client`,

    COALESCE(MAX(c.Adresse), '') AS `Adresse`,

    SUM(COALESCE(q.Quantite, 0)) AS `Quantité`,

    SUM(COALESCE(q.Quantite_Chair, 0)) AS `Quantité CHAIR`,

    SUM(COALESCE(q.Quantite_Ponte, 0)) AS `Quantité PONTE`,

    SUM(COALESCE(q.Quantite_Ruminants, 0)) AS `Quantité RUMINANTS`,

    SUM(COALESCE(q.Quantite_Divers, 0)) AS `Quantité DIVERS`,

    SUM(COALESCE(e.Total, 0)) AS `Total HT`,

    SUM(COALESCE(e.TVA, 0)) AS `TVA`,

    SUM(
        COALESCE(e.Total, 0)
        + COALESCE(e.TVA, 0)
        + COALESCE(e.Timbre, 0)
    ) AS `TTC`

FROM entete_mvts e

LEFT JOIN client c
    ON c.Client_Code = e.Code_FC

LEFT JOIN (
    SELECT
        d.Entete_mvts_Code,
        d.Type_mvt,
        d.Type_bon,

        -- Total quantity
        SUM(d.Quantite) AS Quantite,

        -- CHAIR
        SUM(
            CASE
                WHEN f.Libelle = 'VOLAILLE'
                 AND sf.Libelle = 'CHAIR'
                THEN d.Quantite
                ELSE 0
            END
        ) AS Quantite_Chair,

        -- PONTE
        SUM(
            CASE
                WHEN f.Libelle = 'VOLAILLE'
                 AND sf.Libelle = 'PONTE'
                THEN d.Quantite
                ELSE 0
            END
        ) AS Quantite_Ponte,

        -- RUMINANTS (BOVIN + OVIN)
        SUM(
            CASE
                WHEN f.Libelle = 'RUMINANTS'
                THEN d.Quantite
                ELSE 0
            END
        ) AS Quantite_Ruminants,

        -- DIVERS
        SUM(
            CASE
                WHEN f.Libelle = 'DIVERS'
                 AND sf.Libelle = 'DIVERS'
                THEN d.Quantite
                ELSE 0
            END
        ) AS Quantite_Divers

    FROM detail_mvts d

    LEFT JOIN produit p
        ON p.Code = d.Produit_Code

    LEFT JOIN produit_famille pf
        ON pf.Produit_Code = p.Code

    LEFT JOIN famille f
        ON f.Code = pf.Famille_Code

    LEFT JOIN produit_sousfammille psf
        ON psf.Produit_Code = p.Code

    LEFT JOIN sous_famille sf
        ON sf.Code = psf.SousFamille_Code
       AND sf.Famille_Code = pf.Famille_Code

    GROUP BY
        d.Entete_mvts_Code,
        d.Type_mvt,
        d.Type_bon

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


# --------------------------------------------------
# DATE INPUT
# --------------------------------------------------

def ask_date(prompt):
    while True:
        try:
            return datetime.strptime(
                input(prompt).strip(),
                "%Y-%m-%d"
            ).date()

        except ValueError:
            print(
                "Date invalide. Format attendu : "
                "YYYY-MM-DD (exemple : 2026-09-01)."
            )


# --------------------------------------------------
# MAIN FUNCTION
# --------------------------------------------------

def main():

    print("=== Export des ventes cumulées par client ===")

    start_date = ask_date(
        "Date de début (YYYY-MM-DD) : "
    )

    end_date = ask_date(
        "Date de fin   (YYYY-MM-DD) : "
    )

    if end_date < start_date:
        print(
            "Erreur : la date de fin doit être "
            "postérieure ou égale à la date de début."
        )
        return

    connection = None

    try:
        print("\nConnexion à MariaDB...")

        connection = mysql.connector.connect(
            **DB_CONFIG
        )

        print("Récupération des ventes...")

        cursor = connection.cursor()

        cursor.execute(
            SQL,
            (start_date, end_date)
        )

        rows = cursor.fetchall()

        headers = [
            column[0]
            for column in cursor.description
        ]

        cursor.close()

        if not rows:
            print("Aucune vente trouvée pour cette période.")
            return

        # ------------------------------------------
        # CREATE EXCEL FILE
        # ------------------------------------------

        output_file = (
            Path.cwd()
            / f"Ventes_par_client_{start_date}_{end_date}.xlsx"
        )

        workbook = Workbook()

        worksheet = workbook.active
        worksheet.title = "Ventes par client"

        worksheet.append(headers)

        # ------------------------------------------
        # INSERT DATA
        # ------------------------------------------

        for row in rows:
            worksheet.append(list(row))

        # ------------------------------------------
        # FORMAT HEADER
        # ------------------------------------------

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78"
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        # ------------------------------------------
        # FORMAT NUMBERS
        # ------------------------------------------

        for row in worksheet.iter_rows(min_row=2):

            # Overall quantity
            row[3].number_format = '#,##0.000'

            # CHAIR, PONTE, RUMINANTS, DIVERS
            for cell in row[4:8]:
                cell.number_format = '#,##0.000'

            # Total HT, TVA, TTC
            for cell in row[8:11]:
                cell.number_format = (
                    '#,##0.00;[Red]-#,##0.00'
                )

        # ------------------------------------------
        # COLUMN WIDTHS
        # ------------------------------------------

        widths = [
            18,  # Code client
            35,  # Nom du client
            45,  # Adresse
            16,  # Quantité totale
            20,  # Quantité CHAIR
            20,  # Quantité PONTE
            23,  # Quantité RUMINANTS
            20,  # Quantité DIVERS
            18,  # Total HT
            16,  # TVA
            18,  # TTC
        ]

        for index, width in enumerate(
            widths,
            start=1
        ):
            worksheet.column_dimensions[
                get_column_letter(index)
            ].width = width

        # ------------------------------------------
        # WORKSHEET SETTINGS
        # ------------------------------------------

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 35

        # ------------------------------------------
        # SAVE FILE
        # ------------------------------------------

        workbook.save(output_file)

        print("\nExport terminé avec succès !")
        print(f"Fichier : {output_file}")
        print(f"Nombre de clients : {len(rows)}")

    except mysql.connector.Error as error:

        print(f"Erreur MariaDB : {error}")

        print(
            "Vérifiez DB_CONFIG et assurez-vous "
            "que MariaDB est démarré."
        )

    except Exception as error:

        print(f"Erreur inattendue : {error}")

    finally:

        if (
            connection is not None
            and connection.is_connected()
        ):
            connection.close()


if __name__ == "__main__":
    main()
