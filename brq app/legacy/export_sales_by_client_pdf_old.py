#!/usr/bin/env python3
"""
Export sales accumulated by client to a PDF report.

Install dependencies:
    py -m pip install mysql-connector-python reportlab

Run:
    py export_sales_by_client_pdf.py
"""

from datetime import datetime
from decimal import Decimal
from pathlib import Path

import mysql.connector
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

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
    e.Code_FC AS client_code,

    COALESCE(
        NULLIF(MAX(c.Raison), ''),
        MAX(e.Libelle_Client),
        ''
    ) AS client_name,

    COALESCE(MAX(c.Adresse), '') AS address,

    SUM(COALESCE(q.Quantite, 0)) AS quantity,

    SUM(COALESCE(e.Total, 0)) AS total_ht,

    SUM(COALESCE(e.TVA, 0)) AS tva,

    SUM(COALESCE(e.Timbre, 0)) AS timbre,

    SUM(
        COALESCE(e.Total, 0)
        + COALESCE(e.TVA, 0)
        + COALESCE(e.Timbre, 0)
    ) AS ttc

FROM entete_mvts e

LEFT JOIN client c
    ON c.Client_Code = e.Code_FC

LEFT JOIN (
    SELECT
        Entete_mvts_Code,
        Type_mvt,
        Type_bon,
        SUM(Quantite) AS Quantite
    FROM detail_mvts
    GROUP BY
        Entete_mvts_Code,
        Type_mvt,
        Type_bon
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

ORDER BY MAX(c.Raison), e.Code_FC
"""


# --------------------------------------------------
# DATE INPUT
# --------------------------------------------------

def ask_date(prompt):
    while True:
        value = input(prompt).strip()

        try:
            return datetime.strptime(
                value, "%Y-%m-%d"
            ).date()

        except ValueError:
            print(
                "Format invalide. "
                "Utilisez YYYY-MM-DD "
                "(exemple : 2026-11-01)."
            )


# --------------------------------------------------
# FORMATTING FUNCTIONS
# --------------------------------------------------

def fmt_number(value):
    if value is None:
        value = 0

    return (
        f"{float(value):,.2f}"
        .replace(",", " ")
        .replace(".", ",")
    )


def safe_text(value):
    return (
        str(value or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# --------------------------------------------------
# PDF FOOTER
# --------------------------------------------------

def footer(canvas, doc):
    canvas.saveState()

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)

    canvas.drawString(
        12 * mm,
        8 * mm,
        "État récapitulatif des ventes par client"
    )

    canvas.drawRightString(
        285 * mm,
        8 * mm,
        f"Page {doc.page}"
    )

    canvas.restoreState()


# --------------------------------------------------
# MAIN FUNCTION
# --------------------------------------------------

def main():

    print("=" * 45)
    print("   EXPORT PDF : VENTES PAR CLIENT")
    print("=" * 45)

    start_date = ask_date(
        "Date de début (YYYY-MM-DD) : "
    )

    end_date = ask_date(
        "Date de fin   (YYYY-MM-DD) : "
    )

    if end_date < start_date:
        print(
            "Erreur : la date de fin doit être "
            "égale ou postérieure à la date de début."
        )
        return

    # ----------------------------------------------
    # DATABASE CONNECTION
    # ----------------------------------------------

    connection = None

    try:
        connection = mysql.connector.connect(
            **DB_CONFIG
        )

        cursor = connection.cursor()

        cursor.execute(
            SQL,
            (start_date, end_date)
        )

        rows = cursor.fetchall()

        cursor.close()

    except mysql.connector.Error as exc:
        print(f"Erreur MariaDB : {exc}")
        return

    finally:
        if connection and connection.is_connected():
            connection.close()

    # ----------------------------------------------
    # OUTPUT FILE
    # ----------------------------------------------

    output = (
        Path.cwd()
        / f"Ventes_par_client_{start_date}_{end_date}.pdf"
    )

    # ----------------------------------------------
    # PDF STYLES
    # ----------------------------------------------

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        alignment=TA_CENTER,
        spaceAfter=5,
    )

    subtitle_style = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontSize=9,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#444444"),
    )

    cell_style = ParagraphStyle(
        "Cell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7,
        leading=8,
    )

    header_style = ParagraphStyle(
        "Header",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        alignment=TA_CENTER,
    )

    right_style = ParagraphStyle(
        "RightCell",
        parent=cell_style,
        alignment=TA_RIGHT,
    )

    # ----------------------------------------------
    # PDF DOCUMENT
    # ----------------------------------------------

    doc = SimpleDocTemplate(
        str(output),
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=15 * mm,
    )

    story = []

    story.append(
        Paragraph(
            "ÉTAT RÉCAPITULATIF DES VENTES PAR CLIENT",
            title_style,
        )
    )

    story.append(
        Paragraph(
            f"Période : du "
            f"{start_date.strftime('%d/%m/%Y')} "
            f"au {end_date.strftime('%d/%m/%Y')}",
            subtitle_style,
        )
    )

    story.append(Spacer(1, 7 * mm))

    # ----------------------------------------------
    # TABLE HEADERS
    # ----------------------------------------------

    headers = [
        "Code client",
        "Nom du client",
        "Adresse",
        "Quantité",
        "Total HT",
        "TVA",
        "Timbre",
        "TTC",
    ]

    data = [
        [
            Paragraph(header, header_style)
            for header in headers
        ]
    ]

    # ----------------------------------------------
    # CALCULATE GRAND TOTALS
    # ----------------------------------------------

    totals = [Decimal("0")] * 5

    # ----------------------------------------------
    # ADD CLIENT ROWS
    # ----------------------------------------------

    for row in rows:

        (
            code,
            name,
            address,
            quantity,
            total_ht,
            tva,
            timbre,
            ttc,
        ) = row

        values = [
            quantity,
            total_ht,
            tva,
            timbre,
            ttc,
        ]

        for i, value in enumerate(values):
            totals[i] += Decimal(str(value or 0))

        data.append([
            Paragraph(safe_text(code), cell_style),
            Paragraph(safe_text(name), cell_style),
            Paragraph(safe_text(address), cell_style),

            Paragraph(
                fmt_number(quantity),
                right_style,
            ),

            Paragraph(
                fmt_number(total_ht),
                right_style,
            ),

            Paragraph(
                fmt_number(tva),
                right_style,
            ),

            Paragraph(
                fmt_number(timbre),
                right_style,
            ),

            Paragraph(
                fmt_number(ttc),
                right_style,
            ),
        ])

    # ----------------------------------------------
    # NO DATA MESSAGE
    # ----------------------------------------------

    if not rows:

        data.append([
            Paragraph(
                "Aucune vente trouvée pour cette période.",
                cell_style,
            ),
            "", "", "", "", "", "", "",
        ])

    # ----------------------------------------------
    # GRAND TOTAL ROW
    # ----------------------------------------------

    else:

        data.append([
            Paragraph(
                "<b>TOTAL GÉNÉRAL</b>",
                cell_style,
            ),
            "",
            "",

            Paragraph(
                f"<b>{fmt_number(totals[0])}</b>",
                right_style,
            ),

            Paragraph(
                f"<b>{fmt_number(totals[1])}</b>",
                right_style,
            ),

            Paragraph(
                f"<b>{fmt_number(totals[2])}</b>",
                right_style,
            ),

            Paragraph(
                f"<b>{fmt_number(totals[3])}</b>",
                right_style,
            ),

            Paragraph(
                f"<b>{fmt_number(totals[4])}</b>",
                right_style,
            ),
        ])

    # ----------------------------------------------
    # CREATE PDF TABLE
    # ----------------------------------------------

    table = Table(
        data,
        colWidths=[
            22 * mm,
            47 * mm,
            55 * mm,
            24 * mm,
            27 * mm,
            22 * mm,
            22 * mm,
            27 * mm,
        ],
        repeatRows=1,
        hAlign="LEFT",
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#23415C"),
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white,
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.HexColor("#AAB4BE"),
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE",
            ),
            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                4,
            ),
            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                4,
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5,
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5,
            ),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [
                    colors.white,
                    colors.HexColor("#F2F5F7"),
                ],
            ),
        ])
    )

    # Style the grand total row
    if rows:

        total_row = len(data) - 1

        table.setStyle(
            TableStyle([
                (
                    "BACKGROUND",
                    (0, total_row),
                    (-1, total_row),
                    colors.HexColor("#DCE5EC"),
                ),
                (
                    "LINEABOVE",
                    (0, total_row),
                    (-1, total_row),
                    1,
                    colors.HexColor("#23415C"),
                ),
            ])
        )

    story.append(table)

    story.append(Spacer(1, 4 * mm))

    story.append(
        Paragraph(
            f"Nombre de clients : {len(rows)}",
            cell_style,
        )
    )

    # ----------------------------------------------
    # GENERATE PDF
    # ----------------------------------------------

    doc.build(
        story,
        onFirstPage=footer,
        onLaterPages=footer,
    )

    print()
    print("=" * 45)
    print("PDF créé avec succès !")
    print(f"Fichier : {output}")
    print(f"Nombre de clients : {len(rows)}")
    print("=" * 45)


if __name__ == "__main__":
    main()
