"""Consolidated BRQ desktop application.
Dependencies: pip install pymysql pandas openpyxl reportlab pillow
"""
from __future__ import annotations
import datetime as dt
import os
import queue
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pymysql
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

WORK_FOLDER = r"C:\BRQ"
DEFAULT_SQL_PATH = os.path.join(WORK_FOLDER, "UAB_OULEDHAMLA_COM26.sql")
CLEANED_SQL_PATH = os.path.join(WORK_FOLDER, "_cleaned_import.sql")
MYSQL_EXE_PATH = r"C:\xampp\mysql\bin\mysql.exe"
DB_CONFIG = {"host": "localhost", "user": "root", "password": "",
             "database": "UAB_OULEDHAMLA_COM26", "port": 3306,
             "charset": "utf8mb4", "connect_timeout": 5}

ANNUAL_OBJECTIVE = 380000.00

BRQ_QUERY = """
SELECT
 CASE
  WHEN f.Libelle='VOLAILLE' AND sf.Libelle='CHAIR' THEN 'CHAIR'
  WHEN f.Libelle='VOLAILLE' AND sf.Libelle='PONTE' THEN 'PONTE'
  WHEN f.Libelle='RUMINANTS' AND sf.Libelle IN ('BOVIN','OVIN') THEN 'RUMINANTS'
  WHEN f.Libelle='DIVERS' AND sf.Libelle='DIVERS' THEN 'DIVERS'
 END AS BRQ_Category,
 e.Code AS Invoice_Reference,
 e.Code_FC AS Client_Code,
 e.Libelle_Client AS Client_Name,
 d.Produit_Code,
 d.Libelle AS Product_Label,
 d.Quantite,
 d.Prix AS Unit_Price,
 d.Valeur AS Total_Amount
FROM entete_mvts e
JOIN detail_mvts d ON d.Entete_mvts_Code=e.Code
 AND d.Type_mvt=e.Type_mvt AND d.Type_bon=e.Type_bon
JOIN produit p ON p.Code=d.Produit_Code
LEFT JOIN produit_famille pf ON pf.Produit_Code=p.Code
LEFT JOIN famille f ON f.Code=pf.Famille_Code
LEFT JOIN produit_sousfammille psf ON psf.Produit_Code=p.Code
LEFT JOIN sous_famille sf ON sf.Code=psf.SousFamille_Code
WHERE e.Date >= %(date1)s AND e.Date <= %(date2)s
 AND e.Type_mvt='S' AND e.Type_bon IN ('F','FA')
 AND ((f.Libelle='VOLAILLE' AND sf.Libelle IN ('CHAIR','PONTE'))
 OR (f.Libelle='RUMINANTS' AND sf.Libelle IN ('BOVIN','OVIN'))
 OR (f.Libelle='DIVERS' AND sf.Libelle='DIVERS'))
ORDER BY CASE
 WHEN f.Libelle='VOLAILLE' AND sf.Libelle='CHAIR' THEN 1
 WHEN f.Libelle='VOLAILLE' AND sf.Libelle='PONTE' THEN 2
 WHEN f.Libelle='RUMINANTS' AND sf.Libelle IN ('BOVIN','OVIN') THEN 3
 WHEN f.Libelle='DIVERS' AND sf.Libelle='DIVERS' THEN 4
 END, e.Code, e.Code_FC, e.Libelle_Client, d.Produit_Code
"""

RECAP_QUERY = """
SELECT
 SUM(CASE WHEN f.Libelle='VOLAILLE' AND sf.Libelle='CHAIR' THEN d.Quantite ELSE 0 END) AS CHAIR,
 SUM(CASE WHEN f.Libelle='VOLAILLE' AND sf.Libelle='PONTE' THEN d.Quantite ELSE 0 END) AS PONTE,
 SUM(CASE WHEN f.Libelle='RUMINANTS' AND sf.Libelle IN ('BOVIN','OVIN') THEN d.Quantite ELSE 0 END) AS RUMINANTS,
 SUM(CASE WHEN f.Libelle='DIVERS' AND sf.Libelle='DIVERS' THEN d.Quantite ELSE 0 END) AS DIVERS,
 SUM(CASE WHEN (f.Libelle='VOLAILLE' AND sf.Libelle IN ('CHAIR','PONTE'))
 OR (f.Libelle='RUMINANTS' AND sf.Libelle IN ('BOVIN','OVIN'))
 OR (f.Libelle='DIVERS' AND sf.Libelle='DIVERS') THEN d.Quantite ELSE 0 END) AS TOTAL
FROM entete_mvts e
JOIN detail_mvts d ON d.Entete_mvts_Code=e.Code
 AND d.Type_mvt=e.Type_mvt AND d.Type_bon=e.Type_bon
JOIN produit p ON p.Code=d.Produit_Code
LEFT JOIN produit_famille pf ON pf.Produit_Code=p.Code
LEFT JOIN famille f ON f.Code=pf.Famille_Code
LEFT JOIN produit_sousfammille psf ON psf.Produit_Code=p.Code
LEFT JOIN sous_famille sf ON sf.Code=psf.SousFamille_Code
WHERE e.Date >= %(date1)s AND e.Date <= %(date2)s
 AND e.Type_mvt='S' AND e.Type_bon IN ('F','FA')
"""

def parse_date(value):
    try:
        return dt.datetime.strptime(value.strip(), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("Date invalide. Utilisez le format AAAA-MM-JJ.") from exc

def connect_database():
    """Connect to existing DB; does not import a backup."""
    return pymysql.connect(**DB_CONFIG)

def clean_sql_file(source_path, output_path):
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"Fichier SQL introuvable : {source_path}")
    content = None
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            with open(source_path, "r", encoding=encoding) as source:
                content = source.read()
            break
        except UnicodeDecodeError:
            continue
    if content is None:
        raise UnicodeError("Encodage du fichier SQL non reconnu.")
    marker = content.find("<?xml")
    cleaned = content if marker < 0 else content[:marker]
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as target:
        target.write(cleaned)

def import_sql_into_database(sql_path):
    """Destructive import; called only after explicit user confirmation."""
    if not os.path.isfile(MYSQL_EXE_PATH):
        raise FileNotFoundError(f"mysql.exe introuvable : {MYSQL_EXE_PATH}")
    clean_sql_file(sql_path, CLEANED_SQL_PATH)
    command = [MYSQL_EXE_PATH, "-h", DB_CONFIG["host"], "-P",
               str(DB_CONFIG["port"]), "-u", DB_CONFIG["user"]]
    if DB_CONFIG["password"]:
        command.append(f"-p{DB_CONFIG['password']}")
    database = DB_CONFIG["database"]
    recreate = (f"DROP DATABASE IF EXISTS {database}; "
                f"CREATE DATABASE {database} CHARACTER SET utf8mb4;")
    result = subprocess.run(command + ["-e", recreate], capture_output=True,
                            text=True, check=False)
    if result.returncode:
        raise RuntimeError(result.stderr.strip() or "Échec de création de la base.")
    with open(CLEANED_SQL_PATH, "rb") as sql_file:
        result = subprocess.run(command + [database], stdin=sql_file,
                                capture_output=True, check=False)
    if result.returncode:
        raise RuntimeError(result.stderr.decode(errors="replace").strip()
                           or "Échec de l'import SQL.")

def write_combined_report(details, summary, start_date, end_date, output_path):
    """Write detailed BRQ and cumulative recap together on one worksheet."""
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "BRQ"

    navy = "1F4E78"
    light_blue = "D9EAF7"
    pale_grey = "F2F2F2"
    white = "FFFFFF"
    thin_grey = Side(style="thin", color="D9E1F2")
    header_fill = PatternFill("solid", fgColor=navy)
    section_fill = PatternFill("solid", fgColor=light_blue)

    detail_columns = [
        "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
        "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
        "Unit_Price"
    ]
    detail_headers = [
        "Catégorie", "Référence facture", "Client", "Code client",
        "Code produit", "Quantité", "Montant", "Désignation",
        "Prix unitaire"
    ]
    summary_headers = ["Catégorie", "Jour", "Cumul mois", "Cumul année", "Objectif annuel (%)"]

    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(detail_headers))
    title = sheet.cell(1, 1, "BRQ — ÉTAT DES VENTES")
    title.font = Font(name="Calibri", size=16, bold=True, color=white)
    title.fill = PatternFill("solid", fgColor=navy)
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 28

    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(detail_headers))
    period = sheet.cell(2, 1, f"Période : du {start_date:%d/%m/%Y} au {end_date:%d/%m/%Y}")
    period.font = Font(name="Calibri", size=11, bold=True, color=navy)
    period.alignment = Alignment(horizontal="center")
    sheet.row_dimensions[2].height = 22

    detail_header_row = 4
    for col, label in enumerate(detail_headers, 1):
        cell = sheet.cell(detail_header_row, col, label)
        cell.font = Font(bold=True, color=white)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin_grey)
    sheet.row_dimensions[detail_header_row].height = 30

    detail_count = 0
    category_fills = {
        "CHAIR": "E2F0D9", "PONTE": "FFF2CC",
        "RUMINANTS": "DDEBF7", "DIVERS": "E4DFEC"
    }
    for record in details.reindex(columns=detail_columns).itertuples(index=False, name=None):
        row_number = detail_header_row + 1 + detail_count
        for col, value in enumerate(record, 1):
            if pd.isna(value):
                value = None
            cell = sheet.cell(row_number, col, value)
            cell.border = Border(bottom=thin_grey)
            if col == 1:
                cell.font = Font(bold=True)
                fill = category_fills.get(str(value), pale_grey)
                cell.fill = PatternFill("solid", fgColor=fill)
            elif col in (5, 6, 7, 9):
                cell.number_format = '#,##0.00;[Red]-#,##0.00'
            if col in (2, 4, 5):
                cell.number_format = "@"
        detail_count += 1

    summary_title_row = detail_header_row + detail_count + 3
    sheet.merge_cells(start_row=summary_title_row, start_column=1,
                      end_row=summary_title_row, end_column=len(summary_headers))
    summary_title = sheet.cell(summary_title_row, 1, "RÉCAPITULATIF CUMULÉ")
    summary_title.font = Font(bold=True, size=12, color=white)
    summary_title.fill = header_fill
    summary_title.alignment = Alignment(horizontal="center")
    for col, label in enumerate(summary_headers, 1):
        cell = sheet.cell(summary_title_row + 1, col, label)
        cell.font = Font(bold=True, color=white)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for offset, record in enumerate(summary.itertuples(index=False, name=None), 2):
        row_number = summary_title_row + offset
        for col, value in enumerate(record, 1):
            cell = sheet.cell(row_number, col, value)
            cell.border = Border(bottom=thin_grey)
            if col == 5:
                cell.number_format = '0.00"%"'
            elif col > 1:
                cell.number_format = '#,##0.00;[Red]-#,##0.00'
            if str(record[0]) == "TOTAL":
                cell.font = Font(bold=True)
                cell.fill = section_fill

    widths = [16, 22, 30, 16, 15, 14, 17, 34, 16]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A5"
    if detail_count:
        sheet.auto_filter.ref = (
            f"A{detail_header_row}:I{detail_header_row + detail_count}"
        )
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:4"
    sheet.sheet_properties.pageSetUpPr.autoPageBreaks = False
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.5
    sheet.page_margins.bottom = 0.5
    sheet.print_area = f"A1:I{summary_title_row + len(summary) + 1}"

    workbook.save(output_path)


def _display_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{value:,.2f}"
    return str(value)


def _display_summary_value(column, value):
    """Format the annual objective achievement as a percentage."""
    if column == "Objectif annuel (%)":
        return "" if pd.isna(value) else f"{float(value):.2f}%"
    return _display_value(value)


def write_report_pdf(details, summary, start_date, end_date, output_path):
    """Create a paginated, landscape A3 PDF with details and cumulative recap."""
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, LongTable
    )

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    page_width, _ = landscape(A3)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "BrqTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=16, leading=20, textColor=colors.HexColor("#1F4E78"),
        alignment=TA_CENTER, spaceAfter=5
    )
    period_style = ParagraphStyle(
        "BrqPeriod", parent=styles["Normal"], fontSize=9,
        alignment=TA_CENTER, spaceAfter=12
    )
    cell_style = ParagraphStyle(
        "BrqCell", parent=styles["Normal"], fontSize=6.5, leading=8,
        wordWrap="CJK"
    )
    header_style = ParagraphStyle(
        "BrqHeader", parent=cell_style, fontName="Helvetica-Bold",
        textColor=colors.white, alignment=TA_CENTER
    )

    story = [
        Paragraph("BRQ — ÉTAT DES VENTES", title_style),
        Paragraph(
            f"Période : du {start_date:%d/%m/%Y} au {end_date:%d/%m/%Y}",
            period_style
        ),
        Paragraph("DÉTAIL DES VENTES", styles["Heading2"]),
        Spacer(1, 3)
    ]
    headers = [
        "Catégorie", "Référence facture", "Client", "Code client",
        "Code produit", "Quantité", "Montant", "Désignation",
        "Prix unitaire"
    ]
    table_data = [[Paragraph(escape(label), header_style) for label in headers]]
    detail_columns = [
        "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
        "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
        "Unit_Price"
    ]
    for record in details.reindex(columns=detail_columns).itertuples(index=False, name=None):
        table_data.append([
            Paragraph(escape(_display_value(value)), cell_style)
            for value in record
        ])

    available_width = page_width - 28 * mm
    proportions = [0.09, 0.12, 0.18, 0.09, 0.09, 0.07, 0.10, 0.16, 0.10]
    col_widths = [available_width * value / sum(proportions) for value in proportions]
    detail_table = LongTable(
        table_data, colWidths=col_widths, repeatRows=1, hAlign="LEFT"
    )
    detail_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C8D2DC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4)
    ]
    fills = {
        "CHAIR": "#E2F0D9", "PONTE": "#FFF2CC",
        "RUMINANTS": "#DDEBF7", "DIVERS": "#E4DFEC"
    }
    for row_index, category in enumerate(details.get("BRQ_Category", []), 1):
        fill = fills.get(str(category))
        if fill:
            detail_commands.append(
                ("BACKGROUND", (0, row_index), (0, row_index), colors.HexColor(fill))
            )
    detail_table.setStyle(TableStyle(detail_commands))
    if len(table_data) == 1:
        story.append(Paragraph("Aucun détail pour cette période.", styles["Normal"]))
    else:
        story.append(detail_table)

    story.extend([
        Spacer(1, 12),
        Paragraph("RÉCAPITULATIF CUMULÉ", styles["Heading2"]),
        Spacer(1, 4)
    ])
    summary_headers = ["Catégorie", "Jour", "Cumul mois", "Cumul année", "Objectif annuel (%)"]
    summary_data = [[Paragraph(escape(x), header_style) for x in summary_headers]]
    for record in summary.itertuples(index=False, name=None):
        summary_data.append([
            Paragraph(escape(_display_summary_value(column, value)), cell_style)
            for column, value in zip(summary.columns, record)
        ])
    summary_table = Table(
        summary_data,
        colWidths=[available_width * 0.20] * 5,
        repeatRows=1,
        hAlign="LEFT"
    )
    summary_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C8D2DC")),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5)
    ]
    for row_index, record in enumerate(summary.itertuples(index=False, name=None), 1):
        if str(record[0]) == "TOTAL":
            summary_commands.extend([
                ("BACKGROUND", (0, row_index), (-1, row_index), colors.HexColor("#D9EAF7")),
                ("FONTNAME", (0, row_index), (-1, row_index), "Helvetica-Bold")
            ])
    summary_table.setStyle(TableStyle(summary_commands))
    story.append(summary_table)

    def add_page_number(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawRightString(
            page_width - 14 * mm, 9 * mm,
            f"Page {document.page}"
        )
        canvas.restoreState()

    document = SimpleDocTemplate(
        output_path, pagesize=landscape(A3),
        rightMargin=14 * mm, leftMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=16 * mm,
        title="BRQ - Etat des ventes"
    )
    document.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)


def write_report_png_pages(details, summary, start_date, end_date, output_dir):
    """Export readable, numbered PNG pages; the final page contains the recap."""
    from PIL import Image, ImageDraw, ImageFont

    os.makedirs(output_dir, exist_ok=True)
    base_name = f"BRQ_{start_date}_au_{end_date}"
    headers = [
        "Catégorie", "Réf. facture", "Client", "Code client",
        "Code produit", "Quantité", "Montant", "Désignation", "Prix unit."
    ]
    columns = [
        "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
        "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
        "Unit_Price"
    ]
    widths = [125, 135, 230, 120, 115, 115, 135, 285, 135]
    image_width = sum(widths) + 60
    margin = 30
    row_height = 34
    title_height = 105
    rows_per_page = 22

    font_path = r"C:\Windows\Fonts\arial.ttf"
    bold_path = r"C:\Windows\Fonts\arialbd.ttf"
    try:
        font = ImageFont.truetype(font_path, 15)
        bold_font = ImageFont.truetype(bold_path, 16)
        title_font = ImageFont.truetype(bold_path, 25)
        small_font = ImageFont.truetype(font_path, 15)
    except OSError:
        font = ImageFont.load_default()
        bold_font = font
        title_font = font
        small_font = font

    def draw_cell(draw, x, y, width, height, value, selected_font, fill):
        draw.rectangle((x, y, x + width, y + height), fill=fill,
                       outline="#C8D2DC", width=1)
        text = _display_value(value).replace("\\n", " ")
        max_width = max(10, width - 10)
        while text and draw.textbbox((0, 0), text, font=selected_font)[2] > max_width:
            text = text[:-2] + "…" if len(text) > 2 else ""
        draw.text((x + 5, y + 8), text, fill="#202020", font=selected_font)

    chunks = [
        details.iloc[index:index + rows_per_page]
        for index in range(0, len(details), rows_per_page)
    ]
    if not chunks:
        chunks = [details.iloc[0:0]]
    page_paths = []

    for page_index, chunk in enumerate(chunks, 1):
        image_height = title_height + (len(chunk) + 1) * row_height + 45
        image = Image.new("RGB", (image_width, image_height), "white")
        draw = ImageDraw.Draw(image)
        draw.text((margin, 18), "BRQ — ÉTAT DES VENTES",
                  fill="#1F4E78", font=title_font)
        draw.text(
            (margin, 58),
            f"Période : du {start_date:%d/%m/%Y} au {end_date:%d/%m/%Y} "
            f"| Détails — page {page_index}/{len(chunks)}",
            fill="#333333", font=small_font
        )
        y = title_height
        x = margin
        for width, label in zip(widths, headers):
            draw_cell(draw, x, y, width, row_height, label, bold_font, "#1F4E78")
            x += width
        y += row_height

        for record in chunk.reindex(columns=columns).itertuples(index=False, name=None):
            x = margin
            category = str(record[0])
            row_fill = "#FFFFFF" if y // row_height % 2 else "#F7F9FB"
            for column_index, (width, value) in enumerate(zip(widths, record)):
                fill = {
                    "CHAIR": "#E2F0D9", "PONTE": "#FFF2CC",
                    "RUMINANTS": "#DDEBF7", "DIVERS": "#E4DFEC"
                }.get(category, row_fill) if column_index == 0 else row_fill
                draw_cell(draw, x, y, width, row_height, value, font, fill)
                x += width
            y += row_height

        if chunk.empty:
            draw.text((margin, y + 8), "Aucun détail pour cette période.",
                      fill="#555555", font=font)

        path = os.path.join(output_dir, f"{base_name}_page_{page_index:03d}.png")
        image.save(path, format="PNG", optimize=True)
        page_paths.append(path)

    recap_height = 230
    recap_image = Image.new("RGB", (image_width, recap_height), "white")
    draw = ImageDraw.Draw(recap_image)
    draw.text((margin, 20), "BRQ — RÉCAPITULATIF CUMULÉ",
              fill="#1F4E78", font=title_font)
    draw.text((margin, 62),
              f"Période du {start_date:%d/%m/%Y} au {end_date:%d/%m/%Y}",
              fill="#333333", font=small_font)
    recap_headers = ["Catégorie", "Jour", "Cumul mois", "Cumul année", "Objectif annuel (%)"]
    recap_content_width = image_width - 2 * margin
    recap_widths = [recap_content_width // 5] * 5
    recap_widths[-1] += recap_content_width - sum(recap_widths)
    y = 105
    x = margin
    for width, label in zip(recap_widths, recap_headers):
        draw_cell(draw, x, y, width, row_height, label, bold_font, "#1F4E78")
        x += width
    y += row_height
    for record in summary.itertuples(index=False, name=None):
        x = margin
        is_total = str(record[0]) == "TOTAL"
        for column, width, value in zip(summary.columns, recap_widths, record):
            draw_cell(draw, x, y, width, row_height, _display_summary_value(column, value),
                      bold_font if is_total else font,
                      "#D9EAF7" if is_total else "#FFFFFF")
            x += width
        y += row_height

    recap_path = os.path.join(
        output_dir, f"{base_name}_page_{len(page_paths) + 1:03d}_recap.png"
    )
    recap_image.crop((0, 0, image_width, y + 20)).save(
        recap_path, format="PNG", optimize=True
    )
    page_paths.append(recap_path)
    return page_paths



def fetch_period_totals(connection, start_date, end_date):
    """Return validated category quantities for one inclusive date period."""
    frame = pd.read_sql(
        RECAP_QUERY,
        connection,
        params={"date1": start_date, "date2": end_date}
    )
    categories = ["CHAIR", "PONTE", "RUMINANTS", "DIVERS"]
    if frame.empty:
        return {category: 0.0 for category in categories}
    row = frame.iloc[0]
    return {
        category: (0.0 if pd.isna(row.get(category)) else float(row[category]))
        for category in categories
    }


def build_cumulative_summary(connection, reference_date):
    """
    Calculate daily, month-to-date and year-to-date totals.

    All three periods are anchored to the selected report end date.
    """
    month_start = reference_date.replace(day=1)
    year_start = dt.date(reference_date.year, 1, 1)

    daily = fetch_period_totals(connection, reference_date, reference_date)
    monthly = fetch_period_totals(connection, month_start, reference_date)
    yearly = fetch_period_totals(connection, year_start, reference_date)

    categories = ["CHAIR", "PONTE", "RUMINANTS", "DIVERS"]
    rows = [
        {
            "Catégorie": category,
            "Jour": daily[category],
            "Cumul mois": monthly[category],
            "Cumul année": yearly[category],
            "Objectif annuel (%)": None,
        }
        for category in categories
    ]
    annual_total = sum(yearly.values())
    rows.append({
        "Catégorie": "TOTAL",
        "Jour": sum(daily.values()),
        "Cumul mois": sum(monthly.values()),
        "Cumul année": annual_total,
        "Objectif annuel (%)": annual_total / ANNUAL_OBJECTIVE * 100,
    })
    return pd.DataFrame(rows)


class BrqApp:
    def __init__(self, root):
        self.root = root
        root.title("BRQ Export Manager")
        root.geometry("850x680")
        root.minsize(760, 600)
        self.messages = queue.Queue()
        self.busy = False
        today = dt.date.today()
        self.sql_var = tk.StringVar(value=DEFAULT_SQL_PATH)
        self.from_var = tk.StringVar(value=today.replace(day=1).isoformat())
        self.to_var = tk.StringVar(value=today.isoformat())
        self.status_var = tk.StringVar(value="Vérification de la base...")
        self.output_var = tk.StringVar(value="Excel")
        self.detail_column_definitions = [
            ("BRQ_Category", "Catégorie"),
            ("Invoice_Reference", "Référence facture"),
            ("Client_Name", "Client"),
            ("Client_Code", "Code client"),
            ("Produit_Code", "Code produit"),
            ("Quantite", "Quantité"),
            ("Total_Amount", "Montant"),
            ("Product_Label", "Désignation"),
            ("Unit_Price", "Prix unitaire"),
        ]
        default_columns = {
            "BRQ_Category", "Invoice_Reference", "Client_Name",
            "Client_Code", "Produit_Code", "Quantite", "Total_Amount"
        }
        self.detail_column_vars = {
            column: tk.BooleanVar(value=column in default_columns)
            for column, _label in self.detail_column_definitions
        }
        self.detail_tree = None
        self.build_ui()
        root.after(100, self.poll_messages)
        root.after(200, self.check_connection)

    def build_ui(self):
        pad = {"padx": 12, "pady": 6}
        head = ttk.Frame(self.root, padding=12)
        head.pack(fill="x")
        ttk.Label(head, text="BRQ Export Manager",
                  font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(head, textvariable=self.status_var).pack(anchor="w")
        db = ttk.LabelFrame(self.root, text="Base de données et sauvegarde SQL", padding=10)
        db.pack(fill="x", **pad)
        row = ttk.Frame(db); row.pack(fill="x")
        ttk.Label(row, text="Fichier SQL :", width=13).pack(side="left")
        ttk.Entry(row, textvariable=self.sql_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir...", command=self.browse_sql).pack(side="left")
        row = ttk.Frame(db); row.pack(fill="x", pady=(8, 0))
        self.update_btn = ttk.Button(row, text="Mettre à jour la base",
                                     command=self.on_update)
        self.update_btn.pack(side="left")
        ttk.Label(row, text="Attention : l'import remplace la base actuelle.",
                  foreground="#9a3412").pack(side="left", padx=10)

        period = ttk.LabelFrame(self.root, text="Période du rapport", padding=10)
        period.pack(fill="x", **pad)
        row = ttk.Frame(period); row.pack(fill="x")
        ttk.Label(row, text="Du :", width=7).pack(side="left")
        ttk.Entry(row, textvariable=self.from_var, width=15).pack(side="left")
        ttk.Label(row, text="Au :", padding=(12, 0, 4, 0)).pack(side="left")
        ttk.Entry(row, textvariable=self.to_var, width=15).pack(side="left")

        shortcuts = ttk.Frame(period); shortcuts.pack(fill="x", pady=(8, 0))
        ttk.Label(shortcuts, text="Raccourcis :").pack(side="left", padx=(0, 8))
        ttk.Button(shortcuts, text="Date", command=self.select_date).pack(side="left", padx=3)
        ttk.Button(shortcuts, text="Mois", command=self.select_month).pack(side="left", padx=3)
        ttk.Button(shortcuts, text="Année", command=self.select_year).pack(side="left", padx=3)

        columns_frame = ttk.LabelFrame(
            self.root, text="Colonnes du détail des ventes", padding=(10, 6)
        )
        columns_frame.pack(fill="x", **pad)
        ttk.Label(
            columns_frame,
            text="Choisissez les colonnes à afficher dans l'aperçu (les options peuvent être modifiées ici).",
            foreground="#555555"
        ).pack(anchor="w", pady=(0, 4))
        checkbox_grid = ttk.Frame(columns_frame)
        checkbox_grid.pack(fill="x")
        for index, (_column, label) in enumerate(self.detail_column_definitions):
            variable = self.detail_column_vars[_column]
            ttk.Checkbutton(
                checkbox_grid, text=label, variable=variable,
                command=self.update_detail_columns
            ).grid(row=index // 5, column=index % 5, sticky="w", padx=6, pady=2)

        actions = ttk.Frame(period); actions.pack(fill="x", pady=(8, 0))
        ttk.Label(actions, text="Sortie :").pack(side="left", padx=(0, 6))
        self.output_combo = ttk.Combobox(
            actions, textvariable=self.output_var,
            values=("Visualisation", "Excel", "PDF", "PNG"),
            state="readonly", width=16
        )
        self.output_combo.pack(side="left", padx=(0, 10))
        self.daily_btn = ttk.Button(
            actions, text="Exécuter", command=self.on_daily
        )
        self.daily_btn.pack(side="left")
        ttk.Button(actions, text="Ouvrir le dossier", command=self.open_folder).pack(side="left", padx=8)

        log = ttk.LabelFrame(self.root, text="Journal", padding=10)
        log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(log, height=13, wrap="word", state="disabled")
        scroll = ttk.Scrollbar(log, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def log(self, message):
        self.messages.put(message)

    def poll_messages(self):
        while True:
            try: message = self.messages.get_nowait()
            except queue.Empty: break
            self.log_text.configure(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(100, self.poll_messages)

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for button in (self.update_btn, self.daily_btn):
            button.configure(state=state)

    def run_worker(self, message, callback, *args):
        if self.busy: return
        self.set_busy(True); self.clear_log(); self.log(message)
        def task():
            try:
                callback(*args)
                self.log("Terminé avec succès.")
            except Exception as exc:
                self.log(f"ERREUR : {exc}")
            finally:
                self.root.after(0, lambda: self.set_busy(False))
        threading.Thread(target=task, daemon=True).start()

    def check_connection(self):
        def check():
            try:
                conn = connect_database(); conn.close()
                value = f"Base connectée : {DB_CONFIG['database']}"
            except Exception as exc:
                value = f"Base non disponible : {exc}"
            self.root.after(0, lambda: self.status_var.set(value))
        threading.Thread(target=check, daemon=True).start()

    def _period_anchor_date(self):
        """Use the editable To date as the reference for period shortcuts."""
        try:
            return parse_date(self.to_var.get())
        except ValueError as exc:
            messagebox.showerror("Date invalide", str(exc))
            return None

    def select_date(self):
        anchor = self._period_anchor_date()
        if anchor is not None:
            self.from_var.set(anchor.isoformat())
            self.to_var.set(anchor.isoformat())

    def select_month(self):
        anchor = self._period_anchor_date()
        if anchor is None:
            return
        first_day = anchor.replace(day=1)
        today = dt.date.today()
        last_day = anchor if (anchor.year, anchor.month) == (today.year, today.month) else (
            (first_day.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
            - dt.timedelta(days=1)
        )
        self.from_var.set(first_day.isoformat())
        self.to_var.set(last_day.isoformat())

    def select_year(self):
        anchor = self._period_anchor_date()
        if anchor is None:
            return
        today = dt.date.today()
        end_day = anchor if anchor.year == today.year else dt.date(anchor.year, 12, 31)
        self.from_var.set(dt.date(anchor.year, 1, 1).isoformat())
        self.to_var.set(end_day.isoformat())

    def browse_sql(self):
        path = filedialog.askopenfilename(
            title="Choisir le fichier SQL",
            initialdir=os.path.dirname(self.sql_var.get()) or WORK_FOLDER,
            filetypes=[("Fichiers SQL", "*.sql"), ("Tous les fichiers", "*.*")])
        if path: self.sql_var.set(path)

    def on_update(self):
        path = self.sql_var.get().strip()
        if not os.path.isfile(path):
            messagebox.showerror("Fichier SQL", "Sélectionnez un fichier SQL valide.")
            return
        ok = messagebox.askyesno(
            "Remplacement de la base",
            f"La base {DB_CONFIG['database']} sera supprimée puis recréée à partir du SQL.\\n\\n"
            "Toutes les données actuelles de cette base seront remplacées. Continuer ?",
            icon="warning")
        if ok:
            self.run_worker("Mise à jour de la base...", self.update_database, path)

    def update_database(self, path):
        self.log(f"Import du fichier : {path}")
        import_sql_into_database(path)
        self.log("Import SQL terminé.")
        self.check_connection()

    def _validated_period(self):
        try:
            start = parse_date(self.from_var.get())
            end = parse_date(self.to_var.get())
        except ValueError as exc:
            messagebox.showerror("Dates invalides", str(exc))
            return None
        if start > end:
            messagebox.showerror(
                "Période invalide",
                "La date de début doit être antérieure ou égale à la date de fin."
            )
            return None
        return start, end

    def on_daily(self):
        period = self._validated_period()
        if period is None:
            return
        start, end = period
        mode = self.output_var.get()

        output_path = None
        if mode == "PDF":
            output_path = filedialog.asksaveasfilename(
                title="Enregistrer le rapport PDF",
                initialdir=WORK_FOLDER,
                initialfile=f"BRQ_{start}_au_{end}.pdf",
                defaultextension=".pdf",
                filetypes=[("Fichier PDF", "*.pdf")]
            )
            if not output_path:
                return
        elif mode == "PNG":
            output_path = filedialog.askdirectory(
                title="Choisir le dossier des images PNG",
                initialdir=WORK_FOLDER,
                mustexist=False
            )
            if not output_path:
                return

        self.run_worker(
            f"Génération du BRQ ({mode}) du {start} au {end}...",
            self.generate_report, start, end, mode, output_path
        )

    def generate_report(self, start, end, mode, output_path=None):
        """Fetch report data once, then route it to the selected presentation."""
        self.log("Connexion à la base existante : aucun import SQL.")
        self.log(f"Calcul du BRQ et des cumuls à la date de référence {end}.")
        conn = connect_database()
        try:
            details = pd.read_sql(
                BRQ_QUERY, conn, params={"date1": start, "date2": end}
            )
            summary = build_cumulative_summary(conn, end)
        finally:
            conn.close()

        if mode == "Visualisation":
            self.log(
                f"Aperçu prêt : {len(details)} ligne(s) détaillée(s). "
                "Aucun fichier ne sera créé."
            )
            self.root.after(
                0, lambda: self.show_preview(details, summary, start, end)
            )
        elif mode == "Excel":
            path = os.path.join(WORK_FOLDER, f"BRQ_{start}_au_{end}.xlsx")
            write_combined_report(details, summary, start, end, path)
            self.log(f"Rapport Excel créé : {path}")
        elif mode == "PDF":
            write_report_pdf(details, summary, start, end, output_path)
            self.log(f"Rapport PDF créé : {output_path}")
        elif mode == "PNG":
            paths = write_report_png_pages(
                details, summary, start, end, output_path
            )
            self.log(
                f"{len(paths)} image(s) PNG créée(s) dans : {output_path}"
            )
            for path in paths:
                self.log(path)
        else:
            raise ValueError(f"Mode de sortie inconnu : {mode}")

    def show_preview(self, details, summary, start, end):
        """Display the same fetched report data in a scrollable application window."""
        window = tk.Toplevel(self.root)
        window.title(f"Aperçu BRQ — {start} au {end}")
        window.geometry("1250x700")
        window.minsize(850, 500)

        heading = ttk.Frame(window, padding=10)
        heading.pack(fill="x")
        ttk.Label(
            heading, text="BRQ — ÉTAT DES VENTES",
            font=("Segoe UI", 15, "bold")
        ).pack(anchor="w")
        ttk.Label(
            heading,
            text=f"Période : du {start:%d/%m/%Y} au {end:%d/%m/%Y} "
                 f" | {len(details)} ligne(s)"
        ).pack(anchor="w")

        notebook = ttk.Notebook(window)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        detail_tab = ttk.Frame(notebook, padding=6)
        summary_tab = ttk.Frame(notebook, padding=6)
        notebook.add(detail_tab, text="Détail des ventes")
        notebook.add(summary_tab, text="Récapitulatif cumulé")

        detail_columns = [
            "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
            "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
            "Unit_Price"
        ]
        detail_labels = [
            "Catégorie", "Référence facture", "Client", "Code client",
            "Code produit", "Quantité", "Montant", "Désignation",
            "Prix unitaire"
        ]

        def create_tree(parent, columns, labels, widths):
            container = ttk.Frame(parent)
            container.pack(fill="both", expand=True)
            tree = ttk.Treeview(
                container, columns=columns, show="headings",
                displaycolumns=columns
            )
            for column, label, width in zip(columns, labels, widths):
                tree.heading(column, text=label)
                tree.column(column, width=width, minwidth=70, stretch=True)
            vertical = ttk.Scrollbar(
                container, orient="vertical", command=tree.yview
            )
            horizontal = ttk.Scrollbar(
                container, orient="horizontal", command=tree.xview
            )
            tree.configure(
                yscrollcommand=vertical.set, xscrollcommand=horizontal.set
            )
            tree.grid(row=0, column=0, sticky="nsew")
            vertical.grid(row=0, column=1, sticky="ns")
            horizontal.grid(row=1, column=0, sticky="ew")
            container.rowconfigure(0, weight=1)
            container.columnconfigure(0, weight=1)
            return tree

        column_widths = [95, 125, 190, 110, 100, 95, 110, 220, 110]
        detail_tree = create_tree(
            detail_tab, detail_columns, detail_labels, column_widths
        )
        self.detail_tree = detail_tree
        self.update_detail_columns()
        for record in details.reindex(columns=detail_columns).itertuples(
            index=False, name=None
        ):
            detail_tree.insert(
                "", "end",
                values=tuple(_display_value(value) for value in record)
            )

        summary_columns = list(summary.columns)
        summary_tree = create_tree(
            summary_tab, summary_columns, summary_columns,
            [220, 180, 180, 180, 190]
        )
        for record in summary.itertuples(index=False, name=None):
            summary_tree.insert(
                "", "end",
                values=tuple(
                    _display_summary_value(column, value)
                    for column, value in zip(summary_columns, record)
                )
            )

    def update_detail_columns(self):
        """Apply the main-window column selection to an open detail preview."""
        if self.detail_tree is None or not self.detail_tree.winfo_exists():
            return
        selected = [
            column for column, _label in self.detail_column_definitions
            if self.detail_column_vars[column].get()
        ]
        # Keep at least one visible column so the table never appears empty.
        self.detail_tree["displaycolumns"] = selected or ("BRQ_Category",)

    def open_folder(self):
        try:
            os.makedirs(WORK_FOLDER, exist_ok=True)
            os.startfile(WORK_FOLDER)
        except Exception as exc:
            messagebox.showerror("Ouverture du dossier", str(exc))

def main():
    os.makedirs(WORK_FOLDER, exist_ok=True)
    root = tk.Tk()
    BrqApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
