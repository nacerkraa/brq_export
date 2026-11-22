"""Excel, PDF and PNG report generation."""
import os
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

def write_combined_report(details, summary, start_date, end_date, output_path, selected_columns=None):
    """Write the selected BRQ detail columns and the complete cumulative recap."""
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

    all_detail_columns = [
        "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
        "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
        "Unit_Price"
    ]
    all_detail_headers = [
        "Catégorie", "Référence facture", "Client", "Code client",
        "Code produit", "Quantité", "Montant", "Désignation",
        "Prix unitaire"
    ]
    selected = set(selected_columns) if selected_columns is not None else set(all_detail_columns)
    selected_pairs = [
        (column, label)
        for column, label in zip(all_detail_columns, all_detail_headers)
        if column in selected
    ] or [(all_detail_columns[0], all_detail_headers[0])]
    detail_columns = [column for column, _label in selected_pairs]
    detail_headers = [label for _column, label in selected_pairs]
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
            elif detail_columns[col - 1] in ("Quantite", "Total_Amount", "Unit_Price"):
                cell.number_format = '#,##0.00;[Red]-#,##0.00'
            if detail_columns[col - 1] in ("Invoice_Reference", "Client_Code", "Produit_Code"):
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

    width_by_column = {
        "BRQ_Category": 16, "Invoice_Reference": 22, "Client_Name": 30,
        "Client_Code": 16, "Produit_Code": 15, "Quantite": 14,
        "Total_Amount": 17, "Product_Label": 34, "Unit_Price": 16
    }
    for index, column in enumerate(detail_columns, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width_by_column[column]

    sheet.freeze_panes = "A5"
    if detail_count:
        last_detail_column = get_column_letter(len(detail_columns))
        sheet.auto_filter.ref = (
            f"A{detail_header_row}:{last_detail_column}{detail_header_row + detail_count}"
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
    sheet.print_area = (
        f"A1:{get_column_letter(max(len(detail_columns), len(summary_headers)))}"
        f"{summary_title_row + len(summary) + 1}"
    )

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


def write_report_pdf(details, summary, start_date, end_date, output_path, selected_columns=None):
    """Create a paginated PDF using selected detail columns and the complete recap."""
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
    all_headers = [
        "Catégorie", "Référence facture", "Client", "Code client",
        "Code produit", "Quantité", "Montant", "Désignation",
        "Prix unitaire"
    ]
    all_detail_columns = [
        "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
        "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
        "Unit_Price"
    ]
    selected = set(selected_columns) if selected_columns is not None else set(all_detail_columns)
    selected_pairs = [
        (column, label)
        for column, label in zip(all_detail_columns, all_headers)
        if column in selected
    ] or [(all_detail_columns[0], all_headers[0])]
    detail_columns = [column for column, _label in selected_pairs]
    headers = [label for _column, label in selected_pairs]
    table_data = [[Paragraph(escape(label), header_style) for label in headers]]
    for record in details.reindex(columns=detail_columns).itertuples(index=False, name=None):
        table_data.append([
            Paragraph(escape(_display_value(value)), cell_style)
            for value in record
        ])

    available_width = page_width - 28 * mm
    proportion_by_column = {
        "BRQ_Category": 0.09, "Invoice_Reference": 0.12, "Client_Name": 0.18,
        "Client_Code": 0.09, "Produit_Code": 0.09, "Quantite": 0.07,
        "Total_Amount": 0.10, "Product_Label": 0.16, "Unit_Price": 0.10
    }
    proportions = [proportion_by_column[column] for column in detail_columns]
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
    if "BRQ_Category" in detail_columns:
        category_column_index = detail_columns.index("BRQ_Category")
        for row_index, category in enumerate(details["BRQ_Category"], 1):
            fill = fills.get(str(category))
            if fill:
                detail_commands.append(
                    ("BACKGROUND", (category_column_index, row_index),
                     (category_column_index, row_index), colors.HexColor(fill))
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


def write_report_png_pages(details, summary, start_date, end_date, output_dir, selected_columns=None):
    """Export PNG pages using selected detail columns and the complete recap."""
    from PIL import Image, ImageDraw, ImageFont

    os.makedirs(output_dir, exist_ok=True)
    base_name = f"BRQ_{start_date}_au_{end_date}"
    all_headers = [
        "Catégorie", "Réf. facture", "Client", "Code client",
        "Code produit", "Quantité", "Montant", "Désignation", "Prix unit."
    ]
    all_columns = [
        "BRQ_Category", "Invoice_Reference", "Client_Name", "Client_Code",
        "Produit_Code", "Quantite", "Total_Amount", "Product_Label",
        "Unit_Price"
    ]
    all_widths = [125, 135, 230, 120, 115, 115, 135, 285, 135]
    selected = set(selected_columns) if selected_columns is not None else set(all_columns)
    selected_indices = [index for index, column in enumerate(all_columns) if column in selected]
    if not selected_indices:
        selected_indices = [0]
    columns = [all_columns[index] for index in selected_indices]
    headers = [all_headers[index] for index in selected_indices]
    widths = [all_widths[index] for index in selected_indices]
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
            values_by_column = dict(zip(columns, record))
            category = str(values_by_column.get("BRQ_Category", ""))
            row_fill = "#FFFFFF" if y // row_height % 2 else "#F7F9FB"
            for column_index, (column, width, value) in enumerate(zip(columns, widths, record)):
                fill = {
                    "CHAIR": "#E2F0D9", "PONTE": "#FFF2CC",
                    "RUMINANTS": "#DDEBF7", "DIVERS": "#E4DFEC"
                }.get(category, row_fill) if column == "BRQ_Category" else row_fill
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
