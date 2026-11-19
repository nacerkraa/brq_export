"""
État des ventes — GUI

Pick a period (Du / Au), then export the daily BRQ quantities as
Excel, PDF or PNG, with the title:

    ÉTAT DES VENTES
    Période : du 01/10/2026 au 31/10/2026

Install (once):
    py -m pip install mysql-connector-python pandas openpyxl matplotlib tkcalendar

    (tkcalendar is optional: without it, dates are typed as JJ/MM/AAAA.)

Run:
    py export_daily_brq_gui.py
"""

import os
import subprocess
import sys
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import matplotlib
matplotlib.use("Agg")  # draw to files only; no clash with tkinter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

import mysql.connector
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from tkcalendar import DateEntry
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False

# ---------------------------------------------------------------------------
# Edit these settings to match your XAMPP/MariaDB configuration.
# ---------------------------------------------------------------------------
DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "uab_ouledhamla_com26",
}

TITLE = "ÉTAT DES VENTES"
COMPANY_LINES = ["UAB OULED HAMLA", "DEPT: Controle de Gestion"]   # top left of every export
UNIT_LINES = ["U=DA", "U=QT"]             # shown at the top right of every export
SIGNATURE_TEXT = "Department Controle de Gestion"   # bottom right, under the table
UNIT_ROW_OFFSET = 4                          # units are moved this many rows down
TITLE_ROW_OFFSET = 2                         # title + period are moved this many rows down
QUANTITY_COLUMNS = ["CHAIR", "PONTE", "RUMINANTS", "DIVERS", "TOTAL"]
AMOUNT_COLUMN = "C.A"                        # total value (DZD) of the day
VALUE_COLUMNS = QUANTITY_COLUMNS + [AMOUNT_COLUMN]
HEADER_COLOR = "1F4E78"
TOTAL_COLOR = "DDEBF7"
ZEBRA_COLOR = "F5F9FD"

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
    END) AS TOTAL,
    SUM(CASE
    WHEN (f.Libelle = 'VOLAILLE' AND sf.Libelle IN ('CHAIR', 'PONTE'))
      OR (f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN'))
      OR (f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS')
    THEN SIGN(d.Quantite) * ABS(d.Valeur)
    ELSE 0
END) AS "C.A"
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


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------
def fetch_data(start_date, end_date):
    connection = mysql.connector.connect(**DB_CONFIG)
    try:
        cursor = connection.cursor()
        cursor.execute(SQL, (start_date, end_date))
        rows = cursor.fetchall()
        columns = [d[0] for d in cursor.description]
    finally:
        connection.close()

    df = pd.DataFrame(rows, columns=columns)
    for column in VALUE_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    return df


def period_text(start_date, end_date):
    return f"Période : du {start_date:%d/%m/%Y} au {end_date:%d/%m/%Y}"


def default_file_name(start_date, end_date, extension):
    return f"BRQ_Daily_{start_date:%Y%m%d}_{end_date:%Y%m%d}.{extension}"


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def export_excel(df, start_date, end_date, path):
    # rows 1-2 company (left), rows 3-4 title, units on rows 5-6 (right), then the table
    header_row = 4 + UNIT_ROW_OFFSET       # 8
    first_data_row = header_row + 1        # 9
    last_data_row = first_data_row + len(df) - 1
    total_row = last_data_row + 1
    last_col = get_column_letter(1 + len(VALUE_COLUMNS))

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Daily BRQ", index=False,
                    startrow=header_row - 1)
        ws = writer.sheets["Daily BRQ"]

        # Title block
        title_last_col = get_column_letter(len(VALUE_COLUMNS))   # last column before the units
        # (title starts at column C so the left text in A/B has room)
        t1 = 1 + TITLE_ROW_OFFSET      # title row (3)
        t2 = 2 + TITLE_ROW_OFFSET      # period row (4)
        ws[f"C{t1}"] = TITLE
        ws[f"C{t2}"] = period_text(start_date, end_date)
        ws.merge_cells(f"C{t1}:{title_last_col}{t1}")
        ws.merge_cells(f"C{t2}:{title_last_col}{t2}")
        ws[f"C{t1}"].font = Font(bold=True, size=14, color=HEADER_COLOR)
        ws[f"C{t2}"].font = Font(size=11)
        ws[f"C{t1}"].alignment = Alignment(horizontal="center")
        ws[f"C{t2}"].alignment = Alignment(horizontal="center")

        # Company / department, top left
        for row_no, text in enumerate(COMPANY_LINES, start=1):
            cell = ws[f"A{row_no}"]
            cell.value = text
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="left")

        # Units, right side, moved down
        for row_no, text in enumerate(UNIT_LINES, start=1 + UNIT_ROW_OFFSET):
            cell = ws[f"{last_col}{row_no}"]
            cell.value = text
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="right")

        # Header
        for cell in ws[header_row]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=HEADER_COLOR)
            cell.alignment = Alignment(horizontal="center")

        # Dates
        for row in ws.iter_rows(min_row=first_data_row, max_row=last_data_row, max_col=1):
            for cell in row:
                cell.number_format = "dd/mm/yyyy"
                cell.alignment = Alignment(horizontal="center")

        # Totals row
        ws.cell(row=total_row, column=1, value="TOTAL")
        for col in range(2, 2 + len(VALUE_COLUMNS)):
            letter = get_column_letter(col)
            ws.cell(row=total_row, column=col,
                    value=f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")

        for row in ws.iter_rows(min_row=first_data_row, max_row=total_row,
                                min_col=2, max_col=1 + len(VALUE_COLUMNS)):
            for cell in row:
                cell.number_format = '#,##0.00;[Red]-#,##0.00'

        top_border = Border(top=Side(style="medium"))
        for cell in ws[total_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=TOTAL_COLOR)
            cell.border = top_border
        ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="center")

        # Signature, bottom right (columns D-F, not the very last column)
        sign_row = total_row + 3
        ws[f"D{sign_row}"] = SIGNATURE_TEXT
        ws.merge_cells(f"D{sign_row}:F{sign_row}")
        ws[f"D{sign_row}"].font = Font(bold=True, size=10)
        ws[f"D{sign_row}"].alignment = Alignment(horizontal="center")

        ws.freeze_panes = f"A{first_data_row}"
        ws.auto_filter.ref = f"A{header_row}:{last_col}{last_data_row}"

        for column, width in {"A": 16, "B": 15, "C": 15,
                              "D": 17, "E": 15, "F": 15, "G": 20}.items():
            ws.column_dimensions[column].width = width

        # Print setup: one page wide
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = f"{header_row}:{header_row}"


# ---------------------------------------------------------------------------
# PDF / PNG (drawn with matplotlib)
# ---------------------------------------------------------------------------
A4_W, A4_H = 8.27, 11.69
ROW_H = 0.28              # inches (slightly reduced so 31 days still fit on one A4 page)
COL_WIDTHS = [0.14, 0.13, 0.12, 0.14, 0.11, 0.13, 0.23]   # Date + 6 value columns
TITLE_BLOCK = 1.35 + TITLE_ROW_OFFSET * ROW_H        # inches
ROWS_PER_PAGE = 31          # a 31-day month fits on one page, with header and total


def fmt_number(value):
    """1234.5 -> '1 234,50' (French style, same as a French Excel)."""
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def build_table_rows(df):
    dates = pd.to_datetime(df["Date_mvt"]).dt.strftime("%d/%m/%Y")
    body = [[d] + [fmt_number(v) for v in row]
            for d, row in zip(dates, df[VALUE_COLUMNS].itertuples(index=False))]
    totals = ["TOTAL"] + [fmt_number(df[c].sum()) for c in VALUE_COLUMNS]
    return body, totals


def draw_page(body_rows, total_row, start_date, end_date,
              fig_h, page_no=1, page_count=1):
    fig = plt.figure(figsize=(A4_W, fig_h))

    title_shift = TITLE_ROW_OFFSET * ROW_H
    fig.text(0.5, 1 - (0.45 + title_shift) / fig_h, TITLE, ha="center", va="center",
             fontsize=15, fontweight="bold", color="#" + HEADER_COLOR)
    fig.text(0.5, 1 - (0.85 + title_shift) / fig_h, period_text(start_date, end_date),
             ha="center", va="center", fontsize=11)

    # Company / department, top left
    for i, text in enumerate(COMPANY_LINES):
        fig.text(0.07, 1 - (0.30 + 0.22 * i) / fig_h, text, ha="left", va="center",
                 fontsize=8.5, fontweight="bold")

    # Units, top right, moved down by UNIT_ROW_OFFSET rows
    for i, text in enumerate(UNIT_LINES):
        y = 0.30 + UNIT_ROW_OFFSET * ROW_H + 0.22 * i
        fig.text(0.93, 1 - y / fig_h, text, ha="right", va="center",
                 fontsize=9, fontweight="bold")

    cell_text = list(body_rows) + ([total_row] if total_row else [])
    n_rows = len(cell_text) + 1                       # + header
    table_h = n_rows * ROW_H

    ax = fig.add_axes([0.07, 0, 0.86, 1])
    ax.axis("off")
    table = ax.table(
        cellText=cell_text,
        colLabels=["Date"] + VALUE_COLUMNS,
        colWidths=COL_WIDTHS,
        bbox=[0, 1 - (TITLE_BLOCK + table_h) / fig_h, 1, table_h / fig_h],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)

    last_row = n_rows - 1
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor("#BFBFBF")
        align = "center" if (c == 0 or r == 0) else "right"
        cell._loc = align                              # older matplotlib
        cell.get_text().set_horizontalalignment(align)  # newer matplotlib
        if r == 0:
            cell.set_facecolor("#" + HEADER_COLOR)
            cell.set_text_props(color="white", fontweight="bold")
        elif total_row and r == last_row:
            cell.set_facecolor("#" + TOTAL_COLOR)
            cell.set_text_props(fontweight="bold")
        elif r % 2 == 0:
            cell.set_facecolor("#" + ZEBRA_COLOR)

    # Signature, bottom right of the table (only where the total row is)
    if total_row:
        fig.text(0.76, 1 - (TITLE_BLOCK + table_h + 0.30) / fig_h, SIGNATURE_TEXT,
                 ha="center", va="center", fontsize=10, fontweight="bold")

    if page_count > 1:
        fig.text(0.93, 0.015, f"Page {page_no}/{page_count}",
                 ha="right", va="bottom", fontsize=8, color="gray")
    return fig


def export_pdf(df, start_date, end_date, path):
    body, totals = build_table_rows(df)
    chunks = [body[i:i + ROWS_PER_PAGE] for i in range(0, len(body), ROWS_PER_PAGE)]
    with PdfPages(path) as pdf:
        for i, chunk in enumerate(chunks, start=1):
            is_last = (i == len(chunks))
            fig = draw_page(chunk, totals if is_last else None,
                            start_date, end_date, A4_H,
                            page_no=i, page_count=len(chunks))
            pdf.savefig(fig)
            plt.close(fig)


def export_png(df, start_date, end_date, path):
    body, totals = build_table_rows(df)
    fig_h = TITLE_BLOCK + (len(body) + 2) * ROW_H + 0.9   # header + total + signature + margin
    fig = draw_page(body, totals, start_date, end_date, fig_h)
    dpi = min(200, 30000 / fig_h)          # keep very long periods under image limits
    fig.savefig(path, dpi=dpi, facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


class App:
    FORMATS = {
        "xlsx": ("Excel", export_excel, [("Excel", "*.xlsx")]),
        "pdf": ("PDF", export_pdf, [("PDF", "*.pdf")]),
        "png": ("PNG", export_png, [("Image PNG", "*.png")]),
    }

    def __init__(self, root):
        self.root = root
        root.title("État des ventes")
        root.resizable(False, False)

        pad = {"padx": 12, "pady": 6}

        tk.Label(root, text=TITLE, font=("", 11, "bold"), fg="#" + HEADER_COLOR
                 ).pack(padx=12, pady=(12, 2))

        dates = tk.Frame(root)
        dates.pack(**pad)
        self.start_widget = self._make_date_widget(dates, "Du :", date.today().replace(day=1), 0)
        self.end_widget = self._make_date_widget(dates, "Au :", date.today(), 2)

        if not HAS_CALENDAR:
            tk.Label(root, text="Format des dates : JJ/MM/AAAA", fg="gray").pack()

        buttons = tk.Frame(root)
        buttons.pack(**pad)
        colors = {"xlsx": "#2e7d32", "pdf": "#c62828", "png": "#1565c0"}
        for key, (label, _, _) in self.FORMATS.items():
            tk.Button(buttons, text=f"Exporter en {label}", width=16,
                      bg=colors[key], fg="white",
                      command=lambda k=key: self.on_export(k)).pack(side="left", padx=4)

        self.status = tk.Label(root, text="", fg="gray")
        self.status.pack(padx=12, pady=(0, 12))

    def _make_date_widget(self, parent, label, initial, column):
        tk.Label(parent, text=label).grid(row=0, column=column, padx=(8, 4))
        if HAS_CALENDAR:
            widget = DateEntry(parent, width=12, date_pattern="dd/mm/yyyy")
            widget.set_date(initial)
        else:
            widget = ttk.Entry(parent, width=12)
            widget.insert(0, initial.strftime("%d/%m/%Y"))
        widget.grid(row=0, column=column + 1, padx=(0, 8))
        return widget

    @staticmethod
    def _read_date(widget):
        if HAS_CALENDAR:
            return widget.get_date()
        return datetime.strptime(widget.get().strip(), "%d/%m/%Y").date()

    def on_export(self, fmt):
        label, exporter, filetypes = self.FORMATS[fmt]

        try:
            start_date = self._read_date(self.start_widget)
            end_date = self._read_date(self.end_widget)
        except ValueError:
            messagebox.showerror("Erreur", "Date invalide. Utilisez le format JJ/MM/AAAA.")
            return
        if end_date < start_date:
            messagebox.showerror("Erreur", "La date de fin doit être après la date de début.")
            return

        path = filedialog.asksaveasfilename(
            title=f"Enregistrer en {label}",
            defaultextension=f".{fmt}",
            initialfile=default_file_name(start_date, end_date, fmt),
            filetypes=filetypes,
        )
        if not path:
            return

        self.status.config(text="Traitement en cours...")
        self.root.update_idletasks()
        try:
            df = fetch_data(start_date, end_date)
            if df.empty:
                self.status.config(text="")
                messagebox.showinfo("Aucune donnée", "Aucun mouvement trouvé pour cette période.")
                return
            exporter(df, start_date, end_date, path)
        except mysql.connector.Error as error:
            self.status.config(text="")
            messagebox.showerror(
                "Erreur base de données",
                f"{error}\n\nVérifiez DB_CONFIG et que MariaDB (XAMPP) est démarré.")
            return
        except Exception as error:
            self.status.config(text="")
            messagebox.showerror("Erreur", str(error))
            return

        self.status.config(text=f"Fichier créé : {Path(path).name}")
        if messagebox.askyesno("Terminé", f"Fichier créé :\n{path}\n\nL'ouvrir maintenant ?"):
            open_file(path)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
