"""Tkinter interface for BRQ Export Manager."""
import calendar
import datetime as dt
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from config import DB_CONFIG, DEFAULT_SQL_PATH, OUTPUT_FOLDER, WORK_FOLDER
from database import connect_database, import_sql_into_database
from exporters import write_combined_report, write_report_pdf, write_report_png_pages, _display_value, _display_summary_value
from report_data import build_cumulative_summary
from utils import parse_date

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
        self.month_var = tk.StringVar(value=calendar.month_name[today.month])
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
        ttk.Label(shortcuts, text="Month:").pack(side="left", padx=(10, 4))
        self.month_combo = ttk.Combobox(
            shortcuts, textvariable=self.month_var,
            values=tuple(calendar.month_name[1:]), state="readonly", width=12
        )
        self.month_combo.pack(side="left", padx=3)
        self.month_combo.bind("<<ComboboxSelected>>", self.apply_selected_month)
        ttk.Button(shortcuts, text="Année", command=self.select_year).pack(side="left", padx=3)

        columns_frame = ttk.LabelFrame(
            self.root, text="Colonnes du détail des ventes", padding=(10, 6)
        )
        columns_frame.pack(fill="x", **pad)
        ttk.Label(
            columns_frame,
            text="Choisissez les colonnes à afficher dans l'aperçu et dans les exports Excel, PDF et PNG.",
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
            self.month_var.set(calendar.month_name[anchor.month])

    def apply_selected_month(self, _event=None):
        """Change month while preserving each date's year and day when possible."""
        month_name = self.month_var.get()
        try:
            month = tuple(calendar.month_name).index(month_name)
            start = parse_date(self.from_var.get())
            end = parse_date(self.to_var.get())
        except (ValueError, IndexError):
            messagebox.showerror(
                "Dates invalides",
                "Vérifiez les dates de début et de fin avant de choisir un mois."
            )
            return

        def change_month(value):
            last_day = calendar.monthrange(value.year, month)[1]
            return value.replace(month=month, day=min(value.day, last_day))

        self.from_var.set(change_month(start).isoformat())
        self.to_var.set(change_month(end).isoformat())

    def select_year(self):
        anchor = self._period_anchor_date()
        if anchor is None:
            return
        today = dt.date.today()
        end_day = anchor if anchor.year == today.year else dt.date(anchor.year, 12, 31)
        self.from_var.set(dt.date(anchor.year, 1, 1).isoformat())
        self.to_var.set(end_day.isoformat())
        self.month_var.set(calendar.month_name[anchor.month])

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

        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        selected_columns = self.get_selected_detail_columns()
        output_path = None
        if mode == "PDF":
            output_path = filedialog.asksaveasfilename(
                title="Enregistrer le rapport PDF",
                initialdir=OUTPUT_FOLDER,
                initialfile=f"BRQ_{start}_au_{end}.pdf",
                defaultextension=".pdf",
                filetypes=[("Fichier PDF", "*.pdf")]
            )
            if not output_path:
                return
        elif mode == "PNG":
            output_path = filedialog.askdirectory(
                title="Choisir le dossier des images PNG",
                initialdir=OUTPUT_FOLDER,
                mustexist=False
            )
            if not output_path:
                return

        self.run_worker(
            f"Génération du BRQ ({mode}) du {start} au {end}...",
            self.generate_report, start, end, mode, output_path, selected_columns
        )

    def generate_report(self, start, end, mode, output_path=None, selected_columns=None):
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
                0, lambda: self.show_preview(details, summary, start, end, selected_columns)
            )
        elif mode == "Excel":
            path = os.path.join(OUTPUT_FOLDER, f"BRQ_{start}_au_{end}.xlsx")
            write_combined_report(details, summary, start, end, path, selected_columns)
            self.log(f"Rapport Excel créé : {path}")
        elif mode == "PDF":
            write_report_pdf(details, summary, start, end, output_path, selected_columns)
            self.log(f"Rapport PDF créé : {output_path}")
        elif mode == "PNG":
            paths = write_report_png_pages(
                details, summary, start, end, output_path, selected_columns
            )
            self.log(
                f"{len(paths)} image(s) PNG créée(s) dans : {output_path}"
            )
            for path in paths:
                self.log(path)
        else:
            raise ValueError(f"Mode de sortie inconnu : {mode}")

    def show_preview(self, details, summary, start, end, selected_columns=None):
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
        visible_columns = selected_columns or ["BRQ_Category"]
        detail_tree["displaycolumns"] = visible_columns
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

    def get_selected_detail_columns(self):
        """Return selected detail columns in their original report order."""
        selected = [
            column for column, _label in self.detail_column_definitions
            if self.detail_column_vars[column].get()
        ]
        # Keep at least one visible column so the table never appears empty.
        return selected or ["BRQ_Category"]

    def update_detail_columns(self):
        """Apply the main-window column selection to an open detail preview."""
        if self.detail_tree is None or not self.detail_tree.winfo_exists():
            return
        self.detail_tree["displaycolumns"] = self.get_selected_detail_columns()

    def open_folder(self):
        try:
            os.makedirs(OUTPUT_FOLDER, exist_ok=True)
            os.startfile(OUTPUT_FOLDER)
        except Exception as exc:
            messagebox.showerror("Ouverture du dossier", str(exc))
