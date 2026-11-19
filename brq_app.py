"""Consolidated BRQ desktop application.
Dependencies: pip install pymysql pandas openpyxl
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

WORK_FOLDER = r"C:\BRQ"
DEFAULT_SQL_PATH = os.path.join(WORK_FOLDER, "UAB_OULEDHAMLA_COM26.sql")
CLEANED_SQL_PATH = os.path.join(WORK_FOLDER, "_cleaned_import.sql")
MYSQL_EXE_PATH = r"C:\xampp\mysql\bin\mysql.exe"
DB_CONFIG = {"host": "localhost", "user": "root", "password": "",
             "database": "UAB_OULEDHAMLA_COM26", "port": 3306,
             "charset": "utf8mb4", "connect_timeout": 5}

BRQ_QUERY = """
SELECT
 CASE
  WHEN f.Libelle='VOLAILLE' AND sf.Libelle='CHAIR' THEN 'CHAIR'
  WHEN f.Libelle='VOLAILLE' AND sf.Libelle='PONTE' THEN 'PONTE'
  WHEN f.Libelle='RUMINANTS' AND sf.Libelle IN ('BOVIN','OVIN') THEN 'RUMINANTS'
  WHEN f.Libelle='DIVERS' AND sf.Libelle='DIVERS' THEN 'DIVERS'
 END AS BRQ_Category,
 e.Date AS Date_mvt, e.Libelle_Client AS Client_Name, e.Code_FC AS Code,
 d.Produit_Code, d.Quantite, d.Valeur AS Total_Amount
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
 END, e.Libelle_Client, e.Code, d.Produit_Code
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

def write_excel(dataframe, output_path, sheet_name):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
        sheet = writer.sheets[sheet_name]
        for index, name in enumerate(dataframe.columns, 1):
            lengths = dataframe[name].fillna("").astype(str).map(len).tolist()
            letter = sheet.cell(row=1, column=index).column_letter
            sheet.column_dimensions[letter].width = min(max([len(name), *lengths])+2, 50)

class BrqApp:
    def __init__(self, root):
        self.root = root
        root.title("BRQ Export Manager")
        root.geometry("750x570")
        root.minsize(690, 510)
        self.messages = queue.Queue()
        self.busy = False
        today = dt.date.today()
        self.sql_var = tk.StringVar(value=DEFAULT_SQL_PATH)
        self.from_var = tk.StringVar(value=today.replace(day=1).isoformat())
        self.to_var = tk.StringVar(value=today.isoformat())
        self.status_var = tk.StringVar(value="Vérification de la base...")
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

        actions = ttk.Frame(period); actions.pack(fill="x", pady=(8, 0))
        self.daily_btn = ttk.Button(actions, text="Générer le BRQ", command=self.on_daily)
        self.daily_btn.pack(side="left")
        self.recap_btn = ttk.Button(actions, text="Exporter le récap", command=self.on_recap)
        self.recap_btn.pack(side="left", padx=8)
        ttk.Button(actions, text="Ouvrir le dossier", command=self.open_folder).pack(side="left")

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
        for button in (self.update_btn, self.daily_btn, self.recap_btn):
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
        self.run_worker(
            f"Génération du BRQ du {start} au {end}...",
            self.export_daily, start, end
        )

    def export_daily(self, start, end):
        self.log("Connexion à la base existante : aucun import SQL.")
        conn = connect_database()
        try:
            frame = pd.read_sql(
                BRQ_QUERY, conn, params={"date1": start, "date2": end}
            )
        finally:
            conn.close()
        if frame.empty:
            self.log("Aucune ligne trouvée pour cette période.")
            return
        path = os.path.join(WORK_FOLDER, f"BRQ_Export_{start}_au_{end}.xlsx")
        write_excel(frame, path, "BRQ")
        self.log(f"{len(frame)} lignes exportées : {path}")

    def on_recap(self):
        period = self._validated_period()
        if period is None:
            return
        start, end = period
        self.run_worker(
            f"Récapitulatif du {start} au {end}...",
            self.export_recap, start, end
        )

    def export_recap(self, start, end):
        self.log("Connexion à la base existante : aucun import SQL.")
        conn = connect_database()
        try:
            frame = pd.read_sql(RECAP_QUERY, conn, params={"date1": start, "date2": end})
        finally: conn.close()
        path = os.path.join(WORK_FOLDER, f"BRQ_Recap_{start}_au_{end}.xlsx")
        write_excel(frame, path, "Recap")
        self.log(f"Récapitulatif enregistré : {path}")

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
