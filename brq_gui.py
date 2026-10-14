"""
brq_gui.py

A simple desktop window for the daily BRQ workflow:
  - Pick the .sql file you got today, pick the date, click "Lancer" to run
    export_brq.py (as a real subprocess — no duplicated logic) and watch its
    live output.
  - "Ouvrir le dossier" opens C:\\BRQ in File Explorer.
  - A "Recap" section lets you pick a date range and export the CHAIR /
    PONTE / RUMINANTS / DIVERS / TOTAL quantity summary for that period.

Setup (run once):
    pip install pymysql pandas openpyxl

Requirements:
    export_brq.py must be in the SAME FOLDER as this file (WORK_FOLDER below).

Usage:
    python brq_gui.py
"""

import os
import sys
import shutil
import queue
import threading
import subprocess
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pymysql
import pandas as pd

# ---------------------------------------------------------------------------
# EDIT THESE to match your setup (same values as in export_brq.py)
# ---------------------------------------------------------------------------
WORK_FOLDER = "C:\\BRQ"
SCRIPT_NAME = "export_brq.py"
SQL_FILE_NAME = "UAB_OULEDHAMLA_COM26.sql"  # must match SQL_FILE_NAME in export_brq.py

DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "",                       # XAMPP's default root user has NO password
    "database": "UAB_OULEDHAMLA_COM26",
    "port": 3306,
}

SCRIPT_PATH = os.path.join(WORK_FOLDER, SCRIPT_NAME)
SQL_TARGET_PATH = os.path.join(WORK_FOLDER, SQL_FILE_NAME)

RECAP_QUERY = """
SELECT
    SUM(CASE WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'CHAIR' THEN d.Quantite ELSE 0 END) AS CHAIR,
    SUM(CASE WHEN f.Libelle = 'VOLAILLE' AND sf.Libelle = 'PONTE' THEN d.Quantite ELSE 0 END) AS PONTE,
    SUM(CASE WHEN f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN') THEN d.Quantite ELSE 0 END) AS RUMINANTS,
    SUM(CASE WHEN f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS' THEN d.Quantite ELSE 0 END) AS DIVERS,
    SUM(
        CASE
            WHEN (f.Libelle = 'VOLAILLE' AND sf.Libelle IN ('CHAIR', 'PONTE'))
              OR (f.Libelle = 'RUMINANTS' AND sf.Libelle IN ('BOVIN', 'OVIN'))
              OR (f.Libelle = 'DIVERS' AND sf.Libelle = 'DIVERS')
            THEN d.Quantite ELSE 0
        END
    ) AS TOTAL
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
    e.Date >= %(date1)s
    AND e.Date <= %(date2)s
    AND e.Type_mvt = 'S'
    AND e.Type_bon = 'F';
"""


def open_path_in_explorer(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


class BrqApp:
    def __init__(self, root):
        self.root = root
        root.title("Export BRQ")
        root.geometry("660x560")
        root.resizable(False, False)

        self.log_queue = queue.Queue()
        self.running = False

        pad = {"padx": 10, "pady": 6}

        # --- Daily export section -------------------------------------
        tk.Label(root, text="Export quotidien", font=("", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))

        frame1 = tk.Frame(root)
        frame1.pack(fill="x", **pad)
        tk.Label(frame1, text="Fichier SQL :", width=12, anchor="w").pack(side="left")
        self.sql_path_var = tk.StringVar()
        tk.Entry(frame1, textvariable=self.sql_path_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(frame1, text="Parcourir...", command=self.browse_sql_file).pack(side="left")

        frame2 = tk.Frame(root)
        frame2.pack(fill="x", **pad)
        tk.Label(frame2, text="Date :", width=12, anchor="w").pack(side="left")
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        tk.Entry(frame2, textvariable=self.date_var, width=15).pack(side="left")
        tk.Label(frame2, text="(format AAAA-MM-JJ)", fg="gray").pack(side="left", padx=8)

        frame3 = tk.Frame(root)
        frame3.pack(fill="x", **pad)
        self.run_button = tk.Button(frame3, text="Lancer", command=self.on_run, width=15, bg="#2e7d32", fg="white")
        self.run_button.pack(side="left")
        tk.Button(frame3, text="Ouvrir le dossier", command=self.on_open_folder, width=15).pack(side="left", padx=8)
        self.progress = ttk.Progressbar(frame3, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # --- Separator ---------------------------------------------------
        ttk.Separator(root, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # --- Recap section -------------------------------------------------
        tk.Label(root, text="Récap par période", font=("", 10, "bold")).pack(anchor="w", padx=10)

        frame4 = tk.Frame(root)
        frame4.pack(fill="x", **pad)
        tk.Label(frame4, text="Du :", width=12, anchor="w").pack(side="left")
        self.recap_date1_var = tk.StringVar(value=datetime.date.today().replace(day=1).strftime("%Y-%m-%d"))
        tk.Entry(frame4, textvariable=self.recap_date1_var, width=15).pack(side="left")
        tk.Label(frame4, text="Au :", width=6, anchor="e").pack(side="left", padx=(16, 0))
        self.recap_date2_var = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        tk.Entry(frame4, textvariable=self.recap_date2_var, width=15).pack(side="left")

        frame5 = tk.Frame(root)
        frame5.pack(fill="x", **pad)
        self.recap_button = tk.Button(frame5, text="Exporter le récap", command=self.on_recap, width=18, bg="#1565c0", fg="white")
        self.recap_button.pack(side="left")

        # --- Separator ---------------------------------------------------
        ttk.Separator(root, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # --- Log area ------------------------------------------------------
        tk.Label(root, text="Journal :", anchor="w").pack(fill="x", padx=10)
        log_frame = tk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, height=14, state="disabled", wrap="word")
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.root.after(100, self.poll_log_queue)

    # -----------------------------------------------------------------
    # Shared helpers
    # -----------------------------------------------------------------
    def log(self, message):
        self.log_queue.put(message)

    def poll_log_queue(self):
        while not self.log_queue.empty():
            message = self.log_queue.get_nowait()
            self.log_text.configure(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(100, self.poll_log_queue)

    def on_open_folder(self):
        if not os.path.exists(WORK_FOLDER):
            messagebox.showerror("Erreur", f"Le dossier {WORK_FOLDER} n'existe pas.")
            return
        try:
            open_path_in_explorer(WORK_FOLDER)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le dossier :\n{e}")

    # -----------------------------------------------------------------
    # Daily export
    # -----------------------------------------------------------------
    def browse_sql_file(self):
        path = filedialog.askopenfilename(
            title="Choisir le fichier .sql",
            filetypes=[("Fichiers SQL", "*.sql"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.sql_path_var.set(path)

    def on_run(self):
        if self.running:
            return

        sql_path = self.sql_path_var.get().strip()
        date_str = self.date_var.get().strip()

        if not sql_path or not os.path.exists(sql_path):
            messagebox.showerror("Erreur", "Veuillez choisir un fichier .sql valide.")
            return

        try:
            datetime.datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Erreur", "Date invalide. Utilisez le format AAAA-MM-JJ (ex: 2026-09-24).")
            return

        if not os.path.exists(SCRIPT_PATH):
            messagebox.showerror(
                "Erreur",
                f"Impossible de trouver {SCRIPT_NAME} dans {WORK_FOLDER}.\n"
                "Placez export_brq.py dans ce dossier, à côté de brq_gui.py.",
            )
            return

        self.set_running(True, self.run_button, "En cours...")
        self.clear_log()

        thread = threading.Thread(target=self.export_worker, args=(sql_path, date_str), daemon=True)
        thread.start()

    def export_worker(self, sql_path, date_str):
        try:
            if os.path.abspath(sql_path) != os.path.abspath(SQL_TARGET_PATH):
                self.log(f"Copie du fichier SQL vers {SQL_TARGET_PATH} ...")
                shutil.copyfile(sql_path, SQL_TARGET_PATH)

            self.log(f"Lancement de {SCRIPT_NAME} {date_str} ...")
            self.log("")

            process = subprocess.Popen(
                [sys.executable, SCRIPT_PATH, date_str],
                cwd=WORK_FOLDER,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )

            for line in process.stdout:
                self.log(line.rstrip())

            return_code = process.wait()

            self.log("")
            if return_code == 0:
                self.log("Terminé avec succès.")
            else:
                self.log(f"Le script s'est arrêté avec une erreur (code {return_code}).")

        except Exception as e:
            self.log(f"ERREUR : {e}")
        finally:
            self.root.after(0, lambda: self.set_running(False, self.run_button, "Lancer"))

    # -----------------------------------------------------------------
    # Recap export
    # -----------------------------------------------------------------
    def on_recap(self):
        if self.running:
            return

        date1_str = self.recap_date1_var.get().strip()
        date2_str = self.recap_date2_var.get().strip()

        try:
            date1 = datetime.datetime.strptime(date1_str, "%Y-%m-%d").date()
            date2 = datetime.datetime.strptime(date2_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Erreur", "Dates invalides. Utilisez le format AAAA-MM-JJ.")
            return

        if date1 > date2:
            messagebox.showerror("Erreur", "La date de début doit être avant (ou égale à) la date de fin.")
            return

        self.set_running(True, self.recap_button, "En cours...")
        self.clear_log()

        thread = threading.Thread(target=self.recap_worker, args=(date1, date2), daemon=True)
        thread.start()

    def recap_worker(self, date1, date2):
        try:
            self.log(f"Récap du {date1} au {date2} ...")
            connection = pymysql.connect(**DB_CONFIG)
            try:
                df = pd.read_sql(RECAP_QUERY, connection, params={"date1": date1, "date2": date2})
            finally:
                connection.close()

            os.makedirs(WORK_FOLDER, exist_ok=True)
            filename = os.path.join(WORK_FOLDER, f"BRQ_Recap_{date1}_au_{date2}.xlsx")

            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Recap")
                worksheet = writer.sheets["Recap"]
                for i, col in enumerate(df.columns, start=1):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = max_len

            self.log(f"Récap enregistré : {filename}")
            self.log("Terminé avec succès.")

        except Exception as e:
            self.log(f"ERREUR : {e}")
        finally:
            self.root.after(0, lambda: self.set_running(False, self.recap_button, "Exporter le récap"))

    # -----------------------------------------------------------------
    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def set_running(self, is_running, button, label):
        self.running = is_running
        button.configure(state="disabled" if is_running else "normal", text=label)
        if is_running:
            self.progress.start(12)
        else:
            self.progress.stop()


def main():
    root = tk.Tk()
    BrqApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
