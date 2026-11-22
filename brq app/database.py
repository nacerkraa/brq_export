"""MySQL connection and SQL backup import."""
import os
import subprocess
import pymysql
from config import CLEANED_SQL_PATH, DB_CONFIG, MYSQL_EXE_PATH

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
