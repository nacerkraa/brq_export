"""Application configuration and SQL queries."""
import os

WORK_FOLDER = r"C:\BRQ"
OUTPUT_FOLDER = os.path.join(WORK_FOLDER, "Exports")
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
