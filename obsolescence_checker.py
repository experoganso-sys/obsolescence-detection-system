import pandas as pd
import requests
from datetime import datetime
from tabulate import tabulate
import tkinter as tk
from tkinter import filedialog
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# --- Sélection du fichier inventaire ---
root = tk.Tk()
root.withdraw()

print("📂 Veuillez sélectionner le fichier inventaire...")
inventaire_file = filedialog.askopenfilename(
    title="Sélectionnez le fichier inventaire",
    filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
)

if not inventaire_file:
    print("❌ Aucun fichier sélectionné. Fin du programme.")
    exit()
            if entry.get("cycle") == str(version):
                return {
                    "eol": entry.get("eol"),
                    "link": entry.get("link")
                }
        return None
    except Exception:
        return None

# --- Traitement principal ---
rapport = []

print("\n🔍 Vérification des logiciels sur endoflife.date...\n")

for _, row in inventaire.iterrows():
    logiciel = row["logiciel"].strip().lower().replace(" ", "-")
    version = str(row["version"]).strip()
    machine = row["machine"]

    info = get_eol_info(logiciel, version)

    if info:
        statut = check_status(info["eol"])
        eol_date = info["eol"] or "N/A"
        source = info["link"] or "N/A"
    else:
        statut = "Inconnu"
        eol_date = "N/A"
        source = "N/A"

    rapport.append([machine, row["logiciel"], version, eol_date, statut, source])

# --- Création DataFrame ---
rapport_df = pd.DataFrame(rapport, columns=["Machine", "Logiciel", "Version", "Date EOL", "Statut", "Source"])

# --- Sauvegarde CSV ---
rapport_csv = "rapport_obsolescence.csv"
rapport_df.to_csv(rapport_csv, index=False)
print(f"✅ Rapport CSV généré : {rapport_csv}")

# --- Choix emplacement PDF ---
print("\n💾 Sélectionnez l’emplacement pour enregistrer le rapport PDF...")
pdf_file = filedialog.asksaveasfilename(
    title="Enregistrer le rapport PDF",
    defaultextension=".pdf",
    filetypes=(("PDF files", "*.pdf"),)
)
if not pdf_file:
    print("❌ Aucun emplacement sélectionné. Fin du programme.")
    exit()

# --- Génération du PDF ---
doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4))
elements = []
styles = getSampleStyleSheet()

title = Paragraph("Rapport d’obsolescence des logiciels", styles['Title'])
elements.append(title)
elements.append(Spacer(1, 12))

# Préparer le tableau PDF
data = [rapport_df.columns.tolist()] + rapport_df.values.tolist()
table = Table(data)
table_style = TableStyle([
    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#003366")),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ('BOTTOMPADDING', (0,0), (-1,0), 8),
    ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
    ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
])
table.setStyle(table_style)
elements.append(table)

doc.build(elements)

print(f"✅ Rapport PDF généré avec succès : {pdf_file}")
print("\n📊 Tableau récapitulatif :\n")
print(tabulate(rapport_df, headers="keys", tablefmt="grid"))
