# -*- coding: utf-8 -*-

# IMPORTS
from Autodesk.Revit.DB import *
from pyrevit import revit, DB, forms, script
import os
import xlsxwriter
from datetime import datetime

# VARIABLES
doc 	= __revit__.ActiveUIDocument.Document
uidoc 	= __revit__.ActiveUIDocument
app 	= __revit__.Application
sched 	= revit.active_view

# Vérification des données de la nomenclature
try:
    tableData = sched.GetTableData()
    sectionData = tableData.GetSectionData(DB.SectionType.Body)
    numbRows = sectionData.NumberOfRows
    numbCols = sectionData.NumberOfColumns
except:
    forms.alert("Pas de données dans la table à exporter.", title="Script annulé")
    script.exit()

# Récupération de la date et du numéro du projet
now = datetime.now()
date = now.strftime("%y%m%d")
projName = doc.ProjectInformation.Name or "PRJ"

# Sélection du dossier d'export
destinationFolder = forms.pick_folder()
if not destinationFolder:
    script.exit()

# Récupération des données de la nomenclature
data = []
for i in range(numbRows):
    row = [sched.GetCellText(DB.SectionType.Body, i, j) for j in range(numbCols)]
    if any(row):  # Vérifie que la ligne n'est pas entièrement vide
        data.append(row)

# Calcul des longueurs maximales des colonnes
lengths = [max(len(row[col]) for row in data) for col in range(numbCols)]

# Nom de la feuille et du fichier
worksheetName = sched.Name[:29] if len(sched.Name) >= 30 else sched.Name
filePath = destinationFolder + "\\" + date + "_" + projName + "_" + sched.Name + ".xlsx"

# Création du fichier Excel
workbook = xlsxwriter.Workbook(filePath, {'strings_to_numbers': True})
worksheet = workbook.add_worksheet(worksheetName)

# Définition des formats pour les cellules
fontSize = 12
titleFormat 	= workbook.add_format({"bg_color": '#47bcca',"bold": True, "font_size": fontSize})
subtitleFormat 	= workbook.add_format({"bg_color": "#d2e9fa", "font_size": fontSize,"align": "left"})
cellFormat 		= workbook.add_format({"font_size": fontSize, "align": "left"})

# Définition des largeurs des colonnes
for col_idx, length in enumerate(lengths):
    worksheet.set_column(col_idx, col_idx, length + 2)  # Ajuste automatiquement la largeur avec une marge

# Écriture des données dans le fichier Excel
for row_idx, row_data in enumerate(data):
    first, rest = row_data[0], row_data[1:]
    
    # Détection des sous-titres
    if (first != "" and rest.count("") == len(lengths) - 1) or row_data.count("") == len(lengths):
        row_format = subtitleFormat
    elif row_idx == 0:  # Première ligne (en-tête ou titre)
        row_format = titleFormat
    else:
        row_format = cellFormat
    
    for col_idx, cell_value in enumerate(row_data):
        worksheet.write(row_idx, col_idx, cell_value, row_format)

# Fermeture du fichier Excel
workbook.close()

# Retour utilisateur
forms.alert("Nomenclature exportée", title="Script terminé", warn_icon=False)
