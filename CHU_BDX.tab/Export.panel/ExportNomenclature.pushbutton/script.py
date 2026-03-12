# -*- coding: utf-8 -*-

# IMPORTS
from Autodesk.Revit.DB import *
from pyrevit import revit, DB, forms, script
import os
import re
import xlsxwriter
from datetime import datetime


# VARIABLES
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application


# Fonction pour recuperer toutes les nomenclatures du document
def get_all_schedules():
    schedules = []
    collector = FilteredElementCollector(doc)
    all_views = collector.OfClass(View).ToElements()
    for view in all_views:
        if view.ViewType == ViewType.Schedule:
            schedules.append(view)
    return schedules


# Fonction pour verifier si une nomenclature a des donnees
def has_table_data(schedule):
    try:
        tableData = schedule.GetTableData()
        sectionData = tableData.GetSectionData(DB.SectionType.Body)
        return sectionData.NumberOfRows > 0
    except:
        return False


# Fonction pour exporter une nomenclature en Excel
def export_schedule_to_excel(schedule, destination_folder, date, proj_name):
    try:
        # Recuperation des donnees
        tableData = schedule.GetTableData()
        sectionData = tableData.GetSectionData(DB.SectionType.Body)
        numbRows = sectionData.NumberOfRows
        numbCols = sectionData.NumberOfColumns

        # Recuperation des donnees de la nomenclature
        data = []
        for i in range(numbRows):
            row = [schedule.GetCellText(DB.SectionType.Body, i, j) for j in range(numbCols)]
            if any(row):
                data.append(row)

        if not data:
            return False, "Aucune donnee dans la nomenclature"

        # Calcul des longueurs maximales des colonnes
        lengths = [max(len(row[col]) for row in data) if data else 10 for col in range(numbCols)]

        # Nom de la feuille et du fichier
        worksheetName = schedule.Name[:29] if len(schedule.Name) >= 30 else schedule.Name

        # Nettoyer le nom du fichier
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", schedule.Name)
        filePath = os.path.join(destination_folder, "{}_{}_{}.xlsx".format(date, proj_name, safe_name))

        # Creation du fichier Excel
        workbook = xlsxwriter.Workbook(filePath, {'strings_to_numbers': False})
        worksheet = workbook.add_worksheet(worksheetName)

        # Definition des formats
        fontSize = 12
        titleFormat    = workbook.add_format({"bg_color": "#47bcca", "bold": True,  "font_size": fontSize})
        subtitleFormat = workbook.add_format({"bg_color": "#d2e9fa", "font_size": fontSize, "align": "left"})
        cellFormat     = workbook.add_format({"font_size": fontSize,  "align": "left"})

        # Largeurs des colonnes
        for col_idx, length in enumerate(lengths):
            worksheet.set_column(col_idx, col_idx, min(length + 2, 50))

        # Ecriture des donnees
        for row_idx, row_data in enumerate(data):
            first = row_data[0]
            rest  = row_data[1:]

            if (first != "" and rest.count("") == len(lengths) - 1) or row_data.count("") == len(lengths):
                row_format = subtitleFormat
            elif row_idx == 0:
                row_format = titleFormat
            else:
                row_format = cellFormat

            for col_idx, cell_value in enumerate(row_data):
                worksheet.write(row_idx, col_idx, cell_value, row_format)

        workbook.close()
        return True, filePath

    except Exception as e:
        return False, str(e)


# ============ MAIN ============

# Date et nom du projet
now      = datetime.now()
date     = now.strftime("%y%m%d")
projName = doc.ProjectInformation.Name or "PRJ"

# Recuperer toutes les nomenclatures
all_schedules = get_all_schedules()

if not all_schedules:
    forms.alert("Aucune nomenclature trouvee dans le document.", title="Script annule")
    script.exit()

# Filtrer les nomenclatures avec donnees
valid_schedules = [s for s in all_schedules if has_table_data(s)]

if not valid_schedules:
    forms.alert("Aucune nomenclature avec des donnees trouvee.", title="Script annule")
    script.exit()

# Selection multiple
selected_schedules = forms.SelectFromList.show(
    sorted([s.Name for s in valid_schedules]),
    title="Selectionner les nomenclatures a exporter",
    button_name="Exporter",
    multiselect=True
)

if not selected_schedules:
    script.exit()

# Retrouver les objets View correspondants aux noms selectionnes
selected_views = [s for s in valid_schedules if s.Name in selected_schedules]

# Dossier destination
destinationFolder = forms.pick_folder()
if not destinationFolder:
    script.exit()

# Export
exported_count = 0
failed_count   = 0
failed_list    = []

for schedule in selected_views:
    success, result = export_schedule_to_excel(schedule, destinationFolder, date, projName)
    if success:
        exported_count += 1
    else:
        failed_count += 1
        failed_list.append("{} : {}".format(schedule.Name, result))

# Message final
message = "Nomenclatures exportees : {}".format(exported_count)
if failed_count > 0:
    message += "\n\nEchecs ({}) :\n".format(failed_count) + "\n".join(failed_list)

forms.alert(message, title="Export termine", warn_icon=failed_count > 0)
