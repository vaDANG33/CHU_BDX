# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
from pyrevit import forms, script
import xlrd


output = script.get_output()


def select_excel_file():
    return forms.pick_excel_file(title='Sélectionner un fichier Excel')


def get_sheet_names_xlrd(workbook):
    return workbook.sheet_names()


def select_sheet(sheet_list):
    return forms.SelectFromList.show(sheet_list, title='Choisir une feuille Excel', multiselect=False)


def select_mode():
    return forms.CommandSwitchWindow.show(['Type', 'Occurrence'], message='Appliquer à quel niveau ?')


def filter_parameters_with_values(sheet, headers, mode):
    try:
        type_occ_index = headers.index("Type/Occurrence")
    except ValueError:
        output.print_md("**Erreur :** La colonne `Type/Occurrence` est introuvable.")
        return []


    valid_params = set()
    for row_idx in range(1, sheet.nrows):
        row = sheet.row_values(row_idx)
        cell_value = row[type_occ_index]
        
        # Vérifier si la valeur correspond au mode sélectionné
        if cell_value == mode:
            for col_idx, value in enumerate(row):
                if col_idx != type_occ_index and value not in ["", None]:
                    valid_params.add(headers[col_idx])


    return sorted(valid_params)


def select_parameter(param_list):
    return forms.SelectFromList.show(param_list, title='Choisir un paramètre avec valeurs', multiselect=False)


def filter_values(sheet, headers, param_name, mode):
    try:
        param_index = headers.index(param_name)
        type_occ_index = headers.index("Type/Occurrence")
    except ValueError:
        output.print_md("**Erreur :** La colonne `{0}` ou `Type/Occurrence` est introuvable.".format(param_name))
        return


    filtered_values = [
        sheet.row_values(row_idx)[param_index]
        for row_idx in range(1, sheet.nrows)
        if sheet.row_values(row_idx)[type_occ_index] == mode
    ]


    output.print_md("### Valeurs filtrées pour `{0}` en mode `{1}` :".format(param_name, mode))
    for val in filtered_values:
        output.print_md("- `{0}`".format(val))


# 🔁 Exécution
excel_file = select_excel_file()
if excel_file:
    workbook = xlrd.open_workbook(excel_file)
    sheets = get_sheet_names_xlrd(workbook)
    selected_sheet = select_sheet(sheets)
    if selected_sheet:
        sheet = workbook.sheet_by_name(selected_sheet)
        headers = sheet.row_values(0)
        selected_mode = select_mode()
        if selected_mode:
            filtered_params = filter_parameters_with_values(sheet, headers, selected_mode)
            if filtered_params:
                selected_param = select_parameter(filtered_params)
                if selected_param:
                    filter_values(sheet, headers, selected_param, selected_mode)
            else:
                output.print_md("⚠️ Aucun paramètre avec valeur trouvé pour le mode `{}`.".format(selected_mode))
