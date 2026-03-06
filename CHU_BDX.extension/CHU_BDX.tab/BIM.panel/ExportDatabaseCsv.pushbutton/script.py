# -*- coding: utf-8 -*-
import csv
import codecs
import os
from Autodesk.Revit import DB
from pyrevit import revit, script
from pyrevit.forms import pick_folder, ProgressBar, SelectFromList

output = script.get_output()

# 1️⃣ Choisir le dossier de sauvegarde
folder_path = pick_folder(title="Choisir un dossier pour enregistrer les CSV")
if not folder_path:
    output.print_md("❌ Aucun dossier sélectionné. Export annulé.")
    script.exit()

# 2️⃣ Document Revit actif
doc = revit.doc

# 3️⃣ Sélection de catégories (multiselect activé)
category_names = sorted([cat.Name for cat in doc.Settings.Categories])
selected_cat_names = SelectFromList.show(
    category_names,
    title="Choisir une ou plusieurs catégories à exporter",
    multiselect=True,
    button_name="Exporter"
)
if not selected_cat_names:
    output.print_md("❌ Aucune catégorie sélectionnée. Export annulé.")
    script.exit()

# Normaliser en liste (pyRevit peut retourner un str si une seule sélection)
if isinstance(selected_cat_names, str):
    selected_cat_names = [selected_cat_names]

# 4️⃣ Trouver les catégories sélectionnées
selected_categories = [
    cat for cat in doc.Settings.Categories
    if cat.Name in selected_cat_names
]
if not selected_categories:
    output.print_md("❌ Catégories introuvables. Export annulé.")
    script.exit()

revit_filename = os.path.splitext(os.path.basename(doc.PathName))[0]
export_summary = []

# 5️⃣ Boucle : un CSV par catégorie
for selected_category in selected_categories:
    cat_name = selected_category.Name

    # Chemin du fichier CSV pour cette catégorie
    csv_filename = "{0}_{1}.csv".format(revit_filename, cat_name)
    csv_path = os.path.join(folder_path, csv_filename)

    # Collecteur pour cette catégorie
    collector = DB.FilteredElementCollector(doc).OfCategoryId(selected_category.Id)

    # Premier passage : paramètres + comptage
    param_names_set = set()
    element_count = 0
    type_cache = {}

    for el in collector:
        element_count += 1
        for param in el.Parameters:
            if param.Definition:
                param_names_set.add(param.Definition.Name)

        if not isinstance(el, DB.ElementType):
            type_id = el.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                type_id_int = type_id.IntegerValue
                if type_id_int not in type_cache:
                    type_cache[type_id_int] = doc.GetElement(type_id)
                el_type = type_cache[type_id_int]
                if el_type:
                    for param in el_type.Parameters:
                        if param.Definition:
                            param_names_set.add(param.Definition.Name)

    if element_count == 0:
        output.print_md("⚠️ **{}** : aucun élément trouvé, catégorie ignorée.".format(cat_name))
        continue

    param_names = sorted(param_names_set)
    header = ["Element Unique ID", "Element ID", "Type/Occurrence"] + param_names

    # Écriture du CSV en streaming
    with codecs.open(csv_path, "w", "utf-8-sig") as csvfile:
        writer = csv.writer(csvfile, delimiter=";", lineterminator="\n")
        writer.writerow(header)

        collector = DB.FilteredElementCollector(doc).OfCategoryId(selected_category.Id)

        with ProgressBar(
            title="Export {} en cours...".format(cat_name),
            cancellable=True
        ) as pb:
            idx = 0
            for el in collector:
                if pb.cancelled:
                    output.print_md("❌ Export annulé par l'utilisateur.")
                    script.exit()

                is_type = isinstance(el, DB.ElementType)
                element_id = str(el.Id.IntegerValue)
                row = [el.UniqueId, element_id, "Type" if is_type else "Occurrence"]

                param_dict = {}

                for param in el.Parameters:
                    if param.Definition and param.Definition.Name in param_names_set:
                        param_dict[param.Definition.Name] = param.AsValueString() if param.HasValue else ""

                if not is_type:
                    type_id = el.GetTypeId()
                    if type_id and type_id != DB.ElementId.InvalidElementId:
                        type_id_int = type_id.IntegerValue
                        if type_id_int in type_cache:
                            el_type = type_cache[type_id_int]
                            if el_type:
                                for param in el_type.Parameters:
                                    if param.Definition:
                                        pname = param.Definition.Name
                                        if pname in param_names_set and pname not in param_dict:
                                            param_dict[pname] = param.AsValueString() if param.HasValue else ""

                for pname in param_names:
                    row.append(param_dict.get(pname, ""))

                writer.writerow(row)
                idx += 1
                pb.update_progress(idx, element_count)

    export_summary.append((cat_name, csv_path, element_count, len(param_names)))

# 6️⃣ Résumé final
output.print_md("### ✅ Export terminé — {} fichier(s) généré(s)".format(len(export_summary)))
for cat_name, csv_path, el_count, param_count in export_summary:
    output.print_md(
        "- **{}** → `{}` | {} éléments | {} paramètres".format(
            cat_name, csv_path, el_count, param_count
        )
    )
