# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import forms, revit
import System
from System.Xml.Linq import XDocument, XElement, XAttribute

doc = revit.doc

# ---------------------------------------------------------
# 1. Pièces de la première phase
# ---------------------------------------------------------
first_phase = list(doc.Phases)[0]

def room_label(r):
    num  = r.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or ""
    name = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()   or ""
    return "{} - {}".format(num, name)

class RoomItem(object):
    def __init__(self, room):
        self.room  = room
        self.label = room_label(room)
    def __str__(self):
        return self.label

room_items = sorted([
    RoomItem(r)
    for r in FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_Rooms)
               .WhereElementIsNotElementType()
    if r.get_Parameter(BuiltInParameter.ROOM_PHASE).AsElementId() == first_phase.Id
], key=lambda x: x.label)

# ---------------------------------------------------------
# 2. Sélection pièces + site + chemin (séquence rapide)
# ---------------------------------------------------------
selected = forms.SelectFromList.show(
    room_items,
    title="Sélection des pièces",
    multiselect=True,
    button_name="Suivant"
)
if not selected:
    forms.alert("Aucune pièce sélectionnée.")
    raise SystemExit

site = forms.ask_for_one_item(
    ["TEC-HL", "TEC-PEL", "TEC-SA", "TEC-DG"],
    title="Choisir le site",
    default="TEC-HL",
    prompt="Sélectionnez le site :"
)
if not site:
    forms.alert("Aucun site sélectionné.")
    raise SystemExit

save_path = forms.save_file(
    file_ext="xml",
    default_name="{}.xml".format(doc.Title),
    title="Enregistrer le fichier XML"
)
if not save_path:
    forms.alert("Export annulé.")
    raise SystemExit

# ---------------------------------------------------------
# 3. Construction XML
# ---------------------------------------------------------
date_str = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

def make_box(room, site_code, date):
    num  = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    return XElement("box",
        XAttribute("code", num),
        XElement("box_code",        num),
        XElement("box_description", name),
        XElement("box_site",
            XAttribute("code", site_code),
            site_code
        ),
        XElement("box_WOAuthorized", "true"),
        XElement("box_status",
            XElement("status_changedBy",
                XAttribute("code", "CARLSOURCE"),
                XAttribute("id",   "15")
            ),
            XElement("status_changedDate", date),
            XElement("status_code", "VALIDATE")
        ),
        XElement("box_structure",
            XAttribute("code", "GEOGRAPHIQUE")
        )
    )

entities = XElement("entities",
    XAttribute("exchangeInterface", "BOX_IN"),
    XAttribute("externalSystem",    ""),
    XAttribute("timezone",          "Europe/Paris"),
    XAttribute("language",          "fr_FR"),
    *[make_box(item.room, site, date_str) for item in selected]
)

# ---------------------------------------------------------
# 4. Sauvegarde
# ---------------------------------------------------------
#XDocument(entities).Save(save_path)
#forms.alert("Export XML terminé :\n{}".format(save_path))
