# -*- coding: utf-8 -*-
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import sys
import GUI

# Filtre de sélection inline pour les pièces
class RoomSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Rooms)
    
    def AllowReference(self, reference, point):
        return False

def room_selection(doc, uidoc, select_by, all_rooms_placed):
    try:
        if select_by == 'All':
            # Récupération directe de toutes les pièces placées
            all_rooms = DB.FilteredElementCollector(doc)\
                .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                .WhereElementIsNotElementType()
            rooms = [room for room in all_rooms if room.Area != 0]
            
        elif select_by == 'By Level':
            # Collecte des niveaux
            all_levels = [l for l in DB.FilteredElementCollector(doc)\
                .OfCategory(DB.BuiltInCategory.OST_Levels)\
                .WhereElementIsNotElementType()]
            
            all_level_names = [l.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsValueString() 
                             for l in all_levels]
            
            # Dialogue de sélection des niveaux
            chosen_levels = GUI.user_prompt_get_object_from_names(
                all_levels, 
                all_level_names, 
                title="Choose level to get room data from", 
                multiselect=True
            )
            
            # Filtrage des pièces par niveau
            rooms = [room for room in all_rooms_placed 
                    if room.LookupParameter('Niveau').AsValueString() 
                    in [l.Name for l in chosen_levels]]
            
        elif select_by == 'By Selection':
            # Sélection manuelle avec filtre inline
            rooms = [doc.GetElement(ref.ElementId) 
                    for ref in uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        RoomSelectionFilter(),
                        "Select rooms"
                    )]
        else:
            GUI.task_terminated()
            
    except:
        GUI.task_terminated()
    
    return rooms
