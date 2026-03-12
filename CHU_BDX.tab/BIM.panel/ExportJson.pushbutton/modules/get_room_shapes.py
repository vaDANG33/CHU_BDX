# -*- coding: utf-8 -*-
import clr
clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName('AdWindows')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')

import Autodesk
import Autodesk.Windows as aw
from Autodesk.Revit import DB
from Autodesk.Revit import UI

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uiapp.ActiveUIDocument.Document

import GetSetParameters
import arc_segment_conversion

def get_start_end_point(segment):
    """Extrait les points de début et fin d'un segment en mm"""
    line = segment.GetCurve()
    s_point_x = line.GetEndPoint(0).Multiply(304.8).X
    s_point_y = line.GetEndPoint(0).Multiply(304.8).Y
    e_point_x = line.GetEndPoint(1).Multiply(304.8).X
    e_point_y = line.GetEndPoint(1).Multiply(304.8).Y
    
    return (s_point_x, s_point_y, e_point_x, e_point_y)

def generate_endpoints(segment, is_outer):
    """Génère les points finaux d'un segment (arc ou ligne)"""
    # Si c'est un arc, convertir en série de points
    if segment.GetCurve().GetType() == Autodesk.Revit.DB.Arc:
        endpoints = arc_segment_conversion.arc_segment_conversion(
            segment, 
            full_circle=False, 
            is_outer_boundary=is_outer
        )
        return endpoints
    
    # Si c'est une ligne droite, retourner juste le point final
    coords = get_start_end_point(segment)
    return [[coords[2], coords[3]]]

def get_room_shapes(rooms, parameters, outside_boundary_only=True):
    """
    Extrait la géométrie et les paramètres des pièces Revit
    
    Args:
        rooms: Liste des pièces Revit
        parameters: Liste des paramètres à extraire
        outside_boundary_only: Si True, ignore les contours intérieurs
    
    Returns:
        Dictionnaire {room_id: {params + geometry}}
    """
    output = {}
    
    for room in rooms:
        room_data = {}
        boundary_locations = []
        outer_boundary = True
        
        # Extraction des paramètres
        for param in parameters:
            try:
                param_obj = room.LookupParameter(param)
                if param_obj is None:
                    room_data[param] = ""
                    continue
                
                param_type = GetSetParameters.get_parameter_type(param_obj)
                if param_type == "Double":
                    value = param_obj.AsDouble()
                    room_data[param] = str(value) if value is not None else ""
                else:
                    value = param_obj.AsValueString()
                    room_data[param] = value if value is not None else ""
            except:
                room_data[param] = ""
        
        # Extraction de la géométrie
        boundary_segments = room.GetBoundarySegments(DB.SpatialElementBoundaryOptions())
        
        for boundary_segment in boundary_segments:
            closed_loop = []
            
            # Si c'est un cercle complet (un seul segment arc)
            if len(boundary_segment) == 1 and boundary_segment[0].GetCurve().GetType() == Autodesk.Revit.DB.Arc:
                closed_loop = arc_segment_conversion.arc_segment_conversion(
                    boundary_segment[0],
                    full_circle=True,
                    is_outer_boundary=outer_boundary
                )
            else:
                # Contour multi-segments
                # Pour chaque segment, ajouter ses points (sauf le dernier pour éviter les doublons)
                for i, segment in enumerate(boundary_segment):
                    if segment.GetCurve().GetType() == Autodesk.Revit.DB.Arc:
                        # Arc : récupérer tous les points
                        arc_points = arc_segment_conversion.arc_segment_conversion(
                            segment, 
                            full_circle=False, 
                            is_outer_boundary=outer_boundary
                        )
                        
                        if i == 0:
                            # Premier segment : ajouter tous les points
                            closed_loop.extend(arc_points)
                        else:
                            # Segments suivants : sauter le premier point (déjà présent)
                            closed_loop.extend(arc_points[1:])
                    else:
                        # Ligne droite
                        coords = get_start_end_point(segment)
                        
                        if i == 0:
                            # Premier segment : ajouter début et fin
                            closed_loop.append([coords[0], coords[1]])
                            closed_loop.append([coords[2], coords[3]])
                        else:
                            # Segments suivants : ajouter seulement la fin
                            closed_loop.append([coords[2], coords[3]])
            
            # Vérifier que le polygone est bien fermé
            # GeoJSON nécessite que le premier et dernier point soient identiques
            if len(closed_loop) > 0:
                first_point = closed_loop[0]
                last_point = closed_loop[-1]
                
                # Si les points ne sont pas identiques (tolérance de 0.01mm)
                if abs(first_point[0] - last_point[0]) > 0.01 or abs(first_point[1] - last_point[1]) > 0.01:
                    # Ajouter le premier point à la fin pour fermer le polygone
                    closed_loop.append([first_point[0], first_point[1]])
            
            # Inverser l'ordre pour GeoJSON (sens anti-horaire)
            boundary_locations.append(closed_loop[::-1])
            
            # Si on veut seulement le contour extérieur, arrêter après le premier
            if outside_boundary_only:
                outer_boundary = False
                break
        
        room_data["geometry"] = boundary_locations
        output[str(room.Id.IntegerValue)] = room_data
    
    return output
