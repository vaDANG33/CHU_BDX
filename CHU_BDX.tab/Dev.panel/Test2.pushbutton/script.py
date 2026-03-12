# -*- coding: utf-8 -*-
# Author: Ahmed Helmy Abdelmagid (Optimized Version - Active View Only)
# Description: Creates wall dimensions in active view only

from pyrevit import DB, forms
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import UV
import time

uiapp = __revit__
app = uiapp.Application
doc = uiapp.ActiveUIDocument.Document
active_view = doc.ActiveView

OFFSET_MULTIPLIER_INCREMENT = 10
MAX_OFFSET_ATTEMPTS = 10
TOLERANCE = doc.Application.ShortCurveTolerance


class UIHandler:
    @staticmethod
    def get_user_input():
        """Interface simplifiée pour vue active uniquement"""
        
        # Vérifier que la vue active est une vue en plan
        if not isinstance(active_view, DB.ViewPlan):
            TaskDialog.Show(
                "Erreur", 
                "Cette commande fonctionne uniquement dans une vue en plan.\n"
                "Vue active: {}".format(active_view.ViewType)
            )
            return None, None, None
        
        selected_face = forms.ask_for_one_item(
            ["Internal", "External"], 
            default="External", 
            title="Dimension Face"
        )
        
        if selected_face is None:
            return None, None, None
        
        offset_distance = UIHandler.get_offset_distance()
        
        if offset_distance is None:
            return None, None, None

        offset_distance = DB.UnitUtils.ConvertToInternalUnits(
            offset_distance, DB.UnitTypeId.Millimeters
        )
        
        # NOUVEAU: Sélection du type de cotation
        dimension_type = UIHandler.get_dimension_type()
        
        if dimension_type is None:
            return None, None, None

        return selected_face, offset_distance, dimension_type

    @staticmethod
    def get_offset_distance():
        offset_distance_str = forms.ask_for_string(
            default="500",
            prompt="Enter the offset distance (in millimeters):",
            title="Offset Distance",
        )

        if offset_distance_str:
            try:
                offset_distance = float(offset_distance_str)
                if offset_distance <= 0:
                    TaskDialog.Show(
                        "Error", "Offset distance must be a positive number."
                    )
                    return None
                return offset_distance
            except ValueError:
                TaskDialog.Show(
                    "Error", "Invalid input. Please enter a valid number."
                )
                return None
        else:
            return None

    @staticmethod
    def get_dimension_type():
        """
        NOUVEAU: Permet de sélectionner le type de cotation dans le projet
        """
        # Collecter tous les types de cotation linéaire
        dimension_types = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.DimensionType)
            .ToElements()
        )
        
        # Filtrer pour ne garder que les types de cotation linéaire
        linear_dimension_types = [
            dt for dt in dimension_types 
            if dt.StyleType == DB.DimensionStyleType.Linear
        ]
        
        if not linear_dimension_types:
            TaskDialog.Show(
                "Erreur",
                "Aucun type de cotation linéaire trouvé dans le projet."
            )
            return None
        
        # Créer un dictionnaire nom → type
        dim_type_dict = {}
        for dt in linear_dimension_types:
            # Obtenir le nom complet avec famille
            family_name = dt.FamilyName if hasattr(dt, 'FamilyName') else "Cotation"
            type_name = dt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            full_name = "{} : {}".format(family_name, type_name)
            dim_type_dict[full_name] = dt
        
        # Si un seul type disponible, le sélectionner automatiquement
        if len(dim_type_dict) == 1:
            selected_type = list(dim_type_dict.values())[0]
            print("Type de cotation unique utilisé: {}".format(list(dim_type_dict.keys())[0]))
            return selected_type
        
        # Proposer la sélection à l'utilisateur
        selected_name = forms.SelectFromList.show(
            sorted(dim_type_dict.keys()),
            title="Sélectionner le type de cotation",
            button_name="Sélectionner",
            multiselect=False
        )
        
        if selected_name:
            return dim_type_dict[selected_name]
        else:
            return None


class GeometryHandler:
    # Caches optimisés
    wall_vector_cache = {}
    wall_solid_cache = {}
    wall_edges_cache = {}
    
    @staticmethod
    def clear_caches():
        """Nettoie tous les caches pour libérer la mémoire"""
        GeometryHandler.wall_vector_cache.clear()
        GeometryHandler.wall_solid_cache.clear()
        GeometryHandler.wall_edges_cache.clear()

    @staticmethod
    def get_wall_vectors(wall):
        wall_id = wall.Id.IntegerValue
        if wall_id in GeometryHandler.wall_vector_cache:
            return GeometryHandler.wall_vector_cache[wall_id]

        loc_line = wall.Location.Curve
        wall_dir = loc_line.Direction.Normalize()
        perp_dir = wall_dir.CrossProduct(DB.XYZ.BasisZ)

        GeometryHandler.wall_vector_cache[wall_id] = (wall_dir, perp_dir)
        return wall_dir, perp_dir

    @staticmethod
    def get_wall_solid(wall, options=None):
        """Récupère le solide d'un mur avec cache"""
        wall_id = wall.Id.IntegerValue
        
        if wall_id in GeometryHandler.wall_solid_cache:
            return GeometryHandler.wall_solid_cache[wall_id]
        
        options = options or DB.Options()
        solid = None
        
        for geometry_object in wall.get_Geometry(options):
            if isinstance(geometry_object, DB.Solid) and geometry_object.Faces.Size > 0:
                solid = geometry_object
                break
        
        GeometryHandler.wall_solid_cache[wall_id] = solid
        return solid

    @staticmethod
    def get_wall_outer_edges(wall, opts, dimension_face):
        """Récupère les edges extérieurs avec cache"""
        wall_id = wall.Id.IntegerValue
        cache_key = (wall_id, dimension_face)
        
        if cache_key in GeometryHandler.wall_edges_cache:
            return GeometryHandler.wall_edges_cache[cache_key]
        
        try:
            wall_solid = GeometryHandler.get_wall_solid(wall, opts)
            if not wall_solid:
                GeometryHandler.wall_edges_cache[cache_key] = []
                return []

            edges = []
            for face in wall_solid.Faces:
                for edge_loop in face.EdgeLoops:
                    for edge in edge_loop:
                        try:
                            edge_c = edge.AsCurve()
                            if isinstance(edge_c, DB.Line):
                                if edge_c.Direction.IsAlmostEqualTo(
                                    DB.XYZ.BasisZ
                                ) or edge_c.Direction.IsAlmostEqualTo(
                                    -DB.XYZ.BasisZ
                                ):
                                    edges.append(edge)
                        except:
                            continue

            if dimension_face == "External":
                edge_endpoints = {}
                outermost_edges = []
                for edge in edges:
                    edge_c = edge.AsCurve()
                    start_point = edge_c.GetEndPoint(0)
                    end_point = edge_c.GetEndPoint(1)
                    if (
                        (start_point, end_point) not in edge_endpoints
                        and (end_point, start_point) not in edge_endpoints
                    ):
                        outermost_edges.append(edge)
                        edge_endpoints[(start_point, end_point)] = True
                
                GeometryHandler.wall_edges_cache[cache_key] = outermost_edges
                return outermost_edges
            else:
                GeometryHandler.wall_edges_cache[cache_key] = edges
                return edges
                
        except Exception as e:
            print("Error occurred while processing wall {}: {}".format(wall.Id, str(e)))
            GeometryHandler.wall_edges_cache[cache_key] = []
            return []

    @staticmethod
    def get_reference_position(edge, wall, dimension_line):
        edge_curve = edge.AsCurve()
        edge_midpoint = edge_curve.Evaluate(0.5, True)
        wall_location = wall.Location.Curve
        intersection_result = dimension_line.Project(edge_midpoint)
        projected_point = intersection_result.XYZPoint
        position = wall_location.Project(projected_point).Parameter
        return position


class RevitGeometryUtils:
    @staticmethod
    def collect_walls_in_active_view(doc, view):
        """
        Collecte les murs visibles dans la vue active uniquement
        """
        walls = (
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(DB.BuiltInCategory.OST_Walls)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        return list(walls)

    @staticmethod
    def create_walls_dict(walls):
        """Crée un dictionnaire des murs pour recherche rapide"""
        return {wall.Id.IntegerValue: wall for wall in walls}


class RevitAPIUtils:
    @staticmethod
    def bboxes_intersect(bbox1, bbox2):
        """Test rapide d'intersection de BoundingBox"""
        return not (
            bbox1.Max.X < bbox2.Min.X or bbox1.Min.X > bbox2.Max.X or
            bbox1.Max.Y < bbox2.Min.Y or bbox1.Min.Y > bbox2.Max.Y or
            bbox1.Max.Z < bbox2.Min.Z or bbox1.Min.Z > bbox2.Max.Z
        )

    @staticmethod
    def find_intersecting_walls_optimized(wall, all_walls_dict):
        """
        Version optimisée de recherche des murs intersectants
        Filtre pour garder SEULEMENT les murs perpendiculaires
        """
        bbox = wall.get_BoundingBox(None)
        if not bbox:
            return []
        
        intersecting = []
        wall_id = wall.Id.IntegerValue
        wall_curve = wall.Location.Curve
        wall_direction = wall_curve.Direction.Normalize()
        
        # Tolérance pour considérer deux murs comme perpendiculaires
        PERPENDICULAR_TOLERANCE = 0.1  # ~6 degrés
        
        for other_id, other_wall in all_walls_dict.items():
            if other_id == wall_id:
                continue
            
            other_bbox = other_wall.get_BoundingBox(None)
            if not other_bbox:
                continue
            
            if not RevitAPIUtils.bboxes_intersect(bbox, other_bbox):
                continue
            
            try:
                other_curve = other_wall.Location.Curve
                
                # Vérifier l'intersection géométrique
                result = wall_curve.Intersect(other_curve)
                if result != DB.SetComparisonResult.Overlap:
                    continue
                
                # NOUVEAU: Vérifier que les murs sont perpendiculaires
                other_direction = other_curve.Direction.Normalize()
                dot_product = abs(wall_direction.DotProduct(other_direction))
                
                # Si dot_product proche de 0, les murs sont perpendiculaires
                # Si dot_product proche de 1, les murs sont parallèles
                if dot_product < PERPENDICULAR_TOLERANCE:
                    intersecting.append(other_wall)
                    
            except:
                continue
        
        return intersecting

    @staticmethod
    def is_space_free_for_dimension(doc, line, view):
        start_pt = line.GetEndPoint(0)
        end_pt = line.GetEndPoint(1)

        min_point = DB.XYZ(min(start_pt.X, end_pt.X), start_pt.Y - 0.5, 0)
        max_point = DB.XYZ(max(start_pt.X, end_pt.X), start_pt.Y + 0.5, 0)

        outline = DB.Outline(min_point, max_point)

        intersecting_dimensions = (
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(DB.BuiltInCategory.OST_Dimensions)
            .WherePasses(DB.BoundingBoxIntersectsFilter(outline))
            .WhereElementIsNotElementType()
            .ToElements()
        )

        return len(intersecting_dimensions) == 0

    @staticmethod
    def offset_dimension_line(line, offset_value):
        transform = DB.Transform.CreateTranslation(DB.XYZ(0, offset_value, 0))
        return line.CreateTransformed(transform)


class RevitTransactionManager:
    @staticmethod
    def create_geometry_options(view):
        """Créer les options une seule fois par vue"""
        opts = DB.Options()
        opts.ComputeReferences = True
        opts.IncludeNonVisibleObjects = True
        opts.View = view
        return opts

    @staticmethod
    def create_wall_dimensions(
        doc,
        view,
        selected_face,
        offset_distance,
        tolerance,
        dimension_type,
    ):
        """Version optimisée pour vue active uniquement"""
        start_time = time.time()
        
        stats = {
            'processed_walls': 0,
            'created_dimensions': 0,
            'failed_walls': 0
        }
        
        # Collecter tous les murs dans la vue active
        print("Collecte des murs dans la vue: {}".format(view.Name))
        walls_in_view = RevitGeometryUtils.collect_walls_in_active_view(doc, view)
        print("Total murs dans la vue: {}".format(len(walls_in_view)))
        print("Type de cotation: {}".format(dimension_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
        
        if not walls_in_view:
            TaskDialog.Show(
                "Information",
                "Aucun mur trouvé dans la vue active."
            )
            return stats
        
        # Créer dictionnaire pour recherche rapide
        all_walls_dict = RevitGeometryUtils.create_walls_dict(walls_in_view)
        
        # Créer les geometry options une seule fois
        geometry_options = RevitTransactionManager.create_geometry_options(view)
        
        existing_dim_endpoints = set()
        
        # Barre de progression
        with forms.ProgressBar(title="Création des cotations...") as pb:
            for i, wall in enumerate(walls_in_view):
                pb.update_progress(i + 1, len(walls_in_view))
                stats['processed_walls'] += 1
                
                try:
                    wall_dir, perp_dir = GeometryHandler.get_wall_vectors(wall)
                    existing_line = wall.Location.Curve

                    wall_ext_face_ref = list(
                        DB.HostObjectUtils.GetSideFaces(
                            wall, DB.ShellLayerType.Exterior
                        )
                    )[0]
                    wall_int_face_ref = list(
                        DB.HostObjectUtils.GetSideFaces(
                            wall, DB.ShellLayerType.Interior
                        )
                    )[0]

                    wall_ext_face = wall.GetGeometryObjectFromReference(
                        wall_ext_face_ref
                    )
                    wall_int_face = wall.GetGeometryObjectFromReference(
                        wall_int_face_ref
                    )

                    offset_dir = (
                        -perp_dir
                        if selected_face == "External"
                        else perp_dir
                    )
                    dimensioned_face = (
                        wall_ext_face
                        if selected_face == "External"
                        else wall_int_face
                    )

                    original_off_crv = existing_line.CreateTransformed(
                        DB.Transform.CreateTranslation(
                            offset_dir.Multiply(offset_distance)
                        )
                    )
                    off_crv = original_off_crv

                    vert_edge_sub = DB.ReferenceArray()
                    
                    # Mode "Wall Thickness" uniquement
                    vert_edges = GeometryHandler.get_wall_outer_edges(
                        wall, geometry_options, selected_face
                    )

                    # Murs intersectants optimisés
                    intersecting_walls = RevitAPIUtils.find_intersecting_walls_optimized(
                        wall, all_walls_dict
                    )
                    
                    for int_wall in intersecting_walls:
                        int_edges = GeometryHandler.get_wall_outer_edges(
                            int_wall, geometry_options, selected_face
                        )
                        vert_edges.extend(int_edges)

                    # Tri optimisé
                    ref_point = existing_line.GetEndPoint(0)
                    vert_edges.sort(
                        key=lambda e: e.AsCurve().GetEndPoint(0).DistanceTo(ref_point)
                    )
                    
                    reference_positions = set()

                    for edge in vert_edges:
                        value = round(
                            DB.UnitUtils.ConvertFromInternalUnits(
                                edge.ApproximateLength, DB.UnitTypeId.Millimeters
                            ),
                            2,
                        )
                        if value > tolerance:
                            ref_position = GeometryHandler.get_reference_position(
                                edge, wall, off_crv
                            )
                            if not any(
                                abs(ref_position - existing_position)
                                < tolerance
                                for existing_position in reference_positions
                            ):
                                vert_edge_sub.Append(edge.Reference)
                                reference_positions.add(ref_position)

                    if vert_edge_sub.Size >= 2:
                        line = off_crv
                        offset_multiplier = 1
                        
                        # Limite de sécurité
                        while offset_multiplier <= MAX_OFFSET_ATTEMPTS:
                            if RevitAPIUtils.is_space_free_for_dimension(doc, line, view):
                                break
                            offset_multiplier += 1
                            line = RevitAPIUtils.offset_dimension_line(
                                original_off_crv,
                                OFFSET_MULTIPLIER_INCREMENT * offset_multiplier,
                            )
                        
                        if offset_multiplier > MAX_OFFSET_ATTEMPTS:
                            print("  Impossible de placer dimension pour mur {}".format(wall.Id))
                            stats['failed_walls'] += 1
                            continue

                        dim_line = DB.Line.CreateBound(
                            line.GetEndPoint(0), line.GetEndPoint(1)
                        )
                        dim_tuple = tuple(
                            sorted((dim_line.GetEndPoint(0), dim_line.GetEndPoint(1)))
                        )
                        
                        if dim_tuple not in existing_dim_endpoints:
                            # NOUVEAU: Créer la dimension avec le type sélectionné
                            dim = doc.Create.NewDimension(view, dim_line, vert_edge_sub)
                            
                            # Appliquer le type de cotation sélectionné
                            if dim and dimension_type:
                                dim.DimensionType = dimension_type
                            
                            existing_dim_endpoints.add(dim_tuple)
                            stats['created_dimensions'] += 1

                except Exception as e:
                    stats['failed_walls'] += 1
                    print("  Failed to create dimension for wall {}: {}".format(wall.Id, e))
                    continue
        
        # Nettoyage des caches
        GeometryHandler.clear_caches()
        
        # Statistiques
        elapsed = time.time() - start_time
        print("\n" + "="*50)
        print("COTATION TERMINÉE - VUE: {}".format(view.Name))
        print("="*50)
        print("Temps d'exécution: {:.2f}s".format(elapsed))
        print("Murs traités: {}".format(stats['processed_walls']))
        print("Cotations créées: {}".format(stats['created_dimensions']))
        print("Échecs: {}".format(stats['failed_walls']))
        print("="*50)
        
        return stats


def main():
    try:
        # Vérifier que la vue active est valide
        if not isinstance(active_view, DB.ViewPlan):
            TaskDialog.Show(
                "Erreur",
                "Cette commande fonctionne uniquement dans une vue en plan.\n\n"
                "Vue active: {}\n"
                "Type: {}".format(active_view.Name, active_view.ViewType)
            )
            return
        
        # Interface utilisateur simplifiée
        selected_face, offset_distance, dimension_type = UIHandler.get_user_input()

        if offset_distance is None or selected_face is None or dimension_type is None:
            TaskDialog.Show("Warning", "Operation cancelled by the user.")
        else:
            with DB.Transaction(doc, "Auto Dimension Walls (Active View)") as transaction:
                try:
                    transaction.Start()
                    stats = RevitTransactionManager.create_wall_dimensions(
                        doc,
                        active_view,
                        selected_face,
                        offset_distance,
                        TOLERANCE,
                        dimension_type,
                    )
                    transaction.Commit()
                    
                    # Message de succès
                    if stats['created_dimensions'] > 0:
                        dim_type_name = dimension_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                        success_msg = "Cotation terminée avec succès!\n\n"
                        success_msg += "Vue: {}\n".format(active_view.Name)
                        success_msg += "Type de cotation: {}\n".format(dim_type_name)
                        success_msg += "Murs traités: {}\n".format(stats['processed_walls'])
                        success_msg += "Cotations créées: {}\n".format(stats['created_dimensions'])
                        if stats['failed_walls'] > 0:
                            success_msg += "Échecs: {}".format(stats['failed_walls'])
                        
                        TaskDialog.Show("Succès", success_msg)
                    else:
                        TaskDialog.Show(
                            "Information",
                            "Aucune cotation créée.\n\n"
                            "Vérifiez que les murs ont des edges valides."
                        )
                    
                except Exception as e:
                    transaction.RollBack()
                    TaskDialog.Show(
                        "Error",
                        "An error occurred while dimensioning walls:\n{}".format(str(e)),
                    )
                    raise
    except Exception as e:
        TaskDialog.Show("Error", "An unexpected error occurred:\n{}".format(str(e)))
        raise


if __name__ == "__main__":
    main()