# -*- coding: utf-8 -*-
__title__ = "Export IFC\nARC"
__doc__ = """Export IFC des maquettes *_ARC.rvt
depuis un dossier et ses sous-dossiers.
- Gère les fichiers de versions antérieures (ex: 2023 sur 2024)
- Crée la vue 3D_IFC_EXPORT si absente (phase 1)
- Ferme sans enregistrer dans tous les cas"""

import os
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import IFailuresPreprocessor
from pyrevit import forms, script, HOST_APP

logger = script.get_logger()
output = script.get_output()


# ──────────────────────────────────────────────────────────────
# HELPER : supprimer les avertissements de migration
# ──────────────────────────────────────────────────────────────
class SilentFailuresPreprocessor(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for failure in failuresAccessor.GetFailureMessages():
            if failure.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue


# ──────────────────────────────────────────────────────────────
# HELPER : ouvrir un document
# ──────────────────────────────────────────────────────────────
def open_doc(app, file_path, detach, is_workshared):
    opts = OpenOptions()
    if is_workshared and detach:
        opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
    elif is_workshared and not detach:
        opts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
        wc = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        opts.SetOpenWorksetsConfiguration(wc)
    else:
        opts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
    mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)
    return app.OpenDocumentFile(mp, opts)


# ──────────────────────────────────────────────────────────────
# HELPER : chercher la vue 3D_IFC_EXPORT
# ──────────────────────────────────────────────────────────────
def find_ifc_view(doc):
    for v in FilteredElementCollector(doc).OfClass(View3D):
        if v.Name == "3D_IFC_EXPORT" and not v.IsTemplate:
            return v
    return None


# ──────────────────────────────────────────────────────────────
# HELPER : créer la vue 3D_IFC_EXPORT avec la 1ère phase
# ──────────────────────────────────────────────────────────────
def create_ifc_view(doc):
    if doc.IsReadOnly:
        raise Exception("Document en lecture seule — impossible de créer la vue")

    vft = next(
        (v for v in FilteredElementCollector(doc).OfClass(ViewFamilyType)
         if v.ViewFamily == ViewFamily.ThreeDimensional),
        None
    )
    if not vft:
        raise Exception("Aucun type de vue 3D disponible")

    t = Transaction(doc, "Créer vue 3D_IFC_EXPORT")
    # ✅ Supprimer les avertissements dans la transaction
    t.SetFailureHandlingOptions(
        t.GetFailureHandlingOptions().SetFailuresPreprocessor(SilentFailuresPreprocessor())
    )
    try:
        t.Start()
        view_3d = View3D.CreateIsometric(doc, vft.Id)
        view_3d.Name = "3D_IFC_EXPORT"

        first_phase = doc.Phases.get_Item(0)
        phase_param = view_3d.get_Parameter(BuiltInParameter.VIEW_PHASE)
        if phase_param and not phase_param.IsReadOnly:
            phase_param.Set(first_phase.Id)

        t.Commit()
        output.print_md("  - ✅ Vue créée et commitée")
        return view_3d, first_phase.Name

    except Exception as e:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
        raise Exception("Erreur création vue : {}".format(str(e)))


# ──────────────────────────────────────────────────────────────
# HELPER : export IFC
# ✅ Transaction requise par l'API Revit — RollBack après export
# ──────────────────────────────────────────────────────────────
def export_ifc(doc, view_3d, dir_path, base_name):
    ifc_options = IFCExportOptions()
    ifc_options.FilterViewId = view_3d.Id
    ifc_options.WallAndColumnSplitting = False
    ifc_options.ExportBaseQuantities = False
    ifc_options.SpaceBoundaryLevel = 0
    ifc_options.FileVersion = IFCVersion.IFC2x3CV2
    ifc_options.AddOption("Export2DElements", "false")
    ifc_options.AddOption("ExportInternalRevitPropertySets", "false")
    ifc_options.AddOption("ExportIFCCommonPropertySets", "true")
    ifc_options.AddOption("ExportLinkedFiles", "false")
    ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
    ifc_options.AddOption("UseActiveViewGeometry", "true")

    t = Transaction(doc, "Export IFC")
    # ✅ Supprimer les avertissements pendant l'export
    t.SetFailureHandlingOptions(
        t.GetFailureHandlingOptions().SetFailuresPreprocessor(SilentFailuresPreprocessor())
    )
    t.Start()
    try:
        success = doc.Export(dir_path, base_name, ifc_options)
    finally:
        t.RollBack()  # ✅ RollBack : export écrit sur disque, doc non modifié
    return success


# ──────────────────────────────────────────────────────────────
# 1. Sélection du dossier
# ──────────────────────────────────────────────────────────────
folder = forms.pick_folder(title="Sélectionner le dossier racine des maquettes ARC")
if not folder:
    forms.alert("Aucun dossier sélectionné.", exitscript=True)


# ──────────────────────────────────────────────────────────────
# 2. Récupérer tous les *_ARC.rvt récursivement
# ──────────────────────────────────────────────────────────────
rvt_files = [
    os.path.join(root, f)
    for root, dirs, files in os.walk(folder)
    for f in files
    if "_arc" in f.lower() and f.lower().endswith(".rvt")
]

if not rvt_files:
    forms.alert("Aucun fichier *_ARC.rvt trouvé.", exitscript=True)


# ──────────────────────────────────────────────────────────────
# 3. Aperçu dans l'output
# ──────────────────────────────────────────────────────────────
output.print_md("**Fichiers trouvés :** `{}`\n".format(len(rvt_files)))
for f in sorted(rvt_files)[:5]:
    output.print_md("- `{}`".format(os.path.relpath(f, folder)))
if len(rvt_files) > 5:
    output.print_md("- *... et {} autres fichiers*".format(len(rvt_files) - 5))


# ──────────────────────────────────────────────────────────────
# 4. Sélection dans une liste
# ──────────────────────────────────────────────────────────────
class RvtFileItem(forms.TemplateListItem):
    @property
    def name(self):
        rel = os.path.relpath(self.item, folder)
        return "{:<45} {}".format(
            os.path.basename(self.item),
            os.path.dirname(rel) or "."
        )

options = [RvtFileItem(p) for p in sorted(rvt_files)]

selected_items = forms.SelectFromList.show(
    options,
    title="Sélectionner les maquettes à exporter en IFC",
    button_name="Exporter en IFC",
    multiselect=True
)

if not selected_items:
    forms.alert("Aucun fichier sélectionné.", exitscript=True)

selected_paths = [
    item.item if hasattr(item, 'item') else item
    for item in selected_items
]


# ──────────────────────────────────────────────────────────────
# 5. Application Revit
# ──────────────────────────────────────────────────────────────
app = HOST_APP.uiapp.Application


# ──────────────────────────────────────────────────────────────
# 6. Boucle d'export
# ──────────────────────────────────────────────────────────────
results = []
doc = None

for file_path in selected_paths:
    fname = os.path.basename(file_path)
    output.print_md("---\n### ⚙️ Traitement : `{}`".format(fname))

    try:
        # ── Détection metadata sans ouvrir ────────────────────
        try:
            basic_info = BasicFileInfo.Extract(file_path)
            is_workshared = basic_info.IsWorkshared
            is_current_version = basic_info.SavedInCurrentVersion
        except:
            is_workshared = False
            is_current_version = True

        output.print_md("- Workshared : `{}`".format(is_workshared))
        if not is_current_version:
            output.print_md("- ⚠️ Version antérieure détectée — migration en mémoire")

        # ── Ouverture légère ──────────────────────────────────
        doc = open_doc(app, file_path, detach=False, is_workshared=is_workshared)
        if not doc:
            raise Exception("Impossible d'ouvrir le fichier")

        # ── Chercher la vue ───────────────────────────────────
        view_3d = find_ifc_view(doc)

        if view_3d:
            output.print_md("- ✅ Vue `3D_IFC_EXPORT` trouvée")
        else:
            doc.Close(False)
            doc = None
            output.print_md("- ⚠️ Vue absente — réouverture en mode détaché...")

            doc = open_doc(app, file_path, detach=True, is_workshared=is_workshared)
            if not doc:
                raise Exception("Impossible de rouvrir en mode détaché")

            view_3d, phase_name = create_ifc_view(doc)
            output.print_md("- ✅ Vue créée (phase : *{}*)".format(phase_name))

        # ── Export IFC ────────────────────────────────────────
        dir_path = os.path.dirname(file_path)
        base_name = os.path.splitext(fname)[0]

        output.print_md("- Export IFC en cours...")
        success = export_ifc(doc, view_3d, dir_path, base_name)

        doc.Close(False)
        doc = None

        if success:
            msg = "✅ Export OK → `{}.ifc`".format(base_name)
        else:
            msg = "⚠️ Export échoué (raison inconnue)"
        output.print_md("- " + msg)
        results.append("{} → {}".format(fname, msg))

    except Exception as ex:
        msg = "❌ Erreur : {}".format(str(ex))
        output.print_md("- " + msg)
        results.append("{} → {}".format(fname, msg))
        if doc:
            try:
                if doc.IsValidObject:
                    doc.Close(False)
            except:
                pass
            doc = None


# ──────────────────────────────────────────────────────────────
# 7. Récapitulatif final
# ──────────────────────────────────────────────────────────────
output.print_md("\n---\n## 📋 Récapitulatif")
output.print_md("**Total traité :** {} fichier(s)\n".format(len(selected_paths)))

successes = [r for r in results if "✅" in r]
warnings  = [r for r in results if "⚠️" in r]
errors    = [r for r in results if "❌" in r]

if successes:
    output.print_md("### ✅ Succès ({})".format(len(successes)))
    for r in successes:
        output.print_md("- {}".format(r))
if warnings:
    output.print_md("### ⚠️ Avertissements ({})".format(len(warnings)))
    for r in warnings:
        output.print_md("- {}".format(r))
if errors:
    output.print_md("### ❌ Erreurs ({})".format(len(errors)))
    for r in errors:
        output.print_md("- {}".format(r))

output.print_md("\n---\n**🏁 Traitement terminé !**")
