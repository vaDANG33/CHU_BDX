# -*- coding: utf-8 -*-
import clr, os
clr.AddReference("QRCoder")
clr.AddReference("PdfSharp")

from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script
from System.Drawing import Bitmap, Graphics, Font, FontStyle, Brushes, StringAlignment, StringFormat, Color, Pen, RectangleF
from System.Drawing.Imaging import ImageFormat
from System.IO import MemoryStream
from QRCoder import QRCodeGenerator, PngByteQRCode
from PdfSharp.Pdf import PdfDocument
from PdfSharp.Drawing import XGraphics, XImage, XUnit

# =========================================================================
# CONFIGURATION
# =========================================================================
# Dimensions physiques (cm)
W_CM = 6
H_CM = 3

# Conversion cm → points (1 cm = 28.35 points)
CM_TO_PT = 28.35
W_PT = W_CM * CM_TO_PT  # 170.1 points
H_PT = H_CM * CM_TO_PT  # 85.05 points

# Résolution bitmap (300 DPI pour qualité haute)
DPI = 300
W_PX = int(W_CM * DPI / 2.54)  # 709 pixels
H_PX = int(H_CM * DPI / 2.54)  # 354 pixels

# Positions internes (pixels)
QR_SZ  = H_PX - 20
SEP_X  = QR_SZ + 20
TEXT_X = SEP_X + 20

# Références Revit et pyRevit
doc    = revit.doc
output = script.get_output()

# Objets réutilisables (performance)
_qr_gen = QRCodeGenerator()
_font = Font("Arial", 16, FontStyle.Bold)
_brush_black = Brushes.Black
_pen_sep = Pen(Color.LightSteelBlue, 2)

# =========================================================================
# FONCTION 1 : Récupération des pièces
# =========================================================================
def get_selected_rooms():
    """Récupère et affiche les pièces disponibles pour sélection."""
    first_phase = list(doc.Phases)[0]
    
    rooms = [
        r for r in FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_Rooms)
                  .WhereElementIsNotElementType()
        if r.get_Parameter(BuiltInParameter.ROOM_PHASE).AsElementId() == first_phase.Id
    ]
    
    def room_label(r):
        num  = r.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or ""
        name = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or ""
        return "{} - {}".format(num, name)
    
    class RoomItem(object):
        def __init__(self, room):
            self.room = room
            self.label = room_label(room)
        def __str__(self):
            return self.label
    
    room_items = sorted([RoomItem(r) for r in rooms], key=lambda x: x.label)
    
    selected = forms.SelectFromList.show(
        room_items,
        title="Sélection des pièces",
        multiselect=True,
        button_name="Générer les QR codes"
    )
    
    return selected

# =========================================================================
# FONCTION 2 : Choix du fichier de sortie
# =========================================================================
def get_pdf_output_path():
    """Demande le chemin de sauvegarde du PDF."""
    pdf_output = forms.save_file(
        file_ext="pdf",
        default_name="{}_Etiquettes.pdf".format(doc.Title),
        title="Enregistrer le PDF des étiquettes"
    )
    
    if not pdf_output:
        forms.alert("Enregistrement annulé.")
        return None
    
    return pdf_output

# =========================================================================
# FONCTION 3 : Génération du code QR
# =========================================================================
def make_qr_image(data):
    """Génère un code QR en PNG."""
    qr_code = _qr_gen.CreateQrCode(data, QRCodeGenerator.ECCLevel.Q)
    qr_bytes = PngByteQRCode(qr_code).GetGraphic(20)
    return Bitmap(MemoryStream(qr_bytes))

# =========================================================================
# FONCTION 4 : Création de l'étiquette
# =========================================================================
def create_label(room_number):
    """Crée une étiquette avec QR code et texte."""
    bmp = Bitmap(W_PX, H_PX)
    bmp.SetResolution(DPI, DPI)  # Métadonnées (ignoré par PdfSharp)
    
    with Graphics.FromImage(bmp) as g:
        g.Clear(Color.White)
        
        # QR Code
        qr_img = make_qr_image(room_number)
        g.DrawImage(qr_img, 10, 10, QR_SZ, QR_SZ)
        qr_img.Dispose()
        
        # Ligne séparatrice
        g.DrawLine(_pen_sep, SEP_X, 15, SEP_X, H_PX - 15)
        
        # Texte (numéro de pièce)
        sf = StringFormat()
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        
        text_rect = RectangleF(TEXT_X, 0, W_PX - TEXT_X - 10, H_PX)
        g.DrawString(room_number, _font, _brush_black, text_rect, sf)
    
    return bmp

# =========================================================================
# FONCTION 5 : Assemblage du PDF
# =========================================================================
def build_pdf(items, pdf_path):
    """Génère un PDF avec les étiquettes."""
    pdf = PdfDocument()
    success_count = 0
    error_count = 0
    
    for idx, item in enumerate(items, 1):
        room_num = item.room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        
        try:
            # Créer la bitmap
            bmp = create_label(room_num)
            
            # Convertir en PNG en mémoire
            png_stream = MemoryStream()
            bmp.Save(png_stream, ImageFormat.Png)
            png_stream.Position = 0
            bmp.Dispose()
            
            # Charger dans PdfSharp
            ximg = XImage.FromStream(png_stream)
            
            # Conversion pixels → points (72 DPI logique)
            w_pt = W_PX * 72.0 / DPI
            h_pt = H_PX * 72.0 / DPI
            
            # Créer une page
            page = pdf.AddPage()
            page.Width = XUnit.FromPoint(w_pt)
            page.Height = XUnit.FromPoint(h_pt)
            
            # Dessiner l'image
            gfx = XGraphics.FromPdfPage(page)
            gfx.DrawImage(ximg, 0, 0, w_pt, h_pt)
            
            output.print_md("✅ [{}/{}] **{}**".format(idx, len(items), room_num))
            success_count += 1
            
        except Exception as e:
            output.print_md("❌ [{}/{}] {} : {}".format(idx, len(items), room_num, str(e)))
            error_count += 1
    
    # Sauvegarder
    pdf.Save(pdf_path)
    
    return success_count, error_count

# =========================================================================
# FONCTION 6 : Ouverture optionnelle du PDF
# =========================================================================
def open_pdf_if_requested(pdf_path):
    """Demande et ouvre le PDF si l'utilisateur le souhaite."""
    if forms.ask_for_one_item(
        ["Oui, ouvrir", "Non"],
        title="Ouvrir le PDF ?",
        default="Oui, ouvrir"
    ) == "Oui, ouvrir":
        try:
            os.startfile(pdf_path)
        except Exception as e:
            output.print_md("⚠️ Impossible d'ouvrir le PDF : {}".format(str(e)))

# =========================================================================
# MAIN - EXÉCUTION PRINCIPALE
# =========================================================================
def main():
    """Fonction principale."""
    try:
        # Étape 1 : Sélection des pièces
        output.print_md("**Récupération des pièces...**\n")
        selected = get_selected_rooms()
        
        if not selected:
            output.print_md("❌ Aucune pièce sélectionnée.")
            return False
        
        output.print_md("✅ {} pièce(s) sélectionnée(s)\n".format(len(selected)))
        
        # Étape 2 : Chemin de sauvegarde
        output.print_md("**Choix du fichier de sortie...**\n")
        pdf_path = get_pdf_output_path()
        
        if not pdf_path:
            output.print_md("❌ Opération annulée.")
            return False
        
        # Étape 3 : Génération
        output.print_md("**Génération en cours...**\n")
        success, errors = build_pdf(selected, pdf_path)
        
        # Résumé
        output.print_md("\n---\n")
        output.print_md("📊 **Résumé :**\n")
        output.print_md("- ✅ Réussi : **{}**".format(success))
        output.print_md("- ❌ Erreurs : **{}**".format(errors))
        output.print_md("- 📄 PDF : `{}`\n".format(pdf_path))
        
        # Étape 4 : Ouverture
        open_pdf_if_requested(pdf_path)
        
        output.print_md("---\n**✅ Terminé !**")
        return True
        
    except Exception as e:
        output.print_md("❌ **Erreur critique** : {}".format(str(e)))
        forms.alert("Erreur : {}".format(str(e)))
        return False

# =========================================================================
# EXÉCUTION
# =========================================================================
if __name__ == "__main__":
    main()
