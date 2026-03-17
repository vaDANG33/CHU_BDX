# -*- coding: utf-8 -*-
import clr, os, tempfile
clr.AddReference("QRCoder")
clr.AddReference("PdfSharp")

from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script
from System.Drawing import Bitmap, Graphics, Font, FontStyle, Brushes, Color, Pen, RectangleF, StringFormat, StringAlignment
from System.Drawing.Imaging import ImageFormat
from System.IO import MemoryStream
from QRCoder import QRCodeGenerator, PngByteQRCode
from PdfSharp.Pdf import PdfDocument
from PdfSharp.Drawing import XGraphics, XImage, XUnit

# =========================================================================
# CONFIGURATION
# =========================================================================
W_CM   = 6
H_CM   = 3
DPI    = 300
CM_TO_PT = 28.35

# Points exacts (PdfSharp)
W_PT = W_CM * CM_TO_PT   # 170.1 pt
H_PT = H_CM * CM_TO_PT   # 85.05 pt

# Pixels (bitmap raster)
W_PX = int(W_CM * DPI / 2.54)
H_PX = int(H_CM * DPI / 2.54)

# Layout interne
QR_SZ  = H_PX - 20
SEP_X  = QR_SZ + 20
TEXT_X = SEP_X + 20

doc    = revit.doc
output = script.get_output()

# Sélectionner la première phase (existe toujours)
first_phase = list(doc.Phases)[0]

# =========================================================================
# FONCTION 1 : Sélection des pièces
# =========================================================================
def get_selected_rooms():
    """Récupère les pièces de la première phase existante."""
    
    # Collecter les pièces de la première phase
    rooms = [
        r for r in FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_Rooms)
                  .WhereElementIsNotElementType()
        if r.get_Parameter(BuiltInParameter.ROOM_PHASE).AsElementId() == first_phase.Id
    ]

    if not rooms:
        forms.alert("Aucune pièce trouvée dans la première phase.")
        return None

    class RoomItem(object):
        def __init__(self, r):
            self.room = r
            num = r.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or ""
            name = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or ""
            self.label = "{} - {}".format(num, name)
        
        def __str__(self):
            return self.label

    room_items = sorted([RoomItem(r) for r in rooms], key=lambda x: x.label)

    return forms.SelectFromList.show(
        room_items,
        title="Sélection des pièces",
        multiselect=True,
        button_name="Générer les QR codes"
    )

# =========================================================================
# FONCTION 2 : Génération QR + étiquette bitmap
# =========================================================================
def make_qr_bitmap(data, qr_gen):
    """Génère un code QR en PNG."""
    qr_bytes = PngByteQRCode(
        qr_gen.CreateQrCode(data, QRCodeGenerator.ECCLevel.Q)
    ).GetGraphic(20)
    return Bitmap(MemoryStream(qr_bytes))

def create_label(room_number, qr_gen, font, pen_sep):
    """Crée une étiquette bitmap (300 DPI)."""
    bmp = Bitmap(W_PX, H_PX)
    bmp.SetResolution(DPI, DPI)

    g = Graphics.FromImage(bmp)
    try:
        g.Clear(Color.White)

        # QR code
        qr_img = make_qr_bitmap(room_number, qr_gen)
        try:
            g.DrawImage(qr_img, 10, 10, QR_SZ, QR_SZ)
        finally:
            qr_img.Dispose()

        # Séparateur vertical
        g.DrawLine(pen_sep, SEP_X, 15, SEP_X, H_PX - 15)

        # Texte centré
        sf = StringFormat()
        sf.Alignment     = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        g.DrawString(
            room_number, font, Brushes.Black,
            RectangleF(TEXT_X, 0, W_PX - TEXT_X - 10, H_PX), sf
        )
    finally:
        g.Dispose()

    return bmp

# =========================================================================
# FONCTION 3 : Assemblage PDF
# =========================================================================
def build_pdf(items, pdf_path):
    """Génère un PDF avec les étiquettes."""
    # Ressources GDI+ partagées
    qr_gen  = QRCodeGenerator()
    font    = Font("Arial", 16, FontStyle.Bold)
    pen_sep = Pen(Color.LightSteelBlue, 2)
    tmp_dir = tempfile.mkdtemp()

    pdf = PdfDocument()
    success = errors = 0

    try:
        for idx, item in enumerate(items, 1):
            room_num = item.room.get_Parameter(
                BuiltInParameter.ROOM_NUMBER
            ).AsString() or "SANS_NUM_{}".format(idx)

            try:
                # 1. Bitmap → PNG temporaire
                png_path = os.path.join(tmp_dir, "room_{}.png".format(idx))
                bmp = create_label(room_num, qr_gen, font, pen_sep)
                bmp.Save(png_path, ImageFormat.Png)
                bmp.Dispose()

                # 2. Page PDF aux dimensions exactes
                ximg = XImage.FromFile(png_path)
                page = pdf.AddPage()
                page.Width  = XUnit.FromPoint(W_PT)
                page.Height = XUnit.FromPoint(H_PT)

                # 3. Dessin + nettoyage immédiat
                gfx = XGraphics.FromPdfPage(page)
                try:
                    gfx.DrawImage(ximg, 0, 0, W_PT, H_PT)
                finally:
                    gfx.Dispose()
                    ximg.Dispose()

                # 4. Supprimer le PNG temporaire
                try:
                    os.remove(png_path)
                except Exception:
                    pass

                output.print_md("✅ [{}/{}] **{}**".format(idx, len(items), room_num))
                success += 1

            except Exception as e:
                output.print_md("❌ [{}/{}] {} : {}".format(idx, len(items), room_num, str(e)))
                errors += 1

        pdf.Save(pdf_path)

    finally:
        # Nettoyage GDI+ garanti
        font.Dispose()
        pen_sep.Dispose()
        try:
            os.rmdir(tmp_dir)
        except Exception:
            pass

    return success, errors

# =========================================================================
# MAIN
# =========================================================================
def main():
    try:
        # Étape 1 : Sélection
        output.print_md("**Récupération des pièces de la phase : {}**\n".format(first_phase.Name))
        selected = get_selected_rooms()
        
        if not selected:
            output.print_md("❌ Aucune pièce sélectionnée.")
            return

        output.print_md("✅ {} pièce(s) sélectionnée(s)\n".format(len(selected)))

        # Étape 2 : Destination PDF
        pdf_path = forms.save_file(
            file_ext="pdf",
            default_name="{}_Etiquettes.pdf".format(doc.Title),
            title="Enregistrer le PDF des étiquettes"
        )
        if not pdf_path:
            output.print_md("❌ Enregistrement annulé.")
            return

        # Étape 3 : Génération
        output.print_md("**Génération en cours...**\n")
        success, errors = build_pdf(selected, pdf_path)

        # Résumé
        output.print_md("\n---")
        output.print_md("✅ Réussi : **{}** | ❌ Erreurs : **{}**".format(success, errors))
        output.print_md("📄 PDF : `{}`".format(pdf_path))

        # Étape 4 : Ouverture
        if forms.alert("Ouvrir le PDF généré ?", yes=True, no=True):
            try:
                os.startfile(pdf_path)
            except Exception as e:
                output.print_md("⚠️ Impossible d'ouvrir : {}".format(str(e)))

        output.print_md("\n---\n**✅ Terminé !**")

    except Exception as e:
        output.print_md("❌ **Erreur critique** : {}".format(str(e)))
        forms.alert("Erreur critique : {}".format(str(e)))

if __name__ == "__main__":
    main()
