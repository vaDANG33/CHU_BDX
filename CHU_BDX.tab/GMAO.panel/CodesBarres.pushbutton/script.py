# -*- coding: utf-8 -*-
import clr, os, tempfile
clr.AddReference("QRCoder")
clr.AddReference("PdfSharp")

from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script
from System.Drawing import Bitmap, Graphics, Font, FontStyle, Brushes, StringAlignment, StringFormat, Color, Pen, RectangleF
from System.Drawing.Imaging import ImageFormat
from System.IO import MemoryStream
from QRCoder import QRCodeGenerator, PngByteQRCode
from PdfSharp.Pdf import PdfDocument
from PdfSharp.Drawing import XGraphics, XImage

# ---------------------------------------------------------
# Constantes
# ---------------------------------------------------------
DPI    = 300
W_PX   = int(6 * DPI / 2.54)
H_PX   = int(3 * DPI / 2.54)
QR_SZ  = H_PX - 20
SEP_X  = QR_SZ + 20
TEXT_X = SEP_X + 20

doc    = revit.doc
output = script.get_output()

# ---------------------------------------------------------
# 1. Sélection des pièces
# ---------------------------------------------------------
first_phase = list(doc.Phases)[0]

rooms = [
    r for r in FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_Rooms)
              .WhereElementIsNotElementType()
    if r.get_Parameter(BuiltInParameter.ROOM_PHASE).AsElementId() == first_phase.Id
]

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

room_items = sorted([RoomItem(r) for r in rooms], key=lambda x: x.label)

selected = forms.SelectFromList.show(
    room_items,
    title="Sélection des pièces",
    multiselect=True,
    button_name="Générer les QR codes"
)
if not selected:
    forms.alert("Aucune pièce sélectionnée.")
    raise SystemExit

# ---------------------------------------------------------
# 2. Choix destination PDF
# ---------------------------------------------------------
PDF_OUTPUT = forms.save_file(
    file_ext="pdf",
    default_name="{}_Etiquettes.pdf".format(doc.Title),
    title="Enregistrer le PDF des étiquettes"
)

if not PDF_OUTPUT:
    forms.alert("Enregistrement annulé.")
    raise SystemExit

# Dossier temporaire pour les PNG intermédiaires
TEMP_FOLDER = tempfile.mkdtemp()

# ---------------------------------------------------------
# 3. Génération QR + étiquette (tout en mémoire)
# ---------------------------------------------------------
_qr_gen = QRCodeGenerator()

def make_qr_image(data):
    qr_bytes = PngByteQRCode(_qr_gen.CreateQrCode(data, QRCodeGenerator.ECCLevel.Q)).GetGraphic(20)
    return Bitmap(MemoryStream(qr_bytes))

def create_label(room_number):
    bmp = Bitmap(W_PX, H_PX)
    bmp.SetResolution(DPI, DPI)
    with Graphics.FromImage(bmp) as g:
        g.Clear(Color.White)
        g.DrawImage(make_qr_image(room_number), 10, 10, QR_SZ, QR_SZ)
        g.DrawLine(Pen(Color.LightSteelBlue, 2), SEP_X, 15, SEP_X, H_PX - 15)
        sf = StringFormat()
        sf.Alignment     = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        g.DrawString(room_number, Font("Arial", 16, FontStyle.Bold),
                     Brushes.Black, RectangleF(TEXT_X, 0, W_PX - TEXT_X - 10, H_PX), sf)
    return bmp

# ---------------------------------------------------------
# 4. Assemblage PDF (PNG temporaires supprimés après)
# ---------------------------------------------------------
def build_pdf(items, pdf_path):
    pdf = PdfDocument()
    for item in items:
        room_num = item.room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        try:
            # Générer PNG temporaire
            png_path = os.path.join(TEMP_FOLDER, room_num + ".png")
            bmp = create_label(room_num)
            bmp.Save(png_path, ImageFormat.Png)
            bmp.Dispose()

            # Ajouter au PDF
            ximg      = XImage.FromFile(png_path)
            w_pt, h_pt = ximg.PixelWidth / DPI * 72, ximg.PixelHeight / DPI * 72
            page      = pdf.AddPage()
            page.Width, page.Height = w_pt, h_pt
            XGraphics.FromPdfPage(page).DrawImage(ximg, 0, 0, w_pt, h_pt)

            # Supprimer PNG temporaire immédiatement
            os.remove(png_path)
            output.print_md("✅ **{}**".format(room_num))
            
        except Exception as e:
            output.print_md("❌ {} : {}".format(room_num, e))
    
    pdf.Save(pdf_path)

# ---------------------------------------------------------
# 5. Exécution + ouverture optionnelle
# ---------------------------------------------------------
output.print_md("**Génération en cours...**\n")
build_pdf(selected, PDF_OUTPUT)
output.print_md("\n📄 **PDF :** `{}`".format(PDF_OUTPUT))

if forms.ask_for_one_item(
    ["Oui, ouvrir", "Non"],
    title="Ouvrir le PDF ?",
    default="Oui, ouvrir"
) == "Oui, ouvrir":
    os.startfile(PDF_OUTPUT)

output.print_md("\n---\n**Terminé !**")
