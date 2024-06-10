from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import SimpleDocTemplate, Frame, Paragraph
from reportlab.platypus import BaseDocTemplate, PageTemplate
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT

def create_address_labels(filename, addresses):
    c = canvas.Canvas(filename, pagesize=A4)
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    width, height = A4
    
    label_width = width / 2
    label_height = height / 4

    margin_left = 10 * mm
    margin_top = label_height - 15 * mm
    line_spacing = 25

    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    styleN.fontName = 'HeiseiKakuGo-W5'

    c.setFont('HeiseiKakuGo-W5', 6*mm)

    for i, address in enumerate(addresses):
        row = i // 2
        col = i % 2
        x_offset = col * label_width
        y_offset = height - (row + 1) * label_height

        y_position = y_offset + margin_top

        c.drawString(x_offset + margin_left, y_position, address[0])
        y_position -= 35 * mm

        frames=[Frame(x_offset + margin_left, y_position, 85*mm, 30*mm, showBoundary=0),
                Frame(x_offset + margin_left, y_position-30, 70*mm, 13*mm, showBoundary=0),
                Frame(x_offset + margin_left, y_position-30, 70*mm, 20*mm, showBoundary=0),
                Frame(x_offset + margin_left+70*mm, y_position-30, 13*mm, 13*mm, showBoundary=0),]
        
        styleN.alignment=TA_LEFT
        styleN.fontSize = 5*mm
        styleN.leading = 18
        frames[0].addFromList([Paragraph(address[1], styleN)], c)

        styleN.fontSize = 7*mm
        styleN.leading = 20
        styleN.alignment=TA_RIGHT
        w, h = Paragraph(address[2], styleN).wrap(70*mm, 13*mm)
        if h <= 8*mm:
            frames[1].addFromList([Paragraph(address[2], styleN)], c)
        else:
            frames[2].addFromList([Paragraph(address[2], styleN)], c)

        frames[3].addFromList([Paragraph("様", styleN)], c)

    c.save()


addresses = [
    ["XXX-XXXX", "ABC県DEF市G区XXX-YYYYYYYYネコチャンフンワリマンション Ca棟 XXX-X", "長い名前"] for _ in range(8)
]


create_address_labels("address_labels.pdf", addresses)