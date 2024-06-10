from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import SimpleDocTemplate, Frame, Paragraph
from reportlab.platypus import BaseDocTemplate, PageTemplate
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT

import pandas as pd
import re

csv_path = "./booth_orders.csv"
pattern = re.compile(r'([\d|A-Za-z|\-|\.]+|\d+|[^0-9A-Za-z\-]+|[A-Za-z]+)')

def create_address_labels(filename, address_data):
    c = canvas.Canvas(filename, pagesize=A4)
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    width, height = A4
    
    label_width = width / 2
    label_height = height / 4

    margin_left = 10*mm
    margin_top = label_height - 15*mm
    line_spacing = 9*mm

    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    styleN.fontName = 'HeiseiKakuGo-W5'

    c.setFont('HeiseiKakuGo-W5', 6*mm)
    for i in range(len(address_data)):
        if i%8==0 and i!=0:
            c.showPage()
            c.setFont('HeiseiKakuGo-W5', 6*mm)
        row = (i%8) // 2
        col = i % 2
        x_offset = col * label_width
        y_offset = height - (row + 1) * label_height

        y_position = y_offset + margin_top

        c.drawString(x_offset + margin_left, y_position, address_data["郵便番号"][i])
        y_position -= 35*mm

        frames=[Frame(x_offset + margin_left, y_position, 85*mm, 30*mm, showBoundary=0),
                Frame(x_offset + margin_left, y_position-11*mm, 70*mm, 13*mm, showBoundary=0),
                Frame(x_offset + margin_left, y_position-11*mm, 70*mm, 20*mm, showBoundary=0),
                Frame(x_offset + margin_left+70*mm, y_position-30, 13*mm, 13*mm, showBoundary=0),]

        styleN.alignment=TA_LEFT
        styleN.fontSize = 5*mm
        styleN.leading = 6*mm
        parts = pattern.findall(address_data["住所"][i])
        line1=""
        line2=""
        line3=""
        for p in parts:
            w, h = Paragraph(line1+line2+line3+p, styleN).wrap(81*mm, 30*mm)
            if h<=20:
                line1+=p
            elif h<=40:
                line2+=p
            else:
                line3+=p
        line=line1
        if line2!="":
            line = line + "\n" +line2
        if line3!="":
            line = line + "\n" +line3
        frames[0].addFromList([Paragraph(line, styleN)], c)

        styleN.fontSize = 7*mm
        styleN.leading = 7*mm
        styleN.alignment=TA_RIGHT
        w, h = Paragraph(address_data["氏名"][i], styleN).wrap(70*mm, 13*mm)
        if h <= 8*mm:
            frames[1].addFromList([Paragraph(address_data["氏名"][i], styleN)], c)
        else:
            frames[2].addFromList([Paragraph(address_data["氏名"][i], styleN)], c)

        frames[3].addFromList([Paragraph("様", styleN)], c)

    c.save()


raw_data = pd.read_csv(csv_path)
raw_data = raw_data.dropna(subset=["郵便番号"]).reset_index(drop=True)
raw_data["マンション・建物名・部屋番号"] = raw_data["マンション・建物名・部屋番号"].fillna("")

formatted_data = pd.DataFrame()
formatted_data["郵便番号"] = raw_data["郵便番号"].astype(int).astype(str).str.zfill(7).apply(lambda x: f"{x[:3]}-{x[3:]}")
formatted_data["住所"] = raw_data["都道府県"] + raw_data["市区町村・丁目・番地"] + raw_data["マンション・建物名・部屋番号"]

formatted_data["氏名"] = raw_data["氏名"]

create_address_labels("address_labels.pdf", formatted_data)