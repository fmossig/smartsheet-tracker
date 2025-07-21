import csv, os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

# Farb‑Mapping
COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8"
}

def read_snapshot(date_str):
    snap_file = f"status/status_snapshot_{date_str}.csv"
    data = list(csv.reader(open(snap_file, encoding="utf-8")))[1:]
    counts = {g: 0 for g in COLORS}
    for grp, *_ in data:
        if grp in counts:
            counts[grp] += 1
    return [counts[g] for g in COLORS]

def make_report():
    # Setup
    date_str = datetime.utcnow().strftime("%Y-%m-%d")
    pdf_file = f"status/status_report_{date_str}.pdf"

    # Daten
    groups = list(COLORS.keys())
    values = read_snapshot(date_str)

    # Dokument
    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=A4,
        leftMargin=20*mm,
        rightMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )
    styles = getSampleStyleSheet()
    # Sans‑Serif für besseren Look
    styles.add(ParagraphStyle(name='Title2', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=18))
    styles.add(ParagraphStyle(name='Heading', fontName='Helvetica-Bold', fontSize=14, spaceAfter=6))

    elems = []

    # (Optional) Logo
    # logo_path = "assets/logo.png"
    # if os.path.exists(logo_path):
    #     elems.append(Image(logo_path, width=40*mm, height=15*mm))
    #     elems.append(Spacer(1, 5*mm))

    # Titel
    elems.append(Paragraph("PRODUKTGRUPPEN ÜBERSICHT", styles['Title2']))
    elems.append(Spacer(1, 8*mm))

    # Untertitel / Chart‑Titel
    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe (letzte 30 Tage)", styles['Heading']))
    elems.append(Spacer(1, 4*mm))

    # Chart zeichnen
    drawing = Drawing(180*mm, 100*mm)
    chart = VerticalBarChart()
    chart.x = 15*mm
    chart.y = 15*mm
    chart.height = 80*mm
    chart.width = 150*mm
    chart.data = [values]
    chart.categoryAxis.categoryNames = groups
    chart.categoryAxis.labels.boxAnchor = 'n'
    chart.valueAxis.valueMin = 0
    max_val = max(values) if values else 1
    chart.valueAxis.valueMax = max_val * 1.1
    chart.valueAxis.valueStep = max(1, int(max_val/10) or 1)
    # Gitterlinien
    chart.valueAxis.gridStrokeColor = colors.lightgrey

    # Balkenfarben & keine Umrandung
    for i, grp in enumerate(groups):
        bar = chart.bars[i]
        bar.fillColor = colors.HexColor(COLORS[grp])
        bar.strokeColor = None

        # Daten‑Label
        label = String(
            chart.x + (i+0.5)*(chart.width/len(values)),
            chart.y + (values[i]/chart.valueAxis.valueMax)*chart.height + 4,
            str(values[i]),
            fontName='Helvetica', fontSize=9,
            textAnchor='middle'
        )
        drawing.add(label)

    drawing.add(chart)
    elems.append(drawing)

    # Fußzeile Datum
    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph(f"Erstellt am: {date_str}", styles['Normal']))

    # PDF bauen
    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
