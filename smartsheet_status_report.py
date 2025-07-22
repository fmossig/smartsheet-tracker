import os, csv
from datetime import datetime, timedelta, timezone
from collections import defaultdict

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth


# Pie
from reportlab.graphics.charts.piecharts import Pie

# Smartsheet
import smartsheet
from dotenv import load_dotenv

# ---------- Farben ----------
GROUP_COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8",
}

PHASE_COLORS = {
    1: "#1F77B4",
    2: "#FF7F0E",
    3: "#2CA02C",
    4: "#9467BD",
    5: "#D62728",
}

EMP_COLORS = {
    "DM":  "#223459",
    "EK":  "#6A5AAA",
    "HI":  "#B45082",
    "SM":  "#F9767F",
    "JHU": "#FFB142",
    "LK":  "#FFDE70",
}

CODE_VERSION = "2025-07-22_multi-groups_with_pies"

# ---------- Spalten ----------
COL_ARTIKEL = "Artikel"
COL_LINK    = "Link"
COL_AMAZON  = "Amazon"
PHASE_DATE_COLS = ["Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"]

# ---------- Sheet-IDs ----------
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# ---------- Helpers ----------
def read_snapshot_counts(date_str: str):
    """Zählt alle Phase-Events pro Produktgruppe aus status_snapshot_<date>.csv."""
    snap = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = {g: 0 for g in GROUP_COLORS}
    with open(snap, encoding="utf-8") as f:
        r = csv.reader(f)
        next(r, None)  # Header
        for row in r:
            grp = row[0]
            if grp in counts:
                counts[grp] += 1
    # gleiche Reihenfolge wie GROUP_COLORS
    return [counts[g] for g in GROUP_COLORS]


def phase_distribution_for_group(date_str: str, group_code: str):
    """Liefert dict {1..5: count} für Phasen in der Produktgruppe (30 Tage Snapshot)."""
    snap = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = {p: 0 for p in range(1, 6)}
    with open(snap, encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            if row["Produktgruppe"] == group_code:
                p = int(row["Phase"])
                if p in counts:
                    counts[p] += 1
    return counts


def read_phase_employee_by_group(date_str: str, group_code: str):
    """
    dict: phase -> {emp: count} für die angegebene Produktgruppe,
    basierend auf der Snapshot-CSV.
    """
    snap = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = defaultdict(lambda: defaultdict(int))
    with open(snap, encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            if row["Produktgruppe"] != group_code:
                continue
            phase = int(row["Phase"])
            emp = (row["Mitarbeiter"] or "").strip()
            if emp in EMP_COLORS:
                counts[phase][emp] += 1
    return counts


def build_phase_pie(dist_dict: dict, diam_mm: float):
    """
    Erstellt ein Pie-Drawing mit weißen Rändern und PHASE_COLORS.
    Labels werden nicht im Pie gezeigt (Legende separat).
    """
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib.units import mm

    phases = [1, 2, 3, 4, 5]
    data   = [dist_dict.get(p, 0) for p in phases]

    d = Drawing(diam_mm * mm, diam_mm * mm)
    pie = Pie()
    pie.width = pie.height = diam_mm * mm
    pie.x = pie.y = 0

    pie.data    = data
    pie.labels  = []          # wir nutzen externe Legende
    pie.simpleLabels = 1      # keine Verbindungslinien

    # Weiße Ränder
    pie.slices.strokeColor = colors.white
    pie.slices.strokeWidth = 0.6

    # Farben setzen
    for i, p in enumerate(phases):
        pie.slices[i].fillColor = colors.HexColor(PHASE_COLORS[p])

    d.add(pie)
    return d


def calc_metrics_for_group(client, group_code: str, cutoff_date):
    """
    Berechnet KPIs für eine Produktgruppe direkt aus dem Smartsheet.
    Rückgabe: (artikel_count, mp_count, bearbeitet, pct)
    """
    sheet = client.Sheets.get_sheet(SHEET_IDS[group_code])
    col_map = {c.title: c.id for c in sheet.columns}

    artikel_count = 0
    mp_count = 0
    touched_rows = set()

    for row in sheet.rows:
        artikel_val = ""
        link_val    = ""
        amazon_val  = ""
        has_recent  = False

        # Basisfelder
        for cell in row.cells:
            if cell.column_id == col_map.get(COL_ARTIKEL):
                artikel_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_LINK):
                link_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_AMAZON):
                amazon_val = (cell.display_value or "").strip()

        # Phasen-Datumsprüfung
        for col_name in PHASE_DATE_COLS:
            cid = col_map.get(col_name)
            if not cid:
                continue
            cell = next((c for c in row.cells if c.column_id == cid), None)
            if cell and cell.value:
                try:
                    dt = datetime.fromisoformat(cell.value).date()
                    if dt >= cutoff_date:
                        has_recent = True
                except Exception:
                    pass

        # Zählen
        if artikel_val and not link_val:
            artikel_count += 1
        if artikel_val and link_val and amazon_val:
            mp_count += 1
        if has_recent:
            touched_rows.add(row.id)

    bearbeitet = len(touched_rows)
    pct = (bearbeitet / mp_count * 100) if mp_count else 0.0
    return artikel_count, mp_count, bearbeitet, pct


# ------- Misch-UI-Helper (Header / Legenden) -------
from reportlab.pdfbase.pdfmetrics import stringWidth

def build_group_header(grp_code: str, grp_color_hex: str, period_text="Zeitraum: 30 Tage"):
    """
    Titelblock:
      [roter Chip mit GRP]   [Daten & Kennzahlen / Zeitraum: 30 Tage]
    """
    chip_font   = 'Helvetica-Bold'
    chip_fs     = 18
    chip_pad_x  = 4*mm
    chip_pad_y  = 2*mm

    line1_font  = 'Helvetica-Bold'
    line1_fs    = 14
    line2_font  = 'Helvetica'
    line2_fs    = 9

    gap_between = 6*mm

    CHIP_TEXT_UP = 1*mm   # „NA“ minimal höher
    TEXT_LEFT    = 2*mm   # Textblock leicht nach links

    # Chip-Größe
    chip_text_w = stringWidth(grp_code, chip_font, chip_fs)
    chip_w = chip_text_w + 2*chip_pad_x
    chip_h = chip_fs * 1.2 + 2*chip_pad_y

    line1_w = stringWidth("Daten & Kennzahlen", line1_font, line1_fs)
    line2_w = stringWidth(period_text,          line2_font, line2_fs)
    text_block_w = max(line1_w, line2_w)

    total_w = chip_w + gap_between + text_block_w
    total_h = chip_h

    d = Drawing(total_w, total_h)

    # Chip
    d.add(Rect(0, 0, chip_w, chip_h,
               fillColor=colors.HexColor(grp_color_hex),
               strokeColor=None))
    d.add(String(chip_pad_x,
                 chip_pad_y + chip_fs*0.1 + CHIP_TEXT_UP,
                 grp_code,
                 fontName=chip_font, fontSize=chip_fs,
                 fillColor=colors.white, textAnchor='start'))

    # Textblock
    text_x = chip_w + gap_between - TEXT_LEFT
    center_y = total_h / 2.0
    line1_y = center_y + line1_fs * 0.35
    line2_y = center_y - line2_fs * 1.1

    d.add(String(text_x, line1_y,
                 "Daten & Kennzahlen",
                 fontName=line1_font, fontSize=line1_fs,
                 fillColor=colors.black, textAnchor='start'))
    d.add(String(text_x, line2_y,
                 period_text,
                 fontName=line2_font, fontSize=line2_fs,
                 fillColor=colors.black, textAnchor='start'))

    return d, total_h


def build_emp_legend_banner(
    width,
    emp_items,
    box_size=5*mm,
    font_sz=10,
    gap_item=14*mm,
    banner_color="#F2F2F2"
):
    """
    Graues Banner (volle Breite) mit zentrierter Mitarbeiter-Legende.
    Rückgabe: (Drawing, height_mm)
    """
    banner_h = 14*mm
    d = Drawing(width, banner_h)

    d.add(Rect(0, 0, width, banner_h,
               fillColor=colors.HexColor(banner_color),
               strokeColor=None))

    y_center = banner_h / 2.0
    text_y   = y_center - (font_sz * 0.35)

    item_w   = box_size + 2*mm + gap_item
    total_w  = len(emp_items) * item_w - gap_item
    x_cursor = (width - total_w) / 2.0

    for emp in emp_items:
        d.add(Rect(x_cursor,
                   y_center - box_size/2,
                   box_size, box_size,
                   fillColor=colors.HexColor(EMP_COLORS[emp]),
                   strokeColor=None))
        d.add(String(x_cursor + box_size + 2*mm,
                     text_y,
                     emp,
                     fontName='Helvetica',
                     fontSize=font_sz,
                     textAnchor='start'))
        x_cursor += item_w

    return d, banner_h

# ---------- Report ----------
def make_report():
    load_dotenv()
    token = os.getenv("SMARTSHEET_TOKEN")
    client = smartsheet.Smartsheet(token)

    now = datetime.now(timezone.utc)
    today = now.date()
    date_str = today.isoformat()
    cutoff = today - timedelta(days=30)

    pdf_file = os.path.join("status", f"status_report_{date_str}.pdf")

    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=20*mm, bottomMargin=20*mm
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CoverTitle', parent=styles['Title'],
                              fontName='Helvetica-Bold', fontSize=20, spaceAfter=12))
    styles.add(ParagraphStyle(name='CoverInfo', parent=styles['Normal'],
                              fontName='Helvetica', fontSize=12, spaceAfter=6))
    styles.add(ParagraphStyle(name='ChartTitle', parent=styles['Title'],
                              fontName='Helvetica-Bold', fontSize=14, spaceAfter=6))

    elems = []

    # Deckblatt
    elems.append(Paragraph("Amazon Content Management - Activity Report", styles['CoverTitle']))
    elems.append(Paragraph(f"Erstellungsdatum: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['CoverInfo']))
    elems.append(Paragraph(f"Abgrenzungsdatum: {cutoff.isoformat()}", styles['CoverInfo']))
    elems.append(Paragraph(f"CODE_VERSION: {CODE_VERSION}", styles['CoverInfo']))
    elems.append(PageBreak())

    # ----- Seite: Produktgruppen Balken + Pies -----
    groups = list(GROUP_COLORS.keys())
    values = read_snapshot_counts(date_str)

    elems.append(Paragraph("Produktgruppen Daten (30 Tage)", styles['CoverTitle']))
    elems.append(Spacer(1, 4*mm))
    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe", styles['ChartTitle']))
    elems.append(Spacer(1, 4*mm))

    usable_width = A4[0] - doc.leftMargin - doc.rightMargin

    # >>> Bars dichter an den Seitenrand setzen
    chart_height = 55*mm
    origin_y     = 8*mm          # statt 25mm
    total_gap    = usable_width * 0.10
    gap          = total_gap / (len(groups) + 1)
    bar_width    = (usable_width - total_gap) / len(groups)

    d1       = Drawing(usable_width, chart_height + origin_y)
    max_val  = max(values) if values else 1
    bar_x_ct = []

    for i, grp in enumerate(groups):
        val = values[i]
        x   = gap * (i + 1) + bar_width * i
        h   = (val / max_val) * chart_height

        # Balken
        d1.add(Rect(x, origin_y, bar_width, h,
                    fillColor=colors.HexColor(GROUP_COLORS[grp]),
                    strokeColor=None))
        # Wert
        d1.add(String(x + bar_width/2, origin_y + h + 4,
                      str(val), fontName='Helvetica', fontSize=9,
                      textAnchor='middle'))
        # X‑Label
        label_y = origin_y - 8    # dichter an die Achse
        d1.add(String(x + bar_width/2, label_y,
                      grp, fontName='Helvetica', fontSize=8,
                      textAnchor='middle'))

        bar_x_ct.append(x + bar_width/2)

    elems.append(d1)

    # --- Pies direkt darunter auf grauem Banner ---
    PIE_BANNER_COLOR = colors.HexColor("#F2F2F2")
    PIE_BANNER_H_MM  = 24              # noch schmaler
    pie_diam_mm      = 18

    banner_h = PIE_BANNER_H_MM * mm
    pies_banner = Drawing(usable_width, banner_h)
    pies_banner.add(Rect(0, 0, usable_width, banner_h,
                         fillColor=PIE_BANNER_COLOR, strokeColor=None))

    pie_w = pie_diam_mm * mm
    pie_h = pie_diam_mm * mm
    pie_y = (banner_h - pie_h) / 2.0   # vertikal mittig im Banner

    for idx, grp in enumerate(groups):
        dist     = phase_distribution_for_group(date_str, grp)
        pie_draw = build_phase_pie(dist, pie_diam_mm)  # enthält weißen Stroke
        pie_x    = bar_x_ct[idx] - pie_w/2
        pie_draw.translate(pie_x, pie_y)
        pies_banner.add(pie_draw)

    # KEIN Spacer mehr dazwischen
    elems.append(pies_banner)

    # Pie-Legende
    legend_w = usable_width
    legend_h = 10*mm
    leg      = Drawing(legend_w, legend_h)

    box       = 5*mm
    font_sz   = 8
    gap_item  = 16*mm
    items     = [1, 2, 3, 4, 5]
    item_w    = box + 2*mm + gap_item
    total_w   = len(items)*item_w - gap_item
    x_cursor  = (legend_w - total_w) / 2.0
    y_center  = legend_h / 2.0

    for p in items:
        leg.add(Rect(x_cursor, y_center - box/2, box, box,
                     fillColor=colors.HexColor(PHASE_COLORS[p]),
                     strokeColor=colors.white, strokeWidth=0.4))
        leg.add(String(x_cursor + box + 2*mm,
                       y_center - (font_sz*0.35),
                       f"Phase {p}",
                       fontName='Helvetica', fontSize=font_sz,
                       textAnchor='start'))
        x_cursor += item_w

    elems.append(Spacer(1, 2*mm))
    elems.append(leg)


    # ---- pro Gruppe eigenes Board ----
    phases_sorted = [1, 2, 3, 4, 5]
    legend_items = list(EMP_COLORS.keys())

    banner_h2 = 12*mm
    banner_w2 = usable_width
    legend_h2 = 10*mm
    box_size = 5*mm
    font_sz2  = 10
    gap_item2 = 14*mm

    shrink    = 0.75
    chart_w   = usable_width * 0.45
    left_ax   = 15*mm
    row_h     = 8*mm * shrink
    gap_y     = 4*mm * shrink
    origin_y2 = 10*mm

    for grp in groups:
        elems.append(PageBreak())

        # Banner
        header_draw, header_h = build_group_header(grp, GROUP_COLORS[grp])
        elems.append(header_draw)
        elems.append(Spacer(1, 4*mm))


        # --- Mitarbeiter-Legende im grauen Banner (volle Breite) ---
        leg2_draw, _ = build_emp_legend_banner(
            width=usable_width,
            emp_items=legend_items,
            box_size=5*mm,
            font_sz=10,
            gap_item=14*mm,
            banner_color="#F2F2F2"
        )
        elems.append(leg2_draw)
        elems.append(Spacer(1, 1*mm))  # Abstand zum Chart minimal halten


        # Gestapeltes Chart
        counts = read_phase_employee_by_group(date_str, grp)

        rows_drawn = len(phases_sorted)
        total_h = rows_drawn * (row_h + gap_y) + origin_y2 + 6*mm
        max_h   = A4[1] - doc.topMargin - doc.bottomMargin - 40*mm
        if total_h > max_h:
            scale = max_h / total_h
            row_h     *= scale
            gap_y     *= scale
            origin_y2 *= scale
            total_h    = max_h

        d2 = Drawing(chart_w, total_h)

        global_max = 0
        for p in phases_sorted:
            for emp in legend_items:
                global_max = max(global_max, counts[p][emp])

        avail_w     = chart_w - left_ax - 5*mm
        px_per_unit = avail_w / (global_max if global_max else 1)

        y_index = 0
        for phase in phases_sorted:
            y = origin_y2 + (rows_drawn - 1 - y_index) * (row_h + gap_y)
            y_index += 1
            x = left_ax

            d2.add(String(2*mm, y + row_h/2, f"Phase {phase}",
                          fontName='Helvetica', fontSize=8, textAnchor='start'))

            run_w = 0
            phase_total = sum(counts[phase][e] for e in legend_items)
            if phase_total == 0:
                d2.add(Rect(x, y, 1, row_h, fillColor=None, strokeColor=colors.grey))
                continue

            for emp in legend_items:
                v = counts[phase][emp]
                if v == 0:
                    continue
                seg_w = v * px_per_unit
                rect = Rect(x + run_w, y, seg_w, row_h,
                            fillColor=colors.HexColor(EMP_COLORS[emp]),
                            strokeColor=None)
                d2.add(rect)
                if seg_w > 14:
                    d2.add(String(x + run_w + seg_w/2, y + row_h/2, str(v),
                                  fontName='Helvetica-Bold', fontSize=9,
                                  textAnchor='middle', fillColor=colors.white))
                run_w += seg_w

        elems.append(d2)

        # KPI-Boxen
        artikel, mp, bearbeitet, pct = calc_metrics_for_group(client, grp, cutoff)
        elems.append(Spacer(1, 3*mm))

        red        = colors.HexColor(GROUP_COLORS[grp])
        box_h      = 20*mm
        spacing    = 10*mm
        font_val   = 14
        font_lab   = 9

        kpis = [
            ("Anzahl Artikel",       artikel),
            ("Marktplatzartikel",    mp),
            ("Bearbeitete Artikel",  bearbeitet),
            ("% bearbeitet",         f"{pct:.1f}%"),
        ]

        usable_full = usable_width
        n = len(kpis)
        total_spacing = spacing * (n - 1)
        box_w = (usable_full * 0.88 - total_spacing) / n
        if box_w < 30*mm:
            box_w = 30*mm
        total_w = n * box_w + total_spacing
        start_x = (usable_full - total_w) / 2.0

        kpi_draw = Drawing(usable_full, box_h)
        for i, (label, value) in enumerate(kpis):
            x = start_x + i * (box_w + spacing)
            kpi_draw.add(Rect(x, 0, box_w, box_h,
                              fillColor=red, strokeColor=None))
            kpi_draw.add(String(x + box_w/2, box_h * 0.62,
                                str(value),
                                fontName='Helvetica-Bold', fontSize=font_val,
                                textAnchor='middle', fillColor=colors.white))
            kpi_draw.add(String(x + box_w/2, box_h * 0.28,
                                label,
                                fontName='Helvetica-Bold', fontSize=font_lab,
                                textAnchor='middle', fillColor=colors.white))

        elems.append(kpi_draw)
        elems.append(Spacer(1, 6*mm))

    # Footer
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
