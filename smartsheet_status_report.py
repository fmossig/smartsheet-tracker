import os, csv
from datetime import datetime, timedelta, timezone
from collections import defaultdict

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

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
def read_snapshot_counts(date_str):
    snap = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = {g: 0 for g in GROUP_COLORS}
    with open(snap, encoding="utf-8") as f:
        r = csv.reader(f)
        next(r)
        for row in r:
            grp = row[0]
            if grp in counts:
                counts[grp] += 1
    return [counts[g] for g in GROUP_COLORS]

def phase_distribution_for_group(date_str, group_code):
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

def read_phase_employee_by_group(date_str, group_code):
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

def build_phase_pie(dist_dict, diam_mm):
    phases = [1,2,3,4,5]
    data   = [dist_dict.get(p, 0) for p in phases]

    d = Drawing(diam_mm*mm, diam_mm*mm)
    pie = Pie()
    pie.width  = pie.height = diam_mm*mm
    pie.x = pie.y = 0
    pie.data  = data
    pie.sideLabels  = 0              # keine seitlichen Labels
    pie.simpleLabels = 1             # Standard-Placement
    pie.labelRadius = pie.width/2 - 1*mm   # etwas nach innen
    pie.slices[i].popout = 0         # sicherheitshalber
    pie.slices[i].labelRadius = pie.labelRadius

    pie.labels = []

    # Farben & weiße Ränder
    for i, p in enumerate(phases):
        pie.slices[i].fillColor   = colors.HexColor(PHASE_COLORS[p])
        pie.slices[i].strokeColor = colors.white
        pie.slices[i].strokeWidth = 0.6

    d.add(pie)
    return d

def calc_metrics_for_group(client, group_code, cutoff_date):
    sheet = client.Sheets.get_sheet(SHEET_IDS[group_code])
    col_map = {c.title: c.id for c in sheet.columns}

    artikel_count = 0
    mp_count = 0
    touched_rows = set()

    for row in sheet.rows:
        artikel_val = link_val = amazon_val = ""
        has_recent = False

        for cell in row.cells:
            if cell.column_id == col_map.get(COL_ARTIKEL):
                artikel_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_LINK):
                link_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_AMAZON):
                amazon_val = (cell.display_value or "").strip()

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

        if artikel_val and not link_val:
            artikel_count += 1
        if artikel_val and link_val and amazon_val:
            mp_count += 1
        if has_recent:
            touched_rows.add(row.id)

    bearbeitet = len(touched_rows)
    pct = (bearbeitet / mp_count * 100) if mp_count else 0.0
    return artikel_count, mp_count, bearbeitet, pct

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
    styles.add(ParagraphStyle(name='ChartTitle', parent=styles['Heading2'],
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
    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe", styles['ChartTitle']))

    usable_width = A4[0] - doc.leftMargin - doc.rightMargin
    chart_height = 60*mm
    origin_y = 25*mm
    total_gap = usable_width * 0.1
    gap = total_gap / (len(groups) + 1)
    bar_width = (usable_width - total_gap) / len(groups)

    d1 = Drawing(usable_width, chart_height + origin_y + 2*mm)
    max_val = max(values) if values else 1
    bar_x_centers = []

    for i, grp in enumerate(groups):
        val = values[i]
        x = gap * (i + 1) + bar_width * i
        h = (val / max_val) * chart_height

        d1.add(Rect(x, origin_y, bar_width, h,
                    fillColor=colors.HexColor(GROUP_COLORS[grp]), strokeColor=None))
        d1.add(String(x + bar_width/2, origin_y + h + 4, str(val),
                      fontName='Helvetica', fontSize=9, textAnchor='middle'))
        d1.add(String(x + bar_width/2, origin_y - 10, grp,
                      fontName='Helvetica', fontSize=8, textAnchor='middle'))

        bar_x_centers.append(x + bar_width/2)

    elems.append(d1)

    # --- Pies auf grauem Banner ---
    elems.append(Spacer(1, 2*mm))  # kleiner Abstand

    PIE_BANNER_COLOR = colors.HexColor("#F2F2F2")
    PIE_BANNER_H_MM  = 32
    pie_diam_mm = 20

    banner_h = PIE_BANNER_H_MM * mm
    pies_banner = Drawing(usable_width, banner_h)
    pies_banner.add(Rect(0, 0, usable_width, banner_h,
                         fillColor=PIE_BANNER_COLOR, strokeColor=None))

    pie_w = pie_diam_mm*mm
    pie_h = pie_diam_mm*mm
    pie_y = (banner_h - pie_h) / 2.0  # vertikal mittig

    for idx, grp in enumerate(groups):
        dist = phase_distribution_for_group(date_str, grp)
        pie_draw = build_phase_pie(dist, pie_diam_mm)
        pie_x = bar_x_centers[idx] - pie_w/2
        pie_draw.translate(pie_x, pie_y)
        pies_banner.add(pie_draw)

    elems.append(pies_banner)

    # Pie-Legende
    elems.append(Spacer(1, 4*mm))
    legend_w = usable_width
    legend_h = 12*mm
    leg = Drawing(legend_w, legend_h)
    box = 5*mm
    font_sz = 8
    gap_item = 18*mm
    items = [1,2,3,4,5]
    item_w = box + 2*mm + gap_item
    total_w = len(items)*item_w - gap_item
    x_cursor = (legend_w - total_w)/2
    y_center = legend_h/2
    for p in items:
        leg.add(Rect(x_cursor, y_center - box/2, box, box,
                     fillColor=colors.HexColor(PHASE_COLORS[p]), strokeColor=None))
        leg.add(String(x_cursor + box + 2*mm, y_center - 2,
                       f"Phase {p}", fontName='Helvetica', fontSize=font_sz,
                       textAnchor='start'))
        x_cursor += item_w
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
        banner = Drawing(banner_w2, banner_h2)
        banner.add(Rect(0, 0, banner_w2, banner_h2,
                        fillColor=colors.HexColor(GROUP_COLORS[grp]), strokeColor=None))
        banner.add(String(banner_w2/2, banner_h2/2 - 2*mm,
                          f"Produktgruppe ({grp}, Daten für 30 Tage)",
                          fontName='Helvetica-Bold', fontSize=18,
                          textAnchor='middle', fillColor=colors.white))
        elems.append(banner)
        elems.append(Spacer(1, 4*mm))

        # Emp-Legende
        leg2 = Drawing(banner_w2, legend_h2)
        y_center = legend_h2 / 2
        text_y   = y_center - (font_sz2 * 0.35)
        item_w   = box_size + 2*mm + gap_item2
        total_w  = len(legend_items)*item_w - gap_item2
        x_cursor = (banner_w2 - total_w) / 2.0

        for emp in legend_items:
            leg2.add(Rect(x_cursor,
                          y_center - box_size/2,
                          box_size, box_size,
                          fillColor=colors.HexColor(EMP_COLORS[emp]),
                          strokeColor=None))
            leg2.add(String(x_cursor + box_size + 2*mm,
                            text_y,
                            emp,
                            fontName='Helvetica',
                            fontSize=font_sz2,
                            textAnchor='start'))
            x_cursor += item_w

        elems.append(leg2)
        elems.append(Spacer(1, 2*mm))

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
