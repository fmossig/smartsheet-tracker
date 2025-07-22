   # ---- Chart 2: Gestapelt NA ----
    counts = read_na_phase_employee(date_str)
    phases_sorted = [1, 2, 3, 4, 5]
    emp_sorted = [e for e in EMP_COLORS if any(counts[p][e] for p in phases_sorted)]

    elems.append(PageBreak())
    elems.append(Paragraph("Mitarbeiterbasierte Phasenstatistik (NA, 30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 6*mm))

    # Größen
    chart_w     = (A4[0] - 2*20*mm) * 0.70      # 70% Seitenbreite
    legend_w    = 45*mm
    total_w     = chart_w + legend_w
    left_axis   = 18*mm
    row_h       = 8*mm
    gap_y       = 4*mm
    origin_y2   = 12*mm

    # Höhe berechnen
    total_h = len(phases_sorted) * (row_h + gap_y) + origin_y2 + 6*mm
    max_h   = A4[1] - 2*20*mm
    if total_h > max_h:
        scale = max_h / total_h
        row_h     *= scale
        gap_y     *= scale
        origin_y2 *= scale
        total_h    = max_h

    d2 = Drawing(total_w, total_h)

    for i, phase in enumerate(phases_sorted):
        phase_total = sum(counts[phase][e] for e in emp_sorted)
        if phase_total == 0:
            continue
        y = origin_y2 + (len(phases_sorted)-1-i) * (row_h + gap_y)
        x = left_axis
        d2.add(String(2*mm, y + row_h/2, f"Phase {phase}",
                      fontName='Helvetica', fontSize=8, textAnchor='start'))
        run_w   = 0
        avail_w = chart_w - left_axis - 5*mm
        for emp in emp_sorted:
            v = counts[phase][emp]
            if v == 0:
                continue
            seg_w = max((v / phase_total) * avail_w, 3)  # min-breite
            rect = Rect(x + run_w, y, seg_w, row_h,
                        fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None)
            d2.add(rect)
            if seg_w > 14:
                d2.add(String(x + run_w + seg_w/2, y + row_h/2, str(v),
                              fontName='Helvetica-Bold', fontSize=7,
                              textAnchor='middle', fillColor=colors.white))
            run_w += seg_w

    # Legende
    legend_x = chart_w + 5*mm
    legend_y = total_h - 8*mm
    box_h    = 4.5*mm
    for j, emp in enumerate(emp_sorted):
        yy = legend_y - j * (box_h + 2*mm)
        d2.add(Rect(legend_x, yy, box_h, box_h,
                    fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None))
        d2.add(String(legend_x + box_h + 2*mm, yy + box_h/2,
                      emp, fontName='Helvetica', fontSize=7, textAnchor='start'))

    elems.append(d2)
    elems.append(Spacer(1, 6*mm))
