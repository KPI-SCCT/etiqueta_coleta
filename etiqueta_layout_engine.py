from __future__ import annotations

from typing import Any


Box = tuple[float, float, float, float]


def clamp(value: float, minimum: float, maximum: float) -> float:
    return max(minimum, min(value, maximum))


def inset_box(box: Box, pad_x: float, pad_y: float) -> Box:
    x, y, w, h = box
    inner_w = max(0.0, w - (2 * pad_x))
    inner_h = max(0.0, h - (2 * pad_y))
    return x + pad_x, y + pad_y, inner_w, inner_h


def split_box_vertical(box: Box, heights: list[float], gap: float = 0.0, from_top: bool = True) -> list[Box]:
    x, y, w, h = box
    if not heights:
        return []
    gap = max(0.0, gap)
    total_gaps = gap * max(0, len(heights) - 1)
    heights_sum = sum(max(0.0, value) for value in heights)
    if heights_sum <= 0:
        return [(x, y, w, 0.0) for _ in heights]

    if heights_sum + total_gaps > h:
        available = max(0.0, h - total_gaps)
        factor = available / heights_sum if heights_sum > 0 else 0.0
        heights = [max(0.0, value) * factor for value in heights]

    boxes: list[Box] = []
    if from_top:
        cursor = y + h
        for item_h in heights:
            box_h = max(0.0, item_h)
            cursor -= box_h
            boxes.append((x, cursor, w, box_h))
            cursor -= gap
    else:
        cursor = y
        for item_h in heights:
            box_h = max(0.0, item_h)
            boxes.append((x, cursor, w, box_h))
            cursor += box_h + gap
    return boxes


def split_box_horizontal(box: Box, widths: list[float], gap: float = 0.0, from_left: bool = True) -> list[Box]:
    x, y, w, h = box
    if not widths:
        return []
    gap = max(0.0, gap)
    total_gaps = gap * max(0, len(widths) - 1)
    widths_sum = sum(max(0.0, value) for value in widths)
    if widths_sum <= 0:
        return [(x, y, 0.0, h) for _ in widths]

    if widths_sum + total_gaps > w:
        available = max(0.0, w - total_gaps)
        factor = available / widths_sum if widths_sum > 0 else 0.0
        widths = [max(0.0, value) * factor for value in widths]

    boxes: list[Box] = []
    if from_left:
        cursor = x
        for item_w in widths:
            box_w = max(0.0, item_w)
            boxes.append((cursor, y, box_w, h))
            cursor += box_w + gap
    else:
        cursor = x + w
        for item_w in widths:
            box_w = max(0.0, item_w)
            cursor -= box_w
            boxes.append((cursor, y, box_w, h))
            cursor -= gap
    return boxes


def split_rows(box: Box, count: int, gap: float = 0.0, from_top: bool = True) -> list[Box]:
    if count <= 0:
        return []
    x, y, w, h = box
    gap = max(0.0, gap)
    total_gaps = gap * max(0, count - 1)
    row_h = max(0.0, (h - total_gaps) / count)
    return split_box_vertical((x, y, w, h), [row_h] * count, gap=gap, from_top=from_top)


def _safe_text(value: Any) -> str:
    return "" if value is None else str(value)


def _split_word_to_fit(c: Any, word: str, font_name: str, font_size: float, max_width: float) -> list[str]:
    if max_width <= 0:
        return [word]
    if c.stringWidth(word, font_name, font_size) <= max_width:
        return [word]
    chunks: list[str] = []
    current = ""
    for char in word:
        candidate = f"{current}{char}"
        if current and c.stringWidth(candidate, font_name, font_size) > max_width:
            chunks.append(current)
            current = char
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks or [word]


def _truncate_with_ellipsis(
    c: Any,
    text: str,
    font_name: str,
    font_size: float,
    max_width: float,
    suffix: str = "...",
) -> str:
    if max_width <= 0:
        return suffix
    if c.stringWidth(text, font_name, font_size) <= max_width:
        return text
    base = text.rstrip()
    while base and c.stringWidth(f"{base}{suffix}", font_name, font_size) > max_width:
        base = base[:-1]
    return f"{base}{suffix}" if base else suffix


def _wrap_text_lines(c: Any, text: str, font_name: str, font_size: float, max_width: float) -> list[str]:
    raw = _safe_text(text).replace("\r", "\n")
    if not raw:
        return [""]

    lines: list[str] = []
    for paragraph in raw.split("\n"):
        words = paragraph.split()
        if not words:
            lines.append("")
            continue

        current = ""
        for word in words:
            parts = _split_word_to_fit(c, word, font_name, font_size, max_width)
            for part in parts:
                candidate = part if not current else f"{current} {part}"
                if not current or c.stringWidth(candidate, font_name, font_size) <= max_width:
                    current = candidate
                else:
                    lines.append(current)
                    current = part
        if current:
            lines.append(current)

    return lines or [""]


def _limit_lines(
    c: Any,
    lines: list[str],
    font_name: str,
    font_size: float,
    max_width: float,
    max_lines: int | None,
) -> tuple[list[str], bool]:
    if max_lines is None or max_lines <= 0:
        return lines, False
    if len(lines) <= max_lines:
        return lines, False

    clipped = lines[:max_lines]
    clipped[-1] = _truncate_with_ellipsis(c, clipped[-1], font_name, font_size, max_width)
    return clipped, True


def fit_text_to_box(
    c: Any,
    text: str,
    font_name: str,
    box_w: float,
    box_h: float,
    min_font: float,
    max_font: float,
    max_lines: int | None,
    line_spacing: float = 1.15,
) -> tuple[float, list[str]]:
    if box_w <= 0 or box_h <= 0:
        return max(1.0, min_font), [""]

    min_size = max(1.0, min_font)
    max_size = max(min_size, max_font)

    best_size = min_size
    best_lines = [""]

    low = min_size
    high = max_size
    for _ in range(14):
        mid = (low + high) / 2.0
        wrapped = _wrap_text_lines(c, text, font_name, mid, box_w)
        lines, _ = _limit_lines(c, wrapped, font_name, mid, box_w, max_lines)

        line_h = mid * line_spacing
        text_h = len(lines) * line_h
        widest = max(c.stringWidth(line, font_name, mid) for line in lines) if lines else 0.0
        fits = (text_h <= box_h + 0.2) and (widest <= box_w + 0.2)

        if fits:
            best_size = mid
            best_lines = lines
            low = mid
        else:
            high = mid

        if abs(high - low) < 0.08:
            break

    wrapped = _wrap_text_lines(c, text, font_name, best_size, box_w)
    best_lines, _ = _limit_lines(c, wrapped, font_name, best_size, box_w, max_lines)
    return best_size, best_lines


def draw_text_box(
    c: Any,
    box: Box,
    text: str,
    font_name: str,
    max_font: float,
    min_font: float = 5.0,
    max_lines: int | None = 1,
    line_spacing: float = 1.15,
    align: str = "left",
    valign: str = "center",
    pad_x: float = 0.0,
    pad_y: float = 0.0,
) -> float:
    x, y, w, h = inset_box(box, pad_x, pad_y)
    if w <= 0 or h <= 0:
        return min_font

    font_size, lines = fit_text_to_box(
        c,
        _safe_text(text),
        font_name,
        w,
        h,
        min_font=min_font,
        max_font=max_font,
        max_lines=max_lines,
        line_spacing=line_spacing,
    )

    line_h = font_size * line_spacing
    text_h = len(lines) * line_h
    if valign == "top":
        y_cursor = y + h - line_h
    elif valign == "bottom":
        y_cursor = y + max(0.0, text_h - line_h)
    else:
        y_cursor = y + ((h + text_h) / 2.0) - line_h

    c.setFont(font_name, font_size)
    for line in lines:
        if align == "center":
            c.drawCentredString(x + (w / 2.0), y_cursor, line)
        elif align == "right":
            c.drawRightString(x + w, y_cursor, line)
        else:
            c.drawString(x, y_cursor, line)
        y_cursor -= line_h
    return font_size


def _draw_barcode_area(
    c: Any,
    barcode_module: Any,
    box: Box,
    code_value: str,
    id_text: str,
    scale_ref: float,
    scale_content: float,
) -> None:
    # Mantem uma zona de silencio lateral, mas deixa o codigo usar mais largura util.
    x, y, w, h = inset_box(box, max(5.0, box[2] * 0.04), max(1.4, box[3] * 0.05))
    if w <= 0 or h <= 0:
        return

    code_h = max(6.5, h * 0.16)
    id_h = max(6.5, h * 0.16) if id_text else 0.0
    gap = max(1.4, h * 0.03)
    bar_h = h - code_h - id_h - (gap * (2 if id_text else 1))

    min_bar_h = max(10.0, h * 0.38)
    if bar_h < min_bar_h:
        deficit = min_bar_h - bar_h
        if id_h > 6.0:
            cut_id = min(deficit * 0.6, id_h - 6.0)
            id_h -= max(0.0, cut_id)
            deficit -= cut_id
        if deficit > 0 and code_h > 6.0:
            cut_code = min(deficit, code_h - 6.0)
            code_h -= max(0.0, cut_code)
        bar_h = h - code_h - id_h - (gap * (2 if id_text else 1))
        bar_h = max(6.0, bar_h)

    code_box = (x, y, w, code_h)
    bar_box = (x, y + code_h + gap, w, bar_h)
    id_box = (x, y + code_h + gap + bar_h + gap, w, id_h) if id_text else None

    code = _safe_text(code_value)
    if barcode_module is not None and code:
        modules = max(80, (11 * len(code)) + 35)
        target_w = bar_box[2] * 0.86
        bar_w = clamp(target_w / modules, 0.12, 1.15)
        draw_h = max(4.0, bar_box[3] * 0.88)
        barcode = barcode_module.Code128(code, barHeight=draw_h, barWidth=bar_w)

        for _ in range(18):
            if barcode.width > bar_box[2] * 0.90 and bar_w > 0.11:
                bar_w *= 0.94
                barcode = barcode_module.Code128(code, barHeight=draw_h, barWidth=bar_w)
                continue
            if barcode.width < bar_box[2] * 0.82 and bar_w < 1.2:
                bar_w *= 1.04
                barcode = barcode_module.Code128(code, barHeight=draw_h, barWidth=bar_w)
                continue
            break

        bar_x = bar_box[0] + ((bar_box[2] - barcode.width) / 2.0)
        bar_y = bar_box[1] + ((bar_box[3] - draw_h) / 2.0)
        barcode.drawOn(c, bar_x, bar_y)

    max_id_font = clamp(7.0 * scale_ref * scale_content, 5.2, 18.0)
    max_code_font = clamp(7.8 * scale_ref * scale_content, 5.4, 20.0)
    if id_box:
        draw_text_box(
            c,
            id_box,
            id_text,
            "Helvetica",
            max_font=max_id_font,
            min_font=5.0,
            max_lines=1,
            line_spacing=1.05,
            align="center",
            valign="center",
        )
    draw_text_box(
        c,
        code_box,
        code,
        "Helvetica",
        max_font=max_code_font,
        min_font=5.2,
        max_lines=1,
        line_spacing=1.05,
        align="center",
        valign="center",
    )


def draw_template_padrao(
    c: Any,
    barcode_module: Any,
    x: float,
    y: float,
    largura_pt: float,
    altura_pt: float,
    dados: dict,
    mm_to_points: float,
    app_name: str,
    espacamento_extra: float,
    escala_fonte_usuario: float,
    ajuste_cabecalho: float = 0.0,
) -> None:
    ref_w = 105 * mm_to_points
    ref_h = 148.5 * mm_to_points
    scale_ref = min(largura_pt / ref_w, altura_pt / ref_h)
    scale_content = clamp(escala_fonte_usuario, 0.6, 4.0)

    border = clamp(0.85 * scale_ref, 0.6, 1.4)
    c.setLineWidth(border)
    c.rect(x, y, largura_pt, altura_pt)

    pad = max(3.2 * mm_to_points, 5.4 * mm_to_points * scale_ref)
    content = inset_box((x, y, largura_pt, altura_pt), pad, pad)
    cx, cy, cw, ch = content
    if cw <= 0 or ch <= 0:
        return

    section_gap = max(1.8, 0.9 * mm_to_points * scale_ref) + (max(0.0, espacamento_extra) * 0.32)
    section_gap = min(section_gap, ch * 0.08)
    usable_h = max(6.0, ch - (2 * section_gap))

    header_ratio = clamp(0.15 + (ajuste_cabecalho / max(ch, 1.0)), 0.10, 0.25)
    barcode_ratio = 0.34
    details_ratio = 1.0 - header_ratio - barcode_ratio
    if details_ratio < 0.32:
        details_ratio = 0.32
        barcode_ratio = max(0.20, 1.0 - header_ratio - details_ratio)

    section_heights = [
        usable_h * header_ratio,
        usable_h * details_ratio,
        usable_h - (usable_h * header_ratio) - (usable_h * details_ratio),
    ]
    header_box, details_box, barcode_box = split_box_vertical(
        (cx, cy, cw, ch),
        section_heights,
        gap=section_gap,
        from_top=True,
    )

    title_scale = 1.0 + ((scale_content - 1.0) * 0.30)
    draw_text_box(
        c,
        inset_box(header_box, max(1.2, cw * 0.02), max(1.0, header_box[3] * 0.12)),
        app_name,
        "Helvetica-Bold",
        max_font=clamp(21.0 * scale_ref * title_scale, 10.0, 56.0),
        min_font=8.0,
        max_lines=1,
        line_spacing=1.0,
        align="center",
        valign="center",
    )
    c.setLineWidth(max(0.35, border * 0.7))
    c.line(cx, header_box[1], cx + cw, header_box[1])

    fields = [
        ("ORIGEM", dados.get("origem", "")),
        ("DESTINO", dados.get("destino", "")),
        ("ROMANEIO", dados.get("romaneio", "")),
        ("PROJETO", dados.get("projeto", "")),
        ("NR NF", dados.get("nr_nf", "")),
        ("VOLUME", dados.get("volume", "")),
    ]

    row_gap = max(0.9, 0.45 * mm_to_points * scale_ref) + (max(0.0, espacamento_extra) * 0.42)
    rows = split_rows(details_box, len(fields), gap=row_gap, from_top=True)
    col_gap = max(2.2, 1.1 * mm_to_points * scale_ref)
    label_w = clamp(details_box[2] * 0.34, 40.0, details_box[2] * 0.46)
    max_label_font = clamp(10.0 * scale_ref * scale_content, 6.8, 30.0)
    max_value_font = clamp(10.4 * scale_ref * scale_content, 6.8, 32.0)

    multi_line_fields = {"ORIGEM", "DESTINO", "PROJETO"}
    for row_box, (label, value) in zip(rows, fields):
        rx, ry, rw, rh = row_box
        value_w = max(0.0, rw - label_w - col_gap)
        label_box = (rx, ry, label_w, rh)
        value_box = (rx + label_w + col_gap, ry, value_w, rh)

        draw_text_box(
            c,
            label_box,
            f"{label}:",
            "Helvetica-Bold",
            max_font=max_label_font,
            min_font=5.3,
            max_lines=1,
            line_spacing=1.05,
            align="left",
            valign="center",
        )
        draw_text_box(
            c,
            value_box,
            _safe_text(value),
            "Helvetica",
            max_font=max_value_font,
            min_font=5.3,
            max_lines=2 if label in multi_line_fields else 1,
            line_spacing=1.08,
            align="left",
            valign="center",
        )

    _draw_barcode_area(
        c,
        barcode_module,
        barcode_box,
        code_value=_safe_text(dados.get("codigo_barras", "")),
        id_text=_safe_text(dados.get("id_fedex_data", "")),
        scale_ref=scale_ref,
        scale_content=scale_content,
    )


def draw_template_rede(
    c: Any,
    barcode_module: Any,
    x: float,
    y: float,
    largura_pt: float,
    altura_pt: float,
    dados: dict,
    mm_to_points: float,
    espacamento_extra: float,
    escala_fonte_usuario: float,
    ajuste_cabecalho: float = 0.0,
    ajuste_rodape: float = 0.0,
) -> None:
    ref_w = 90 * mm_to_points
    ref_h = 100 * mm_to_points
    scale_ref = min(largura_pt / ref_w, altura_pt / ref_h)
    scale_content = clamp(escala_fonte_usuario, 0.6, 4.0)

    border = clamp(0.9 * scale_ref, 0.6, 1.4)
    c.setLineWidth(border)
    c.rect(x, y, largura_pt, altura_pt)

    pad = max(3.0 * mm_to_points, 5.0 * mm_to_points * scale_ref)
    content = inset_box((x, y, largura_pt, altura_pt), pad, pad)
    cx, cy, cw, ch = content
    if cw <= 0 or ch <= 0:
        return

    section_gap = max(1.6, 0.85 * mm_to_points * scale_ref) + (max(0.0, espacamento_extra) * 0.3)
    section_gap = min(section_gap, ch * 0.06)
    usable_h = max(8.0, ch - (3 * section_gap))

    header_ratio = clamp(0.14 + (ajuste_cabecalho / max(ch, 1.0)), 0.10, 0.24)
    footer_ratio = clamp(0.10 + (ajuste_rodape / max(ch, 1.0)), 0.07, 0.18)
    barcode_ratio = 0.29
    details_ratio = 1.0 - header_ratio - footer_ratio - barcode_ratio
    if details_ratio < 0.28:
        details_ratio = 0.28
        barcode_ratio = max(0.18, 1.0 - header_ratio - footer_ratio - details_ratio)

    section_heights = [
        usable_h * header_ratio,
        usable_h * details_ratio,
        usable_h * footer_ratio,
        usable_h - (usable_h * header_ratio) - (usable_h * details_ratio) - (usable_h * footer_ratio),
    ]
    header_box, details_box, footer_box, barcode_box = split_box_vertical(
        (cx, cy, cw, ch),
        section_heights,
        gap=section_gap,
        from_top=True,
    )

    header_gap_x = max(2.0, 1.0 * mm_to_points * scale_ref)
    header_left_w = cw * 0.62
    header_left, header_right = split_box_horizontal(
        header_box,
        [header_left_w, max(0.0, cw - header_left_w - header_gap_x)],
        gap=header_gap_x,
        from_left=True,
    )
    draw_text_box(
        c,
        header_left,
        "OPERACAO REVERSA",
        "Helvetica-Bold",
        max_font=clamp(12.0 * scale_ref * (1.0 + ((scale_content - 1.0) * 0.25)), 7.0, 32.0),
        min_font=6.0,
        max_lines=1,
        line_spacing=1.0,
        align="left",
        valign="center",
    )
    draw_text_box(
        c,
        header_right,
        f"Tecnologia: {_safe_text(dados.get('tecnologia', ''))}",
        "Helvetica-Bold",
        max_font=clamp(9.2 * scale_ref * scale_content, 6.0, 24.0),
        min_font=5.6,
        max_lines=1,
        line_spacing=1.0,
        align="right",
        valign="center",
    )
    c.setLineWidth(max(0.35, border * 0.7))
    c.line(cx, header_box[1], cx + cw, header_box[1])

    fields = [
        ("Origem", dados.get("origem", "")),
        ("Destino", dados.get("destino", "")),
        ("Numero CRED", dados.get("numero_cred", dados.get("numero_cred_label", ""))),
        ("Nota Fiscal", dados.get("nota_fiscal", "")),
        ("Data Emissao", dados.get("data_emissao", "")),
        ("OS", dados.get("os", "")),
        ("Volume", dados.get("volume", "")),
    ]
    row_gap = max(0.9, 0.4 * mm_to_points * scale_ref) + (max(0.0, espacamento_extra) * 0.42)
    rows = split_rows(details_box, len(fields), gap=row_gap, from_top=True)
    col_gap = max(2.0, 1.0 * mm_to_points * scale_ref)
    label_w = clamp(details_box[2] * 0.36, 45.0, details_box[2] * 0.48)
    max_label_font = clamp(8.8 * scale_ref * scale_content, 5.9, 26.0)
    max_value_font = clamp(9.2 * scale_ref * scale_content, 6.0, 28.0)

    multi_line_fields = {"Origem", "Destino"}
    for row_box, (label, value) in zip(rows, fields):
        rx, ry, rw, rh = row_box
        value_w = max(0.0, rw - label_w - col_gap)
        label_box = (rx, ry, label_w, rh)
        value_box = (rx + label_w + col_gap, ry, value_w, rh)

        draw_text_box(
            c,
            label_box,
            f"{label}:",
            "Helvetica-Bold",
            max_font=max_label_font,
            min_font=5.2,
            max_lines=1,
            line_spacing=1.04,
            align="left",
            valign="center",
        )
        draw_text_box(
            c,
            value_box,
            _safe_text(value),
            "Helvetica",
            max_font=max_value_font,
            min_font=5.2,
            max_lines=2 if label in multi_line_fields else 1,
            line_spacing=1.06,
            align="left",
            valign="center",
        )

    c.setLineWidth(max(0.35, border * 0.7))
    c.line(cx, footer_box[1] + footer_box[3], cx + cw, footer_box[1] + footer_box[3])
    foot_gap = max(2.0, 1.0 * mm_to_points * scale_ref)
    foot_left_w = footer_box[2] * 0.58
    foot_left, foot_right = split_box_horizontal(
        footer_box,
        [foot_left_w, max(0.0, footer_box[2] - foot_left_w - foot_gap)],
        gap=foot_gap,
        from_left=True,
    )
    draw_text_box(
        c,
        foot_left,
        f"Ordem Servico: {_safe_text(dados.get('os', ''))}",
        "Helvetica-Bold",
        max_font=clamp(8.8 * scale_ref * scale_content, 5.6, 24.0),
        min_font=5.0,
        max_lines=1,
        line_spacing=1.0,
        align="left",
        valign="center",
    )
    draw_text_box(
        c,
        foot_right,
        f"Nota Fiscal: {_safe_text(dados.get('nota_fiscal', ''))}",
        "Helvetica-Bold",
        max_font=clamp(8.8 * scale_ref * scale_content, 5.6, 24.0),
        min_font=5.0,
        max_lines=1,
        line_spacing=1.0,
        align="right",
        valign="center",
    )

    _draw_barcode_area(
        c,
        barcode_module,
        barcode_box,
        code_value=_safe_text(dados.get("codigo_barras", "")),
        id_text=_safe_text(dados.get("id_fedex", "")),
        scale_ref=scale_ref,
        scale_content=scale_content,
    )
