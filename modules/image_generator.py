"""
image_generator.py
Builds a PNG image of the sprint summary block for sharing in chat/email.
"""

import io
import textwrap
from datetime import date
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

WHITE = "#FFFFFF"
BLACK = "#000000"
SCALE = 2

# Fonts vendored in the repo so the image renders identically on every host
# (local Windows and Linux/Streamlit Cloud) - no dependency on system fonts.
_FONT_DIR = Path(__file__).resolve().parent.parent / "assets" / "fonts"


def _font(size: int, bold: bool = False):
    candidates = [
        # Vendored font first, so output is identical on every platform.
        str(_FONT_DIR / ("DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf")),
        # Fallbacks if the vendored file is somehow unavailable.
        "arialbd.ttf" if bold else "arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    # Last resort: Pillow's built-in scalable default (Pillow >= 10.1) so text
    # is at least correctly sized rather than the tiny fixed bitmap.
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _safe_date(value) -> str:
    if isinstance(value, date):
        return value.strftime("%d %b %Y")
    return "" if value is None else str(value)


def _draw_cell(draw, xy, text, fill, fg=BLACK, font=None, align="center", border="#C9CED6", wrap=True):
    x, y, w, h = xy
    draw.rectangle([x, y, x + w, y + h], fill=fill, outline=border, width=1)
    font = font or _font(18)
    text = "" if text is None else str(text)
    avail = max(w - 8, 4)  # inner horizontal padding

    # First try to keep the text on ONE line by shrinking the font to fit the
    # cell (down to a floor). This makes headers fit without wrapping - matching
    # the intended layout - regardless of the font's width metrics.
    base_size = getattr(font, "size", 14)
    min_size = max(int(base_size * 0.6), 9)
    if text and hasattr(font, "font_variant"):
        size = base_size
        while size > min_size and draw.textlength(text, font=font) > avail:
            size -= 1
            font = font.font_variant(size=size)

    # If it still overflows and wrapping is allowed, wrap at the fitted size.
    if text and wrap and draw.textlength(text, font=font) > avail:
        fs = getattr(font, "size", base_size)
        max_chars = max(int(avail / max(fs * 0.55, 1)), 4)
        lines = textwrap.wrap(text, width=max_chars) or [text]
    else:
        lines = [text] if text else [""]

    line_h = getattr(font, "size", 14) + 3
    total_h = len(lines) * line_h
    ty = y + max((h - total_h) / 2, 3)
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        tw = bbox[2] - bbox[0]
        tx = x + 5
        if align == "center":
            tx = x + max((w - tw) / 2, 5)
        draw.text((tx, ty), line, fill=fg, font=font)
        ty += line_h


def _row(draw, y, widths, values, fills, height, font, text_colors=None, aligns=None):
    x = 0
    text_colors = text_colors or [BLACK] * len(values)
    aligns = aligns or ["center"] * len(values)
    for width, value, fill, fg, align in zip(widths, values, fills, text_colors, aligns):
        _draw_cell(draw, (x, y, width, height), value, fill, fg=fg, font=font, align=align)
        x += width


def build_summary_image(form_data: dict, parsed: dict) -> bytes:
    kpis = parsed["kpis"]
    width = 1728 * SCALE
    row_h = 30 * SCALE
    major_h = 38 * SCALE
    height = 360 * SCALE

    image = Image.new("RGB", (width, height), "#F2F2F2")
    draw = ImageDraw.Draw(image)

    header_font = _font(12 * SCALE, bold=True)
    value_font = _font(13 * SCALE, bold=True)
    small_font = _font(11 * SCALE, bold=True)
    body_font = _font(11 * SCALE)

    meta_w = [w * SCALE for w in [202, 234, 288, 80, 334, 92, 75, 139, 92, 92]]
    meta_headers = [
        "Sprint Number", "Sprint Start Date", "Sprint Development Release", "Sprint QA Release",
        "Production Release", "Tech Debt Release", "Sprint End Date", "Total No. of Days",
        "Scrum Master", "Grooming session",
    ]
    meta_values = [
        form_data["sprint_number"],
        _safe_date(form_data["sprint_start"]),
        _safe_date(form_data["dev_release"]),
        _safe_date(form_data["qa_release"]),
        _safe_date(form_data["prod_release"]),
        "",
        _safe_date(form_data["sprint_end"]),
        form_data["total_days"],
        form_data["scrum_master"],
        "",
    ]
    meta_colors = ["#000000", "#1F4E79", "#2E75B6", "#00B0F0", "#C00000", "#C00000", "#375623", "#595959", "#7030A0", "#7030A0"]
    y = 0
    _row(draw, y, meta_w, meta_headers, meta_colors, row_h, header_font, text_colors=[WHITE] * len(meta_headers))
    y += row_h
    _row(draw, y, meta_w, meta_values, ["#FFF2CC"] * len(meta_values), row_h, value_font)
    y += row_h + (6 * SCALE)

    kpi_w = [w * SCALE for w in [202, 234, 288, 80, 334, 92, 75, 139, 92, 92]]
    kpi_headers = [
        "No of Days Left in Sprint", "Action Items", "Completed - QA", "Completion - QA %",
        "Pending %", "Not Initiated %", "Production Release %", "", "", "",
    ]
    kpi_values = [
        form_data["days_left"], kpis["action_items"], kpis["completed_qa"],
        kpis["completion_qa_pct"], kpis["pending_pct"], kpis["not_initiated_pct"],
        kpis["production_release_pct"], "", "", "",
    ]
    kpi_colors = ["#000000", "#1F3864", "#00B050", "#00B050", "#FFC000", "#ED7D31", "#375623", "#D9D9D9", "#D9D9D9", "#D9D9D9"]
    _row(draw, y, kpi_w, kpi_headers, kpi_colors, row_h, header_font, text_colors=[WHITE, WHITE, WHITE, WHITE, WHITE, WHITE, WHITE, BLACK, BLACK, BLACK])
    y += row_h
    _row(draw, y, kpi_w, kpi_values, [WHITE] * len(kpi_values), row_h, value_font)
    y += row_h + (6 * SCALE)

    stat_w = [w * SCALE for w in [202, 234, 288, 80, 334, 92, 75, 139, 92, 92, 92]]
    daily_task = round(kpis["action_items"] / form_data["total_days"], 2) if form_data["total_days"] > 0 else 0
    stat_headers = [
        "Daily Task Count", "Pending Action Items", "Not Initiated", "In Progress", "Staging",
        "QA Review", "QA Deployed", "QA Approved", "Production", "On Hold",
        "To Be Picked In Another Sprint",
    ]
    stat_values = [
        daily_task, kpis["pending_action_items"], kpis["not_initiated"], kpis["in_progress"],
        kpis["staging"], kpis["qa_review"], kpis["qa_deployed"], kpis["qa_approved"],
        kpis["production"], kpis["on_hold"], kpis["to_be_picked"],
    ]
    stat_colors = ["#595959", "#F4B942", "#ED7D31", "#00B0F0", "#BF8F00", "#FFC000", "#70AD47", "#00B050", "#375623", "#A6A6A6", "#7030A0"]
    _row(draw, y, stat_w, stat_headers, stat_colors, row_h + 6, small_font, text_colors=[WHITE] * len(stat_headers))
    y += row_h + 6
    _row(draw, y, stat_w, stat_values, [WHITE] * len(stat_values), row_h, value_font)
    y += row_h + (6 * SCALE)

    goal_w = [202 * SCALE, 234 * SCALE]
    _row(draw, y, goal_w, ["Sprint Goal", "Major Sprint Items"], ["#7030A0", "#1F3864"], row_h, header_font, text_colors=[WHITE, WHITE])
    y += row_h
    major_rows = [
        [form_data.get("sprint_goal", ""), form_data.get("major_item_1", "")],
        ["", form_data.get("major_item_2", "")],
        ["", form_data.get("major_item_3", "")],
    ]
    for row_values in major_rows:
        _row(
            draw,
            y,
            goal_w,
            row_values,
            ["#FAE5D3", "#FFF2CC"],
            major_h,
            body_font,
            aligns=["left", "left"],
        )
        y += major_h

    out = io.BytesIO()
    image.save(out, format="PNG", optimize=True)
    return out.getvalue()
