"""
pdf_generator.py
Builds a landscape PDF version of the sprint report without external PDF
dependencies. The table is scaled to one page width and continues vertically
across as many pages as needed.
"""

import io
from datetime import date
from typing import Iterable

JIRA_BASE = "https://jira-zigram.atlassian.net/browse"

A3_LANDSCAPE = (1190.55, 841.89)
MARGIN = 18

WHITE = "FFFFFF"
BLACK = "000000"
BLUE = "1F3864"
LIGHT_GRID = "D9E2F3"
MUTED_GRID = "E2E8F0"


def _rgb(hex_color: str) -> tuple[float, float, float]:
    value = hex_color.strip("#")
    return (
        int(value[0:2], 16) / 255,
        int(value[2:4], 16) / 255,
        int(value[4:6], 16) / 255,
    )


def _pdf_text(value) -> str:
    text = "" if value is None else str(value)
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _safe_date(value) -> str:
    if isinstance(value, date):
        return value.strftime("%d %b %Y")
    return "" if value is None else str(value)


def _wrap_text(value, width: float, font_size: float) -> list[str]:
    text = "" if value is None else str(value)
    text = " ".join(text.replace("\r", " ").replace("\n", " ").split())
    if not text:
        return [""]

    max_chars = max(int(width / max(font_size * 0.48, 1)), 4)
    lines = []
    for paragraph in text.splitlines() or [text]:
        current = ""
        for word in paragraph.split(" "):
            if not word:
                continue
            while len(word) > max_chars:
                if current:
                    lines.append(current)
                    current = ""
                lines.append(word[:max_chars])
                word = word[max_chars:]
            candidate = word if not current else f"{current} {word}"
            if len(candidate) <= max_chars:
                current = candidate
            else:
                lines.append(current)
                current = word
        if current:
            lines.append(current)
    return lines or [""]


class _PDF:
    def __init__(self, size=A3_LANDSCAPE):
        self.width, self.height = size
        self.pages: list[list[str]] = []
        self.current: list[str] | None = None

    def new_page(self) -> None:
        self.current = []
        self.pages.append(self.current)

    def _cmd(self, command: str) -> None:
        if self.current is None:
            self.new_page()
        self.current.append(command)

    def rect(self, x, y, w, h, fill=WHITE, stroke="CCCCCC", line_width=0.35):
        fr, fg, fb = _rgb(fill)
        sr, sg, sb = _rgb(stroke)
        self._cmd(
            f"{line_width:.2f} w {fr:.3f} {fg:.3f} {fb:.3f} rg "
            f"{sr:.3f} {sg:.3f} {sb:.3f} RG {x:.2f} {y:.2f} {w:.2f} {h:.2f} re B"
        )

    def text(self, x, y, value, size=8, color=BLACK, bold=False):
        r, g, b = _rgb(color)
        font = "F2" if bold else "F1"
        self._cmd(
            f"BT {r:.3f} {g:.3f} {b:.3f} rg /{font} {size:.2f} Tf "
            f"1 0 0 1 {x:.2f} {y:.2f} Tm ({_pdf_text(value)}) Tj ET"
        )

    def cell(self, x, y, w, h, value, fill=WHITE, text_color=BLACK, size=7, bold=False, align="left"):
        self.rect(x, y, w, h, fill=fill)
        lines = _wrap_text(value, w - 5, size)
        line_h = size + 1.7
        max_lines = max(int((h - 4) / line_h), 1)
        lines = lines[:max_lines]
        if len(lines) == max_lines and len(_wrap_text(value, w - 5, size)) > max_lines:
            lines[-1] = lines[-1][:-3] + "..." if len(lines[-1]) > 3 else "..."
        start_y = y + h - size - 3
        for i, line in enumerate(lines):
            tx = x + 2.6
            if align == "center":
                approx_w = len(line) * size * 0.46
                tx = x + max((w - approx_w) / 2, 2.6)
            self.text(tx, start_y - i * line_h, line, size=size, color=text_color, bold=bold)

    def bytes(self) -> bytes:
        objects: list[bytes] = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"",  # pages placeholder
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
        ]

        kids = []
        for page in self.pages:
            content = "\n".join(page).encode("latin-1", errors="replace")
            stream = b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream"
            page_obj_num = len(objects) + 1
            content_obj_num = len(objects) + 2
            kids.append(f"{page_obj_num} 0 R")
            objects.append(
                (
                    f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {self.width:.2f} {self.height:.2f}] "
                    f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                    f"/Contents {content_obj_num} 0 R >>"
                ).encode()
            )
            objects.append(stream)

        objects[1] = f"<< /Type /Pages /Kids [{' '.join(kids)}] /Count {len(kids)} >>".encode()

        out = io.BytesIO()
        out.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for i, obj in enumerate(objects, start=1):
            offsets.append(out.tell())
            out.write(f"{i} 0 obj\n".encode())
            out.write(obj)
            out.write(b"\nendobj\n")
        xref = out.tell()
        out.write(f"xref\n0 {len(objects) + 1}\n".encode())
        out.write(b"0000000000 65535 f \n")
        for offset in offsets[1:]:
            out.write(f"{offset:010d} 00000 n \n".encode())
        out.write(
            (
                f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
                f"startxref\n{xref}\n%%EOF"
            ).encode()
        )
        return out.getvalue()


def _draw_row(pdf: _PDF, x: float, y: float, widths: list[float], values: Iterable, fills: list[str], height: float, size=6.2, bold=False):
    cursor = x
    for width, value, fill in zip(widths, values, fills):
        text_color = WHITE if fill in {"1F3864", "2E75B6", "7030A0", "375623", "C00000", "000000"} else BLACK
        pdf.cell(cursor, y, width, height, value, fill=fill, text_color=text_color, size=size, bold=bold)
        cursor += width


def _row_height(values: list, widths: list[float], font_size: float, min_height=15, max_height=90) -> float:
    line_counts = [_wrap_text(v, w - 5, font_size) for v, w in zip(values, widths)]
    max_lines = max(len(lines) for lines in line_counts)
    return min(max(min_height, 5 + max_lines * (font_size + 1.8)), max_height)


def build_pdf(form_data: dict, parsed: dict) -> bytes:
    pdf = _PDF()
    pdf.new_page()

    page_w, page_h = pdf.width, pdf.height
    table_x = MARGIN
    table_w = page_w - (MARGIN * 2)
    y = page_h - MARGIN

    kpis = parsed["kpis"]
    hierarchy = parsed["hierarchy"]

    pdf.rect(MARGIN, y - 38, table_w, 38, fill=BLUE, stroke=BLUE)
    pdf.text(MARGIN + 12, y - 16, f"Sprint Report - {form_data['project_name']}", size=16, color=WHITE, bold=True)
    pdf.text(MARGIN + 12, y - 30, f"Sprint #{form_data['sprint_number']} | Generated report", size=8.5, color="BDD7EE")
    y -= 52

    meta_headers = [
        "Sprint Number", "Sprint Start Date", "Sprint Development Release", "Sprint QA Release",
        "Production Release", "Sprint End Date", "Total No. of Days", "Scrum Master",
    ]
    meta_values = [
        form_data["sprint_number"],
        _safe_date(form_data["sprint_start"]),
        _safe_date(form_data["dev_release"]),
        _safe_date(form_data["qa_release"]),
        _safe_date(form_data["prod_release"]),
        _safe_date(form_data["sprint_end"]),
        form_data["total_days"],
        form_data["scrum_master"],
    ]
    meta_w = [table_w / len(meta_headers)] * len(meta_headers)
    _draw_row(pdf, table_x, y - 18, meta_w, meta_headers, ["1F3864"] * len(meta_headers), 18, size=6.3, bold=True)
    y -= 36
    _draw_row(pdf, table_x, y, meta_w, meta_values, ["FFF2CC"] * len(meta_values), 18, size=6.5)
    y -= 30

    kpi_headers = ["Days Left", "Action Items", "Completed - QA", "Completion - QA %", "Pending %", "Not Initiated %", "Production Release %"]
    kpi_values = [
        form_data["days_left"],
        kpis["action_items"],
        kpis["completed_qa"],
        kpis["completion_qa_pct"],
        kpis["pending_pct"],
        kpis["not_initiated_pct"],
        kpis["production_release_pct"],
    ]
    kpi_colors = ["000000", "1F3864", "00B050", "00B050", "FFC000", "ED7D31", "375623"]
    kpi_w = [table_w / len(kpi_headers)] * len(kpi_headers)
    _draw_row(pdf, table_x, y - 18, kpi_w, kpi_headers, kpi_colors, 18, size=6.2, bold=True)
    y -= 36
    _draw_row(pdf, table_x, y, kpi_w, kpi_values, [WHITE] * len(kpi_values), 18, size=7, bold=True)
    y -= 30

    stat_headers = [
        "Daily Task Count", "Pending Action Items", "Not Initiated", "In Progress", "Staging",
        "QA Review", "QA Deployed", "QA Approved", "Production", "On Hold", "Another Sprint",
    ]
    daily_task = round(kpis["action_items"] / form_data["total_days"], 2) if form_data["total_days"] > 0 else 0
    stat_values = [
        daily_task, kpis["pending_action_items"], kpis["not_initiated"], kpis["in_progress"],
        kpis["staging"], kpis["qa_review"], kpis["qa_deployed"], kpis["qa_approved"],
        kpis["production"], kpis["on_hold"], kpis["to_be_picked"],
    ]
    stat_colors = ["595959", "F4B942", "ED7D31", "00B0F0", "BF8F00", "FFC000", "70AD47", "00B050", "375623", "A6A6A6", "7030A0"]
    stat_w = [table_w / len(stat_headers)] * len(stat_headers)
    _draw_row(pdf, table_x, y - 20, stat_w, stat_headers, stat_colors, 20, size=5.6, bold=True)
    y -= 39
    _draw_row(pdf, table_x, y, stat_w, stat_values, [WHITE] * len(stat_values), 18, size=7, bold=True)
    y -= 32

    goal_w = [table_w * 0.38, table_w * 0.62]
    major_rows = [
        [form_data.get("sprint_goal", ""), form_data.get("major_item_1", "")],
        ["", form_data.get("major_item_2", "")],
        ["", form_data.get("major_item_3", "")],
    ]
    goal_heights = [
        _row_height(row, goal_w, 6.5, min_height=20, max_height=44)
        for row in major_rows
    ]
    _draw_row(pdf, table_x, y - 18, goal_w, ["Sprint Goal", "Major Sprint Items"], ["7030A0", "1F3864"], 18, size=6.5, bold=True)
    y -= 18
    for row, row_h in zip(major_rows, goal_heights):
        y -= row_h
        _draw_row(pdf, table_x, y, goal_w, row, ["FAE5D3", "FFF2CC"], row_h, size=6.5)
    y -= 22

    headers = [
        "S.No", "Issue Key", "Jira Link / Confluence Document Link", "Issue Type",
        "Summary / Title", "Status", "Priority", "Assignee", "Start Date", "End Date",
        "Revised Start Date", "Revised End Date", "Comment",
    ]
    proportions = [0.035, 0.055, 0.13, 0.055, 0.19, 0.065, 0.045, 0.075, 0.055, 0.055, 0.06, 0.06, 0.12]
    widths = [table_w * p for p in proportions]

    rows = []
    epic_counter = 0
    story_counter = 0
    sub_counters = {}
    last_epic_num = 0
    last_story_num = 0
    last_story_key = None
    for item in hierarchy:
        if item["level"] == 0:
            epic_counter += 1
            story_counter = 0
            last_story_key = None
            item_sno = str(epic_counter)
            last_epic_num = epic_counter
        elif item["level"] == 1:
            story_counter += 1
            last_story_key = item["issue_key"]
            sub_counters[last_story_key] = 0
            item_sno = f"{last_epic_num}.{story_counter}"
            last_story_num = story_counter
        else:
            if last_story_key and last_story_key in sub_counters:
                sub_counters[last_story_key] += 1
                sub_num = sub_counters[last_story_key]
            else:
                sub_num = 1
            item_sno = f"{last_epic_num}.{last_story_num}.{sub_num}"

        indent = "" if item["level"] == 0 else ("  > " if item["level"] == 1 else "    - ")
        rows.append([
            item_sno,
            item["issue_key"],
            f"{JIRA_BASE}/{item['issue_key']}" if item["issue_key"] != "-" else "",
            item["issue_type"],
            indent + item["summary"],
            item["status"],
            item["priority"],
            item["assignee"],
            item["target_start"],
            item["target_end"],
            "",
            "",
            item.get("latest_comment", ""),
            item["level"],
        ])

    def draw_task_header(current_y: float) -> float:
        _draw_row(pdf, table_x, current_y - 22, widths, headers, ["D1BBF0"] * len(headers), 22, size=5.6, bold=True)
        return current_y - 22

    if y < 150:
        pdf.new_page()
        y = page_h - MARGIN
    y = draw_task_header(y)

    for row in rows:
        values = row[:-1]
        level = row[-1]
        height = _row_height(values, widths, 5.4, min_height=15, max_height=80)
        if y - height < MARGIN:
            pdf.new_page()
            y = page_h - MARGIN
            y = draw_task_header(y)

        if level == 0:
            fills = ["C39BD3"] * len(values)
            bold = True
        else:
            fills = [WHITE] * len(values)
            status = str(values[5])
            fills[5] = "C6EFCE" if any(s in status for s in ["Done", "Production", "Released"]) else (
                "FFEB9C" if "Progress" in status else ("FFF2CC" if "Staging" in status else "FCE4D6")
            )
            bold = level == 1
        _draw_row(pdf, table_x, y - height, widths, values, fills, height, size=5.4, bold=bold)
        y -= height

    return pdf.bytes()
