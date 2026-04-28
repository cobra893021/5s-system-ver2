"""5S診断レポートPDF生成モジュール"""
from __future__ import annotations

from io import BytesIO
import zipfile
from datetime import datetime
from typing import Any

from PIL import Image

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader


PAGE_W, PAGE_H = A4
MARGIN = 10 * mm
CONTENT_W = PAGE_W - MARGIN * 2

FONT = "HeiseiKakuGo-W5"
pdfmetrics.registerFont(UnicodeCIDFont(FONT))

COLORS = {
    "main": colors.HexColor("#346D99"),
    "navy": colors.HexColor("#0B2E5F"),
    "light_bg": colors.HexColor("#EEF5FB"),
    "line": colors.HexColor("#C9D3E3"),
    "text": colors.HexColor("#1e293b"),
    "sub": colors.HexColor("#475569"),
    "green_bg": colors.HexColor("#EDF7ED"),
    "green": colors.HexColor("#2F855A"),
    "A": colors.HexColor("#2563eb"),
    "B": colors.HexColor("#16a34a"),
    "C": colors.HexColor("#f97316"),
    "D": colors.HexColor("#ef4444"),
}

GRADE_MASTER = {
    "A": {
        "label": "とても良好",
        "text": "2S（整理、整頓）が高いレベルであり、ムダが少ない現場です。維持管理（習慣化）が課題となります。",
    },
    "B": {
        "label": "良好",
        "text": "大きな問題は少ないものの、一部に改善余地があります。改善を行い、現場の収益力を高めましょう。",
    },
    "C": {
        "label": "要改善",
        "text": "作業効率や安全面に影響する課題が見られ、早めの対応が必要です。改善を行うことで10％程度の生産性、収益性の改善が見込まれます。",
    },
    "D": {
        "label": "早急な改善が必要",
        "text": "探す時間、歩行などが多く発生して、生産性、収益性を大きく下げており、至急改善が必要です。改善を行うことで20％以上の生産性、収益性の改善が見込まれます。",
    },
}


def pstyle(size: float = 8, leading: float = 10, color=None, align: int = 0) -> ParagraphStyle:
    return ParagraphStyle(
        name="base",
        fontName=FONT,
        fontSize=size,
        leading=leading,
        textColor=color or COLORS["text"],
        alignment=align,
        wordWrap="CJK",
        spaceAfter=0,
        spaceBefore=0,
    )


def para(c, text: str, x: float, y: float, w: float, h: float, size: float = 8, leading: float = 10, color=None) -> float:
    text = str(text or "").replace("\n", "<br/>")
    p = Paragraph(text, pstyle(size=size, leading=leading, color=color))
    _, ph = p.wrap(w, h)
    p.drawOn(c, x, y + h - ph)
    return ph


def measure_para_height(text: str, w: float, size: float = 8, leading: float = 10, color=None) -> float:
    text = str(text or "").replace("\n", "<br/>")
    p = Paragraph(text, pstyle(size=size, leading=leading, color=color))
    _, ph = p.wrap(w, 10000)
    return ph


def rounded_card(c, x: float, y: float, w: float, h: float, radius: float = 4, stroke=COLORS["line"], fill=None, width: float = 0.7):
    c.saveState()
    c.setStrokeColor(stroke)
    c.setLineWidth(width)
    if fill:
        c.setFillColor(fill)
        c.roundRect(x, y, w, h, radius, stroke=1, fill=1)
    else:
        c.roundRect(x, y, w, h, radius, stroke=1, fill=0)
    c.restoreState()


def draw_summary_icon(c, x: float, y: float, w: float, h: float):
    c.saveState()
    c.setFillColor(COLORS["navy"])
    c.roundRect(x, y, w, h, 2, stroke=0, fill=1)
    c.setStrokeColor(colors.white)
    c.setLineWidth(1.1)
    bubble_x = x + 3.0 * mm
    bubble_y = y + 3.3 * mm
    bubble_w = w - 6.0 * mm
    bubble_h = h - 7.0 * mm
    c.roundRect(bubble_x, bubble_y, bubble_w, bubble_h, 1.6 * mm, stroke=1, fill=0)
    tail = [
        (bubble_x + 5.0 * mm, bubble_y),
        (bubble_x + 7.3 * mm, bubble_y - 2.0 * mm),
        (bubble_x + 8.6 * mm, bubble_y),
    ]
    path = c.beginPath()
    path.moveTo(*tail[0])
    path.lineTo(*tail[1])
    path.lineTo(*tail[2])
    c.drawPath(path, stroke=1, fill=0)
    for dy in (bubble_y + bubble_h - 3.0 * mm, bubble_y + bubble_h - 5.5 * mm):
        c.line(bubble_x + 2.0 * mm, dy, bubble_x + bubble_w - 2.0 * mm, dy)
    c.restoreState()


def navy_label(c, x: float, y: float, text: str, w: float = 32 * mm, h: float = 6 * mm):
    c.saveState()
    c.setFillColor(COLORS["navy"])
    c.roundRect(x, y, w, h, 2, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(FONT, 8.5)
    c.drawString(x + 3 * mm, y + 1.8 * mm, text)
    c.restoreState()


def fit_image(c, image_bytes: bytes | None, x: float, y: float, w: float, h: float):
    if not image_bytes:
        c.setStrokeColor(COLORS["line"])
        c.rect(x, y, w, h)
        c.setFont(FONT, 10)
        c.setFillColor(COLORS["sub"])
        c.drawCentredString(x + w / 2, y + h / 2, "画像なし")
        return

    img = Image.open(BytesIO(image_bytes)).convert("RGB")
    img.thumbnail((1280, 960))

    bio = BytesIO()
    img.save(bio, format="JPEG", quality=58, optimize=True, progressive=True)
    bio.seek(0)

    iw, ih = img.size
    scale = min(w / iw, h / ih)
    dw, dh = iw * scale, ih * scale
    dx = x + (w - dw) / 2
    dy = y + (h - dh) / 2

    c.drawImage(ImageReader(bio), dx, dy, dw, dh, preserveAspectRatio=True, mask="auto")


def grade_from_score(score) -> str:
    try:
        score = float(score)
    except Exception:
        return "C"

    if score >= 80:
        return "A"
    if score >= 60:
        return "B"
    if score >= 40:
        return "C"
    return "D"


def draw_header(c, data: dict[str, Any], y: float) -> float:
    h = 14 * mm

    c.setFillColor(COLORS["navy"])
    c.setFont(FONT, 18)
    c.drawString(MARGIN, y - 9 * mm, "5S 診断レポート")

    meta_x = MARGIN + 82 * mm
    meta_w = CONTENT_W - 82 * mm
    item_w = meta_w / 3

    items = [
        ("診断日", data.get("date", "")),
        ("会社名", data.get("company", "")),
        ("診断場所", data.get("location", "")),
    ]

    for i, (label, value) in enumerate(items):
        x = meta_x + item_w * i
        c.setFont(FONT, 7)
        c.setFillColor(COLORS["text"])
        c.drawString(x, y - 5.2 * mm, label)
        c.setFont(FONT, 8)
        c.drawString(x + 12 * mm, y - 5.2 * mm, str(value or ""))
        c.setStrokeColor(COLORS["line"])
        c.line(x + 12 * mm, y - 6.2 * mm, x + item_w - 3 * mm, y - 6.2 * mm)

    c.setStrokeColor(COLORS["navy"])
    c.setLineWidth(1.4)
    c.line(MARGIN, y - h, PAGE_W - MARGIN, y - h)

    return y - h - 2 * mm


def draw_top_section(c, data: dict[str, Any], y: float) -> float:
    h = 90 * mm
    gap = 3 * mm
    col_w = (CONTENT_W - gap) / 2
    x1 = MARGIN
    x2 = MARGIN + col_w + gap
    bottom = y - h

    rounded_card(c, x1, bottom, col_w, h)
    navy_label(c, x1 + 3 * mm, y - 7 * mm, "診断画像", w=24 * mm)

    img_x = x1 + 3 * mm
    img_y = bottom + 4 * mm
    img_w = col_w - 6 * mm
    img_h = h - 13 * mm
    fit_image(c, data.get("image_bytes"), img_x, img_y, img_w, img_h)

    rounded_card(c, x2, bottom, col_w, h)
    navy_label(c, x2 + 3 * mm, y - 7 * mm, "グレード評価（4段階評価）", w=45 * mm)

    selected = data.get("overall_grade") or grade_from_score(data.get("overall_score", 0))

    inner_x = x2 + 4 * mm
    inner_y = bottom + 4 * mm
    inner_w = col_w - 8 * mm
    row_gap = 1.4 * mm
    row_h = (h - 13 * mm - row_gap * 3) / 4

    for idx, grade in enumerate(["A", "B", "C", "D"]):
        gy = inner_y + (3 - idx) * (row_h + row_gap)
        color = COLORS[grade]
        is_active = grade == selected

        fill = colors.HexColor("#FFF7ED") if is_active and grade == "C" else None
        stroke = color if is_active else COLORS["line"]
        rounded_card(c, inner_x, gy, inner_w, row_h, radius=3, stroke=stroke, fill=fill, width=1.6 if is_active else 0.7)

        left_w = 24 * mm
        c.setStrokeColor(color)
        c.setLineWidth(0.8)
        c.line(inner_x + left_w, gy + 2 * mm, inner_x + left_w, gy + row_h - 2 * mm)

        c.setFillColor(color)
        c.setFont(FONT, 24)
        c.drawCentredString(inner_x + left_w / 2, gy + row_h - 10.2 * mm, grade)

        c.setFont(FONT, 8.5)
        c.drawCentredString(inner_x + left_w / 2, gy + 3.0 * mm, GRADE_MASTER[grade]["label"])

        text_x = inner_x + left_w + 5 * mm
        text_y = gy + 3 * mm
        text_w = inner_w - left_w - 8 * mm
        text_h = row_h - 5 * mm
        para(c, GRADE_MASTER[grade]["text"], text_x, text_y, text_w, text_h, size=6.9, leading=8.4, color=COLORS["text"])

    return bottom - 1.2 * mm


def draw_summary(c, data: dict[str, Any], y: float) -> float:
    text = data.get("edited_summary") or data.get("summary", "")
    text_h = measure_para_height(text, CONTENT_W - 44 * mm, size=8.4, leading=10.6)
    h = max(28 * mm, text_h + 11 * mm)
    x = MARGIN
    bottom = y - h

    rounded_card(c, x, bottom, CONTENT_W, h)
    draw_summary_icon(c, x + 4 * mm, bottom + h - 16 * mm, 12 * mm, 12 * mm)

    c.setFillColor(COLORS["navy"])
    c.setFont(FONT, 10)
    c.drawString(x + 21 * mm, bottom + h - 9 * mm, "総評")

    para(c, text, x + 38 * mm, bottom + 4.5 * mm, CONTENT_W - 44 * mm, h - 8.5 * mm, size=8.4, leading=10.6)

    return bottom - 1.2 * mm


def draw_2s_detail(c, data: dict[str, Any], y: float) -> float:
    label_w = 34 * mm
    comment_w = CONTENT_W - label_w - 8 * mm
    rows = [
        ("整理", "Seiri", data.get("seiri", {})),
        ("整頓", "Seiton", data.get("seiton", {})),
    ]
    row_heights: list[float] = []
    for _jp, _en, item in rows:
        comment = item.get("comment", "")
        comment_h = measure_para_height(comment, comment_w, size=7.9, leading=10.2)
        row_heights.append(max(18 * mm, comment_h + 12 * mm))

    card_h = sum(row_heights)
    h = card_h + 8 * mm
    x = MARGIN
    bottom = y - h

    c.setFillColor(COLORS["navy"])
    c.roundRect(x, y - 7 * mm, 34 * mm, 7 * mm, 2, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(FONT, 9.5)
    c.drawString(x + 3 * mm, y - 5 * mm, "2S 診断詳細")

    card_y = bottom
    rounded_card(c, x, card_y, CONTENT_W, card_h)

    current_row_top = card_y + card_h
    for i, (jp, en, item) in enumerate(rows):
        row_h = row_heights[i]
        ry = current_row_top - row_h

        if i == 1:
            c.setStrokeColor(COLORS["line"])
            c.line(x, current_row_top, x + CONTENT_W, current_row_top)

        c.setFillColor(COLORS["green_bg"])
        c.circle(x + 13 * mm, ry + row_h / 2, 8 * mm, stroke=0, fill=1)

        c.setFillColor(COLORS["green"])
        c.setFont(FONT, 10)
        c.drawCentredString(x + 13 * mm, ry + row_h / 2 + 1 * mm, jp)
        c.setFont(FONT, 7.5)
        c.drawCentredString(x + 13 * mm, ry + row_h / 2 - 4 * mm, f"({en})")

        c.setStrokeColor(COLORS["line"])
        c.line(x + label_w, ry + 3 * mm, x + label_w, ry + row_h - 3 * mm)

        score_grade = item.get("grade") or grade_from_score(item.get("score", 0))
        priority = item.get("priority", "")
        comment = item.get("comment", "")

        tx = x + label_w + 5 * mm
        c.setFont(FONT, 8.5)
        c.setFillColor(COLORS["navy"])
        c.drawString(tx, ry + row_h - 8 * mm, f"Grade：{score_grade}")
        c.setFillColor(COLORS["sub"])
        c.drawString(tx + 24 * mm, ry + row_h - 8 * mm, f"優先度：{priority}")
        c.setFillColor(COLORS["text"])

        para(c, comment, tx, ry + 3.5 * mm, CONTENT_W - label_w - 8 * mm, row_h - 13 * mm, size=7.9, leading=10.2)
        current_row_top = ry

    return bottom - 1.2 * mm


def draw_actions(c, data: dict[str, Any], y: float) -> float:
    actions = data.get("edited_actions") or data.get("actions", [])
    actions = actions[:3] + [""] * max(0, 3 - len(actions))
    action_w = CONTENT_W - 20 * mm
    row_heights = [
        max(9.5 * mm, measure_para_height(action, action_w, size=7.9, leading=10.0) + 4 * mm)
        for action in actions[:3]
    ]
    h = 7 * mm + sum(row_heights) + 2 * mm
    x = MARGIN
    bottom = y - h

    rounded_card(c, x, bottom, CONTENT_W, h)
    c.setFillColor(COLORS["navy"])
    c.roundRect(x, bottom + h - 7 * mm, CONTENT_W, 7 * mm, 2, stroke=0, fill=1)

    c.setFillColor(colors.white)
    c.setFont(FONT, 9.5)
    c.drawString(x + 4 * mm, bottom + h - 5 * mm, "すぐに実行できる改善アクション")

    current_row_top = bottom + h - 7 * mm - 1 * mm
    for i, action in enumerate(actions[:3]):
        row_h = row_heights[i]
        ry = current_row_top - row_h

        c.setFillColor(COLORS["navy"])
        c.roundRect(x + 4 * mm, ry + row_h - 6.5 * mm, 5 * mm, 5 * mm, 1.5, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont(FONT, 8)
        c.drawCentredString(x + 6.5 * mm, ry + row_h - 5 * mm, str(i + 1))

        para(c, action, x + 14 * mm, ry + 1.2 * mm, CONTENT_W - 20 * mm, row_h - 1.5 * mm, size=7.9, leading=10.0)

        if i < 2:
            c.setStrokeColor(COLORS["line"])
            c.line(x + 12 * mm, ry, x + CONTENT_W - 4 * mm, ry)
        current_row_top = ry

    return bottom - 1.2 * mm


def draw_learning(c, y: float) -> float:
    h = 24 * mm
    x = MARGIN
    bottom = y - h

    rounded_card(c, x, bottom, CONTENT_W, h, fill=COLORS["light_bg"], stroke=colors.white)

    c.setFillColor(COLORS["navy"])
    c.setFont(FONT, 9.5)
    c.drawString(x + 4 * mm, bottom + h - 6 * mm, "2S（整理、整頓）の具体的なやり方を学ぶ")

    col_w = CONTENT_W / 2
    items = [("整理についての\n動画を見る", "QR\nコード"), ("整頓についての\n動画を見る", "QR\nコード")]

    for i, (label, _qr) in enumerate(items):
        cx = x + col_w * i + 4 * mm
        cy = bottom + 2.5 * mm
        cw = col_w - 8 * mm
        ch = 12 * mm

        rounded_card(c, cx, cy, cw, ch, radius=3, stroke=colors.white, fill=colors.white)
        c.setFillColor(COLORS["green_bg"])
        c.circle(cx + 10 * mm, cy + ch / 2, 6 * mm, stroke=0, fill=1)

        c.setFillColor(COLORS["green"])
        c.setFont(FONT, 9)
        c.drawString(cx + 22 * mm, cy + 6.5 * mm, label.split("\n")[0])
        c.drawString(cx + 22 * mm, cy + 2.5 * mm, label.split("\n")[1])

        qr_w = 21 * mm
        qx = cx + cw - qr_w - 4 * mm
        qy = cy + 1.2 * mm
        rounded_card(c, qx, qy, qr_w, ch - 2.4 * mm, radius=2)
        c.setFillColor(COLORS["text"])
        c.setFont(FONT, 8.2)
        c.drawCentredString(qx + qr_w / 2, qy + 6.3 * mm, "QR")
        c.drawCentredString(qx + qr_w / 2, qy + 2.6 * mm, "コード")

    return bottom


def ensure_space(c, data: dict[str, Any], y: float, needed_h: float) -> float:
    if y - needed_h >= MARGIN:
        return y
    c.showPage()
    return draw_header(c, data, PAGE_H - MARGIN)


def estimate_summary_height(data: dict[str, Any]) -> float:
    text = data.get("edited_summary") or data.get("summary", "")
    text_h = measure_para_height(text, CONTENT_W - 44 * mm, size=8.4, leading=10.6)
    return max(28 * mm, text_h + 11 * mm) + 1.2 * mm


def estimate_2s_height(data: dict[str, Any]) -> float:
    label_w = 34 * mm
    comment_w = CONTENT_W - label_w - 8 * mm
    total_rows_h = 0.0
    for item in (data.get("seiri", {}), data.get("seiton", {})):
        comment_h = measure_para_height(item.get("comment", ""), comment_w, size=7.9, leading=10.2)
        total_rows_h += max(18 * mm, comment_h + 12 * mm)
    return total_rows_h + 8 * mm + 1.2 * mm


def estimate_actions_height(data: dict[str, Any]) -> float:
    actions = data.get("edited_actions") or data.get("actions", [])
    actions = actions[:3] + [""] * max(0, 3 - len(actions))
    action_w = CONTENT_W - 20 * mm
    total_rows_h = sum(
        max(9.5 * mm, measure_para_height(action, action_w, size=7.9, leading=10.0) + 4 * mm)
        for action in actions[:3]
    )
    return 7 * mm + total_rows_h + 2 * mm + 1.2 * mm


def estimate_learning_height() -> float:
    return 24 * mm


def generate_pdf(
    result: dict[str, Any],
    image_bytes: bytes,
    filename: str,
    company: str = "",
    location: str = "",
    edited_summary: str = "",
    edited_actions: list[str] | None = None,
    seiri_video_url: str = "",
    seiton_video_url: str = "",
) -> bytes:
    data = {
        "date": datetime.now().strftime("%Y/%m/%d"),
        "company": company,
        "location": location,
        "image_bytes": image_bytes,
        "overall_score": result.get("overall_score", 0),
        "overall_grade": grade_from_score(result.get("overall_score", 0)),
        "summary": result.get("summary", ""),
        "edited_summary": edited_summary,
        "seiri": {
            "score": result.get("seiri", {}).get("score", 0),
            "grade": result.get("seiri", {}).get("grade") or grade_from_score(result.get("seiri", {}).get("score", 0)),
            "priority": result.get("seiri", {}).get("priority", ""),
            "comment": result.get("seiri", {}).get("comment", ""),
        },
        "seiton": {
            "score": result.get("seiton", {}).get("score", 0),
            "grade": result.get("seiton", {}).get("grade") or grade_from_score(result.get("seiton", {}).get("score", 0)),
            "priority": result.get("seiton", {}).get("priority", ""),
            "comment": result.get("seiton", {}).get("comment", ""),
        },
        "actions": result.get("action_items", []) or [],
        "edited_actions": edited_actions or [],
    }

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    y = PAGE_H - MARGIN
    y = draw_header(c, data, y)
    y = draw_top_section(c, data, y)
    y = ensure_space(c, data, y, estimate_summary_height(data))
    y = draw_summary(c, data, y)
    y = ensure_space(c, data, y, estimate_2s_height(data))
    y = draw_2s_detail(c, data, y)
    y = ensure_space(c, data, y, estimate_actions_height(data))
    y = draw_actions(c, data, y)
    y = ensure_space(c, data, y, estimate_learning_height())
    draw_learning(c, y)
    c.showPage()
    c.save()
    return buf.getvalue()


def generate_zip(reports: list[tuple[str, bytes]]) -> bytes:
    """(filename, pdf_bytes)のリストからZIPを生成してbytesで返す"""
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (fname, pdf_bytes) in enumerate(reports, 1):
            base = fname.replace(".jpg", "").replace(".png", "").replace(".jpeg", "")
            zip_name = f"{i:02d}_{base}_診断レポート.pdf"
            zf.writestr(zip_name, pdf_bytes)
    return zip_buf.getvalue()
