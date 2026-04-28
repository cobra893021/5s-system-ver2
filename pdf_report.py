"""5S診断レポートPDF生成モジュール"""
from __future__ import annotations
import io
import zipfile
from datetime import datetime
from typing import Any
from PIL import Image as PILImage

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, Image, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.lib.utils import ImageReader

# 日本語フォント登録
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
FONT = 'HeiseiKakuGo-W5'

PAGE_W, PAGE_H = A4
MARGIN = 12 * mm

PRIMARY = colors.HexColor('#346D99')
NAVY = colors.HexColor('#0B2E5F')
LIGHT_BG = colors.HexColor('#EEF5FB')
GRAY = colors.HexColor('#475569')
DARK = colors.HexColor('#1e293b')
BORDER = colors.HexColor('#cbd5e1')
LINE = colors.HexColor('#C9D3E3')
PDF_IMAGE_MAX_W_PX = 1280
PDF_IMAGE_MAX_H_PX = 960
PDF_IMAGE_QUALITY = 58


def _styles():
    return {
        'title': ParagraphStyle(
            'title', fontName=FONT, fontSize=18,
            textColor=NAVY, spaceAfter=4, spaceBefore=0
        ),
        'subtitle': ParagraphStyle(
            'subtitle', fontName=FONT, fontSize=7,
            textColor=GRAY, spaceAfter=6
        ),
        'section': ParagraphStyle(
            'section', fontName=FONT, fontSize=10,
            textColor=NAVY, spaceBefore=5, spaceAfter=3,
            fontWeight='bold'
        ),
        'box_text': ParagraphStyle(
            'box_text', fontName=FONT, fontSize=7.5,
            textColor=DARK, leading=12, spaceAfter=0
        ),
        'small': ParagraphStyle(
            'small', fontName=FONT, fontSize=6.5,
            textColor=GRAY, spaceAfter=2
        ),
        'score': ParagraphStyle(
            'score', fontName=FONT, fontSize=28,
            textColor=PRIMARY, alignment=1
        ),
        'footer': ParagraphStyle(
            'footer', fontName=FONT, fontSize=7,
            textColor=GRAY
        ),
        'grade_title': ParagraphStyle(
            'grade_title', fontName=FONT, fontSize=10,
            textColor=colors.white, alignment=0
        ),
        'label_white': ParagraphStyle(
            'label_white', fontName=FONT, fontSize=8.5,
            textColor=colors.white
        ),
        'meta': ParagraphStyle(
            'meta', fontName=FONT, fontSize=7,
            textColor=GRAY, alignment=1
        ),
        'summary_text': ParagraphStyle(
            'summary_text', fontName=FONT, fontSize=8.5,
            textColor=DARK, leading=14
        ),
        'detail_text': ParagraphStyle(
            'detail_text', fontName=FONT, fontSize=8,
            textColor=DARK, leading=12
        ),
        'grade_small': ParagraphStyle(
            'grade_small', fontName=FONT, fontSize=8.5,
            textColor=GRAY
        ),
        'grade_desc': ParagraphStyle(
            'grade_desc', fontName=FONT, fontSize=7.6,
            textColor=DARK, leading=11
        ),
        'qr_label': ParagraphStyle(
            'qr_label', fontName=FONT, fontSize=8,
            textColor=DARK, alignment=1
        ),
        'label_dark': ParagraphStyle(
            'label_dark', fontName=FONT, fontSize=8.5,
            textColor=NAVY
        ),
    }


GRADE_DEFINITIONS = [
    (
        "A",
        "とても良好",
        "2S（整理、整頓）が高いレベルであり、ムダが少ない現場です。<br/>維持管理（習慣化）が課題となります。",
        colors.HexColor("#2563eb"),
    ),
    (
        "B",
        "良好",
        "大きな問題は少ないものの、一部に改善余地があります。<br/>改善を行い、現場の収益力を高めましょう。",
        colors.HexColor("#16a34a"),
    ),
    (
        "C",
        "要改善",
        "作業効率や安全面に影響する課題が見られ、早めの対応が必要です。<br/>改善を行うことで10％程度の生産性、収益性の改善が見込まれます。",
        colors.HexColor("#f97316"),
    ),
    (
        "D",
        "早急な改善が必要",
        "探す時間、歩行などが多く発生して、生産性、収益性を大きく下げており、至急改善が必要です。<br/>改善を行うことで20％以上の生産性、収益性の改善が見込まれます。",
        colors.HexColor("#ef4444"),
    ),
]


def _box(text: str, style, bg: colors.Color = LIGHT_BG, width_mm: float = 170) -> Table:
    """テキストを枠付きボックスで表示する"""
    p = Paragraph(text, style)
    t = Table([[p]], colWidths=[width_mm * mm])
    t.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, BORDER),
        ('BACKGROUND', (0, 0), (-1, -1), bg),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [bg]),
    ]))
    return t


def _grade(score: int) -> tuple[str, colors.Color]:
    if score >= 80: return 'A', colors.HexColor('#2563eb')
    if score >= 60: return 'B', colors.HexColor('#16a34a')
    if score >= 40: return 'C', colors.HexColor('#f97316')
    return 'D', colors.HexColor('#ef4444')


def _grade_description(grade: str) -> str:
    for code, _title, desc, _color in GRADE_DEFINITIONS:
        if code == grade:
            return desc.replace("<br/>", " ")
    return ""


def _grade_title(grade: str) -> str:
    for code, title, _desc, _color in GRADE_DEFINITIONS:
        if code == grade:
            return title
    return ""


def _prepare_pdf_image(image_bytes: bytes) -> tuple[io.BytesIO, float, float] | tuple[None, None, None]:
    if not image_bytes:
        return None, None, None

    pil = PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
    pil.thumbnail((PDF_IMAGE_MAX_W_PX, PDF_IMAGE_MAX_H_PX), PILImage.LANCZOS)

    max_w = 80 * mm
    ratio = max_w / pil.width
    new_h = pil.height * ratio
    if new_h > 84 * mm:
        ratio = (84 * mm) / pil.height
        max_w = pil.width * ratio
        new_h = 84 * mm

    img_buf = io.BytesIO()
    pil.save(
        img_buf,
        format="JPEG",
        quality=PDF_IMAGE_QUALITY,
        optimize=True,
        progressive=True,
    )
    img_buf.seek(0)
    return img_buf, max_w, new_h


def _draw_paragraph(c: pdfcanvas.Canvas, text: str, style: ParagraphStyle, x: float, y_top: float, width: float, height: float) -> None:
    para = Paragraph(text.replace("\n", "<br/>"), style)
    _w, used_h = para.wrap(width, height)
    draw_y = y_top - used_h
    para.drawOn(c, x, draw_y)


def _draw_round_card(c: pdfcanvas.Canvas, x: float, y: float, w: float, h: float, radius: float = 6, fill=colors.white, stroke=BORDER, line_width: float = 0.8) -> None:
    c.setLineWidth(line_width)
    c.setStrokeColor(stroke)
    c.setFillColor(fill)
    c.roundRect(x, y, w, h, radius, stroke=1, fill=1)


def _draw_label(c: pdfcanvas.Canvas, x: float, y: float, text: str, width: float = 28 * mm, height: float = 7 * mm) -> None:
    c.setFillColor(NAVY)
    c.setStrokeColor(NAVY)
    c.roundRect(x, y, width, height, 3, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(FONT, 8)
    c.drawString(x + 4, y + 2.1, text)


def _fit_rect(src_w: float, src_h: float, box_w: float, box_h: float) -> tuple[float, float]:
    if src_w <= 0 or src_h <= 0:
        return box_w, box_h
    ratio = min(box_w / src_w, box_h / src_h)
    return src_w * ratio, src_h * ratio


def generate_pdf(
    result: dict[str, Any],
    image_bytes: bytes,
    filename: str,
    company: str = "",
    location: str = "",
    edited_summary: str = "",
    edited_actions: list[str] = None,
    seiri_video_url: str = "",
    seiton_video_url: str = "",
) -> bytes:
    """診断結果からA4帳票デザインPDFを生成してbytesで返す"""
    safe_title = (filename or "5S診断レポート").rsplit(".", 1)[0]
    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=A4)
    c.setTitle(safe_title)
    c.setSubject("5S診断レポート")
    c.setAuthor("5S アドバイスシステム")
    s = _styles()

    page_w, page_h = A4
    margin = MARGIN
    content_w = page_w - margin * 2
    y = page_h - margin

    c.setFillColor(colors.white)
    c.rect(0, 0, page_w, page_h, stroke=0, fill=1)

    # outer subtle frame
    _draw_round_card(c, margin - 2, margin - 2, content_w + 4, page_h - (margin * 2) + 4, radius=8, fill=colors.white, stroke=BORDER, line_width=0.8)

    # header
    c.setFillColor(NAVY)
    c.setFont(FONT, 18)
    c.drawString(margin + 4, y - 8, "5S 診断レポート")

    meta_w = 34 * mm
    meta_gap = 3 * mm
    meta_y_top = y - 2
    meta_x = page_w - margin - (meta_w * 3 + meta_gap * 2)
    meta_items = [
        ("診断日", datetime.now().strftime("%Y/%m/%d")),
        ("会社名", company or "未入力"),
        ("診断場所", location or "未入力"),
    ]
    for idx, (label, value) in enumerate(meta_items):
        x = meta_x + idx * (meta_w + meta_gap)
        _draw_paragraph(c, label, s['meta'], x, meta_y_top, meta_w, 4 * mm)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.8)
        c.line(x, meta_y_top - 6.5 * mm, x + meta_w, meta_y_top - 6.5 * mm)
        _draw_paragraph(c, value, s['small'], x, meta_y_top - 1.5 * mm, meta_w, 4 * mm)

    header_line_y = y - 12 * mm
    c.setStrokeColor(NAVY)
    c.setLineWidth(1.8)
    c.line(margin, header_line_y, page_w - margin, header_line_y)

    # top area
    top_y = header_line_y - 4 * mm
    gap = 4 * mm
    col_w = (content_w - gap) / 2
    top_h = 104 * mm
    left_x = margin
    right_x = margin + col_w + gap
    top_bottom = top_y - top_h

    _draw_round_card(c, left_x, top_bottom, col_w, top_h, radius=8)
    _draw_round_card(c, right_x, top_bottom, col_w, top_h, radius=8)
    _draw_label(c, left_x + 4, top_y - 7 * mm, "診断画像", width=23 * mm, height=6 * mm)
    _draw_label(c, right_x + 4, top_y - 7 * mm, "グレード評価（4段階評価）", width=42 * mm, height=6 * mm)

    # image
    image_box_x = left_x + 6
    image_box_y = top_bottom + 6
    image_box_w = col_w - 12
    image_box_h = top_h - 16
    if image_bytes:
        pil = PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
        img_w, img_h = _fit_rect(pil.width, pil.height, image_box_w, image_box_h)
        img_x = image_box_x + (image_box_w - img_w) / 2
        img_y = image_box_y + (image_box_h - img_h) / 2
        c.drawImage(ImageReader(pil), img_x, img_y, width=img_w, height=img_h, preserveAspectRatio=True, anchor='c', mask='auto')

    # grade list
    overall = result.get("overall_score", 0)
    selected_grade, _grade_color = _grade(overall)
    rows_top = top_y - 12 * mm
    row_h = 22.5 * mm
    grade_left_w = 20 * mm
    for idx, (code, title, desc, color) in enumerate(GRADE_DEFINITIONS):
        row_y = rows_top - (idx + 1) * row_h
        if code == selected_grade:
            c.setFillColor(colors.HexColor('#F7FAFF'))
            c.setStrokeColor(color)
            c.roundRect(right_x + 6, row_y + 1, col_w - 12, row_h - 2, 4, stroke=1, fill=1)
        c.setStrokeColor(LINE)
        if idx < len(GRADE_DEFINITIONS) - 1:
            c.line(right_x + 6, row_y, right_x + col_w - 6, row_y)
        c.setStrokeColor(color)
        c.setLineWidth(1)
        c.line(right_x + grade_left_w + 11, row_y + 3, right_x + grade_left_w + 11, row_y + row_h - 3)
        c.setFillColor(color)
        c.setFont(FONT, 25)
        c.drawCentredString(right_x + 15, row_y + row_h - 10, code)
        c.setFont(FONT, 8.5)
        c.drawCentredString(right_x + 15, row_y + 6, title)
        _draw_paragraph(c, desc, s['grade_desc'], right_x + grade_left_w + 16, row_y + row_h - 4, col_w - grade_left_w - 24, row_h - 8)

    # summary card
    summary_top = top_bottom - 4 * mm
    summary_h = 24 * mm
    summary_bottom = summary_top - summary_h
    _draw_round_card(c, margin, summary_bottom, content_w, summary_h, radius=8)
    _draw_label(c, margin + 4, summary_top - 7 * mm, "総評", width=14 * mm, height=6 * mm)
    summary = edited_summary or str(result.get("summary") or "")
    _draw_paragraph(c, summary, s['summary_text'], margin + 8, summary_top - 10 * mm, content_w - 16, summary_h - 12)

    # 2S detail
    detail_top = summary_bottom - 4 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT, 10)
    c.drawString(margin + 2, detail_top - 4, "2S 診断詳細")
    detail_h = 36 * mm
    detail_bottom = detail_top - detail_h
    _draw_round_card(c, margin, detail_bottom, content_w, detail_h, radius=8)
    detail_row_h = detail_h / 2
    c.setStrokeColor(LINE)
    c.line(margin + 1, detail_bottom + detail_row_h, margin + content_w - 1, detail_bottom + detail_row_h)
    detail_items = [("seiri", "整理（Seiri）"), ("seiton", "整頓（Seiton）")]
    for idx, (key, label) in enumerate(detail_items):
        item = result.get(key, {})
        item_grade, _ = _grade(item.get("score", 0))
        priority = item.get("priority", "中")
        comment = str(item.get("comment") or "")
        row_top = detail_top - idx * detail_row_h
        row_bottom = row_top - detail_row_h
        icon_x = margin + 6
        icon_y = row_bottom + 5
        icon_w = 24 * mm
        icon_h = detail_row_h - 10
        c.setFillColor(colors.HexColor('#EDF7ED'))
        c.setStrokeColor(BORDER)
        c.roundRect(icon_x, icon_y, icon_w, icon_h, 4, stroke=1, fill=1)
        _draw_paragraph(c, label, ParagraphStyle('detail_label', fontName=FONT, fontSize=8.3, textColor=colors.HexColor('#2F855A'), alignment=1), icon_x + 2, row_top - 6, icon_w - 4, icon_h)
        c.setFillColor(DARK)
        c.setFont(FONT, 8)
        c.drawString(margin + 35 * mm, row_top - 8, f"Grade：{item_grade}")
        c.drawString(margin + 58 * mm, row_top - 8, f"優先度：{priority}")
        _draw_paragraph(c, comment, s['detail_text'], margin + 35 * mm, row_top - 11, content_w - 40 * mm, detail_row_h - 12)

    # actions
    actions_top = detail_bottom - 4 * mm
    header_h = 7 * mm
    actions = edited_actions or result.get("action_items") or []
    action_h = 30 * mm
    action_bottom = actions_top - header_h - action_h
    c.setFillColor(NAVY)
    c.setStrokeColor(NAVY)
    c.roundRect(margin, actions_top - header_h, content_w, header_h, 4, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(FONT, 9)
    c.drawString(margin + 6, actions_top - 5, "すぐに実行できる改善アクション")
    _draw_round_card(c, margin, action_bottom, content_w, action_h, radius=8)
    action_row_h = action_h / max(3, len(actions) or 3)
    for idx in range(3):
        row_top = actions_top - header_h - idx * action_row_h
        row_bottom = row_top - action_row_h
        if idx < 2:
            c.setStrokeColor(LINE)
            c.line(margin + 1, row_bottom, margin + content_w - 1, row_bottom)
        c.setFillColor(NAVY)
        c.roundRect(margin + 6, row_bottom + 4, 5 * mm, 5 * mm, 2, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont(FONT, 8)
        c.drawCentredString(margin + 8.5 * mm, row_bottom + 5.7, str(idx + 1))
        text = str(actions[idx]) if idx < len(actions) else ""
        _draw_paragraph(c, text, s['detail_text'], margin + 14 * mm, row_top - 4, content_w - 18 * mm, action_row_h - 6)

    # learn section
    learn_top = action_bottom - 4 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT, 10)
    c.drawString(margin + 2, learn_top - 4, "2S（整理、整頓）の具体的なやり方を学ぶ")
    learn_h = 25 * mm
    learn_bottom = learn_top - learn_h
    learn_gap = 4 * mm
    learn_col_w = (content_w - learn_gap) / 2
    _draw_round_card(c, margin, learn_bottom, learn_col_w, learn_h, radius=8)
    _draw_round_card(c, margin + learn_col_w + learn_gap, learn_bottom, learn_col_w, learn_h, radius=8)
    for idx, label in enumerate(["整理", "整頓"]):
        x = margin + idx * (learn_col_w + learn_gap)
        c.setFillColor(colors.HexColor('#EDF7ED'))
        c.setStrokeColor(BORDER)
        c.roundRect(x + 6, learn_bottom + 5, 22 * mm, learn_h - 10, 4, stroke=1, fill=1)
        _draw_paragraph(c, label, ParagraphStyle('learn_text', fontName=FONT, fontSize=9, textColor=colors.HexColor('#2F855A'), alignment=1), x + 8, learn_top - 7, 18 * mm, learn_h - 8)
        qr_x = x + learn_col_w - 30 * mm
        qr_y = learn_bottom + 5
        qr_w = 22 * mm
        qr_h = learn_h - 10
        c.setFillColor(colors.white)
        c.setStrokeColor(BORDER)
        c.roundRect(qr_x, qr_y, qr_w, qr_h, 4, stroke=1, fill=1)
        _draw_paragraph(c, "QR<br/>コード", s['qr_label'], qr_x, learn_top - 8, qr_w, qr_h)

    c.showPage()
    c.save()
    return buf.getvalue()


def generate_zip(reports: list[tuple[str, bytes]]) -> bytes:
    """(filename, pdf_bytes)のリストからZIPを生成してbytesで返す"""
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for i, (fname, pdf_bytes) in enumerate(reports, 1):
            base = fname.replace('.jpg', '').replace('.png', '').replace('.jpeg', '')
            zip_name = f"{i:02d}_{base}_診断レポート.pdf"
            zf.writestr(zip_name, pdf_bytes)
    return zip_buf.getvalue()
