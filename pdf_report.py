"""5S診断レポートPDF生成モジュール"""
from __future__ import annotations
import io
import base64
import html
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


def _prepare_embedded_image(image_bytes: bytes) -> str:
    if not image_bytes:
        return ""
    pil = PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
    # PDF上の表示サイズに対して十分な解像度だけ残し、HTML→PDF変換を軽くする
    pil.thumbnail((900, 900), PILImage.LANCZOS)
    out = io.BytesIO()
    pil.save(out, format="JPEG", quality=52, optimize=True, progressive=True)
    return f"data:image/jpeg;base64,{base64.b64encode(out.getvalue()).decode('ascii')}"


def _escape_html(text: str) -> str:
    return html.escape(text or "").replace("\n", "<br>")


def _build_report_html(
    result: dict[str, Any],
    image_data_url: str,
    filename: str,
    company: str,
    location: str,
    edited_summary: str,
    edited_actions: list[str],
) -> str:
    now_str = datetime.now().strftime("%Y/%m/%d")
    summary = edited_summary or str(result.get("summary") or "")
    actions = edited_actions or result.get("action_items") or []
    overall = result.get("overall_score", 0)
    selected_grade, _ = _grade(overall)

    grade_rows = []
    for code, title, desc, color in GRADE_DEFINITIONS:
        selected_class = " selected" if code == selected_grade else ""
        grade_rows.append(
            f"""
            <div class="grade-row{selected_class}">
              <div class="grade-left" style="color:{color};">
                <div class="grade-letter">{code}</div>
                <div class="grade-title">{_escape_html(title)}</div>
              </div>
              <div class="grade-desc">{desc}</div>
            </div>
            """
        )

    detail_rows = []
    for key, label in [("seiri", "整理（Seiri）"), ("seiton", "整頓（Seiton）")]:
        item = result.get(key, {})
        item_grade, _ = _grade(item.get("score", 0))
        priority = str(item.get("priority") or "中")
        comment = str(item.get("comment") or "")
        detail_rows.append(
            f"""
            <div class="detail-row">
              <div class="detail-label">{_escape_html(label)}</div>
              <div class="detail-body">
                <div class="detail-meta">Grade：{_escape_html(item_grade)}　　優先度：{_escape_html(priority)}</div>
                <div class="detail-text">{_escape_html(comment)}</div>
              </div>
            </div>
            """
        )

    action_rows = []
    for idx in range(3):
        text = actions[idx] if idx < len(actions) else ""
        action_rows.append(
            f"""
            <div class="action-row">
              <div class="action-no">{idx + 1}</div>
              <div class="action-text">{_escape_html(str(text))}</div>
            </div>
            """
        )

    image_html = f'<img src="{image_data_url}" alt="診断画像" />' if image_data_url else '<div class="image-placeholder">画像なし</div>'

    return f"""
    <!doctype html>
    <html lang="ja">
    <head>
      <meta charset="utf-8" />
      <style>
        @page {{
          size: A4;
          margin: 10mm;
        }}
        * {{
          box-sizing: border-box;
        }}
        html, body {{
          margin: 0;
          padding: 0;
          background: #ffffff;
          color: #1e293b;
          font-family: "Noto Sans CJK JP", "Hiragino Sans", "Yu Gothic", sans-serif;
          font-size: 11px;
          line-height: 1.5;
        }}
        body {{
          width: 100%;
        }}
        .report {{
          width: 100%;
          padding: 2px 3px 3px;
          background: #fff;
        }}
        .header {{
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 10px;
          padding-bottom: 4px;
          border-bottom: 2px solid #0B2E5F;
        }}
        .header-title {{
          color: #0B2E5F;
          font-size: 20px;
          font-weight: 700;
          white-space: nowrap;
          line-height: 1.1;
        }}
        .header-meta {{
          display: flex;
          gap: 8px;
          flex: 1;
          justify-content: flex-end;
          align-items: flex-start;
        }}
        .meta-item {{
          min-width: 86px;
        }}
        .meta-label {{
          color: #64748b;
          font-size: 8px;
          margin-bottom: 2px;
        }}
        .meta-value {{
          border-bottom: 1px solid #C9D3E3;
          color: #1e293b;
          font-size: 9px;
          padding-bottom: 1px;
          min-height: 14px;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }}
        .top-section {{
          display: flex;
          gap: 10px;
          margin-top: 4px;
          margin-bottom: 7px;
        }}
        .photo-card,
        .grade-card {{
          width: 50%;
          border: 1px solid #C9D3E3;
          border-radius: 10px;
          padding: 6px;
          background: #fff;
          min-height: 282px;
        }}
        .section-label {{
          display: inline-block;
          background: #0B2E5F;
          color: #fff;
          font-size: 9px;
          font-weight: 700;
          border-radius: 6px;
          padding: 3px 9px;
          margin-bottom: 4px;
        }}
        .photo-frame {{
          width: 100%;
          height: 245px;
          border: 1px solid #D9E1EC;
          border-radius: 8px;
          background: #fff;
          display: flex;
          align-items: center;
          justify-content: center;
        }}
        img {{
          width: 100%;
          height: 100%;
          object-fit: contain;
          display: block;
        }}
        .image-placeholder {{
          color: #94a3b8;
          font-size: 11px;
        }}
        .grade-list {{
          display: flex;
          flex-direction: column;
          gap: 3px;
        }}
        .grade-row {{
          display: grid;
          grid-template-columns: 90px 1fr;
          gap: 10px;
          border: 1px solid #D9E1EC;
          border-radius: 8px;
          padding: 5px 6px;
          min-height: 52px;
          background: #fff;
        }}
        .grade-row.selected {{
          background: #F7FAFF;
          border-color: #9FB8DA;
        }}
        .grade-left {{
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          border-right: 1px solid #C9D3E3;
          padding-right: 8px;
        }}
        .grade-letter {{
          font-size: 24px;
          font-weight: 700;
          line-height: 1;
          margin-bottom: 1px;
        }}
        .grade-title {{
          font-size: 9px;
          font-weight: 700;
          line-height: 1.2;
          text-align: center;
        }}
        .grade-desc {{
          font-size: 8.4px;
          line-height: 1.35;
          color: #1e293b;
          align-self: center;
        }}
        .card {{
          border: 1px solid #C9D3E3;
          border-radius: 10px;
          background: #fff;
          margin-top: 7px;
        }}
        .card-header {{
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 5px 8px 0;
          color: #0B2E5F;
          font-size: 12px;
          font-weight: 700;
        }}
        .card-body {{
          padding: 4px 8px 6px;
        }}
        .summary-text {{
          font-size: 9px;
          line-height: 1.42;
          color: #1e293b;
        }}
        .detail-card .card-body {{
          padding-top: 2px;
        }}
        .detail-row {{
          display: grid;
          grid-template-columns: 94px 1fr;
          gap: 10px;
          padding: 5px 0;
          border-top: 1px solid #E3EAF4;
        }}
        .detail-row:first-child {{
          border-top: none;
        }}
        .detail-label {{
          background: #EDF7ED;
          border: 1px solid #C9D3E3;
          border-radius: 999px;
          color: #2F855A;
          font-size: 9px;
          font-weight: 700;
          display: flex;
          align-items: center;
          justify-content: center;
          text-align: center;
          padding: 5px;
          min-height: 34px;
        }}
        .detail-meta {{
          color: #0B2E5F;
          font-size: 8.6px;
          font-weight: 700;
          margin-bottom: 2px;
        }}
        .detail-text {{
          font-size: 8.8px;
          line-height: 1.42;
          color: #1e293b;
          word-break: break-word;
        }}
        .action-card .card-header.bar {{
          background: #0B2E5F;
          color: #fff;
          padding: 5px 8px;
          margin: 0;
          display: block;
        }}
        .action-card .card-body {{
          padding-top: 0;
        }}
        .action-row {{
          display: grid;
          grid-template-columns: 18px 1fr;
          gap: 10px;
          padding: 5px 0;
          border-top: 1px solid #E3EAF4;
        }}
        .action-row:first-child {{
          border-top: none;
        }}
        .action-no {{
          background: #0B2E5F;
          color: #fff;
          font-size: 9px;
          font-weight: 700;
          border-radius: 4px;
          width: 18px;
          height: 18px;
          display: flex;
          align-items: center;
          justify-content: center;
          margin-top: 2px;
        }}
        .action-text {{
          font-size: 8.8px;
          line-height: 1.42;
          color: #1e293b;
          word-break: break-word;
        }}
        .learning-card .card-body {{
          padding-top: 2px;
        }}
        .learning-grid {{
          display: flex;
          gap: 10px;
        }}
        .learning-col {{
          width: 50%;
          border: 1px solid #D9E1EC;
          border-radius: 8px;
          padding: 6px 8px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          min-height: 50px;
        }}
        .learning-label {{
          color: #2F855A;
          font-size: 10px;
          font-weight: 700;
          line-height: 1.25;
        }}
        .qr-box {{
          width: 46px;
          height: 46px;
          border: 1px solid #C9D3E3;
          border-radius: 6px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: #475569;
          font-size: 8px;
          text-align: center;
          line-height: 1.2;
          background: #fff;
          flex-shrink: 0;
        }}
      </style>
    </head>
    <body>
      <div class="report">
        <header class="header">
          <div class="header-title">5S 診断レポート</div>
          <div class="header-meta">
            <div class="meta-item">
              <div class="meta-label">診断日</div>
              <div class="meta-value">{_escape_html(now_str)}</div>
            </div>
            <div class="meta-item">
              <div class="meta-label">会社名</div>
              <div class="meta-value">{_escape_html(company or "未入力")}</div>
            </div>
            <div class="meta-item">
              <div class="meta-label">診断場所</div>
              <div class="meta-value">{_escape_html(location or "未入力")}</div>
            </div>
          </div>
        </header>

        <section class="top-section">
          <div class="photo-card">
            <div class="section-label">診断画像</div>
            <div class="photo-frame">{image_html}</div>
          </div>
          <div class="grade-card">
            <div class="section-label">グレード評価（4段階評価）</div>
            <div class="grade-list">
              {''.join(grade_rows)}
            </div>
          </div>
        </section>

        <section class="summary-card card">
          <div class="card-header">総評</div>
          <div class="card-body">
            <div class="summary-text">{_escape_html(summary)}</div>
          </div>
        </section>

        <section class="detail-card card">
          <div class="card-header">2S 診断詳細</div>
          <div class="card-body">
            {''.join(detail_rows)}
          </div>
        </section>

        <section class="action-card card">
          <div class="card-header bar">すぐに実行できる改善アクション</div>
          <div class="card-body">
            {''.join(action_rows)}
          </div>
        </section>

        <section class="learning-card card">
          <div class="card-header">2S（整理、整頓）の具体的なやり方を学ぶ</div>
          <div class="card-body">
            <div class="learning-grid">
              <div class="learning-col">
                <div class="learning-label">整理</div>
                <div class="qr-box">QR<br>コード</div>
              </div>
              <div class="learning-col">
                <div class="learning-label">整頓</div>
                <div class="qr-box">QR<br>コード</div>
              </div>
            </div>
          </div>
        </section>
      </div>
    </body>
    </html>
    """


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
    """reportlab の固定レイアウトでA4帳票PDFを生成してbytesで返す"""
    styles = _styles()
    overall_score = int(result.get("overall_score", 0) or 0)
    selected_grade, _selected_color = _grade(overall_score)
    selected_grade_description = _grade_description(selected_grade)
    summary = edited_summary or str(result.get("summary") or "")
    actions = edited_actions or result.get("action_items") or []
    now_str = datetime.now().strftime("%Y/%m/%d")

    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=A4)
    c.setTitle("5S診断レポート")

    page_x = MARGIN
    page_y = MARGIN
    page_w = PAGE_W - (MARGIN * 2)
    page_h = PAGE_H - (MARGIN * 2)
    current_top = PAGE_H - MARGIN

    def draw_para(text: str, style: ParagraphStyle, x: float, y_top: float, width: float, height: float) -> float:
        para = Paragraph((text or "").replace("\n", "<br/>"), style)
        _w, used_h = para.wrap(width, height)
        para.drawOn(c, x, y_top - used_h)
        return used_h

    def draw_card_header(title: str, x: float, y_top: float, w: float, dark_bar: bool = False) -> float:
        bar_h = 8 * mm if dark_bar else 6.5 * mm
        if dark_bar:
            c.setFillColor(NAVY)
            c.roundRect(x + 0.5, y_top - bar_h, w - 1.0, bar_h, 6, stroke=0, fill=1)
            c.setFillColor(colors.white)
            c.setFont(FONT, 9.5)
            c.drawString(x + 8, y_top - bar_h + 5.5, title)
        else:
            c.setFillColor(NAVY)
            c.setFont(FONT, 10)
            c.drawString(x + 8, y_top - bar_h + 4.5, title)
        return bar_h

    # Header
    header_h = 15 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT, 18)
    c.drawString(page_x, current_top - 6 * mm, "5S 診断レポート")

    meta_x = page_x + 82 * mm
    meta_w = page_w - (meta_x - page_x)
    item_w = meta_w / 3.0
    meta_items = [
        ("診断日", now_str),
        ("会社名", company or "未入力"),
        ("診断場所", location or "未入力"),
    ]
    for idx, (label, value) in enumerate(meta_items):
        item_x = meta_x + (item_w * idx)
        c.setFillColor(GRAY)
        c.setFont(FONT, 7)
        c.drawString(item_x, current_top - 4.5 * mm, label)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.8)
        c.line(item_x, current_top - 8.2 * mm, item_x + item_w - 5, current_top - 8.2 * mm)
        c.setFillColor(DARK)
        c.setFont(FONT, 8)
        c.drawString(item_x + 2, current_top - 7.2 * mm, str(value)[:26])

    c.setStrokeColor(NAVY)
    c.setLineWidth(1.8)
    c.line(page_x, current_top - header_h, page_x + page_w, current_top - header_h)
    current_top -= header_h + (2 * mm)

    # Top section
    gap = 3.5 * mm
    col_w = (page_w - gap) / 2.0
    top_h = 84 * mm
    left_x = page_x
    right_x = page_x + col_w + gap
    top_y = current_top - top_h

    _draw_round_card(c, left_x, top_y, col_w, top_h, radius=7, stroke=LINE, line_width=0.9)
    _draw_round_card(c, right_x, top_y, col_w, top_h, radius=7, stroke=LINE, line_width=0.9)
    _draw_label(c, left_x + 4, current_top - 8 * mm, "診断画像", width=24 * mm, height=6 * mm)
    _draw_label(c, right_x + 4, current_top - 8 * mm, "グレード評価（4段階評価）", width=45 * mm, height=6 * mm)

    image_box_x = left_x + 6
    image_box_y = top_y + 6
    image_box_w = col_w - 12
    image_box_h = top_h - 16
    c.setStrokeColor(colors.HexColor("#D9E1EC"))
    c.setLineWidth(0.8)
    c.roundRect(image_box_x, image_box_y, image_box_w, image_box_h, 6, stroke=1, fill=0)
    if image_bytes:
        img_buf, _tmp_w, _tmp_h = _prepare_pdf_image(image_bytes)
        if img_buf:
            image_reader = ImageReader(img_buf)
            iw, ih = image_reader.getSize()
            fit_w, fit_h = _fit_rect(iw, ih, image_box_w - 4, image_box_h - 4)
            draw_x = image_box_x + ((image_box_w - fit_w) / 2)
            draw_y = image_box_y + ((image_box_h - fit_h) / 2)
            c.drawImage(image_reader, draw_x, draw_y, width=fit_w, height=fit_h, preserveAspectRatio=True, mask='auto')
    else:
        c.setFillColor(GRAY)
        c.setFont(FONT, 9)
        c.drawCentredString(image_box_x + (image_box_w / 2), image_box_y + (image_box_h / 2), "画像なし")

    grade_inner_x = right_x + 6
    grade_inner_w = col_w - 12
    row_gap = 2.4 * mm
    row_h = (top_h - 18 - (row_gap * 3)) / 4.0
    row_top = current_top - 11 * mm
    for code, title, desc, color in GRADE_DEFINITIONS:
        row_y = row_top - row_h
        row_fill = colors.HexColor("#F7FAFF") if code == selected_grade else colors.white
        row_stroke = colors.HexColor("#9FB8DA") if code == selected_grade else colors.HexColor("#D9E1EC")
        _draw_round_card(c, grade_inner_x, row_y, grade_inner_w, row_h, radius=5, fill=row_fill, stroke=row_stroke, line_width=0.8)
        left_col_w = 24 * mm
        c.setStrokeColor(LINE)
        c.setLineWidth(0.6)
        c.line(grade_inner_x + left_col_w, row_y + 4, grade_inner_x + left_col_w, row_y + row_h - 4)

        c.setFillColor(color)
        c.setFont(FONT, 22)
        c.drawCentredString(grade_inner_x + (left_col_w / 2), row_y + row_h - 10.5 * mm, code)
        c.setFont(FONT, 8.5)
        c.drawCentredString(grade_inner_x + (left_col_w / 2), row_y + 4.5 * mm, title)

        draw_para(
            desc,
            ParagraphStyle(
                "grade_desc_fixed",
                fontName=FONT,
                fontSize=7.2,
                leading=10,
                textColor=DARK,
            ),
            grade_inner_x + left_col_w + 6,
            row_y + row_h - 5,
            grade_inner_w - left_col_w - 12,
            row_h - 10,
        )
        row_top = row_y - row_gap

    current_top = top_y - (2.6 * mm)

    # Summary
    summary_h = 22 * mm
    summary_y = current_top - summary_h
    _draw_round_card(c, page_x, summary_y, page_w, summary_h, radius=7, stroke=LINE, line_width=0.9)
    draw_card_header("総評", page_x, current_top - 1.5, page_w)
    draw_para(
        summary,
        ParagraphStyle(
            "summary_fixed",
            fontName=FONT,
            fontSize=8.5,
            leading=12,
            textColor=DARK,
        ),
        page_x + 8,
        summary_y + summary_h - 12,
        page_w - 16,
        summary_h - 16,
    )
    current_top = summary_y - (2.4 * mm)

    # Detail
    detail_h = 32 * mm
    detail_y = current_top - detail_h
    _draw_round_card(c, page_x, detail_y, page_w, detail_h, radius=7, stroke=LINE, line_width=0.9)
    draw_card_header("2S 診断詳細", page_x, current_top - 1.5, page_w)
    detail_row_h = (detail_h - 12) / 2.0
    detail_top = detail_y + detail_h - 8
    for idx, (key, label) in enumerate([("seiri", "整理（Seiri）"), ("seiton", "整頓（Seiton）")]):
        item = result.get(key, {}) or {}
        item_grade, _ = _grade(int(item.get("score", 0) or 0))
        priority = str(item.get("priority") or "中")
        comment = str(item.get("comment") or "")
        row_y = detail_top - detail_row_h
        if idx > 0:
            c.setStrokeColor(colors.HexColor("#E3EAF4"))
            c.setLineWidth(0.6)
            c.line(page_x + 8, row_y + detail_row_h + 2, page_x + page_w - 8, row_y + detail_row_h + 2)
        label_w = 26 * mm
        c.setFillColor(colors.HexColor("#EDF7ED"))
        c.setStrokeColor(LINE)
        c.roundRect(page_x + 8, row_y + 4, label_w, detail_row_h - 8, 12, stroke=1, fill=1)
        c.setFillColor(colors.HexColor("#2F855A"))
        c.setFont(FONT, 8.5)
        c.drawCentredString(page_x + 8 + (label_w / 2), row_y + detail_row_h - 8, label)
        meta_x = page_x + 8 + label_w + 8
        c.setFillColor(NAVY)
        c.setFont(FONT, 8)
        c.drawString(meta_x, row_y + detail_row_h - 10, f"Grade：{item_grade}")
        c.drawString(meta_x + 32 * mm, row_y + detail_row_h - 10, f"優先度：{priority}")
        draw_para(
            comment,
            ParagraphStyle(
                "detail_fixed",
                fontName=FONT,
                fontSize=8,
                leading=10.5,
                textColor=DARK,
            ),
            meta_x,
            row_y + detail_row_h - 14,
            page_w - (meta_x - page_x) - 10,
            detail_row_h - 16,
        )
        detail_top = row_y
    current_top = detail_y - (2.4 * mm)

    # Actions
    action_h = 28 * mm
    action_y = current_top - action_h
    _draw_round_card(c, page_x, action_y, page_w, action_h, radius=7, stroke=LINE, line_width=0.9)
    bar_h = 7 * mm
    c.setFillColor(NAVY)
    c.roundRect(page_x + 0.5, current_top - bar_h - 1, page_w - 1, bar_h, 6, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(FONT, 9.5)
    c.drawString(page_x + 8, current_top - bar_h + 4.8, "すぐに実行できる改善アクション")
    action_row_h = (action_h - bar_h - 4) / 3.0
    action_top = current_top - bar_h - 2
    for idx in range(3):
        text = str(actions[idx]) if idx < len(actions) else ""
        row_y = action_top - action_row_h
        if idx > 0:
            c.setStrokeColor(colors.HexColor("#E3EAF4"))
            c.setLineWidth(0.6)
            c.line(page_x + 8, row_y + action_row_h, page_x + page_w - 8, row_y + action_row_h)
        no_x = page_x + 8
        no_y = row_y + action_row_h - 14
        c.setFillColor(NAVY)
        c.roundRect(no_x, no_y, 12, 12, 3, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont(FONT, 8)
        c.drawCentredString(no_x + 6, no_y + 3.4, str(idx + 1))
        draw_para(
            text,
            ParagraphStyle(
                "action_fixed",
                fontName=FONT,
                fontSize=8,
                leading=10.5,
                textColor=DARK,
            ),
            no_x + 18,
            row_y + action_row_h - 3,
            page_w - 34,
            action_row_h - 4,
        )
        action_top = row_y
    current_top = action_y - (2.4 * mm)

    # Learning
    learning_h = 22 * mm
    learning_y = current_top - learning_h
    _draw_round_card(c, page_x, learning_y, page_w, learning_h, radius=7, stroke=LINE, line_width=0.9)
    draw_card_header("2S（整理、整頓）の具体的なやり方を学ぶ", page_x, current_top - 1.5, page_w)
    learn_gap = 3.5 * mm
    learn_w = (page_w - 16 - learn_gap) / 2.0
    learn_y = learning_y + 6
    for idx, label in enumerate(["整理", "整頓"]):
        lx = page_x + 8 + (idx * (learn_w + learn_gap))
        _draw_round_card(c, lx, learn_y, learn_w, learning_h - 12, radius=5, stroke=colors.HexColor("#D9E1EC"), line_width=0.8)
        c.setFillColor(colors.HexColor("#2F855A"))
        c.setFont(FONT, 9)
        c.drawString(lx + 8, learn_y + (learning_h - 12) / 2, label)
        c.setStrokeColor(LINE)
        c.roundRect(lx + learn_w - 18 * mm, learn_y + 4, 14 * mm, learning_h - 20, 4, stroke=1, fill=0)
        c.setFillColor(GRAY)
        c.setFont(FONT, 8)
        c.drawCentredString(lx + learn_w - 11 * mm, learn_y + (learning_h - 12) / 2 + 3, "QR")
        c.drawCentredString(lx + learn_w - 11 * mm, learn_y + (learning_h - 12) / 2 - 6, "コード")

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
