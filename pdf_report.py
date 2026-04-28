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

# 日本語フォント登録
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
FONT = 'HeiseiKakuGo-W5'

PAGE_W, PAGE_H = A4
MARGIN = 12 * mm

PRIMARY = colors.HexColor('#346D99')
LIGHT_BG = colors.HexColor('#EEF5FB')
GRAY = colors.HexColor('#475569')
DARK = colors.HexColor('#1e293b')
BORDER = colors.HexColor('#cbd5e1')
PDF_IMAGE_MAX_W_PX = 1280
PDF_IMAGE_MAX_H_PX = 960
PDF_IMAGE_QUALITY = 58


def _styles():
    return {
        'title': ParagraphStyle(
            'title', fontName=FONT, fontSize=17,
            textColor=PRIMARY, spaceAfter=4, spaceBefore=0
        ),
        'subtitle': ParagraphStyle(
            'subtitle', fontName=FONT, fontSize=7,
            textColor=GRAY, spaceAfter=6
        ),
        'section': ParagraphStyle(
            'section', fontName=FONT, fontSize=10,
            textColor=PRIMARY, spaceBefore=5, spaceAfter=3,
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
    """診断結果からA4 PDFを生成してbytesで返す"""
    safe_title = (filename or "5S診断レポート").rsplit(".", 1)[0]
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN
    )
    s = _styles()
    story = []

    # ── ヘッダー ──
    story.append(Paragraph("5S 診断レポート", s['title']))
    now_str = datetime.now().strftime("%Y/%m/%d")
    meta = f"診断日：{now_str}　　会社名：{company or '未入力'}　　部門：{location or '未入力'}"
    story.append(Paragraph(meta, s['subtitle']))
    story.append(HRFlowable(width="100%", color=PRIMARY, thickness=1.5))
    story.append(Spacer(1, 4))

    # ── 写真 ＋ Grade 評価 ──
    overall = result.get("overall_score", 0)
    grade, grade_color = _grade(overall)

    img_el = None
    prepared_img, max_w, new_h = _prepare_pdf_image(image_bytes)
    if prepared_img:
        img_el = Image(prepared_img, width=max_w, height=new_h)

    image_card_parts: list[Any] = [
        Table(
            [[Paragraph("診断画像", s['label_white'])]],
            colWidths=[80 * mm],
            style=TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), PRIMARY),
                ('BOX', (0, 0), (-1, -1), 0.5, PRIMARY),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ])
        ),
        Spacer(1, 3),
    ]
    if img_el:
        image_card_parts.append(img_el)
    else:
        image_card_parts.append(Paragraph("画像なし", s['small']))
    image_card = Table([[image_card_parts]], colWidths=[84 * mm])
    image_card.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    grade_rows: list[list[Any]] = [[
        Table(
            [[Paragraph("グレード評価（4段階評価）", s['label_white'])]],
            colWidths=[80 * mm],
            style=TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), PRIMARY),
                ('BOX', (0, 0), (-1, -1), 0.5, PRIMARY),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ])
        )
    ]]
    for idx, (code, title, desc, color) in enumerate(GRADE_DEFINITIONS):
        is_selected = code == grade
        left = Table([[
            Paragraph(
                f"<para align='center'><font color='{color.hexval()}' size='26'><b>{code}</b></font><br/><font color='{color.hexval()}' size='8'><b>{title}</b></font></para>",
                s['grade_desc']
            )
        ]], colWidths=[18 * mm])
        left.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ]))
        right = Paragraph(desc, s['grade_desc'])
        row = Table([[left, right]], colWidths=[22 * mm, 55 * mm])
        row_style = [
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LINEBELOW', (0, 0), (-1, -1), 0.4, BORDER if idx < len(GRADE_DEFINITIONS) - 1 else colors.white),
        ]
        if is_selected:
            row_style.extend([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f8fbff')),
                ('BOX', (0, 0), (-1, -1), 0.7, color),
            ])
        row.setStyle(TableStyle(row_style))
        grade_rows.append([row])
    grade_card = Table(grade_rows, colWidths=[84 * mm])
    grade_card.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))

    top_table = Table([[image_card, grade_card]], colWidths=[85 * mm, 85 * mm])
    top_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(top_table)
    story.append(Spacer(1, 5))

    # ── 総評 ──
    summary = edited_summary or str(result.get("summary") or "")
    story.append(Paragraph("■ 総評", s['section']))
    story.append(_box(summary, s['summary_text']))
    story.append(Spacer(1, 4))

    # ── 2S診断詳細 ──
    story.append(Paragraph("■ 2S 診断詳細", s['section']))
    detail_rows = []
    for key, label in [("seiri", "整理（Seiri）"), ("seiton", "整頓（Seiton）")]:
        item = result.get(key, {})
        score_val = item.get("score", 0)
        item_grade, _ = _grade(score_val)
        comment = item.get("comment", "")
        priority = item.get("priority", "中")
        item_icon = Table(
            [[Paragraph(label, ParagraphStyle(
                'item_icon', fontName=FONT, fontSize=8.5, textColor=colors.HexColor('#2f855a'), alignment=1
            ))]],
            colWidths=[22 * mm],
        )
        item_icon.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.5, BORDER),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#edf7ed')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        meta = Paragraph(
            f"Grade：{item_grade}　　優先度：{priority}<br/>{comment}",
            s['detail_text']
        )
        detail_rows.append([item_icon, meta])
    detail_table = Table(detail_rows, colWidths=[28 * mm, 142 * mm])
    detail_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('LINEBELOW', (0, 0), (-1, 0), 0.4, BORDER),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(detail_table)
    story.append(Spacer(1, 4))

    # ── 改善アクション ──
    actions = edited_actions or result.get("action_items") or []
    story.append(Paragraph("■ すぐに実行できる改善アクション", s['section']))
    action_rows = []
    for i, action in enumerate(actions):
        num = Paragraph(str(i + 1), ParagraphStyle(
            'num', fontName=FONT, fontSize=8,
            textColor=colors.white, alignment=1
        ))
        num_cell = Table([[num]], colWidths=[5 * mm], rowHeights=[5 * mm])
        num_cell.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), PRIMARY),
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
            ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
        ]))
        action_para = Paragraph(action, s['detail_text'])
        action_rows.append([num_cell, action_para])
    if action_rows:
        action_table = Table(action_rows, colWidths=[10 * mm, 160 * mm])
        action_table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
            ('BACKGROUND', (0, 0), (-1, -1), colors.white),
            ('LINEBELOW', (0, 0), (-1, -2), 0.4, BORDER),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(action_table)

    story.append(Spacer(1, 5))

    # ── 2S 学習セクション ──
    story.append(Paragraph("■ 2S（整理、整頓）の具体的なやり方を学ぶ", s['section']))
    qr_cell = Table(
        [[Paragraph("QRコード", s['qr_label'])]],
        colWidths=[22 * mm],
        rowHeights=[18 * mm],
    )
    qr_cell.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ]))
    learn_left = Table([[
        Paragraph("整理", ParagraphStyle('learn_label', fontName=FONT, fontSize=9, textColor=DARK, alignment=1)),
        qr_cell
    ]], colWidths=[45 * mm, 30 * mm])
    learn_right = Table([[
        Paragraph("整頓", ParagraphStyle('learn_label2', fontName=FONT, fontSize=9, textColor=DARK, alignment=1)),
        qr_cell
    ]], colWidths=[45 * mm, 30 * mm])
    learn_left.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    learn_right.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.6, BORDER),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    learn_table = Table([[learn_left, learn_right]], colWidths=[85 * mm, 85 * mm])
    learn_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(learn_table)

    def _set_pdf_meta(canvas, _doc) -> None:
        canvas.setTitle(safe_title)
        canvas.setSubject("5S診断レポート")
        canvas.setAuthor("5S アドバイスシステム")

    doc.build(story, onFirstPage=_set_pdf_meta, onLaterPages=_set_pdf_meta)
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
