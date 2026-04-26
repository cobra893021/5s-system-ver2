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
MARGIN = 20 * mm

PRIMARY = colors.HexColor('#346D99')
LIGHT_BG = colors.HexColor('#EEF5FB')
GRAY = colors.HexColor('#475569')
DARK = colors.HexColor('#1e293b')


def _styles():
    return {
        'title': ParagraphStyle(
            'title', fontName=FONT, fontSize=16,
            textColor=PRIMARY, spaceAfter=8, spaceBefore=0
        ),
        'subtitle': ParagraphStyle(
            'subtitle', fontName=FONT, fontSize=8,
            textColor=GRAY, spaceAfter=16
        ),
        'section': ParagraphStyle(
            'section', fontName=FONT, fontSize=11,
            textColor=PRIMARY, spaceBefore=8, spaceAfter=4,
            fontWeight='bold'
        ),
        'box_text': ParagraphStyle(
            'box_text', fontName=FONT, fontSize=8,
            textColor=DARK, leading=14, spaceAfter=0
        ),
        'small': ParagraphStyle(
            'small', fontName=FONT, fontSize=7,
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
    }


def _box(text: str, style) -> Table:
    """テキストを枠付きボックスで表示する"""
    p = Paragraph(text, style)
    t = Table([[p]], colWidths=[170 * mm])
    t.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
        ('BACKGROUND', (0, 0), (-1, -1), LIGHT_BG),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [LIGHT_BG]),
    ]))
    return t


def _grade(score: int) -> tuple[str, colors.Color]:
    if score >= 80: return 'A', colors.HexColor('#2563eb')
    if score >= 60: return 'B', colors.HexColor('#16a34a')
    if score >= 40: return 'C', colors.HexColor('#f97316')
    return 'D', colors.HexColor('#ef4444')


def _grade_description(grade: str) -> str:
    descriptions = {
        "A": "2S（整理、整頓）が高いレベルであり、ムダが少ない現場です。維持管理（習慣化）が課題となります。",
        "B": "大きな問題は少ないものの、一部に改善余地があります。改善を行い、現場の収益力を高めましょう。",
        "C": "作業効率や安全面に影響する課題が見られ、早めの対応が必要です。改善を行うことで10％程度の生産性、収益性の改善が見込まれます。",
        "D": "探す時間、歩行などが多く発生して、生産性、収益性を大きく下げており、至急改善が必要です。改善を行うことで20％以上の生産性、収益性の改善が見込まれます。",
    }
    return descriptions.get(grade, "")


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
    story.append(Spacer(1, 4))
    now_str = datetime.now().strftime("%Y/%m/%d")
    meta = f"診断日：{now_str}　　会社名：{company or '未入力'}　　部門：{location or '未入力'}"
    story.append(Paragraph(meta, s['subtitle']))
    story.append(HRFlowable(width="100%", color=PRIMARY, thickness=1.5))
    story.append(Spacer(1, 6))

    # ── 写真 ＋ Grade 横並び ──
    overall = result.get("overall_score", 0)
    grade, grade_color = _grade(overall)
    grade_description = _grade_description(grade)

    img_el = None
    if image_bytes:
        pil = PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
        max_w = 85 * mm
        ratio = max_w / pil.width
        new_h = pil.height * ratio
        if new_h > 60 * mm:
            ratio = (60 * mm) / pil.height
            max_w = pil.width * ratio
            new_h = 60 * mm
        img_buf = io.BytesIO()
        pil.save(img_buf, format="JPEG", quality=85)
        img_buf.seek(0)
        img_el = Image(img_buf, width=max_w, height=new_h)

    grade_content = [
        [Paragraph("TOTAL GRADE", s['small'])],
        [Paragraph(grade, ParagraphStyle(
            'grade', fontName=FONT, fontSize=32, leading=36,
            textColor=grade_color, alignment=1
        ))],
        [Paragraph(
            grade_description,
            ParagraphStyle(
                'grade_desc_single',
                fontName=FONT,
                fontSize=8,
                leading=14,
                textColor=DARK,
                alignment=0,
            )
        )],
    ]
    grade_table = Table(grade_content, colWidths=[70 * mm])
    grade_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, -1), LIGHT_BG),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 2), (0, 2), 'LEFT'),
    ]))

    if img_el:
        top_table = Table([[img_el, grade_table]],
                          colWidths=[95 * mm, 75 * mm])
        top_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (0, 0), 0),
            ('RIGHTPADDING', (1, 0), (1, 0), 0),
        ]))
        story.append(top_table)
    else:
        story.append(grade_table)

    story.append(Spacer(1, 8))

    # ── 総評 ──
    summary = edited_summary or str(result.get("summary") or "")
    story.append(Paragraph("■ 総評", s['section']))
    story.append(_box(summary, s['box_text']))
    story.append(Spacer(1, 6))

    # ── 2S診断詳細 ──
    story.append(Paragraph("■ 2S 診断詳細", s['section']))
    for key, label in [("seiri", "整理（Seiri）"), ("seiton", "整頓（Seiton）")]:
        item = result.get(key, {})
        score_val = item.get("score", 0)
        item_grade, _ = _grade(score_val)
        comment = item.get("comment", "")
        priority = item.get("priority", "中")
        story.append(Paragraph(
            f"{label}　Grade {item_grade}　／　優先度：{priority}",
            ParagraphStyle(
                'item_title', fontName=FONT, fontSize=9,
                textColor=DARK, spaceBefore=4, spaceAfter=2
            )
        ))
        story.append(_box(comment, s['box_text']))
        story.append(Spacer(1, 3))

    story.append(Spacer(1, 4))

    # ── 改善アクション ──
    actions = edited_actions or result.get("action_items") or []
    story.append(Paragraph("■ すぐに実行できる改善アクション", s['section']))
    priority_colors = [
        colors.HexColor('#ef4444'),
        colors.HexColor('#64748b'),
        colors.HexColor('#94a3b8'),
    ]
    for i, action in enumerate(actions):
        c = priority_colors[i] if i < len(priority_colors) else priority_colors[-1]
        num = Paragraph(str(i + 1), ParagraphStyle(
            'num', fontName=FONT, fontSize=8,
            textColor=colors.white, alignment=1
        ))
        num_cell = Table([[num]], colWidths=[5 * mm], rowHeights=[5 * mm])
        num_cell.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), c),
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
            ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
        ]))
        action_para = Paragraph(action, s['box_text'])
        action_row = Table([[num_cell, action_para]],
                           colWidths=[8 * mm, 162 * mm])
        action_row.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(action_row)

    # ── フッター ──
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", color=colors.HexColor('#e2e8f0')))
    story.append(Spacer(1, 3))
    story.append(Paragraph(
        f"ファイル名：{filename}　　生成日時：{datetime.now().strftime('%Y/%m/%d %H:%M')}",
        s['footer']
    ))

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
