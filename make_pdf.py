#!/usr/bin/env python3
"""塗料業界_業界構造.md を PDF に変換する。

- 日本語フォントに IPAGothic (等幅) を使用し、罫線アートも整列するようにする。
- 簡易 Markdown パーサで # 見出し / ``` コードブロック / | テーブル / 段落 を扱う。
"""

import re
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Preformatted,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether,
)

FONT_REG = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
FONT_PROP = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"

pdfmetrics.registerFont(TTFont("IPAGothic", FONT_REG))
pdfmetrics.registerFont(TTFont("IPAPGothic", FONT_PROP))

SRC = Path(__file__).parent / "塗料業界_業界構造.md"
OUT = Path(__file__).parent / "塗料業界_業界構造.pdf"


# ---------- スタイル定義 ----------
styles = getSampleStyleSheet()

style_body = ParagraphStyle(
    "Body", parent=styles["Normal"],
    fontName="IPAPGothic", fontSize=10.5, leading=16, spaceAfter=6,
)
style_h1 = ParagraphStyle(
    "H1", parent=styles["Heading1"],
    fontName="IPAPGothic", fontSize=20, leading=26,
    textColor=colors.HexColor("#1f3864"),
    spaceBefore=10, spaceAfter=12,
)
style_h2 = ParagraphStyle(
    "H2", parent=styles["Heading2"],
    fontName="IPAPGothic", fontSize=15, leading=22,
    textColor=colors.HexColor("#2e5597"),
    spaceBefore=14, spaceAfter=8,
    borderPadding=4,
    leftIndent=0,
)
style_h3 = ParagraphStyle(
    "H3", parent=styles["Heading3"],
    fontName="IPAPGothic", fontSize=12, leading=18,
    textColor=colors.HexColor("#2e5597"),
    spaceBefore=8, spaceAfter=4,
)
style_quote = ParagraphStyle(
    "Quote", parent=style_body,
    leftIndent=10, textColor=colors.HexColor("#555555"),
    fontSize=10, leading=15,
)
style_code = ParagraphStyle(
    "Code", parent=styles["Code"],
    fontName="IPAGothic", fontSize=8.2, leading=11,
    leftIndent=0, rightIndent=0,
    backColor=colors.HexColor("#f3f5f9"),
    borderColor=colors.HexColor("#d6dce8"),
    borderWidth=0.5, borderPadding=6,
    spaceBefore=6, spaceAfter=8,
)


def md_inline(text: str) -> str:
    """インライン Markdown を ReportLab マークアップに変換。"""
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)\*", r"<i>\1</i>", text)
    text = re.sub(r"`([^`]+)`", r'<font name="IPAGothic" backColor="#eef0f4">\1</font>', text)
    return text


def make_table(rows):
    """マークダウン table 行列から Table フローを生成。"""
    n_cols = len(rows[0])
    # 列幅は均等
    page_width = A4[0] - 30 * mm
    col_w = [page_width / n_cols] * n_cols

    table_data = []
    for r_idx, row in enumerate(rows):
        cells = []
        for cell in row:
            style = style_body
            if r_idx == 0:
                # ヘッダ
                style = ParagraphStyle(
                    "TH", parent=style_body, fontName="IPAPGothic",
                    textColor=colors.white, fontSize=10, leading=14,
                )
            cells.append(Paragraph(md_inline(cell.strip()), style))
        table_data.append(cells)

    t = Table(table_data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3864")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#aab2c2")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f6f8fc")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def parse_markdown(text):
    """Markdown をパースして flowable のリストを返す。"""
    flow = []
    lines = text.splitlines()
    i = 0

    while i < len(lines):
        line = lines[i]

        # 水平線
        if line.strip() == "---":
            flow.append(Spacer(1, 4))
            from reportlab.platypus import HRFlowable
            flow.append(HRFlowable(width="100%", thickness=0.6,
                                   color=colors.HexColor("#bcc3d4"),
                                   spaceBefore=2, spaceAfter=6))
            i += 1
            continue

        # コードブロック
        if line.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1  # 終了 ```
            code_text = "\n".join(buf)
            flow.append(Preformatted(code_text, style_code))
            continue

        # 見出し
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            content = md_inline(m.group(2))
            if level == 1:
                flow.append(Paragraph(content, style_h1))
            elif level == 2:
                flow.append(Paragraph(content, style_h2))
            else:
                flow.append(Paragraph(content, style_h3))
            i += 1
            continue

        # 引用
        if line.startswith(">"):
            content = md_inline(line.lstrip("> ").strip())
            flow.append(Paragraph(content, style_quote))
            i += 1
            continue

        # テーブル
        if line.strip().startswith("|") and i + 1 < len(lines) and re.match(r"^\s*\|[\s\-:|]+\|\s*$", lines[i + 1]):
            tbl_rows = []
            header = [c.strip() for c in line.strip().strip("|").split("|")]
            tbl_rows.append(header)
            i += 2  # ヘッダ + 区切り
            while i < len(lines) and lines[i].strip().startswith("|"):
                row = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                tbl_rows.append(row)
                i += 1
            flow.append(make_table(tbl_rows))
            flow.append(Spacer(1, 6))
            continue

        # リスト（簡易・ネストなし）
        if re.match(r"^\s*[-*]\s+", line):
            content = re.sub(r"^\s*[-*]\s+", "", line)
            content = md_inline(content)
            bullet_style = ParagraphStyle(
                "Bullet", parent=style_body,
                leftIndent=14, bulletIndent=2, spaceAfter=2,
            )
            flow.append(Paragraph(f"<bullet>&bull;</bullet> {content}", bullet_style))
            i += 1
            continue

        # 空行
        if line.strip() == "":
            flow.append(Spacer(1, 4))
            i += 1
            continue

        # 通常段落
        flow.append(Paragraph(md_inline(line), style_body))
        i += 1

    return flow


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("IPAPGothic", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawString(15 * mm, 10 * mm, "塗料業界 全体構造ガイド")
    canvas.drawRightString(A4[0] - 15 * mm, 10 * mm, f"- {doc.page} -")
    canvas.setStrokeColor(colors.HexColor("#bcc3d4"))
    canvas.line(15 * mm, 12 * mm, A4[0] - 15 * mm, 12 * mm)
    canvas.restoreState()


def main():
    text = SRC.read_text(encoding="utf-8")
    flow = parse_markdown(text)

    doc = BaseDocTemplate(
        str(OUT), pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=18 * mm,
        title="塗料業界 全体構造ガイド",
        author="Industry Briefing",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="normal")
    doc.addPageTemplates([PageTemplate(id="all", frames=frame,
                                       onPage=header_footer)])
    doc.build(flow)
    print(f"OK: {OUT}  ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
