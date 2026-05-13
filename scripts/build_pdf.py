"""Convert a Markdown file to a Japanese-capable PDF.

Usage:
    python3 scripts/build_pdf.py                       # default report
    python3 scripts/build_pdf.py <input.md>            # auto output: same stem + .pdf
    python3 scripts/build_pdf.py <input.md> <out.pdf>  # explicit output
"""
import re
import sys
from pathlib import Path

from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SRC = ROOT / "research" / "paint-industry-demand-side.md"

FONT_REGULAR = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
FONT_PROPORTIONAL = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"

pdfmetrics.registerFont(TTFont("JP", FONT_REGULAR))
pdfmetrics.registerFont(TTFont("JP-P", FONT_PROPORTIONAL))

BASE = getSampleStyleSheet()
STYLES = {
    "h1": ParagraphStyle(
        "h1", parent=BASE["Heading1"], fontName="JP", fontSize=20,
        leading=28, spaceBefore=12, spaceAfter=10, textColor=HexColor("#1a365d"),
    ),
    "h2": ParagraphStyle(
        "h2", parent=BASE["Heading2"], fontName="JP", fontSize=15,
        leading=22, spaceBefore=14, spaceAfter=8, textColor=HexColor("#2c5282"),
    ),
    "h3": ParagraphStyle(
        "h3", parent=BASE["Heading3"], fontName="JP", fontSize=12,
        leading=18, spaceBefore=10, spaceAfter=6, textColor=HexColor("#2b6cb0"),
    ),
    "body": ParagraphStyle(
        "body", parent=BASE["BodyText"], fontName="JP-P", fontSize=10,
        leading=16, spaceAfter=6, textColor=black,
    ),
    "bullet": ParagraphStyle(
        "bullet", parent=BASE["BodyText"], fontName="JP-P", fontSize=10,
        leading=15, leftIndent=14, bulletIndent=2, spaceAfter=3,
    ),
    "quote": ParagraphStyle(
        "quote", parent=BASE["BodyText"], fontName="JP-P", fontSize=10,
        leading=15, leftIndent=12, textColor=HexColor("#4a5568"),
        borderColor=HexColor("#cbd5e0"), borderWidth=0,
        spaceBefore=4, spaceAfter=6,
    ),
    "cell": ParagraphStyle(
        "cell", parent=BASE["BodyText"], fontName="JP-P", fontSize=9,
        leading=13, textColor=black,
    ),
    "cell_head": ParagraphStyle(
        "cell_head", parent=BASE["BodyText"], fontName="JP", fontSize=9,
        leading=13, textColor=white,
    ),
}


def inline(text: str) -> str:
    """Convert Markdown inline syntax to ReportLab mini-HTML."""
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"`([^`]+?)`",
        lambda m: f'<font face="Courier">{m.group(1)}</font>',
        text,
    )
    text = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        r'<link href="\2" color="blue"><u>\1</u></link>',
        text,
    )
    return text


def parse_table(lines):
    """Parse a Markdown table starting at lines[0]; returns (rows, consumed)."""
    rows = []
    consumed = 0
    for line in lines:
        if not line.strip().startswith("|"):
            break
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        rows.append(cells)
        consumed += 1
    if len(rows) >= 2 and all(re.match(r"^:?-+:?$", c.replace(" ", "")) for c in rows[1]):
        del rows[1]
    return rows, consumed


def make_table(rows):
    """Build a reportlab Table from list-of-list rows."""
    if not rows:
        return None
    data = [
        [Paragraph(inline(cell), STYLES["cell_head" if i == 0 else "cell"]) for cell in row]
        for i, row in enumerate(rows)
    ]
    ncols = max(len(r) for r in data)
    for r in data:
        while len(r) < ncols:
            r.append(Paragraph("", STYLES["cell"]))

    page_w = A4[0] - 36 * mm
    col_w = page_w / ncols
    t = Table(data, colWidths=[col_w] * ncols, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#2c5282")),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("GRID", (0, 0), (-1, -1), 0.4, HexColor("#a0aec0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, HexColor("#f7fafc")]),
    ]))
    return t


def render(md_text: str):
    """Parse the Markdown text into a stream of reportlab flowables."""
    flowables = []
    lines = md_text.splitlines()
    i = 0
    pending_bullets = []

    def flush_bullets():
        nonlocal pending_bullets
        if pending_bullets:
            items = [
                ListItem(Paragraph(inline(b), STYLES["bullet"]), leftIndent=12)
                for b in pending_bullets
            ]
            flowables.append(ListFlowable(items, bulletType="bullet", start="•"))
            flowables.append(Spacer(1, 4))
            pending_bullets = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            flush_bullets()
            i += 1
            continue

        if stripped.startswith("# "):
            flush_bullets()
            flowables.append(Paragraph(inline(stripped[2:]), STYLES["h1"]))
            i += 1
            continue
        if stripped.startswith("## "):
            flush_bullets()
            flowables.append(Paragraph(inline(stripped[3:]), STYLES["h2"]))
            i += 1
            continue
        if stripped.startswith("### "):
            flush_bullets()
            flowables.append(Paragraph(inline(stripped[4:]), STYLES["h3"]))
            i += 1
            continue

        if stripped.startswith("---"):
            flush_bullets()
            flowables.append(Spacer(1, 4))
            flowables.append(HRFlowable(width="100%", thickness=0.6,
                                        color=HexColor("#cbd5e0")))
            flowables.append(Spacer(1, 6))
            i += 1
            continue

        if stripped.startswith("|"):
            flush_bullets()
            rows, consumed = parse_table(lines[i:])
            tbl = make_table(rows)
            if tbl is not None:
                flowables.append(tbl)
                flowables.append(Spacer(1, 8))
            i += consumed
            continue

        if stripped.startswith("> "):
            flush_bullets()
            flowables.append(Paragraph("▎ " + inline(stripped[2:]), STYLES["quote"]))
            i += 1
            continue

        m = re.match(r"^(\d+)\.\s+(.*)", stripped)
        if m:
            flush_bullets()
            flowables.append(Paragraph(f"{m.group(1)}. {inline(m.group(2))}",
                                       STYLES["bullet"]))
            i += 1
            continue

        if stripped.startswith("- "):
            pending_bullets.append(stripped[2:])
            i += 1
            continue

        flush_bullets()
        flowables.append(Paragraph(inline(stripped), STYLES["body"]))
        i += 1

    flush_bullets()
    return flowables


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("JP", 8)
    canvas.setFillColor(HexColor("#718096"))
    canvas.drawString(18 * mm, 10 * mm,
                      "塗料業界 調査レポート（需要側フォーカス）")
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm,
                           f"Page {doc.page}")
    canvas.restoreState()


def extract_title(md_text: str) -> str:
    for line in md_text.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return "Project_Py Report"


def make_header_footer(title: str):
    def _hf(canvas, doc):
        canvas.saveState()
        canvas.setFont("JP", 8)
        canvas.setFillColor(HexColor("#718096"))
        canvas.drawString(18 * mm, 10 * mm, title)
        canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {doc.page}")
        canvas.restoreState()
    return _hf


def main():
    if len(sys.argv) >= 2:
        src = Path(sys.argv[1])
    else:
        src = DEFAULT_SRC
    if len(sys.argv) >= 3:
        dst = Path(sys.argv[2])
    else:
        dst = src.with_suffix(".pdf")

    md_text = src.read_text(encoding="utf-8")
    title = extract_title(md_text)
    flowables = render(md_text)

    doc = SimpleDocTemplate(
        str(dst),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=title,
        author="Project_Py",
    )
    hf = make_header_footer(title)
    doc.build(flowables, onFirstPage=hf, onLaterPages=hf)
    print(f"Wrote: {dst}")


if __name__ == "__main__":
    main()
