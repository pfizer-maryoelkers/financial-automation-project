"""Build the Pfizer-branded PDF version of each guide in docs/ from its .md file.

Run after editing a guide, then commit the refreshed PDFs:

    pip install reportlab
    python tools/build_docs_pdf.py [path to pfizer-brand-pdf skill folder]

The Pfizer fonts are taken from the pfizer-brand-pdf skill's assets/fonts folder
(argument, or the PFIZER_PDF_SKILL environment variable); without them the PDFs
fall back to Helvetica.  Logos come from static/.  The app itself only serves the
finished PDFs, so reportlab is not needed to run it.
"""
import html
import os
import re
import sys
from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, KeepTogether, ListFlowable, ListItem, PageTemplate,
    Paragraph, Preformatted, Spacer, Table, TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
LOGO_WHITE = ROOT / "static" / "Pfizer_Logo_White_RGB.png"
LOGO_ASPECT = 6276 / 2591

BLUE_70, BLUE_90 = HexColor("#0000C9"), HexColor("#000067")
CYAN_50, BLUE_05, BLUE_03 = HexColor("#0095FF"), HexColor("#EEF5FB"), HexColor("#ECF6FF")
GRAY_90, GRAY_50 = HexColor("#1A1A1A"), HexColor("#767676")

GUIDES = (
    ("user_guide", "User Guide"),
    ("technical_guide", "Technical Guide"),
    ("governance", "Governance"),
)


def register_fonts(skill_dir: str | None) -> tuple[str, str, str]:
    """Register the Pfizer fonts if available; return (body, body bold, heading) names."""
    font_dir = Path(skill_dir) / "assets" / "fonts" if skill_dir else None
    if font_dir and font_dir.exists():
        for name, filename in (
            ("Diatype", "PfizerDiatype-Regular.ttf"), ("Diatype-Medium", "PfizerDiatype-Medium.ttf"),
            ("Diatype-Bold", "PfizerDiatype-Bold.ttf"), ("PfizerTomorrow-Bold", "PfizerTomorrow-Bold.ttf"),
        ):
            pdfmetrics.registerFont(TTFont(name, str(font_dir / filename)))
        pdfmetrics.registerFontFamily("Diatype", normal="Diatype", bold="Diatype-Bold",
                                      italic="Diatype-Medium", boldItalic="Diatype-Bold")
        return "Diatype", "Diatype-Bold", "PfizerTomorrow-Bold"
    print("Pfizer fonts not found - falling back to Helvetica")
    return "Helvetica", "Helvetica-Bold", "Helvetica-Bold"


def make_styles(body: str, bold: str, head: str) -> dict:
    return {
        "eyebrow": ParagraphStyle("eyebrow", fontName=bold, fontSize=9, textColor=CYAN_50, spaceAfter=4),
        "title": ParagraphStyle("title", fontName=head, fontSize=24, leading=28, textColor=BLUE_70, spaceAfter=8),
        "h2": ParagraphStyle("h2", fontName=head, fontSize=16, leading=20, textColor=BLUE_70, spaceBefore=16, spaceAfter=6),
        "h3": ParagraphStyle("h3", fontName=head, fontSize=12.5, leading=16, textColor=BLUE_70, spaceBefore=10, spaceAfter=4),
        "body": ParagraphStyle("body", fontName=body, fontSize=10, leading=14.5, textColor=GRAY_90, spaceAfter=6, alignment=TA_LEFT),
        "cell": ParagraphStyle("cell", fontName=body, fontSize=8.5, leading=11.5, textColor=GRAY_90),
        "cellh": ParagraphStyle("cellh", fontName=bold, fontSize=8.5, leading=11.5, textColor=white),
        "code": ParagraphStyle("code", fontName="Courier", fontSize=7.8, leading=10, textColor=GRAY_90,
                               backColor=BLUE_05, borderPadding=6, spaceBefore=4, spaceAfter=10),
    }


def inline(text: str) -> str:
    """Markdown inline syntax -> ReportLab paragraph markup."""
    text = text.replace("\\*", "\u0001")
    text = html.escape(text, quote=False)
    codes: list[str] = []
    text = re.sub(r"`([^`]+)`", lambda m: codes.append(m.group(1)) or f"\u0002{len(codes) - 1}\u0002", text)
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)",
                  lambda m: f'<a href="{html.unescape(m.group(2))}" color="#0000C9"><u>{m.group(1)}</u></a>', text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<i>\1</i>", text)
    text = re.sub(r"\u0002(\d+)\u0002", lambda m: f'<font face="Courier" size="8.5">{codes[int(m.group(1))]}</font>', text)
    return text.replace("\u0001", "*")


def build_table(rows: list[list[str]], st: dict, width: float) -> Table:
    header, body_rows = rows[0], rows[1:]
    ncols = len(header)
    lengths = [max(len(re.sub(r"[`*\[\]]", "", r[i])) if i < len(r) else 0 for r in rows) for i in range(ncols)]
    weights = [min(max(n, 10), 60) for n in lengths]
    colw = [width * w / sum(weights) for w in weights]
    data = [[Paragraph(inline(c) or "&nbsp;", st["cellh"]) for c in header]]
    for r in body_rows:
        r = (r + [""] * ncols)[:ncols]
        data.append([Paragraph(inline(c), st["cell"]) for c in r])
    t = Table(data, colWidths=colw, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLUE_70),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, HexColor("#C9D6E8")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(2, len(data), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), BLUE_03))
    t.setStyle(TableStyle(style))
    return t


def parse(md: str, st: dict, width: float) -> list:
    lines = md.splitlines()
    flow, para, i = [], [], 0

    def flush():
        if para:
            flow.append(Paragraph(inline(" ".join(para)), st["body"]))
            para.clear()

    first_text = True
    while i < len(lines):
        line = lines[i]
        s = line.strip()
        if s.startswith("```"):
            flush()
            i += 1
            block = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                block.append(lines[i])
                i += 1
            flow.append(Preformatted("\n".join(block), st["code"]))
        elif s.startswith("|"):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            flow.append(build_table(rows, st, width))
            flow.append(Spacer(1, 8))
            continue
        elif re.match(r"^(#{1,3}) ", s):
            flush()
            level = len(s.split(" ")[0])
            flow.append(Paragraph(inline(s[level + 1:]), st["title" if level == 1 else "h2" if level == 2 else "h3"]))
        elif re.match(r"^(- |\d+\. )", s):
            flush()
            ordered = bool(re.match(r"^\d+\. ", s))
            items = []
            while i < len(lines) and re.match(r"^(- |\d+\. )", lines[i].strip()):
                items.append(ListItem(Paragraph(inline(re.sub(r"^(- |\d+\. )", "", lines[i].strip())), st["body"]),
                                      leftIndent=16))
                i += 1
            flow.append(ListFlowable(items, bulletType="1" if ordered else "bullet",
                                     bulletFontName=st["body"].fontName, bulletFontSize=9, leftIndent=16,
                                     bulletColor=BLUE_70))
            flow.append(Spacer(1, 4))
            continue
        elif not s:
            flush()
        else:
            if first_text:  # the line above the title ("User Guide") is an eyebrow label
                flow.append(Paragraph(inline(s).upper(), st["eyebrow"]))
                first_text = False
            else:
                para.append(s)
        if s:
            first_text = False
        i += 1
    flush()
    return flow


def build(stem: str, title: str, fonts: tuple[str, str, str]) -> Path:
    body, bold, head = fonts
    st = make_styles(body, bold, head)
    out = DOCS / f"{stem}.pdf"
    page_w, page_h = letter
    margin, bar_h = 0.8 * inch, 0.55 * inch

    def chrome(c, doc):
        c.saveState()
        c.setFillColor(BLUE_70)
        c.rect(0, page_h - bar_h, page_w, bar_h, stroke=0, fill=1)
        if LOGO_WHITE.exists():
            lw = 1.05 * inch
            lh = lw / LOGO_ASPECT
            c.drawImage(ImageReader(str(LOGO_WHITE)), margin, page_h - bar_h + (bar_h - lh) / 2,
                        width=lw, height=lh, mask="auto")
        c.setFillColor(BLUE_90)
        c.rect(0, 0, page_w, 0.45 * inch, stroke=0, fill=1)
        c.setFillColor(white)
        c.setFont(body, 8)
        c.drawString(margin, 0.17 * inch, f"Financial Automation - {title}")
        c.drawRightString(page_w - margin, 0.17 * inch, f"Page {doc.page}")
        c.restoreState()

    doc = BaseDocTemplate(str(out), pagesize=letter, title=f"Financial Automation - {title}",
                          author="Mary Oelkers", leftMargin=margin, rightMargin=margin,
                          topMargin=bar_h + 0.35 * inch, bottomMargin=0.45 * inch + 0.3 * inch)
    frame = Frame(margin, doc.bottomMargin, page_w - 2 * margin, page_h - doc.topMargin - doc.bottomMargin, id="f")
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=chrome)])
    md = (DOCS / f"{stem}.md").read_text(encoding="utf-8")
    doc.build(parse(md, st, page_w - 2 * margin))
    return out


if __name__ == "__main__":
    skill = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("PFIZER_PDF_SKILL")
    fonts = register_fonts(skill)
    for stem, title in GUIDES:
        print("built", build(stem, title, fonts))
