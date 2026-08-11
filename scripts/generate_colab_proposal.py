#!/usr/bin/env python3
"""CoLab Point proposal — full content, small side icons (not large photos)."""

import io
import urllib.request
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
    PageBreak, Table, TableStyle, ListFlowable, ListItem,
)

ASSETS = Path("/workspace/assets/proposal-icons")
OUT_DOCX = "/workspace/CoLab_Space_Point_Proposal_Photos.docx"
OUT_PDF = "/workspace/CoLab_Space_Point_Proposal_Photos.pdf"
OUT_DOCX_ALT = "/workspace/CoLab_Space_Point_Digital_Agency_Proposal.docx"

BLACK = RGBColor(0, 0, 0)
TEAL = RGBColor(0x06, 0xAC, 0xBA)
NAVY = RGBColor(0x04, 0x24, 0x3C)
FONT = "Arial"


def ascii_safe(text: str) -> str:
    return text.replace("\u2014", "-").replace("\u2013", "-")


# Shared proposal content (Word + PDF)
WEBSITE_PACKAGES = [
    {
        "title": "BASIC WEBSITE",
        "price": "50,000",
        "icon": "basic",
        "audience": "Small businesses and startups",
        "items": [
            "Professional business website",
            "Up to 5 pages",
            "Responsive mobile design",
            "WordPress CMS",
            "Custom UI",
            "Contact form",
            "WhatsApp integration",
            "Social media profile links",
            "Google Analytics setup",
            "SSL configuration",
            "30 days support",
        ],
        "note": "IMPORTANT: SEO and e-commerce are NOT included in this package.",
    },
    {
        "title": "STANDARD WEBSITE",
        "price": "80,000",
        "icon": "standard",
        "audience": "Growing brands and online sellers",
        "items": [
            "Everything in Basic Website, plus:",
            "Up to 10 pages",
            "Premium UI/UX design",
            "Blog section",
            "WooCommerce online store",
            "Product upload (initial batch)",
            "Payment gateway integration",
            "Google Search Console setup",
            "Facebook Pixel",
            "Advanced SEO",
            "Speed optimization",
            "60 days support",
        ],
        "note": None,
    },
    {
        "title": "PREMIUM WEBSITE",
        "price": "120,000",
        "icon": "premium",
        "audience": "Established businesses and enterprises",
        "items": [
            "Everything in Standard Website, plus:",
            "Unlimited pages (agreed scope)",
            "Fully custom design",
            "Advanced WooCommerce setup",
            "Unlimited product catalog setup",
            "CRM integration",
            "Booking / appointment system",
            "API integrations",
            "Premium security and performance optimization",
            "Admin training",
            "90 days priority support",
        ],
        "note": None,
    },
]

MARKETING_PACKAGES = [
    ("BASIC", "15,000", ["Meta Ads", "Audience targeting", "Campaign optimization", "Monthly reporting"]),
    (
        "STANDARD",
        "30,000",
        [
            "Google Ads + Meta Ads",
            "Lead generation",
            "Conversion tracking",
            "Landing page recommendations",
            "Monthly reports",
        ],
    ),
    (
        "PREMIUM",
        "50,000",
        [
            "Google Ads + Meta Ads",
            "SEO",
            "Remarketing",
            "Conversion optimization",
            "Marketing strategy",
            "Weekly meetings",
            "Detailed reports",
        ],
    ),
]

COMPANY = "CoLab Space Point"
WEB = "www.colabpoint.com"
PHONE = "+92 349 7684322  |  +92 478 986460"
EMAIL = "colabpoint@gmail.com  |  hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"

LOGO_URL = "https://colabpoint.com/wp-content/uploads/2024/05/Web-Logo-II.png"

# Small content-related icons (~96px)
ICONS = {
    "logo": LOGO_URL,
    "company": "https://img.icons8.com/fluency/96/company.png",
    "website": "https://img.icons8.com/fluency/96/domain.png",
    "basic": "https://img.icons8.com/fluency/96/home-page.png",
    "standard": "https://img.icons8.com/fluency/96/shopping-cart.png",
    "premium": "https://img.icons8.com/fluency/96/crown.png",
    "marketing": "https://img.icons8.com/fluency/96/combo-chart.png",
    "contact": "https://img.icons8.com/fluency/96/phone.png",
}


def fetch_icon(key: str) -> Path | None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    dest = ASSETS / f"{key}.png"
    if dest.exists() and dest.stat().st_size > 500:
        return dest
    url = ICONS.get(key)
    if not url:
        return None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        dest.write_bytes(urllib.request.urlopen(req, timeout=20).read())
        return dest
    except Exception:
        return None


def fetch_all_icons() -> dict:
    return {k: fetch_icon(k) for k in ICONS}


def force_white_page(doc):
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), "FFFFFF")
    doc.element.insert(0, bg)


def write_run(p, text, size=12, bold=False, color=BLACK):
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = FONT
    r.font.color.rgb = color


def line(doc, text="", size=12, bold=False, color=BLACK, center=False, space=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space)
    p.paragraph_format.space_before = Pt(2)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text:
        write_run(p, text, size, bold, color)


def section_with_icon(doc, title: str, icon_path: Path | None, subtitle=""):
    """Section heading with small icon on the left."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.columns[0].width = Cm(1.6)
    table.columns[1].width = Cm(15)
    left, right = table.rows[0].cells[0], table.rows[0].cells[1]
    left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if icon_path and icon_path.exists():
        left.paragraphs[0].add_run().add_picture(str(icon_path), width=Inches(0.55))
    p = right.paragraphs[0]
    write_run(p, title, 18, True, BLACK)
    if subtitle:
        p2 = right.add_paragraph()
        write_run(p2, subtitle, 11, False, NAVY)
    doc.add_paragraph()


def bullets(doc, items):
    for item in items:
        line(doc, f"  •  {item}", 12, False, BLACK, space=3)


def package_page(doc, title, price, audience, items, note, icon_path):
    doc.add_page_break()
    section_with_icon(doc, title, icon_path, f"Package Price: PKR {price}")
    line(doc, f"Best for: {audience}", 12, True, BLACK, space=8)
    line(doc, "WHAT IS INCLUDED IN THIS PRICE:", 13, True, TEAL, space=6)
    bullets(doc, items)
    if note:
        line(doc, note, 12, True, NAVY, space=10)


def build_docx(icons: dict):
    doc = Document()
    force_white_page(doc)
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(12)
    normal.font.color.rgb = BLACK

    # Cover — logo only, no big banner
    if icons.get("logo") and icons["logo"].exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(icons["logo"]), width=Inches(1.4))
    line(doc, "CoLab Point", 14, True, TEAL, center=True, space=4)
    line(doc, COMPANY, 26, True, BLACK, center=True, space=6)
    line(doc, "Digital Agency Proposal", 16, True, BLACK, center=True, space=10)
    line(doc, "Website Designing & Development  |  Digital Marketing", 12, False, BLACK, center=True)
    line(doc, WEB, 12, True, NAVY, center=True, space=8)
    line(doc, PHONE, 11, False, BLACK, center=True, space=3)
    line(doc, EMAIL, 11, False, BLACK, center=True, space=3)
    line(doc, ADDRESS, 11, False, BLACK, center=True, space=3)

    doc.add_page_break()
    section_with_icon(
        doc,
        "Company Profile",
        icons.get("company"),
        "Digital services from CoLab Point, Gujrat",
    )
    line(
        doc,
        f"{COMPANY} is the digital services arm of CoLab Point, Gujrat. We build websites, "
        "e-commerce stores, and run performance marketing campaigns.",
        12,
        space=8,
    )
    line(
        doc,
        "WordPress is our primary development platform. Custom solutions are available on request.",
        12,
        space=10,
    )

    section_with_icon(doc, "Website Designing & Development", icons.get("website"))

    for pkg in WEBSITE_PACKAGES:
        package_page(
            doc,
            f"{pkg['title']} — PKR {pkg['price']}",
            pkg["price"],
            pkg["audience"],
            pkg["items"],
            pkg["note"],
            icons.get(pkg["icon"]),
        )

    doc.add_page_break()
    section_with_icon(doc, "Digital Marketing", icons.get("marketing"), "Monthly packages")

    for name, price, items in [
        (f"{n} — PKR {p} / month", p, items) for n, p, items in MARKETING_PACKAGES
    ]:
        line(doc, name, 14, True, BLACK, space=8)
        line(doc, "What is included:", 12, True, TEAL, space=4)
        bullets(doc, items)
        line(doc, "", space=6)

    doc.add_page_break()
    section_with_icon(doc, "Contact", icons.get("contact"))
    line(doc, f"Company: {COMPANY}", 12, space=4)
    line(doc, f"Website: {WEB}", 12, space=4)
    line(doc, f"Phone: {PHONE}", 12, space=4)
    line(doc, f"Email: {EMAIL}", 12, space=4)
    line(doc, f"Address: {ADDRESS}", 12, space=4)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    write_run(fp, f"{COMPANY} | {WEB}", 9, False, BLACK)

    doc.save(OUT_DOCX)
    doc.save(OUT_DOCX_ALT)
    print(f"Word: {OUT_DOCX} ({Path(OUT_DOCX).stat().st_size // 1024} KB)")


def _pdf_styles():
    return {
        "cover_brand": ParagraphStyle(
            "cover_brand", fontName="Helvetica", fontSize=11, textColor=HexColor("#06ACBA"), alignment=TA_CENTER,
        ),
        "cover_title": ParagraphStyle(
            "cover_title", fontName="Helvetica-Bold", fontSize=24, textColor=black, alignment=TA_CENTER, spaceAfter=8,
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub", fontName="Helvetica-Bold", fontSize=14, textColor=black, alignment=TA_CENTER, spaceAfter=12,
        ),
        "cover_body": ParagraphStyle(
            "cover_body", fontName="Helvetica", fontSize=11, textColor=black, alignment=TA_CENTER, leading=16,
        ),
        "section": ParagraphStyle(
            "section", fontName="Helvetica-Bold", fontSize=16, textColor=black, spaceBefore=6, spaceAfter=4,
        ),
        "subtitle": ParagraphStyle(
            "subtitle", fontName="Helvetica", fontSize=10, textColor=HexColor("#04243C"), spaceAfter=8,
        ),
        "body": ParagraphStyle(
            "body", fontName="Helvetica", fontSize=11, textColor=black, leading=15, spaceAfter=6,
        ),
        "bold": ParagraphStyle(
            "bold", fontName="Helvetica-Bold", fontSize=11, textColor=black, leading=15, spaceAfter=4,
        ),
        "teal": ParagraphStyle(
            "teal", fontName="Helvetica-Bold", fontSize=11, textColor=HexColor("#06ACBA"), spaceBefore=6, spaceAfter=4,
        ),
        "note": ParagraphStyle(
            "note", fontName="Helvetica-Bold", fontSize=10, textColor=HexColor("#04243C"), spaceAfter=8,
        ),
        "pkg_title": ParagraphStyle(
            "pkg_title", fontName="Helvetica-Bold", fontSize=14, textColor=black, spaceAfter=4,
        ),
        "bullet": ParagraphStyle(
            "bullet", fontName="Helvetica", fontSize=11, textColor=black, leftIndent=18, bulletIndent=8, leading=14,
        ),
    }


def _pdf_icon(path: Path | None, size=0.45):
    if path and path.exists():
        return RLImage(str(path), width=size * cm, height=size * cm)
    return Spacer(size * cm, size * cm)


def _pdf_section_row(icon_path, title, subtitle="", styles=None):
    styles = styles or _pdf_styles()
    icon = _pdf_icon(icon_path)
    text_bits = [Paragraph(ascii_safe(title), styles["section"])]
    if subtitle:
        text_bits.append(Paragraph(ascii_safe(subtitle), styles["subtitle"]))
    inner = Table([[text_bits]], colWidths=[14 * cm])
    inner.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    row = Table([[icon, inner]], colWidths=[1.2 * cm, 14.5 * cm])
    row.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ])
    )
    return row


def _pdf_bullets(items, styles):
    return ListFlowable(
        [ListItem(Paragraph(ascii_safe(i), styles["bullet"])) for i in items],
        bulletType="bullet",
        start="•",
        leftIndent=12,
    )


def _pdf_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(HexColor("#06ACBA"))
    canvas.setLineWidth(0.5)
    canvas.line(2 * cm, 1.6 * cm, A4[0] - 2 * cm, 1.6 * cm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(HexColor("#666666"))
    canvas.drawString(2 * cm, 1 * cm, f"{COMPANY}  |  {WEB}")
    canvas.drawRightString(A4[0] - 2 * cm, 1 * cm, f"Page {doc.page}")
    canvas.restoreState()


def build_pdf(icons: dict):
    styles = _pdf_styles()
    story = []

    # Cover
    if icons.get("logo") and icons["logo"].exists():
        logo = RLImage(str(icons["logo"]), width=3.2 * cm, height=3.2 * cm, kind="proportional")
        t = Table([[logo]], colWidths=[17 * cm])
        t.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))
        story.append(t)
        story.append(Spacer(1, 0.4 * cm))

    story += [
        Paragraph("CoLab Point", styles["cover_brand"]),
        Paragraph(COMPANY, styles["cover_title"]),
        Paragraph("Digital Agency Proposal", styles["cover_sub"]),
        Paragraph("Website Designing &amp; Development  |  Digital Marketing", styles["cover_body"]),
        Spacer(1, 0.3 * cm),
        Paragraph(WEB, styles["cover_body"]),
        Paragraph(PHONE, styles["cover_body"]),
        Paragraph(EMAIL, styles["cover_body"]),
        Paragraph(ADDRESS, styles["cover_body"]),
        PageBreak(),
    ]

    # Company profile
    story.append(_pdf_section_row(icons.get("company"), "Company Profile", "Digital services from CoLab Point, Gujrat"))
    story += [
        Paragraph(
            ascii_safe(
                f"{COMPANY} is the digital services arm of CoLab Point, Gujrat. We build websites, "
                "e-commerce stores, and run performance marketing campaigns."
            ),
            styles["body"],
        ),
        Paragraph(
            "WordPress is our primary development platform. Custom solutions are available on request.",
            styles["body"],
        ),
        Spacer(1, 0.3 * cm),
        _pdf_section_row(icons.get("website"), "Website Designing & Development"),
        Spacer(1, 0.2 * cm),
    ]

    for pkg in WEBSITE_PACKAGES:
        story.append(PageBreak())
        story.append(
            _pdf_section_row(
                icons.get(pkg["icon"]),
                f"{pkg['title']} - PKR {pkg['price']}",
                f"Package Price: PKR {pkg['price']}",
            )
        )
        story.append(Paragraph(f"Best for: {pkg['audience']}", styles["bold"]))
        story.append(Paragraph("WHAT IS INCLUDED IN THIS PRICE:", styles["teal"]))
        story.append(_pdf_bullets(pkg["items"], styles))
        if pkg["note"]:
            story.append(Spacer(1, 0.15 * cm))
            story.append(Paragraph(ascii_safe(pkg["note"]), styles["note"]))

    story.append(PageBreak())
    story.append(_pdf_section_row(icons.get("marketing"), "Digital Marketing", "Monthly packages"))
    story.append(Spacer(1, 0.2 * cm))

    for name, price, items in MARKETING_PACKAGES:
        story.append(Paragraph(f"{name} - PKR {price} / month", styles["pkg_title"]))
        story.append(Paragraph("What is included:", styles["teal"]))
        story.append(_pdf_bullets(items, styles))
        story.append(Spacer(1, 0.25 * cm))

    story.append(PageBreak())
    story.append(_pdf_section_row(icons.get("contact"), "Contact"))
    for row in [
        f"Company: {COMPANY}",
        f"Website: {WEB}",
        f"Phone: {PHONE}",
        f"Email: {EMAIL}",
        f"Address: {ADDRESS}",
    ]:
        story.append(Paragraph(ascii_safe(row), styles["body"]))

    doc = SimpleDocTemplate(
        OUT_PDF,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2.2 * cm,
        title=f"{COMPANY} Proposal",
        author=COMPANY,
    )
    doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
    print(f"PDF: {OUT_PDF} ({Path(OUT_PDF).stat().st_size // 1024} KB)")


def main():
    icons = fetch_all_icons()
    print("Icons:", sum(1 for v in icons.values() if v), "/", len(icons))
    build_docx(icons)
    build_pdf(icons)


if __name__ == "__main__":
    main()
