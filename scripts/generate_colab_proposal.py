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

COMPANY = "CoLab Space Point"
WEB = "www.colabpoint.com"
PHONE = "+92 349 7684322  |  +92 478 986460"
EMAIL = "colabpoint@gmail.com  |  hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"

COMPANY_STATS = [
    ("Est. 2021", "CoLab Point ecosystem"),
    ("Location", "Gujrat, Pakistan"),
    ("Platform", "WordPress & WooCommerce"),
    ("Focus", "Websites + Digital Marketing"),
]

COMPANY_SERVICES = [
    "Website Designing & Development",
    "WordPress Development",
    "E-commerce Solutions",
    "Branding & Visual Identity",
    "Digital Marketing",
    "SEO (Search Engine Optimization)",
    "Google Ads",
    "Meta Ads (Facebook & Instagram)",
    "Business Growth Strategy",
]

WHY_CHOOSE = [
    "Experienced multidisciplinary team",
    "Professional support and clear communication",
    "Business-focused, results-driven solutions",
    "Modern technology stack",
    "Creative design and performance marketing",
    "Transparent pricing and long-term partnership",
]

COMPANY_ABOUT = (
    f"{COMPANY} is the digital services division of CoLab Point — a trusted innovation and "
    "coworking hub in Gujrat since 2021. We help businesses establish a strong online presence, "
    "sell through e-commerce, and acquire customers through data-driven marketing."
)

COMPANY_MISSION = (
    "Our mission is to deliver premium digital experiences that convert visitors into customers — "
    "with transparent pricing, professional delivery, and ongoing support after every launch."
)

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


def shade_cell(cell, hex_color: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def add_profile_header_docx(doc, icons):
    """Styled company profile page header."""
    bar = doc.add_table(rows=1, cols=1)
    c = bar.rows[0].cells[0]
    shade_cell(c, "04243C")
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    write_run(p, "COMPANY PROFILE", 20, True, RGBColor(0xFF, 0xFF, 0xFF))
    p2 = c.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    write_run(p2, "CoLab Space Point  |  Digital Agency  |  Gujrat", 11, False, TEAL)
    doc.add_paragraph()

    intro = doc.add_table(rows=1, cols=2)
    intro.columns[0].width = Cm(2.2)
    intro.columns[1].width = Cm(14)
    lc, rc = intro.rows[0].cells[0], intro.rows[0].cells[1]
    shade_cell(rc, "E8F7F9")
    if icons.get("logo") and icons["logo"].exists():
        lc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        lc.paragraphs[0].add_run().add_picture(str(icons["logo"]), width=Inches(0.9))
    rp = rc.paragraphs[0]
    write_run(rp, "About Us", 14, True, NAVY)
    rp2 = rc.add_paragraph()
    write_run(rp2, COMPANY_ABOUT, 11, False, BLACK)
    rp3 = rc.add_paragraph()
    write_run(rp3, COMPANY_MISSION, 11, False, BLACK)
    doc.add_paragraph()


def add_stats_row_docx(doc):
    t = doc.add_table(rows=2, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (label, val) in enumerate(COMPANY_STATS):
        shade_cell(t.rows[0].cells[i], "06ACBA")
        shade_cell(t.rows[1].cells[i], "F4FBFC")
        t.rows[0].cells[i].text = ""
        t.rows[1].cells[i].text = ""
        p0 = t.rows[0].cells[i].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        write_run(p0, label, 11, True, RGBColor(0xFF, 0xFF, 0xFF))
        p1 = t.rows[1].cells[i].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        write_run(p1, val, 10, False, BLACK)
    doc.add_paragraph()


def add_services_grid_docx(doc, icons):
    line(doc, "Our Core Services", 14, True, NAVY, space=8)
    half = (len(COMPANY_SERVICES) + 1) // 2
    t = doc.add_table(rows=half, cols=2)
    for i in range(half):
        for j in range(2):
            idx = i + j * half
            cell = t.rows[i].cells[j]
            shade_cell(cell, "FFFFFF")
            cell.text = ""
            if idx < len(COMPANY_SERVICES):
                write_run(cell.paragraphs[0], f"  ✓  {COMPANY_SERVICES[idx]}", 11, False, BLACK)
    doc.add_paragraph()


def add_why_choose_docx(doc):
    box = doc.add_table(rows=1, cols=1)
    c = box.rows[0].cells[0]
    shade_cell(c, "E8F7F9")
    p = c.paragraphs[0]
    write_run(p, "Why Choose CoLab Space Point", 13, True, NAVY)
    for w in WHY_CHOOSE:
        bp = c.add_paragraph()
        write_run(bp, f"  •  {w}", 11, False, BLACK)
    doc.add_paragraph()


def add_company_profile_docx(doc, icons):
    add_profile_header_docx(doc, icons)
    add_stats_row_docx(doc)
    add_services_grid_docx(doc, icons)
    add_why_choose_docx(doc)
    line(
        doc,
        "WordPress is our primary development platform. Fully custom solutions are available on request.",
        11,
        False,
        BLACK,
        space=8,
    )


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
    add_company_profile_docx(doc, icons)

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


def _pdf_company_profile(story, icons, styles):
    # Header bar
    hdr = Table(
        [[Paragraph("COMPANY PROFILE", ParagraphStyle(
            "ph", fontName="Helvetica-Bold", fontSize=18, textColor=white, alignment=TA_CENTER,
        ))],
         [Paragraph("CoLab Space Point  |  Digital Agency  |  Gujrat", ParagraphStyle(
            "ps", fontName="Helvetica", fontSize=10, textColor=HexColor("#B8E8EE"), alignment=TA_CENTER,
        ))]],
        colWidths=[17 * cm],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#04243C")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#06ACBA")),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.35 * cm))

    # About box with logo
    logo_cell = Spacer(1.8 * cm, 1.8 * cm)
    if icons.get("logo") and icons["logo"].exists():
        logo_cell = RLImage(str(icons["logo"]), width=1.8 * cm, height=1.8 * cm, kind="proportional")
    about_text = [
        Paragraph("<b>About Us</b>", styles["bold"]),
        Paragraph(ascii_safe(COMPANY_ABOUT), styles["body"]),
        Paragraph(ascii_safe(COMPANY_MISSION), styles["body"]),
    ]
    about_inner = Table([[about_text]], colWidths=[13.5 * cm])
    about_inner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#E8F7F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#06ACBA")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    about_row = Table([[logo_cell, about_inner]], colWidths=[2.2 * cm, 14 * cm])
    about_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(about_row)
    story.append(Spacer(1, 0.35 * cm))

    # Stats row
    stat_hdr = [Paragraph(f"<b>{ascii_safe(l)}</b>", ParagraphStyle(
        "sh", fontName="Helvetica-Bold", fontSize=9, textColor=white, alignment=TA_CENTER,
    )) for l, _ in COMPANY_STATS]
    stat_val = [Paragraph(ascii_safe(v), ParagraphStyle(
        "sv", fontName="Helvetica", fontSize=9, textColor=black, alignment=TA_CENTER,
    )) for _, v in COMPANY_STATS]
    stats = Table([stat_hdr, stat_val], colWidths=[4.25 * cm] * 4)
    stats.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#06ACBA")),
        ("BACKGROUND", (0, 1), (-1, 1), HexColor("#F4FBFC")),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#06ACBA")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#06ACBA")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(stats)
    story.append(Spacer(1, 0.35 * cm))

    # Services
    story.append(Paragraph("Our Core Services", styles["bold"]))
    half = (len(COMPANY_SERVICES) + 1) // 2
    rows = []
    for i in range(half):
        left = COMPANY_SERVICES[i] if i < len(COMPANY_SERVICES) else ""
        right = COMPANY_SERVICES[i + half] if i + half < len(COMPANY_SERVICES) else ""
        rows.append([
            Paragraph(ascii_safe(f"✓  {left}") if left else "", styles["body"]),
            Paragraph(ascii_safe(f"✓  {right}") if right else "", styles["body"]),
        ])
    svc = Table(rows, colWidths=[8.5 * cm, 8.5 * cm])
    svc.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#CCCCCC")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#EEEEEE")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(svc)
    story.append(Spacer(1, 0.3 * cm))

    # Why choose box
    why_content = [Paragraph("<b>Why Choose CoLab Space Point</b>", styles["bold"])]
    why_content += [Paragraph(ascii_safe(f"•  {w}"), styles["body"]) for w in WHY_CHOOSE]
    why = Table([[why_content]], colWidths=[17 * cm])
    why.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#E8F7F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#06ACBA")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(why)
    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph(
        "WordPress is our primary development platform. Fully custom solutions are available on request.",
        styles["body"],
    ))


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

    # Company profile (dedicated designed page)
    _pdf_company_profile(story, icons, styles)
    story.append(PageBreak())
    story.append(_pdf_section_row(icons.get("website"), "Website Designing & Development"))
    story.append(Spacer(1, 0.2 * cm))

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
