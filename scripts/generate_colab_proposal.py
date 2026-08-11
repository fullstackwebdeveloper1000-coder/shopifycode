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

# PDF layout constants (A4 with 2cm margins → 17cm content)
PDF_M = 2 * cm
PDF_W = A4[0] - 4 * cm  # 17cm

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
BANNER_URL = "https://colabpoint.com/wp-content/uploads/2024/06/popup-bg-1-1024x374-1.jpg"

# Small content-related icons (~96px)
ICONS = {
    "logo": LOGO_URL,
    "banner": BANNER_URL,
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
    ext = ".jpg" if key == "banner" else ".png"
    dest = ASSETS / f"{key}{ext}"
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


# ─── PDF (aligned layout, icons, banner) ───────────────────────────────────

def _pdf_styles():
    def S(name, **kw):
        defaults = dict(fontName="Helvetica", fontSize=11, textColor=black, leading=15)
        defaults.update(kw)
        return ParagraphStyle(name, **defaults)

    return {
        "cover_brand": S("cover_brand", fontSize=12, textColor=HexColor("#06ACBA"), alignment=TA_CENTER),
        "cover_title": S("cover_title", fontName="Helvetica-Bold", fontSize=26, alignment=TA_CENTER, spaceAfter=6, leading=30),
        "cover_sub": S("cover_sub", fontName="Helvetica-Bold", fontSize=15, alignment=TA_CENTER, spaceAfter=10),
        "cover_body": S("cover_body", fontSize=11, alignment=TA_CENTER, leading=17),
        "section": S("section", fontName="Helvetica-Bold", fontSize=15, textColor=HexColor("#04243C"), spaceAfter=2),
        "subtitle": S("subtitle", fontSize=10, textColor=HexColor("#06ACBA"), spaceAfter=6),
        "body": S("body", spaceAfter=4),
        "bold": S("bold", fontName="Helvetica-Bold", spaceAfter=4),
        "teal": S("teal", fontName="Helvetica-Bold", textColor=HexColor("#06ACBA"), spaceBefore=4, spaceAfter=4),
        "note": S("note", fontName="Helvetica-Bold", fontSize=10, textColor=HexColor("#04243C"), spaceAfter=6),
        "price": S("price", fontName="Helvetica-Bold", fontSize=16, textColor=HexColor("#04243C"), alignment=TA_CENTER),
        "pkg_name": S("pkg_name", fontName="Helvetica-Bold", fontSize=13, textColor=white),
        "center": S("center", alignment=TA_CENTER),
        "bullet": S("bullet", leftIndent=14, bulletIndent=6, leading=14, spaceAfter=2),
        "stat_label": S("stat_label", fontName="Helvetica-Bold", fontSize=9, textColor=white, alignment=TA_CENTER),
        "stat_val": S("stat_val", fontSize=9, alignment=TA_CENTER, leading=12),
    }


def _tbl_style_box(bg=None, border=HexColor("#06ACBA"), grid=False):
    cmd = [
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    if bg:
        cmd.append(("BACKGROUND", (0, 0), (-1, -1), bg))
    if grid:
        cmd.append(("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#D0E8EB")))
    return TableStyle(cmd)


def _pdf_img(path, w_cm, h_cm=None):
    if path and Path(path).exists():
        if h_cm:
            return RLImage(str(path), width=w_cm * cm, height=h_cm * cm, kind="proportional")
        return RLImage(str(path), width=w_cm * cm, height=w_cm * cm, kind="proportional")
    return Spacer(w_cm * cm, (h_cm or w_cm) * cm)


def _pdf_icon_cell(path, size=1.0):
    return _pdf_img(path, size, size)


def _pdf_bullets(items, styles):
    return ListFlowable(
        [ListItem(Paragraph(ascii_safe(i), styles["bullet"])) for i in items],
        bulletType="bullet", start="•", leftIndent=10,
    )


def _pdf_section_bar(title, subtitle="", icon_path=None, styles=None):
    """Navy section bar with optional icon — full content width."""
    styles = styles or _pdf_styles()
    icon_w = 1.3 * cm
    text_w = PDF_W - icon_w
    icon = _pdf_icon_cell(icon_path, 1.0) if icon_path else Spacer(1 * cm, 1 * cm)
    lines = [Paragraph(ascii_safe(title), styles["section"])]
    if subtitle:
        lines.append(Paragraph(ascii_safe(subtitle), styles["subtitle"]))
    inner = Table([[lines]], colWidths=[text_w])
    inner.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    row = Table([[icon, inner]], colWidths=[icon_w, text_w])
    row.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#E8F7F9")),
        ("BOX", (0, 0), (-1, -1), 0.6, HexColor("#06ACBA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (0, 0), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return row


def _pdf_page_decor(canvas, doc):
    canvas.saveState()
    # Top accent
    canvas.setFillColor(HexColor("#04243C"))
    canvas.rect(PDF_M, A4[1] - 1.1 * cm, PDF_W, 0.45 * cm, fill=1, stroke=0)
    canvas.setFillColor(HexColor("#06ACBA"))
    canvas.rect(PDF_M, A4[1] - 1.1 * cm, PDF_W, 0.08 * cm, fill=1, stroke=0)
    if doc.page > 1:
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(HexColor("#04243C"))
        canvas.drawString(PDF_M, A4[1] - 0.85 * cm, ascii_safe(COMPANY))
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(HexColor("#06ACBA"))
        canvas.drawRightString(A4[0] - PDF_M, A4[1] - 0.85 * cm, WEB)
    # Footer
    canvas.setStrokeColor(HexColor("#06ACBA"))
    canvas.setLineWidth(0.4)
    canvas.line(PDF_M, 1.55 * cm, A4[0] - PDF_M, 1.55 * cm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(HexColor("#888888"))
    canvas.drawString(PDF_M, 1.05 * cm, f"{COMPANY}  |  {WEB}")
    canvas.drawRightString(A4[0] - PDF_M, 1.05 * cm, f"Page {doc.page}")
    canvas.restoreState()


def _pdf_cover(story, icons, styles):
    banner = icons.get("banner")
    if banner and banner.exists():
        story.append(_pdf_img(banner, 17, 4.5))
        story.append(Spacer(1, 0.25 * cm))

    logo_row = Table(
        [[_pdf_img(icons.get("logo"), 3.5)]],
        colWidths=[PDF_W],
        rowHeights=[3.8 * cm],
    )
    logo_row.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story.append(logo_row)

    title_box = Table(
        [[Paragraph("CoLab Point", styles["cover_brand"])],
         [Paragraph(COMPANY, styles["cover_title"])],
         [Paragraph("Digital Agency Proposal", styles["cover_sub"])],
         [Paragraph("Website Designing &amp; Development  |  Digital Marketing", styles["cover_body"])]],
        colWidths=[PDF_W],
    )
    title_box.setStyle(_tbl_style_box(HexColor("#04243C"), HexColor("#06ACBA")))
    for i in range(4):
        title_box.setStyle(TableStyle([("TEXTCOLOR", (0, i), (-1, i), white if i > 0 else HexColor("#06ACBA"))]))
    story.append(title_box)
    story.append(Spacer(1, 0.35 * cm))

    contact_box = Table(
        [[Paragraph(WEB, styles["cover_body"])],
         [Paragraph(PHONE, styles["cover_body"])],
         [Paragraph(EMAIL, styles["cover_body"])],
         [Paragraph(ascii_safe(ADDRESS), styles["cover_body"])]],
        colWidths=[PDF_W],
    )
    contact_box.setStyle(_tbl_style_box(HexColor("#F4FBFC")))
    story.append(contact_box)


def _pdf_company_profile(story, icons, styles):
    # Page title
    hdr = Table(
        [[Paragraph("COMPANY PROFILE", ParagraphStyle(
            "cph", fontName="Helvetica-Bold", fontSize=18, textColor=white, alignment=TA_CENTER,
        ))],
         [Paragraph("CoLab Space Point  |  Digital Agency  |  Gujrat, Pakistan", ParagraphStyle(
            "cps", fontSize=10, textColor=HexColor("#B8E8EE"), alignment=TA_CENTER,
        ))]],
        colWidths=[PDF_W],
    )
    hdr.setStyle(_tbl_style_box(HexColor("#04243C")))
    story.append(hdr)
    story.append(Spacer(1, 0.3 * cm))

    # About + logo
    about_lines = [
        Paragraph("<b>About Us</b>", styles["bold"]),
        Paragraph(ascii_safe(COMPANY_ABOUT), styles["body"]),
        Paragraph(ascii_safe(COMPANY_MISSION), styles["body"]),
    ]
    about_tbl = Table(
        [[_pdf_icon_cell(icons.get("logo"), 1.6), about_lines]],
        colWidths=[2.0 * cm, PDF_W - 2.0 * cm],
    )
    about_tbl.setStyle(_tbl_style_box(HexColor("#E8F7F9")))
    story.append(about_tbl)
    story.append(Spacer(1, 0.3 * cm))

    # Stats 4-col aligned
    col_w = PDF_W / 4
    stat_tbl = Table(
        [
            [Paragraph(f"<b>{ascii_safe(l)}</b>", styles["stat_label"]) for l, _ in COMPANY_STATS],
            [Paragraph(ascii_safe(v), styles["stat_val"]) for _, v in COMPANY_STATS],
        ],
        colWidths=[col_w] * 4,
    )
    stat_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#06ACBA")),
        ("BACKGROUND", (0, 1), (-1, 1), HexColor("#F4FBFC")),
        ("BOX", (0, 0), (-1, -1), 0.6, HexColor("#06ACBA")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#06ACBA")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(stat_tbl)
    story.append(Spacer(1, 0.3 * cm))

    # Services grid with company icon header
    story.append(_pdf_section_bar("Our Core Services", icon_path=icons.get("company")))
    story.append(Spacer(1, 0.15 * cm))
    half = (len(COMPANY_SERVICES) + 1) // 2
    col_half = PDF_W / 2
    svc_rows = []
    for i in range(half):
        l = COMPANY_SERVICES[i] if i < len(COMPANY_SERVICES) else ""
        r = COMPANY_SERVICES[i + half] if i + half < len(COMPANY_SERVICES) else ""
        svc_rows.append([
            Paragraph(ascii_safe(f"✓  {l}") if l else "", styles["body"]),
            Paragraph(ascii_safe(f"✓  {r}") if r else "", styles["body"]),
        ])
    svc_tbl = Table(svc_rows, colWidths=[col_half, col_half])
    svc_tbl.setStyle(_tbl_style_box(grid=True))
    story.append(svc_tbl)
    story.append(Spacer(1, 0.25 * cm))

    # Why choose
    why_lines = [Paragraph("<b>Why Choose CoLab Space Point</b>", styles["bold"])]
    why_lines += [Paragraph(ascii_safe(f"•  {w}"), styles["body"]) for w in WHY_CHOOSE]
    why_tbl = Table([[why_lines]], colWidths=[PDF_W])
    why_tbl.setStyle(_tbl_style_box(HexColor("#E8F7F9")))
    story.append(why_tbl)


def _pdf_package_page(story, pkg, icons, styles):
    icon_path = icons.get(pkg["icon"])
    # Package header card
    hdr = Table(
        [[
            _pdf_icon_cell(icon_path, 1.2),
            [
                Paragraph(ascii_safe(pkg["title"]), styles["pkg_name"]),
                Paragraph(f"PKR {pkg['price']}", ParagraphStyle(
                    "pr", fontName="Helvetica-Bold", fontSize=14, textColor=HexColor("#06ACBA"),
                )),
            ],
        ]],
        colWidths=[1.5 * cm, PDF_W - 1.5 * cm],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#04243C")),
        ("TEXTCOLOR", (0, 0), (-1, -1), white),
        ("BOX", (0, 0), (-1, -1), 0.6, HexColor("#06ACBA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (0, 0), 8),
        ("LEFTPADDING", (1, 0), (1, 0), 4),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.2 * cm))

    body_tbl = Table(
        [[
            Paragraph(f"<b>Best for:</b> {ascii_safe(pkg['audience'])}", styles["body"]),
        ]],
        colWidths=[PDF_W],
    )
    body_tbl.setStyle(_tbl_style_box(HexColor("#F4FBFC")))
    story.append(body_tbl)
    story.append(Spacer(1, 0.15 * cm))

    incl = Table(
        [[
            Paragraph("WHAT IS INCLUDED IN THIS PRICE:", styles["teal"]),
            _pdf_bullets(pkg["items"], styles),
        ]],
        colWidths=[PDF_W],
    )
    incl.setStyle(_tbl_style_box(grid=True))
    story.append(incl)
    if pkg.get("note"):
        story.append(Spacer(1, 0.15 * cm))
        note = Table([[Paragraph(ascii_safe(pkg["note"]), styles["note"])]], colWidths=[PDF_W])
        note.setStyle(_tbl_style_box(HexColor("#FFF8E1"), HexColor("#F0AD4E")))
        story.append(note)


def _pdf_marketing_page(story, icons, styles):
    story.append(_pdf_section_bar("Digital Marketing", "Monthly packages", icons.get("marketing")))
    story.append(Spacer(1, 0.2 * cm))

    for name, price, items in MARKETING_PACKAGES:
        card = Table(
            [[
                _pdf_icon_cell(icons.get("marketing"), 0.9),
                [
                    Paragraph(ascii_safe(f"{name} - PKR {price} / month"), styles["bold"]),
                    Paragraph("What is included:", styles["teal"]),
                    _pdf_bullets(items, styles),
                ],
            ]],
            colWidths=[1.3 * cm, PDF_W - 1.3 * cm],
        )
        card.setStyle(_tbl_style_box(HexColor("#FAFEFF")))
        story.append(card)
        story.append(Spacer(1, 0.2 * cm))


def _pdf_contact_page(story, icons, styles):
    story.append(_pdf_section_bar("Contact", "Get in touch", icons.get("contact")))
    story.append(Spacer(1, 0.25 * cm))
    rows = [
        ("Company", COMPANY),
        ("Website", WEB),
        ("Phone", PHONE),
        ("Email", EMAIL),
        ("Address", ADDRESS),
    ]
    contact_data = []
    for label, val in rows:
        contact_data.append([
            Paragraph(f"<b>{label}</b>", styles["bold"]),
            Paragraph(ascii_safe(val), styles["body"]),
        ])
    ct = Table(contact_data, colWidths=[3.5 * cm, PDF_W - 3.5 * cm])
    ct.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), HexColor("#E8F7F9")),
        ("BOX", (0, 0), (-1, -1), 0.6, HexColor("#06ACBA")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#D0E8EB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(ct)


def build_pdf(icons: dict):
    styles = _pdf_styles()
    story = []

    _pdf_cover(story, icons, styles)
    story.append(PageBreak())

    _pdf_company_profile(story, icons, styles)
    story.append(PageBreak())

    story.append(_pdf_section_bar("Website Designing & Development", "WordPress primary platform", icons.get("website")))
    story.append(Spacer(1, 0.2 * cm))

    for i, pkg in enumerate(WEBSITE_PACKAGES):
        if i > 0:
            story.append(PageBreak())
        _pdf_package_page(story, pkg, icons, styles)

    story.append(PageBreak())
    _pdf_marketing_page(story, icons, styles)

    story.append(PageBreak())
    _pdf_contact_page(story, icons, styles)

    doc = SimpleDocTemplate(
        OUT_PDF,
        pagesize=A4,
        leftMargin=PDF_M,
        rightMargin=PDF_M,
        topMargin=2.3 * cm,
        bottomMargin=2.0 * cm,
        title=f"{COMPANY} Proposal",
        author=COMPANY,
    )
    doc.build(story, onFirstPage=_pdf_page_decor, onLaterPages=_pdf_page_decor)
    print(f"PDF: {OUT_PDF} ({Path(OUT_PDF).stat().st_size // 1024} KB)")


def main():
    icons = fetch_all_icons()
    print("Icons:", sum(1 for v in icons.values() if v), "/", len(icons))
    build_docx(icons)
    build_pdf(icons)


if __name__ == "__main__":
    main()
