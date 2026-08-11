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
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
    PageBreak, Table, TableStyle, ListFlowable, ListItem,
)

ASSETS = Path("/workspace/assets/proposal-icons")
OUT_DOCX = "/workspace/CoLab_Space_Point_Proposal_Photos.docx"
OUT_PDF = "/workspace/CoLab_Space_Point_Proposal_Final.pdf"
OUT_PDF_ALT = "/workspace/CoLab_Space_Point_Proposal_Photos.pdf"
OUT_DOCX_ALT = "/workspace/CoLab_Space_Point_Digital_Agency_Proposal.docx"

# PDF layout constants (A4 with 2cm margins → 17cm content)
PDF_M = 2 * cm
PDF_W = A4[0] - 4 * cm  # 17cm
COVER_BAND_H = 5.4 * cm

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
    icons = {k: fetch_icon(k) for k in ICONS}
    logo = icons.get("logo")
    if logo and logo.exists():
        icons["logo_cover"] = prepare_cover_logo(logo)
    return icons


def prepare_cover_logo(logo_path: Path) -> Path:
    """Logo on navy — Web-Logo-II is light grey and fades on white backgrounds."""
    from PIL import Image

    dest = ASSETS / "logo_cover.png"
    if dest.exists() and dest.stat().st_mtime >= logo_path.stat().st_mtime:
        return dest
    im = Image.open(logo_path).convert("RGBA")
    pad_x, pad_y = 56, 36
    canvas = Image.new("RGBA", (im.size[0] + pad_x * 2, im.size[1] + pad_y * 2), (4, 36, 60, 255))
    canvas.paste(im, (pad_x, pad_y), im)
    canvas.save(dest, "PNG")
    return dest


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

    # Cover — logo on navy band (visible on dark, not washed out on white)
    logo_cover = icons.get("logo_cover") or icons.get("logo")
    if logo_cover and logo_cover.exists():
        bar = doc.add_table(rows=1, cols=1)
        c = bar.rows[0].cells[0]
        shade_cell(c, "04243C")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(logo_cover), width=Inches(5.8))
        p2 = c.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        write_run(p2, "CoLab Point  |  Digital Agency  |  Gujrat, Pakistan", 10, False, TEAL)
    line(doc, "", space=8)
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


# ─── PDF — clean corporate layout (no broken table cells) ─────────────────

def _pdf_styles():
    def S(name, **kw):
        base = dict(fontName="Helvetica", fontSize=11, textColor=black, leading=16)
        base.update(kw)
        return ParagraphStyle(name, **base)

    return {
        "h1": S("h1", fontName="Helvetica-Bold", fontSize=20, textColor=HexColor("#04243C"), spaceAfter=10, spaceBefore=4),
        "h2": S("h2", fontName="Helvetica-Bold", fontSize=15, textColor=HexColor("#04243C"), spaceAfter=6, spaceBefore=8),
        "h3": S("h3", fontName="Helvetica-Bold", fontSize=13, textColor=HexColor("#06ACBA"), spaceAfter=4),
        "cover_co": S("cover_co", fontSize=12, textColor=HexColor("#06ACBA"), alignment=TA_CENTER),
        "cover_title": S("cover_title", fontName="Helvetica-Bold", fontSize=28, alignment=TA_CENTER, spaceAfter=8, leading=32),
        "cover_sub": S("cover_sub", fontName="Helvetica-Bold", fontSize=15, alignment=TA_CENTER, spaceAfter=14),
        "cover_info": S("cover_info", fontSize=11, alignment=TA_CENTER, leading=18),
        "body": S("body", spaceAfter=6),
        "bold": S("bold", fontName="Helvetica-Bold", spaceAfter=4),
        "price_big": S("price_big", fontName="Helvetica-Bold", fontSize=22, textColor=HexColor("#04243C"), spaceAfter=8),
        "muted": S("muted", fontSize=10, textColor=HexColor("#555555"), spaceAfter=4),
        "bullet": S("bullet", leftIndent=16, bulletIndent=8, leading=15, spaceAfter=3),
    }


def _pdf_bullets(items, styles):
    return ListFlowable(
        [ListItem(Paragraph(ascii_safe(i), styles["bullet"])) for i in items],
        bulletType="bullet", start="•", leftIndent=12,
    )


def _pdf_icon(path, size_cm=0.75):
    if path and Path(path).exists():
        return RLImage(str(path), width=size_cm * cm, height=size_cm * cm, kind="proportional")
    return Spacer(size_cm * cm, size_cm * cm)


def _pdf_logo(path, width_cm=12.0):
    """Wide logo — use width-based sizing, not square box."""
    if not path or not Path(path).exists():
        return Spacer(width_cm * cm, 1.5 * cm)
    from PIL import Image

    w, h = Image.open(path).size
    height_cm = width_cm * h / w if w else 2.0
    return RLImage(str(path), width=width_cm * cm, height=height_cm * cm)


def _pdf_hr(story, color="#06ACBA"):
    t = Table([[""]], colWidths=[PDF_W], rowHeights=[2])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor(color)),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2 * cm))


def _pdf_heading(story, title, subtitle="", icon_path=None, styles=None):
    styles = styles or _pdf_styles()
    title_p = Paragraph(ascii_safe(title), styles["h1"])
    if icon_path and Path(icon_path).exists():
        sub = Paragraph(ascii_safe(subtitle), styles["muted"]) if subtitle else Spacer(1, 2)
        row = Table(
            [[_pdf_icon(icon_path, 0.85), Table([[title_p], [sub]], colWidths=[PDF_W - 1.4 * cm])]],
            colWidths=[1.2 * cm, PDF_W - 1.2 * cm],
        )
        row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
        story.append(row)
    else:
        story.append(title_p)
        if subtitle:
            story.append(Paragraph(ascii_safe(subtitle), styles["muted"]))
    _pdf_hr(story)


def _pdf_cover_page(canvas, doc):
    """Full-bleed navy band with large logo — logo PNG is light grey and needs dark bg."""
    canvas.saveState()
    navy = HexColor("#04243C")
    teal = HexColor("#06ACBA")
    band_h = COVER_BAND_H

    canvas.setFillColor(navy)
    canvas.rect(0, A4[1] - band_h, A4[0], band_h, fill=1, stroke=0)
    canvas.setFillColor(teal)
    canvas.rect(0, A4[1] - band_h, A4[0], 0.12 * cm, fill=1, stroke=0)

    logo_path = getattr(doc, "logo_cover_path", None)
    if logo_path and Path(logo_path).exists():
        img = ImageReader(str(logo_path))
        iw, ih = img.getSize()
        target_w = min(A4[0] - 2.4 * cm, 13.5 * cm)
        target_h = target_w * ih / iw
        x = (A4[0] - target_w) / 2
        y = A4[1] - band_h + (band_h - target_h) / 2 + 0.35 * cm
        canvas.drawImage(img, x, y, width=target_w, height=target_h, mask="auto")

    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(teal)
    canvas.drawCentredString(
        A4[0] / 2, A4[1] - band_h + 0.6 * cm,
        "CoLab Point  |  Digital Agency  |  Gujrat, Pakistan",
    )
    canvas.restoreState()


def _pdf_footer(canvas, doc):
    canvas.saveState()
    if doc.page == 1:
        _pdf_cover_page(canvas, doc)
    elif doc.page > 1:
        canvas.setStrokeColor(HexColor("#06ACBA"))
        canvas.setLineWidth(0.8)
        canvas.line(PDF_M, A4[1] - 1.55 * cm, A4[0] - PDF_M, A4[1] - 1.55 * cm)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(HexColor("#04243C"))
        canvas.drawString(PDF_M, A4[1] - 1.25 * cm, ascii_safe(COMPANY))
    canvas.setStrokeColor(HexColor("#06ACBA"))
    canvas.setLineWidth(0.4)
    canvas.line(PDF_M, 1.55 * cm, A4[0] - PDF_M, 1.55 * cm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(HexColor("#777777"))
    canvas.drawString(PDF_M, 1.05 * cm, f"{WEB}  |  +92 349 7684322")
    canvas.drawRightString(A4[0] - PDF_M, 1.05 * cm, f"Page {doc.page}")
    canvas.restoreState()


def _pdf_package_block(story, pkg, styles, icons):
    """Full-width package card — no broken columns."""
    title_w = ParagraphStyle(
        "pt", fontName="Helvetica-Bold", fontSize=14, textColor=white, leading=18,
    )
    price_w = ParagraphStyle(
        "pp", fontName="Helvetica-Bold", fontSize=16, textColor=HexColor("#06ACBA"),
        alignment=TA_RIGHT, leading=18,
    )
    hdr = Table(
        [[
            Paragraph(ascii_safe(pkg["title"]), title_w),
            Paragraph(f"PKR {pkg['price']}", price_w),
        ]],
        colWidths=[PDF_W * 0.62, PDF_W * 0.38],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#04243C")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
    ]))
    story.append(hdr)

    best = Table(
        [[Paragraph(f"<b>Best for:</b> {ascii_safe(pkg['audience'])}", styles["body"])]],
        colWidths=[PDF_W],
    )
    best.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#E8F7F9")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
    ]))
    story.append(best)
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph("What is included in this price:", styles["h3"]))
    story.append(_pdf_bullets(pkg["items"], styles))
    if pkg.get("note"):
        story.append(Spacer(1, 0.15 * cm))
        note = Table([[Paragraph(ascii_safe(pkg["note"]), styles["bold"])]], colWidths=[PDF_W])
        note.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), HexColor("#FFF3CD")),
            ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#F0AD4E")),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ]))
        story.append(note)


def build_pdf(icons: dict):
    styles = _pdf_styles()
    story = []
    logo_cover = icons.get("logo_cover") or icons.get("logo")

    # ── Cover: large logo drawn on full-bleed navy band (see _pdf_cover_page) ──
    story.append(Spacer(1, COVER_BAND_H - 2.6 * cm + 0.35 * cm))

    story.append(Paragraph(COMPANY, ParagraphStyle(
        "ct", fontName="Helvetica-Bold", fontSize=30, textColor=HexColor("#04243C"),
        alignment=TA_CENTER, spaceAfter=10, leading=34,
    )))
    story.append(Paragraph("Digital Agency Proposal", ParagraphStyle(
        "cs", fontName="Helvetica", fontSize=16, textColor=HexColor("#555555"),
        alignment=TA_CENTER, spaceAfter=16,
    )))
    _pdf_hr(story)
    story.append(Paragraph("Website Designing &amp; Development", styles["cover_info"]))
    story.append(Paragraph("Digital Marketing", styles["cover_info"]))
    story.append(Spacer(1, 1.2 * cm))

    contact_rows = [
        ["Website", WEB],
        ["Phone", PHONE],
        ["Email", EMAIL],
        ["Address", ADDRESS],
    ]
    ct = Table(
        [[Paragraph(f"<b>{k}</b>", ParagraphStyle("ck", fontName="Helvetica-Bold", fontSize=10, textColor=HexColor("#04243C"))),
          Paragraph(ascii_safe(v), ParagraphStyle("cv", fontSize=10, leading=14))] for k, v in contact_rows],
        colWidths=[2.8 * cm, PDF_W - 2.8 * cm],
    )
    ct.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#06ACBA")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#D0E8EB")),
        ("BACKGROUND", (0, 0), (0, -1), HexColor("#F7FBFC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(ct)
    story.append(PageBreak())

    # ── Company Profile ──
    _pdf_heading(story, "Company Profile", "CoLab Space Point - Gujrat, Pakistan", icons.get("logo"), styles)
    story.append(Paragraph("<b>About Us</b>", styles["bold"]))
    story.append(Paragraph(ascii_safe(COMPANY_ABOUT), styles["body"]))
    story.append(Paragraph(ascii_safe(COMPANY_MISSION), styles["body"]))
    story.append(Spacer(1, 0.25 * cm))

    col_w = PDF_W / 4
    stats = Table(
        [[Paragraph(f"<b>{ascii_safe(a)}</b>", ParagraphStyle("sl", fontName="Helvetica-Bold", fontSize=9, textColor=white, alignment=TA_CENTER)) for a, _ in COMPANY_STATS],
         [Paragraph(ascii_safe(b), ParagraphStyle("sv", fontSize=9, alignment=TA_CENTER, leading=13)) for _, b in COMPANY_STATS]],
        colWidths=[col_w] * 4,
    )
    stats.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#06ACBA")),
        ("BACKGROUND", (0, 1), (-1, 1), HexColor("#F7FBFC")),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#06ACBA")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#06ACBA")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(stats)
    story.append(Spacer(1, 0.35 * cm))
    story.append(Paragraph("Our Core Services", styles["h2"]))
    half = (len(COMPANY_SERVICES) + 1) // 2
    svc_rows = []
    for i in range(half):
        l = f"• {COMPANY_SERVICES[i]}" if i < len(COMPANY_SERVICES) else ""
        r = f"• {COMPANY_SERVICES[i + half]}" if i + half < len(COMPANY_SERVICES) else ""
        svc_rows.append([Paragraph(ascii_safe(l), styles["body"]), Paragraph(ascii_safe(r), styles["body"])])
    svc = Table(svc_rows, colWidths=[PDF_W / 2, PDF_W / 2])
    svc.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(svc)
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Why Choose CoLab Space Point", styles["h2"]))
    story.append(_pdf_bullets(WHY_CHOOSE, styles))
    story.append(PageBreak())

    # ── Website packages — one per page, full-width cards ──
    for i, pkg in enumerate(WEBSITE_PACKAGES):
        if i > 0:
            story.append(PageBreak())
        if i == 0:
            _pdf_heading(story, "Website Designing & Development", "WordPress primary platform", icons.get("website"), styles)
        _pdf_package_block(story, pkg, styles, icons)

    story.append(PageBreak())
    _pdf_heading(story, "Digital Marketing", "Monthly packages", icons.get("marketing"), styles)
    for name, price, items in MARKETING_PACKAGES:
        m_hdr = Table(
            [[Paragraph(ascii_safe(f"{name} - PKR {price} / month"), ParagraphStyle(
                "mh", fontName="Helvetica-Bold", fontSize=13, textColor=white,
            ))]],
            colWidths=[PDF_W],
        )
        m_hdr.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), HexColor("#04243C")),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ]))
        story.append(m_hdr)
        story.append(Paragraph("What is included:", styles["h3"]))
        story.append(_pdf_bullets(items, styles))
        story.append(Spacer(1, 0.3 * cm))

    story.append(PageBreak())
    _pdf_heading(story, "Contact", "Get in touch with our team", icons.get("contact"), styles)
    for label, val in [("Company", COMPANY), ("Website", WEB), ("Phone", PHONE), ("Email", EMAIL), ("Address", ADDRESS)]:
        story.append(Paragraph(f"<b>{label}:</b> {ascii_safe(val)}", styles["body"]))

    doc = SimpleDocTemplate(
        OUT_PDF, pagesize=A4,
        leftMargin=PDF_M, rightMargin=PDF_M,
        topMargin=2.6 * cm, bottomMargin=2.0 * cm,
        title=f"{COMPANY} Proposal", author=COMPANY,
    )
    doc.logo_cover_path = str(logo_cover) if logo_cover else None
    doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)

    # Copy to legacy names so all links work
    import shutil
    shutil.copy(OUT_PDF, OUT_PDF_ALT)
    print(f"PDF: {OUT_PDF} ({Path(OUT_PDF).stat().st_size // 1024} KB)")


def main():
    icons = fetch_all_icons()
    print("Icons:", sum(1 for v in icons.values() if v), "/", len(icons))
    build_docx(icons)
    build_pdf(icons)


if __name__ == "__main__":
    main()
