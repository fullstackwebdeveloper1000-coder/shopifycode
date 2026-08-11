#!/usr/bin/env python3
"""CoLab Point proposal — WHITE page, BLACK large text (readable everywhere)."""

import io
import urllib.request
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BLACK = RGBColor(0, 0, 0)
TEAL = RGBColor(0x06, 0xAC, 0xBA)
NAVY = RGBColor(0x04, 0x24, 0x3C)
FONT = "Arial"

OUT = "/workspace/CoLab_Space_Point_Digital_Agency_Proposal.docx"
LOGO_URL = "https://colabpoint.com/wp-content/uploads/2024/05/Web-Logo-II-150x145.png"

COMPANY = "CoLab Space Point"
WEB = "www.colabpoint.com"
PHONE = "+92 349 7684322  |  +92 478 986460"
EMAIL = "colabpoint@gmail.com  |  hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"


def force_white_page(doc):
    """Page background always white (fixes dark preview)."""
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), "FFFFFF")
    doc.element.insert(0, bg)
    for sec in doc.sections:
        sectPr = sec._sectPr
        pg = sectPr.find(qn("w:pgSz"))
        if pg is not None:
            pass
        # Remove any page color fill
        for el in sectPr.findall(qn("w:pgMar")):
            pass


def write_run(paragraph, text, size=14, bold=False, color=BLACK):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = FONT
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    for tag in ("w:rFonts", "w:color"):
        for old in rPr.findall(qn(tag)):
            rPr.remove(old)
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), FONT)
    fonts.set(qn("w:hAnsi"), FONT)
    fonts.set(qn("w:cs"), FONT)
    rPr.append(fonts)
    col = OxmlElement("w:color")
    col.set(qn("w:val"), "000000" if color == BLACK else f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
    rPr.append(col)
    return run


def line(doc, text="", size=14, bold=False, color=BLACK, center=False, space=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space)
    p.paragraph_format.space_before = Pt(4)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text:
        write_run(p, text, size, bold, color)
    return p


def section_title(doc, text):
    """Big black heading on white — no dark boxes."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(8)
    pPr_elem = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), "06ACBA")
    pBdr.append(left)
    pPr_elem.append(pBdr)
    write_run(p, text, 20, True, BLACK)


def price_line(doc, text):
    p = line(doc, text, 18, True, NAVY, space=6)


def includes_header(doc):
    line(doc, "WHAT IS INCLUDED IN THIS PRICE:", 16, True, TEAL, space=8)


def bullet(doc, text):
    p = line(doc, f"  •  {text}", 14, False, BLACK, space=4)


def package(doc, title, price, for_who, items, note=None):
    doc.add_page_break()
    section_title(doc, title)
    price_line(doc, f"Package Price: PKR {price}")
    line(doc, f"Best for: {for_who}", 14, False, BLACK, space=12)
    includes_header(doc)
    for item in items:
        bullet(doc, item)
    if note:
        line(doc, note, 14, True, NAVY, space=12)


def monthly_block(doc, name, price, items):
    line(doc, f"{name} — PKR {price} / MONTH", 16, True, BLACK, space=8)
    line(doc, "What is included:", 14, True, TEAL, space=6)
    for item in items:
        bullet(doc, item)
    line(doc, "—" * 40, 10, False, BLACK, space=14)


def build():
    doc = Document()
    force_white_page(doc)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(14)
    normal.font.color.rgb = BLACK

    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    # Cover
    try:
        data = urllib.request.urlopen(LOGO_URL, timeout=15).read()
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lp.add_run().add_picture(io.BytesIO(data), width=Inches(1.6))
    except Exception:
        pass

    line(doc, "CoLab Point", 16, True, TEAL, center=True, space=8)
    line(doc, COMPANY, 32, True, BLACK, center=True, space=8)
    line(doc, "Digital Agency Proposal", 20, True, BLACK, center=True, space=16)
    for s in [
        "Website Designing & Development",
        "Digital Marketing",
    ]:
        line(doc, s, 15, False, BLACK, center=True, space=6)
    line(doc, WEB, 15, True, NAVY, center=True, space=12)
    line(doc, ADDRESS, 14, False, BLACK, center=True, space=4)
    line(doc, PHONE, 14, False, BLACK, center=True, space=4)
    line(doc, EMAIL, 14, False, BLACK, center=True, space=4)

    doc.add_page_break()
    section_title(doc, "Company Profile")
    line(
        doc,
        f"{COMPANY} delivers website design, WordPress development, e-commerce, branding, "
        "SEO, Google Ads, Meta Ads, and business growth solutions.",
        14,
        False,
        BLACK,
        space=10,
    )
    line(
        doc,
        "Each section shows the package price in bold, followed by a clear list of everything "
        "included at that price.",
        14,
        True,
        BLACK,
        space=12,
    )

    section_title(doc, "Why Choose CoLab Space Point")
    for point in [
        "Experienced multidisciplinary team",
        "Professional support and clear communication",
        "Business-focused, results-driven solutions",
        "Modern technology stack",
        "Creative design and performance marketing",
        "Transparent pricing and long-term partnership",
    ]:
        bullet(doc, point)

    section_title(doc, "Website Designing & Development")
    line(
        doc,
        "WordPress is our primary platform. Fully custom development is also available on request.",
        14,
        space=10,
    )

    package(
        doc,
        "BASIC WEBSITE",
        "50,000",
        "Small businesses and startups",
        [
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
        "IMPORTANT: SEO and e-commerce are NOT included in this package.",
    )

    package(
        doc,
        "STANDARD WEBSITE",
        "80,000",
        "Growing brands and online sellers",
        [
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
    )

    package(
        doc,
        "PREMIUM WEBSITE",
        "120,000",
        "Established businesses and enterprises",
        [
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
    )

    doc.add_page_break()
    section_title(doc, "Digital Marketing")
    line(doc, "Monthly packages — pricing and included services", 14, True, BLACK, space=12)

    monthly_block(
        doc,
        "BASIC",
        "15,000",
        ["Meta Ads", "Audience targeting", "Campaign optimization", "Monthly reporting"],
    )
    monthly_block(
        doc,
        "STANDARD",
        "30,000",
        [
            "Google Ads + Meta Ads",
            "Lead generation",
            "Conversion tracking",
            "Landing page recommendations",
            "Monthly reports",
        ],
    )
    monthly_block(
        doc,
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
    )

    doc.add_page_break()
    section_title(doc, "Contact")
    line(doc, f"Company: {COMPANY}", 14, space=6)
    line(doc, f"Website: {WEB}", 14, space=6)
    line(doc, f"Phone: {PHONE}", 14, space=6)
    line(doc, f"Email: {EMAIL}", 14, space=6)
    line(doc, f"Address: {ADDRESS}", 14, space=6)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    write_run(fp, f"{COMPANY} | {WEB}", 10, False, BLACK)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
