#!/usr/bin/env python3
"""CoLab Point proposal — white background, images, clear pricing."""

import io
import os
import urllib.request
from pathlib import Path

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
ASSETS = Path("/workspace/assets/proposal")

COMPANY = "CoLab Space Point"
WEB = "www.colabpoint.com"
PHONE = "+92 349 7684322  |  +92 478 986460"
EMAIL = "colabpoint@gmail.com  |  hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"

# Brand + professional digital agency imagery
IMAGES = {
    "logo": "https://colabpoint.com/wp-content/uploads/2024/05/Web-Logo-II.png",
    "cover_banner": "https://colabpoint.com/wp-content/uploads/2024/06/popup-bg-1-1024x374-1.jpg",
    "office": "https://images.unsplash.com/photo-1497366216548-37526070297c?w=1200&q=80",
    "web_design": "https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=1200&q=80",
    "ecommerce": "https://images.unsplash.com/photo-1556742049-0cfed4f6a45d?w=1200&q=80",
    "premium_web": "https://images.unsplash.com/photo-1498050108023-c5249f4df085?w=1200&q=80",
    "digital_marketing": "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=1200&q=80",
    "contact": "https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1200&q=80",
}


def fetch_image(key: str) -> bytes | None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    ext = ".jpg" if "jpg" in IMAGES[key] else ".png"
    cache = ASSETS / f"{key}{ext}"
    if cache.exists():
        return cache.read_bytes()
    try:
        req = urllib.request.Request(IMAGES[key], headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=20).read()
        cache.write_bytes(data)
        return data
    except Exception:
        return None


def force_white_page(doc):
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), "FFFFFF")
    doc.element.insert(0, bg)


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


def add_image(doc, key: str, width_in=5.5, space_after=10):
    data = fetch_image(key)
    if not data:
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(space_after)
    p.add_run().add_picture(io.BytesIO(data), width=Inches(width_in))


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
    line(doc, text, 18, True, NAVY, space=6)


def includes_header(doc):
    line(doc, "WHAT IS INCLUDED IN THIS PRICE:", 16, True, TEAL, space=8)


def bullet(doc, text):
    line(doc, f"  •  {text}", 14, False, BLACK, space=4)


def package(doc, title, price, for_who, items, image_key=None, note=None):
    doc.add_page_break()
    section_title(doc, title)
    if image_key:
        add_image(doc, image_key, width_in=5.0, space_after=8)
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
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2)
    sec.right_margin = Cm(2)

    # ── Cover ──
    add_image(doc, "cover_banner", width_in=6.2, space_after=12)
    add_image(doc, "logo", width_in=1.8, space_after=10)
    line(doc, "CoLab Point", 16, True, TEAL, center=True, space=6)
    line(doc, COMPANY, 30, True, BLACK, center=True, space=8)
    line(doc, "Digital Agency Proposal", 20, True, BLACK, center=True, space=14)
    for s in ["Website Designing & Development", "Digital Marketing"]:
        line(doc, s, 15, False, BLACK, center=True, space=5)
    line(doc, WEB, 15, True, NAVY, center=True, space=10)
    line(doc, ADDRESS, 13, False, BLACK, center=True, space=4)
    line(doc, PHONE, 13, False, BLACK, center=True, space=4)
    line(doc, EMAIL, 13, False, BLACK, center=True, space=4)

    # ── Company Profile ──
    doc.add_page_break()
    section_title(doc, "Company Profile")
    add_image(doc, "office", width_in=5.5, space_after=10)
    line(
        doc,
        f"{COMPANY} is the digital services arm of CoLab Point — a trusted innovation hub in "
        "Gujrat since 2021. We deliver website design, WordPress development, e-commerce, "
        "branding, SEO, Google Ads, Meta Ads, and business growth solutions.",
        14,
        space=10,
    )
    line(
        doc,
        "Each package price is shown in bold, followed by a complete list of what is included.",
        14,
        True,
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

    # ── Websites ──
    doc.add_page_break()
    section_title(doc, "Website Designing & Development")
    add_image(doc, "web_design", width_in=5.5, space_after=10)
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
        image_key="web_design",
        note="IMPORTANT: SEO and e-commerce are NOT included in this package.",
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
        image_key="ecommerce",
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
        image_key="premium_web",
    )

    # ── Digital Marketing ──
    doc.add_page_break()
    section_title(doc, "Digital Marketing")
    add_image(doc, "digital_marketing", width_in=5.5, space_after=10)
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

    # ── Contact ──
    doc.add_page_break()
    section_title(doc, "Contact")
    add_image(doc, "contact", width_in=4.5, space_after=10)
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
