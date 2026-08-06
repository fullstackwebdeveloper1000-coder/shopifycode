#!/usr/bin/env python3
"""CoLab Point proposal — Word file with maximum visibility (black text, borders)."""

import io
import urllib.request
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BLACK = RGBColor(0x00, 0x00, 0x00)
BRAND_DARK = RGBColor(0x04, 0x24, 0x3C)
BRAND_TEAL = RGBColor(0x06, 0xAC, 0xBA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HEX_DARK = "04243C"
HEX_TEAL = "06ACBA"
HEX_LIGHT = "D4EEF2"
FONT = "Calibri"

COMPANY = "CoLab Space Point"
BRAND = "CoLab Point"
WEB = "www.colabpoint.com"
PHONE = "+92 349 7684322  |  +92 332 4384322"
EMAIL = "colabpoint@gmail.com  |  hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"
LOGO_URL = "https://colabpoint.com/wp-content/uploads/2024/05/Web-Logo-II-150x145.png"
OUT = "/workspace/CoLab_Space_Point_Digital_Agency_Proposal.docx"


def set_cell_border(cell, color="04243C", size="8"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def set_table_borders(table, color="04243C"):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, color)


def shade(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_run_font(run, size=12, bold=False, color=BLACK):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = FONT
    run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)
    rf.set(qn("w:cs"), FONT)
    rPr.append(rf)
    # Explicit Word color (fixes invisible text in some viewers)
    for old in rPr.findall(qn("w:color")):
        rPr.remove(old)
    col = OxmlElement("w:color")
    col.set(qn("w:val"), f"{color.rgb:06X}" if hasattr(color, "rgb") else "000000")
    rPr.append(col)


def rgb_hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def set_run_font_ex(run, size=12, bold=False, color=BLACK):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = FONT
    run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)
    rf.set(qn("w:cs"), FONT)
    rPr.append(rf)
    for old in rPr.findall(qn("w:color")):
        rPr.remove(old)
    col = OxmlElement("w:color")
    col.set(qn("w:val"), rgb_hex(color))
    rPr.append(col)


def p_text(cell_or_doc, text, size=12, bold=False, color=BLACK, align=None, is_cell=False):
    if is_cell:
        p = cell_or_doc.paragraphs[0] if cell_or_doc.paragraphs else cell_or_doc.add_paragraph()
    else:
        p = cell_or_doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.space_before = Pt(3)
    set_run_font_ex(p.add_run(text), size, bold, color)
    return p


def heading_block(doc, title, subtitle=""):
    doc.add_paragraph()
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.rows[0].cells[0]
    shade(c, HEX_DARK)
    set_cell_border(c, "06ACBA", "12")
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_run_font_ex(p.add_run(title), 18, True, WHITE)
    if subtitle:
        p2 = c.add_paragraph()
        set_run_font_ex(p2.add_run(subtitle), 12, False, WHITE)
    set_table_borders(t)
    doc.add_paragraph()


def body(doc, text, bold=False):
    p_text(doc, text, 12, bold, BLACK)


def package_page(doc, name, price, suitable, features, note=None):
    doc.add_page_break()
    heading_block(doc, name, f"Package Price: PKR {price}  |  {suitable}")

    p = doc.add_paragraph()
    set_run_font_ex(p.add_run("IS PRICE ME YE SHAMIL HOGA:"), 14, True, BRAND_DARK)

    for item in features:
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Cm(0.5)
        set_run_font_ex(bp.add_run(f"• {item}"), 12, False, BLACK)

    if note:
        doc.add_paragraph()
        np = doc.add_paragraph()
        set_run_font_ex(np.add_run(f"IMPORTANT: {note}"), 12, True, BRAND_DARK)


def monthly_packages(doc, section_title, packs):
    doc.add_page_break()
    heading_block(doc, section_title, "Monthly price + included services")

    for title, price, features in packs:
        doc.add_paragraph()
        t = doc.add_table(rows=2, cols=1)
        c1 = t.rows[0].cells[0]
        shade(c1, HEX_TEAL)
        set_cell_border(c1)
        p1 = c1.paragraphs[0]
        set_run_font_ex(
            p1.add_run(f"{title}  —  PKR {price} per month"),
            14,
            True,
            BLACK,
        )

        c2 = t.rows[1].cells[0]
        shade(c2, HEX_LIGHT)
        set_cell_border(c2)
        p2 = c2.paragraphs[0]
        set_run_font_ex(p2.add_run("Is price me ye shamil hai:"), 12, True, BLACK)
        for f in features:
            fp = c2.add_paragraph()
            set_run_font_ex(fp.add_run(f"• {f}"), 11, False, BLACK)
        set_table_borders(t)


def setup_doc(doc):
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2)
    sec.right_margin = Cm(2)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(12)
    normal.font.color.rgb = BLACK
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font_ex(fp.add_run(f"{COMPANY} | {WEB} | Gujrat"), 9, False, BLACK)


def build():
    doc = Document()
    setup_doc(doc)

    # Cover
    try:
        data = urllib.request.urlopen(LOGO_URL, timeout=15).read()
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lp.add_run().add_picture(io.BytesIO(data), width=Inches(1.5))
    except Exception:
        pass

    p_text(doc, BRAND, 14, True, BRAND_TEAL, WD_ALIGN_PARAGRAPH.CENTER)
    p_text(doc, COMPANY, 26, True, BLACK, WD_ALIGN_PARAGRAPH.CENTER)
    p_text(doc, "Digital Agency Proposal", 16, True, BLACK, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    for line in [
        "Website Designing & Development",
        "Social Media Management",
        "Digital Marketing",
    ]:
        p_text(doc, line, 13, False, BLACK, WD_ALIGN_PARAGRAPH.CENTER)
    p_text(doc, WEB, 12, True, BRAND_DARK, WD_ALIGN_PARAGRAPH.CENTER)
    p_text(doc, ADDRESS, 11, False, BLACK, WD_ALIGN_PARAGRAPH.CENTER)
    p_text(doc, PHONE, 11, False, BLACK, WD_ALIGN_PARAGRAPH.CENTER)
    p_text(doc, EMAIL, 11, False, BLACK, WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_page_break()
    heading_block(doc, "Company Profile", "CoLab Point — Gujrat")
    body(
        doc,
        f"{COMPANY} aap ko website design, WordPress, e-commerce, branding, social media, "
        "SEO, Google Ads, Meta Ads aur business growth ki complete services deta hai.",
    )
    body(
        doc,
        "Har section me pehle PACKAGE PRICE likhi hai, phir detail me bataya gaya hai "
        "ke is price me kya kya shamil hoga.",
        True,
    )

    heading_block(doc, "Why Choose Us", "")
    for x in [
        "Experienced team",
        "Professional support",
        "Business focused solutions",
        "Modern technology",
        "Creative design",
        "Performance marketing",
        "Transparent communication",
        "Long term partnership",
    ]:
        bp = doc.add_paragraph()
        set_run_font_ex(bp.add_run(f"• {x}"), 12, False, BLACK)

    doc.add_page_break()
    heading_block(
        doc,
        "Website Designing & Development",
        "WordPress primary — custom development available",
    )

    package_page(
        doc,
        "BASIC WEBSITE — PKR 60,000",
        "60,000",
        "Small business & startups",
        [
            "Professional business website",
            "Up to 5 pages",
            "Responsive mobile design",
            "WordPress CMS",
            "Custom UI",
            "Contact form",
            "WhatsApp integration",
            "Social media integration",
            "Basic SEO",
            "Google Analytics",
            "Speed optimization",
            "SSL configuration",
            "Security setup",
            "Training",
            "30 days support",
        ],
        "E-Commerce is package me shamil NAHI hai.",
    )

    package_page(
        doc,
        "STANDARD WEBSITE — PKR 120,000",
        "120,000",
        "Growing brands",
        [
            "Up to 10 pages",
            "Premium UI/UX",
            "Blog",
            "WooCommerce store",
            "Product upload",
            "Payment gateway",
            "Google Search Console",
            "Facebook Pixel",
            "Advanced SEO",
            "Speed optimization",
            "60 days support",
        ],
    )

    package_page(
        doc,
        "PREMIUM WEBSITE — PKR 250,000",
        "250,000",
        "Enterprise clients",
        [
            "Unlimited pages (scope ke andar)",
            "Fully custom design",
            "Advanced WooCommerce",
            "Unlimited products",
            "Payment gateway",
            "CRM integration",
            "Booking system",
            "API integration",
            "Advanced SEO",
            "Premium security",
            "Performance optimization",
            "Admin training",
            "90 days support",
        ],
    )

    monthly_packages(
        doc,
        "Social Media Management",
        [
            (
                "BASIC",
                "25,000",
                [
                    "Facebook",
                    "Instagram",
                    "12 posts",
                    "Captions",
                    "Hashtags",
                    "Monthly report",
                ],
            ),
            (
                "STANDARD",
                "45,000",
                [
                    "Facebook, Instagram, LinkedIn",
                    "20 posts",
                    "Stories",
                    "Reels planning",
                    "Community management",
                    "Analytics",
                ],
            ),
            (
                "PREMIUM",
                "75,000",
                [
                    "Facebook, Instagram, LinkedIn, TikTok",
                    "30+ posts",
                    "Daily stories",
                    "Reels strategy",
                    "Community management",
                    "Weekly reports",
                ],
            ),
        ],
    )

    monthly_packages(
        doc,
        "Digital Marketing",
        [
            (
                "BASIC",
                "35,000",
                [
                    "Meta Ads",
                    "Audience targeting",
                    "Campaign optimization",
                    "Monthly reporting",
                ],
            ),
            (
                "STANDARD",
                "65,000",
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
                "120,000",
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
        ],
    )

    doc.add_page_break()
    heading_block(doc, "Add-On Services", "Optional — alag price")
    addons = [
        ("Logo Design", "PKR 15,000 se"),
        ("Brand Identity", "PKR 45,000 se"),
        ("Landing Page", "PKR 35,000 se"),
        ("Website Maintenance", "PKR 8,000 / month se"),
        ("Content Writing", "PKR 3,500 / page"),
        ("Graphic Design", "PKR 2,500"),
        ("Video Editing", "PKR 5,000 / minute"),
        ("SEO Audit", "PKR 25,000"),
        ("Product Upload", "PKR 500 / product"),
        ("Business Email", "PKR 5,000"),
        ("Domain & Hosting Help", "Cost + PKR 3,000"),
    ]
    tbl = doc.add_table(rows=1, cols=2)
    hdr = tbl.rows[0].cells
    shade(hdr[0], HEX_TEAL)
    shade(hdr[1], HEX_TEAL)
    set_run_font_ex(hdr[0].paragraphs[0].add_run("Service"), 12, True, BLACK)
    set_run_font_ex(hdr[1].paragraphs[0].add_run("Price"), 12, True, BLACK)
    for s, pr in addons:
        row = tbl.add_row().cells
        set_run_font_ex(row[0].paragraphs[0].add_run(s), 11, False, BLACK)
        set_run_font_ex(row[1].paragraphs[0].add_run(pr), 11, True, BLACK)
    set_table_borders(tbl)

    doc.add_page_break()
    heading_block(doc, "Our Process", "")
    for step in [
        "Discovery",
        "Planning",
        "Design",
        "Development",
        "Testing",
        "Launch",
        "Support",
    ]:
        sp = doc.add_paragraph()
        set_run_font_ex(sp.add_run(f"{step}"), 12, True, BRAND_DARK)

    heading_block(doc, "Contact", COMPANY)
    for label, val in [
        ("Website", WEB),
        ("Phone", PHONE),
        ("Email", EMAIL),
        ("Address", ADDRESS),
    ]:
        cp = doc.add_paragraph()
        set_run_font_ex(cp.add_run(f"{label}: "), 12, True, BLACK)
        set_run_font_ex(cp.add_run(val), 12, False, BLACK)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
