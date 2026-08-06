#!/usr/bin/env python3
"""CoLab Point — simple, visual pricing proposal (Word)."""

import io
import urllib.request
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Colab Point brand (from colabpoint.com)
BRAND_DARK = RGBColor(0x04, 0x24, 0x3C)
BRAND_TEAL = RGBColor(0x06, 0xAC, 0xBA)
BRAND_GREEN = RGBColor(0x00, 0xD0, 0x84)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x33, 0x33, 0x33)
MUTED = RGBColor(0x66, 0x66, 0x66)

HEX_DARK = "04243C"
HEX_TEAL = "06ACBA"
HEX_LIGHT = "E8F7F9"
HEX_WHITE = "FFFFFF"

COMPANY = "CoLab Space Point"
BRAND = "CoLab Point"
WEB = "www.colabpoint.com"
PHONE = "+92 349 7684322  |  +92 332 4384322"
EMAIL = "colabpoint@gmail.com  |  hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"
LOGO_URL = "https://colabpoint.com/wp-content/uploads/2024/05/Web-Logo-II-150x145.png"
OUT = "/workspace/CoLab_Space_Point_Digital_Agency_Proposal.docx"


def shade(cell, hex_color: str):
    el = OxmlElement("w:shd")
    el.set(qn("w:fill"), hex_color)
    el.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(el)


def run_style(run, size=11, bold=False, color=TEXT, name="Segoe UI"):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def para_space(p, before=0, after=6):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.2


def add_title_bar(doc, title, subtitle=""):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.rows[0].cells[0]
    shade(c, HEX_DARK)
    p = c.paragraphs[0]
    para_space(p, 10, 4)
    r = p.add_run(title)
    run_style(r, 20, True, WHITE)
    if subtitle:
        p2 = c.add_paragraph()
        para_space(p2, 0, 8)
        r2 = p2.add_run(subtitle)
        run_style(r2, 11, False, BRAND_TEAL)
    doc.add_paragraph()


def add_text(doc, text, size=11, bold=False, color=TEXT, center=False):
    p = doc.add_paragraph()
    para_space(p, 0, 8)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    run_style(r, size, bold, color)
    return p


def add_checks(doc, items):
    for item in items:
        p = doc.add_paragraph()
        para_space(p, 0, 4)
        r = p.add_run(f"✓  {item}")
        run_style(r, 11, False, TEXT)


def price_card(doc, package_name, price_pkr, suitable, included_label, items, note=None):
    """One clear pricing block: price + 'is price me ye hoga' list."""
    doc.add_page_break()
    # Header strip
    t = doc.add_table(rows=1, cols=1)
    c = t.rows[0].cells[0]
    shade(c, HEX_TEAL)
    p = c.paragraphs[0]
    para_space(p, 8, 8)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(package_name.upper())
    run_style(r, 18, True, WHITE)

    # Price box
    t2 = doc.add_table(rows=1, cols=1)
    c2 = t2.rows[0].cells[0]
    shade(c2, HEX_DARK)
    p2 = c2.paragraphs[0]
    para_space(p2, 14, 4)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("PACKAGE PRICE")
    run_style(r2, 10, False, BRAND_TEAL)
    p3 = c2.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_space(p3, 0, 12)
    r3 = p3.add_run(f"PKR {price_pkr}")
    run_style(r3, 32, True, WHITE)
    p4 = c2.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_space(p4, 0, 10)
    r4 = p4.add_run(f"Suitable for: {suitable}")
    run_style(r4, 10, False, RGBColor(0xCC, 0xEE, 0xF2))

    doc.add_paragraph()
    # Included section
    t3 = doc.add_table(rows=1, cols=1)
    c3 = t3.rows[0].cells[0]
    shade(c3, HEX_LIGHT)
    p5 = c3.paragraphs[0]
    para_space(p5, 10, 6)
    r5 = p5.add_run(included_label)
    run_style(r5, 13, True, BRAND_DARK)
    p6 = c3.add_paragraph()
    para_space(p6, 0, 10)
    for item in items:
        bp = c3.add_paragraph()
        para_space(bp, 0, 5)
        br = bp.add_run(f"✓  {item}")
        run_style(br, 11, False, TEXT)

    if note:
        doc.add_paragraph()
        tn = doc.add_table(rows=1, cols=1)
        cn = tn.rows[0].cells[0]
        shade(cn, "FFF3CD")
        pn = cn.paragraphs[0]
        para_space(pn, 8, 8)
        rn = pn.add_run(f"Note: {note}")
        run_style(rn, 10, True, RGBColor(0x85, 0x60, 0x00))


def three_column_pricing(doc, section_title, packages):
    """packages: list of (name, price, features list)"""
    doc.add_page_break()
    add_title_bar(doc, section_title, "Monthly packages — price aur included services")
    for name, price, features in packages:
        doc.add_paragraph()
        t = doc.add_table(rows=2, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Row 1: name + price
        c0 = t.rows[0].cells[0]
        shade(c0, HEX_DARK)
        p = c0.paragraphs[0]
        para_space(p, 8, 2)
        r1 = p.add_run(f"{name}  —  PKR {price} / Month")
        run_style(r1, 14, True, WHITE)
        # Row 2: included
        c1 = t.rows[1].cells[0]
        shade(c1, HEX_WHITE)
        p2 = c1.paragraphs[0]
        para_space(p2, 8, 4)
        r2 = p2.add_run("Is price me ye shamil hai:")
        run_style(r2, 11, True, BRAND_TEAL)
        for f in features:
            fp = c1.add_paragraph()
            para_space(fp, 0, 3)
            fr = fp.add_run(f"✓  {f}")
            run_style(fr, 10, False, TEXT)
        blank = c1.add_paragraph()
        para_space(blank, 0, 6)


def cover(doc):
    try:
        data = urllib.request.urlopen(LOGO_URL, timeout=15).read()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para_space(p, 40, 12)
        p.add_run().add_picture(io.BytesIO(data), width=Inches(1.4))
    except Exception:
        pass

    add_text(doc, BRAND, 14, True, BRAND_TEAL, center=True)
    add_text(doc, COMPANY, 28, True, BRAND_DARK, center=True)
    add_text(doc, "Digital Agency Proposal", 16, False, MUTED, center=True)
    doc.add_paragraph()

    bar = doc.add_table(rows=1, cols=1)
    bc = bar.rows[0].cells[0]
    shade(bc, HEX_TEAL)
    bp = bc.paragraphs[0]
    para_space(bp, 12, 12)
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for line in [
        "Website Designing & Development",
        "Social Media Management",
        "Digital Marketing",
    ]:
        lp = bc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para_space(lp, 2, 2)
        lr = lp.add_run(f"◆  {line}")
        run_style(lr, 12, True, WHITE)

    add_text(doc, WEB, 11, False, MUTED, center=True)
    add_text(doc, "Gujrat, Pakistan", 11, False, MUTED, center=True)


def intro(doc):
    doc.add_page_break()
    add_title_bar(doc, "Company Profile", "CoLab Point — Gujrat")
    add_text(
        doc,
        f"{COMPANY} CoLab Point ke sath milkar complete digital solutions deta hai: website design "
        "aur development, WordPress, e-commerce, branding, social media, SEO, Google Ads, Meta Ads, "
        "aur business growth.",
    )
    add_text(
        doc,
        "Neeche har service ka package price clear likha hai — aur har price ke sath detail mein "
        "bataya gaya hai ke is price me exactly kya kya shamil hoga.",
        11,
        True,
        BRAND_DARK,
    )


def why_us(doc):
    add_title_bar(doc, "Why CoLab Space Point", "Highlights")
    points = [
        "Experienced team — design, development, marketing",
        "Professional support aur clear communication",
        "Business-focused solutions — sales & leads focus",
        "Modern tech: WordPress, WooCommerce, Analytics",
        "Creative design + performance marketing",
        "Long-term partnership — support after launch",
    ]
    add_checks(doc, points)


def website_section(doc):
    doc.add_page_break()
    add_title_bar(
        doc,
        "Website Designing & Development",
        "WordPress primary platform — custom solutions bhi available",
    )
    add_text(
        doc,
        "Teen packages neeche diye gaye hain. Har page par pehle PRICE, phir "
        "'Is Price Me Ye Shamil Hoga' ki full list.",
    )

    price_card(
        doc,
        "Basic Website",
        "60,000",
        "Small business & startups",
        "Is Price Me Ye Shamil Hoga:",
        [
            "Professional business website",
            "Up to 5 pages",
            "Mobile responsive design",
            "WordPress CMS",
            "Custom UI design",
            "Contact form",
            "WhatsApp integration",
            "Social media links",
            "Basic SEO setup",
            "Google Analytics",
            "Speed optimization",
            "SSL configuration",
            "Security setup",
            "1 training session",
            "30 days support after launch",
        ],
        note="E-Commerce is package me shamil NAHI hai.",
    )

    price_card(
        doc,
        "Standard Website",
        "120,000",
        "Growing brands & online sellers",
        "Is Price Me Ye Shamil Hoga:",
        [
            "Up to 10 pages",
            "Premium UI/UX design",
            "Blog section",
            "WooCommerce online store",
            "Product upload (initial batch)",
            "Payment gateway integration",
            "Google Search Console",
            "Facebook Pixel",
            "Advanced SEO",
            "Speed optimization",
            "60 days support after launch",
        ],
    )

    price_card(
        doc,
        "Premium Website",
        "250,000",
        "Enterprise & high-growth business",
        "Is Price Me Ye Shamil Hoga:",
        [
            "Unlimited pages (agreed scope)",
            "Fully custom design",
            "Advanced WooCommerce store",
            "Unlimited products setup",
            "Payment gateway",
            "CRM integration",
            "Booking / appointment system",
            "API integrations",
            "Advanced SEO + schema",
            "Premium security",
            "Performance optimization",
            "Admin training",
            "90 days priority support",
        ],
    )

    # Comparison
    doc.add_page_break()
    add_title_bar(doc, "Website Packages — Quick Compare", "")
    headers = ["Feature", "Basic\n60,000", "Standard\n120,000", "Premium\n250,000"]
    tbl = doc.add_table(rows=1, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        shade(tbl.rows[0].cells[i], HEX_DARK)
        tbl.rows[0].cells[i].text = ""
        pr = tbl.rows[0].cells[i].paragraphs[0]
        run_style(pr.add_run(h), 9, True, WHITE)

    rows = [
        ("Pages", "5", "10", "Unlimited*"),
        ("E-Commerce", "No", "Yes", "Advanced"),
        ("Blog", "No", "Yes", "Yes"),
        ("Payment Gateway", "No", "Yes", "Yes"),
        ("CRM / Booking", "No", "No", "Yes"),
        ("Support", "30 days", "60 days", "90 days"),
    ]
    for row in rows:
        cells = tbl.add_row().cells
        for i, val in enumerate(row):
            if i == 0:
                shade(cells[i], HEX_LIGHT)
            cells[i].text = ""
            run_style(cells[i].paragraphs[0].add_run(val), 9, i == 0, TEXT if i else MUTED)


def social_section(doc):
    three_column_pricing(
        doc,
        "Social Media Management",
        [
            (
                "Basic Package",
                "25,000",
                [
                    "Facebook management",
                    "Instagram management",
                    "12 posts per month",
                    "Captions writing",
                    "Hashtag research",
                    "Monthly report",
                ],
            ),
            (
                "Standard Package",
                "45,000",
                [
                    "Facebook, Instagram, LinkedIn",
                    "20 posts per month",
                    "Stories",
                    "Reels planning",
                    "Community management",
                    "Analytics report",
                ],
            ),
            (
                "Premium Package",
                "75,000",
                [
                    "Facebook, Instagram, LinkedIn, TikTok",
                    "30+ posts per month",
                    "Daily stories",
                    "Reels strategy",
                    "Community management",
                    "Weekly reports",
                ],
            ),
        ],
    )


def marketing_section(doc):
    three_column_pricing(
        doc,
        "Digital Marketing",
        [
            (
                "Basic Package",
                "35,000",
                [
                    "Meta (Facebook/Instagram) Ads",
                    "Audience targeting",
                    "Campaign optimization",
                    "Monthly reporting",
                    "(Ad spend alag — platform ko direct)",
                ],
            ),
            (
                "Standard Package",
                "65,000",
                [
                    "Google Ads + Meta Ads",
                    "Lead generation campaigns",
                    "Conversion tracking",
                    "Landing page recommendations",
                    "Monthly reports",
                ],
            ),
            (
                "Premium Package",
                "120,000",
                [
                    "Google Ads + Meta Ads",
                    "SEO support",
                    "Remarketing",
                    "Conversion optimization",
                    "Marketing strategy",
                    "Weekly meetings",
                    "Detailed reports",
                ],
            ),
        ],
    )


def addons(doc):
    doc.add_page_break()
    add_title_bar(doc, "Add-On Services (Optional)", "Alag se price — jab zaroorat ho")
    data = [
        ("Logo Design", "From PKR 15,000"),
        ("Brand Identity", "From PKR 45,000"),
        ("Landing Page Design", "From PKR 35,000"),
        ("Website Maintenance", "From PKR 8,000 / month"),
        ("Content Writing", "From PKR 3,500 / page"),
        ("Graphic Design", "From PKR 2,500 / design"),
        ("Video Editing", "From PKR 5,000 / minute"),
        ("SEO Audit", "From PKR 25,000"),
        ("Product Upload", "From PKR 500 / product"),
        ("Business Email Setup", "From PKR 5,000"),
        ("Domain & Hosting Help", "Cost + PKR 3,000 setup"),
    ]
    tbl = doc.add_table(rows=1, cols=2)
    shade(tbl.rows[0].cells[0], HEX_TEAL)
    shade(tbl.rows[0].cells[1], HEX_TEAL)
    tbl.rows[0].cells[0].text = ""
    tbl.rows[0].cells[1].text = ""
    run_style(tbl.rows[0].cells[0].paragraphs[0].add_run("Service"), 11, True, WHITE)
    run_style(tbl.rows[0].cells[1].paragraphs[0].add_run("Price"), 11, True, WHITE)
    for svc, pr in data:
        row = tbl.add_row().cells
        row[0].text = ""
        row[1].text = ""
        run_style(row[0].paragraphs[0].add_run(svc), 10, False, TEXT)
        run_style(row[1].paragraphs[0].add_run(pr), 10, True, BRAND_DARK)


def process_contact(doc):
    doc.add_page_break()
    add_title_bar(doc, "Our Process", "")
    steps = [
        "Discovery — goals & requirements",
        "Planning — sitemap & timeline",
        "Design — UI approval",
        "Development — build & content",
        "Testing — mobile & speed check",
        "Launch — go live",
        "Support — training & help",
    ]
    add_checks(doc, steps)

    doc.add_page_break()
    add_title_bar(doc, "Contact", "Next step: discovery call")
    t = doc.add_table(rows=5, cols=2)
    rows = [
        ("Company", COMPANY),
        ("Website", WEB),
        ("Phone", PHONE),
        ("Email", EMAIL),
        ("Address", ADDRESS),
    ]
    for i, (k, v) in enumerate(rows):
        shade(t.rows[i].cells[0], HEX_LIGHT)
        t.rows[i].cells[0].text = ""
        t.rows[i].cells[1].text = ""
        run_style(t.rows[i].cells[0].paragraphs[0].add_run(k), 10, True, BRAND_DARK)
        run_style(t.rows[i].cells[1].paragraphs[0].add_run(v), 10, False, TEXT)


def setup(doc):
    s = doc.sections[0]
    s.top_margin = Cm(1.5)
    s.bottom_margin = Cm(1.5)
    s.left_margin = Cm(1.8)
    s.right_margin = Cm(1.8)
    fp = s.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_style(fp.add_run(f"{COMPANY}  |  {WEB}  |  Colab Point Gujrat"), 8, False, MUTED)


def main():
    doc = Document()
    setup(doc)
    cover(doc)
    intro(doc)
    why_us(doc)
    website_section(doc)
    social_section(doc)
    marketing_section(doc)
    addons(doc)
    process_contact(doc)
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
