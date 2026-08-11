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
from fpdf import FPDF
from PIL import Image

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

    package_page(
        doc,
        "BASIC WEBSITE — PKR 50,000",
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
        icons.get("basic"),
    )

    package_page(
        doc,
        "STANDARD WEBSITE — PKR 80,000",
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
        None,
        icons.get("standard"),
    )

    package_page(
        doc,
        "PREMIUM WEBSITE — PKR 120,000",
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
        None,
        icons.get("premium"),
    )

    doc.add_page_break()
    section_with_icon(doc, "Digital Marketing", icons.get("marketing"), "Monthly packages")

    for name, price, items in [
        (
            "BASIC — PKR 15,000 / month",
            "15,000",
            ["Meta Ads", "Audience targeting", "Campaign optimization", "Monthly reporting"],
        ),
        (
            "STANDARD — PKR 30,000 / month",
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
            "PREMIUM — PKR 50,000 / month",
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


class PDF(FPDF):
    def footer(self):
        self.set_y(-10)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 8, f"{COMPANY} | {WEB}", align="C")


def pdf_section(pdf, title, icon_path: Path | None, body_lines=None):
    pdf.ln(3)
    y = pdf.get_y()
    if icon_path and icon_path.exists():
        pdf.image(str(icon_path), x=12, y=y, w=10)
        pdf.set_xy(26, y)
    else:
        pdf.set_x(12)
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 7, ascii_safe(title), new_x="LMARGIN", new_y="NEXT")
    if body_lines:
        pdf.set_x(12)
        pdf.set_font("Helvetica", "", 10)
        for bl in body_lines:
            pdf.multi_cell(186, 5, ascii_safe(bl))
    pdf.ln(2)


def build_pdf(icons: dict):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    if icons.get("logo") and icons["logo"].exists():
        pdf.image(str(icons["logo"]), x=88, w=32)
        pdf.ln(22)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(6, 172, 186)
    pdf.cell(0, 6, "CoLab Point", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, COMPANY, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "Digital Agency Proposal", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 6, "Website Designing & Development  |  Digital Marketing", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, WEB, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, PHONE, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, EMAIL, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.multi_cell(0, 6, ADDRESS, align="C")

    pdf.add_page()
    pdf_section(
        pdf,
        "Company Profile",
        icons.get("company"),
        [
            f"{COMPANY} is the digital services arm of CoLab Point, Gujrat.",
            "We build websites, e-commerce stores, and run performance marketing campaigns.",
        ],
    )

    packages = [
        (
            "BASIC WEBSITE — PKR 50,000",
            icons.get("basic"),
            "Best for: Small businesses and startups",
            [
                "Professional business website | Up to 5 pages | Responsive design",
                "WordPress CMS | Custom UI | Contact form | WhatsApp integration",
                "Social media links | Google Analytics | SSL | 30 days support",
                "SEO and e-commerce NOT included.",
            ],
        ),
        (
            "STANDARD WEBSITE — PKR 80,000",
            icons.get("standard"),
            "Best for: Growing brands and online sellers",
            [
                "Everything in Basic, plus: Up to 10 pages | Premium UI/UX | Blog",
                "WooCommerce | Product upload | Payment gateway | Search Console",
                "Facebook Pixel | Advanced SEO | Speed optimization | 60 days support",
            ],
        ),
        (
            "PREMIUM WEBSITE — PKR 120,000",
            icons.get("premium"),
            "Best for: Established businesses and enterprises",
            [
                "Everything in Standard, plus: Unlimited pages | Fully custom design",
                "Advanced WooCommerce | CRM | Booking system | API integrations",
                "Premium security | Performance optimization | Admin training | 90 days support",
            ],
        ),
    ]

    for title, icon, audience, items in packages:
        pdf.add_page()
        pdf_section(pdf, title, icon, [audience, "WHAT IS INCLUDED IN THIS PRICE:"] + items)

    pdf.add_page()
    pdf_section(pdf, "Digital Marketing", icons.get("marketing"), ["Monthly packages"])
    for block in [
        ("BASIC — PKR 15,000 / month", ["Meta Ads", "Audience targeting", "Campaign optimization", "Monthly reporting"]),
        (
            "STANDARD — PKR 30,000 / month",
            ["Google + Meta Ads", "Lead generation", "Conversion tracking", "Landing page recommendations", "Monthly reports"],
        ),
        (
            "PREMIUM — PKR 50,000 / month",
            ["Google + Meta Ads", "SEO", "Remarketing", "Conversion optimization", "Marketing strategy", "Weekly meetings", "Detailed reports"],
        ),
    ]:
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_x(12)
        pdf.cell(0, 6, ascii_safe(block[0]), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        pdf.set_x(16)
        pdf.cell(0, 5, "What is included:", new_x="LMARGIN", new_y="NEXT")
        for it in block[1]:
            pdf.set_x(18)
            pdf.cell(0, 5, f"- {it}", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

    pdf.add_page()
    pdf_section(
        pdf,
        "Contact",
        icons.get("contact"),
        [
            f"Company: {COMPANY}",
            f"Website: {WEB}",
            f"Phone: {PHONE}",
            f"Email: {EMAIL}",
            f"Address: {ADDRESS}",
        ],
    )

    pdf.output(OUT_PDF)
    print(f"PDF: {OUT_PDF} ({Path(OUT_PDF).stat().st_size // 1024} KB)")


def main():
    icons = fetch_all_icons()
    print("Icons:", sum(1 for v in icons.values() if v), "/", len(icons))
    build_docx(icons)
    build_pdf(icons)


if __name__ == "__main__":
    main()
