#!/usr/bin/env python3
"""Generate CoLab Space Point premium corporate business proposal (.docx)."""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

NAVY = RGBColor(0x0F, 0x2B, 0x46)
NAVY_MID = RGBColor(0x1A, 0x4A, 0x7A)
ACCENT = RGBColor(0x2E, 0x6B, 0xA4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x5A, 0x5A, 0x5A)
LIGHT_BG = "E8EEF4"

COMPANY = "CoLab Space Point"
TAGLINE = "Complete Digital Business Solutions"
WEB = "www.colabpoint.com"
PHONE1 = "+92 349 7684322"
PHONE2 = "+92 332 4384322"
EMAIL1 = "colabpoint@gmail.com"
EMAIL2 = "hello@colabpoint.com"
ADDRESS = "2nd Floor Anwar Center, Madina Road Near Gymkhana, Gujrat, Pakistan"


def set_cell_shading(cell, fill_hex: str):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def set_paragraph_shading(paragraph, fill_hex: str):
    pPr = paragraph._p.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill_hex)
    shading.set(qn("w:val"), "clear")
    pPr.append(shading)


def add_horizontal_rule(doc, color_hex="0F2B46"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)


def style_run(run, size=11, bold=False, color=None, font="Calibri"):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text.upper() if level == 1 else text)
    if level == 1:
        style_run(run, 22, True, NAVY, "Calibri Light")
    elif level == 2:
        style_run(run, 16, True, NAVY_MID)
    else:
        style_run(run, 13, True, ACCENT)
    return p


def add_body(doc, text, space_after=8):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    style_run(run, 11, False, GRAY)
    return p


def add_bullets(doc, items, icon="▸"):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(4)
        p.clear()
        run = p.add_run(f"{icon}  {item}")
        style_run(run, 11, False, GRAY)


def add_callout(doc, title, body):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, LIGHT_BG)
    p = cell.paragraphs[0]
    r = p.add_run(title + "\n")
    style_run(r, 12, True, NAVY)
    r2 = p.add_run(body)
    style_run(r2, 11, False, GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def page_break(doc):
    doc.add_page_break()


def cover_page(doc):
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(COMPANY)
    style_run(r, 36, True, NAVY, "Calibri Light")

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("DIGITAL AGENCY PROPOSAL")
    style_run(r2, 18, True, ACCENT)

    add_horizontal_rule(doc)
    services = [
        "Website Designing & Development",
        "Digital Marketing",
        "Social Media Management",
    ]
    for s in services:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(f"◆  {s}")
        style_run(run, 13, False, NAVY_MID)

    doc.add_paragraph()
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run(TAGLINE)
    style_run(r3, 11, False, GRAY)

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run(f"{WEB}  |  Gujrat, Pakistan")
    style_run(r4, 10, False, GRAY)

    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(48)
    r5 = p5.add_run("Confidential — Prepared for Valued Partners")
    style_run(r5, 9, True, GRAY)


def executive_summary(doc):
    page_break(doc)
    section_header_bar(doc, "Executive Summary", "Strategic overview")
    add_body(
        doc,
        "This proposal outlines how CoLab Space Point will partner with your organization to design, build, "
        "and grow a high-performing digital ecosystem. It is structured as a strategic engagement document — "
        "not a line-item quotation — so decision-makers can evaluate capabilities, methodology, investment "
        "tiers, and long-term value in one cohesive narrative.",
    )
    add_body(
        doc,
        "Our recommendations span three interconnected disciplines: web experience (design & development), "
        "always-on social presence, and performance marketing. Each discipline is available independently "
        "or as an integrated program with unified reporting and governance.",
    )
    add_callout(
        doc,
        "Engagement Highlights",
        "• Agency-grade delivery from Gujrat's leading innovation hub (CoLab Point)\n"
        "• WordPress-first development with enterprise-ready options\n"
        "• Transparent packages with clear upgrade paths\n"
        "• Measurable marketing aligned to leads, sales, and brand equity",
    )
    add_heading(doc, "Recommended Path", 2)
    add_body(
        doc,
        "Most clients begin with a website foundation (Basic, Standard, or Premium), then layer Social Media "
        "Management for brand consistency and Digital Marketing for scalable acquisition. Add-on services "
        "such as brand identity, content, and maintenance can be activated at any milestone.",
    )
    add_body(
        doc,
        "The following sections detail scope, deliverables, comparison matrices, process, governance, and "
        "commercial terms. We welcome the opportunity to tailor any package to your industry, compliance "
        "requirements, and growth targets.",
    )


def table_of_contents(doc):
    page_break(doc)
    section_header_bar(doc, "Contents", "Proposal structure")
    toc = [
        ("01", "Executive Summary"),
        ("02", "Company Profile"),
        ("03", "Why Choose CoLab Space Point"),
        ("04", "Website Designing & Development — Overview"),
        ("05", "Website Package Detail — Basic"),
        ("06", "Website Package Detail — Standard"),
        ("07", "Website Package Detail — Premium"),
        ("08", "Website Package Comparison"),
        ("09", "Social Media Management"),
        ("10", "Digital Marketing"),
        ("11", "Add-On Services"),
        ("12", "Technology & Quality Standards"),
        ("13", "Our Process"),
        ("14", "Why Clients Trust Us"),
        ("15", "Commercial Terms & Assumptions"),
        ("16", "Frequently Asked Questions"),
        ("17", "Contact & Next Steps"),
    ]
    for num, title in toc:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r1 = p.add_run(f"{num}    ")
        style_run(r1, 11, True, ACCENT)
        r2 = p.add_run(title)
        style_run(r2, 11, False, NAVY)


def section_header_bar(doc, title, subtitle=None):
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, "0F2B46")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(title)
    style_run(r, 20, True, WHITE, "Calibri Light")
    if subtitle:
        r2 = p.add_run(f"\n{subtitle}")
        style_run(r2, 11, False, RGBColor(0xCC, 0xDD, 0xEE))
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def pricing_table_header(table, headers):
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        r = p.add_run(h)
        style_run(r, 10, True, WHITE)
        set_cell_shading(hdr[i], "0F2B46")


def format_table_cell(cell, text, bold=False, header=False):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    style_run(r, 10, bold or header, NAVY if header else GRAY)
    if header:
        set_cell_shading(cell, "D6E4F0")


def company_profile(doc):
    page_break(doc)
    section_header_bar(doc, "Company Profile", "Who We Are")
    add_body(
        doc,
        f"{COMPANY} is the digital services division of the CoLab Point ecosystem — a trusted name in "
        "Gujrat's innovation landscape since 2021. Rooted in a culture of collaboration, entrepreneurship, "
        "and technology excellence at our flagship coworking hub, we extend that same commitment to businesses "
        "seeking world-class digital presence and measurable growth.",
    )
    add_body(
        doc,
        "We deliver end-to-end digital business solutions that align strategy, design, technology, and "
        "marketing under one accountable partner. From first impression to final conversion, every touchpoint "
        "is engineered for clarity, credibility, and performance.",
    )
    add_callout(
        doc,
        "Our Mission",
        "To empower organizations of every scale with premium digital experiences and data-driven marketing "
        "that translate ambition into sustainable revenue and brand equity.",
    )
    add_heading(doc, "Core Capabilities", 2)
    capabilities = [
        "Website Designing & Development",
        "WordPress Development",
        "E-commerce Solutions",
        "Branding & Visual Identity",
        "Social Media Management",
        "Digital Marketing (Full-Funnel)",
        "Search Engine Optimization (SEO)",
        "Google Ads & Performance Campaigns",
        "Meta Ads (Facebook & Instagram)",
        "Business Growth & Conversion Strategy",
    ]
    add_bullets(doc, capabilities, "●")
    add_body(
        doc,
        "Whether you are launching a startup, scaling an established brand, or modernizing legacy systems, "
        f"{COMPANY} combines agency-grade creativity with enterprise discipline — transparent processes, "
        "documented deliverables, and dedicated support.",
    )


def why_choose(doc):
    page_break(doc)
    section_header_bar(doc, "Why Choose CoLab Space Point", "Your Strategic Digital Partner")
    reasons = [
        (
            "Experienced Team",
            "Multidisciplinary specialists in design, development, content, and paid media — aligned on your KPIs.",
        ),
        (
            "Professional Support",
            "Structured onboarding, milestone reviews, and responsive communication throughout every engagement.",
        ),
        (
            "Business-Focused Solutions",
            "We prioritize outcomes — leads, sales, and retention — not vanity metrics.",
        ),
        (
            "Modern Technologies",
            "WordPress, WooCommerce, analytics stacks, and automation tools chosen for longevity and scale.",
        ),
        (
            "Creative Design",
            "Premium UI/UX that reflects your brand positioning and builds trust with your audience.",
        ),
        (
            "Performance Marketing",
            "Campaign architecture built on audience insight, testing, and continuous optimization.",
        ),
        (
            "Transparent Communication",
            "Clear scopes, realistic timelines, and reporting you can share with stakeholders.",
        ),
        (
            "Long-Term Partnership",
            "We grow with you — maintenance, iterations, and strategic counsel beyond launch day.",
        ),
    ]
    for title, desc in reasons:
        add_heading(doc, title, 3)
        add_body(doc, desc, 6)


def website_packages(doc):
    page_break(doc)
    section_header_bar(
        doc,
        "Section 1 — Website Designing & Development",
        "WordPress-led builds with fully custom solutions on demand",
    )
    add_body(
        doc,
        "WordPress is one of our primary development platforms — trusted for flexibility, security, and "
        "merchant-friendly content management. We also deliver bespoke stacks and integrations when your "
        "roadmap requires them. Every package includes responsive design, performance baseline, and "
        "professional quality assurance.",
    )

    packages = [
        (
            "Basic Website",
            "PKR 60,000",
            "Small Businesses & Startups",
            [
                "Professional business website",
                "Up to 5 pages",
                "Fully responsive design",
                "WordPress CMS",
                "Custom UI tailored to your brand",
                "Contact form",
                "WhatsApp integration",
                "Social media integration",
                "Basic on-page SEO",
                "Google Analytics setup",
                "Speed optimization",
                "SSL configuration",
                "Security hardening",
                "Client training session",
                "30 days post-launch support",
            ],
            "Important: This package does not include e-commerce functionality.",
        ),
        (
            "Standard Website",
            "PKR 120,000",
            "Growing brands & online sellers",
            [
                "Up to 10 pages",
                "Premium UI/UX design",
                "Blog module",
                "WooCommerce store",
                "Product upload (initial batch)",
                "Payment gateway integration",
                "Google Search Console",
                "Facebook Pixel",
                "Advanced SEO configuration",
                "Speed & Core Web Vitals optimization",
                "60 days post-launch support",
            ],
            None,
        ),
        (
            "Premium Website",
            "PKR 250,000",
            "Enterprises & high-growth businesses",
            [
                "Unlimited pages (within agreed scope)",
                "Fully custom design system",
                "Advanced WooCommerce",
                "Unlimited product catalog setup",
                "Payment gateway(s)",
                "CRM integration",
                "Booking / appointment system",
                "Third-party API integrations",
                "Advanced SEO & schema",
                "Premium security suite",
                "Performance optimization (CDN-ready)",
                "Comprehensive admin training",
                "90 days priority support",
            ],
            None,
        ),
    ]

    package_narratives = {
        "Basic Website": (
            "The Basic Website package establishes a credible digital headquarters for emerging businesses. "
            "We focus on clarity of message, mobile-first layouts, and essential trust signals — contact paths, "
            "social proof placeholders, and analytics — so you can begin generating inquiries from day one. "
            "This tier is intentionally non-transactional: no shopping cart, no product catalog — keeping "
            "scope lean and launch velocity high.",
        ),
        "Standard Website": (
            "The Standard Website package bridges brand storytelling and revenue. With WooCommerce, payment "
            "gateways, and advanced SEO, you are equipped to sell online while publishing thought leadership "
            "through an integrated blog. Ideal for retailers, service providers productizing offers, and "
            "regional brands expanding beyond Gujrat.",
        ),
        "Premium Website": (
            "The Premium Website package is our flagship build — bespoke design systems, unlimited structural "
            "pages, deep integrations (CRM, booking, APIs), and hardened security. Suited to organizations "
            "with complex funnels, multi-stakeholder approval, or aggressive growth plans requiring a "
            "platform that scales for years.",
        ),
    }

    for name, price, suitable, features, note in packages:
        page_break(doc)
        section_header_bar(doc, f"Website Package — {name}", price)
        add_body(doc, package_narratives.get(name, ""))
        t = doc.add_table(rows=2, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        format_table_cell(t.rows[0].cells[0], "Investment", True, True)
        format_table_cell(t.rows[0].cells[1], price, True)
        format_table_cell(t.rows[1].cells[0], "Ideal For", True, True)
        format_table_cell(t.rows[1].cells[1], suitable)
        doc.add_paragraph()
        add_heading(doc, "Included Deliverables", 3)
        add_bullets(doc, features)
        if note:
            add_callout(doc, "Package Note", note)
        add_heading(doc, "Outcomes You Can Expect", 3)
        outcomes = {
            "Basic Website": [
                "Professional first impression for investors, partners, and customers",
                "Self-service content updates via WordPress",
                "Measurable traffic baseline via Google Analytics",
            ],
            "Standard Website": [
                "Online sales channel with secure checkout",
                "Improved discoverability through advanced SEO",
                "Retargeting-ready presence via Facebook Pixel",
            ],
            "Premium Website": [
                "Unified customer journey from marketing to CRM",
                "Operational efficiency through booking and automation",
                "Enterprise resilience with premium security posture",
            ],
        }
        add_bullets(doc, outcomes.get(name, []), "✓")
        doc.add_paragraph()

    page_break(doc)
    add_heading(doc, "Website Package Comparison", 2)
    comp = doc.add_table(rows=1, cols=4)
    comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    pricing_table_header(comp, ["Feature", "Basic", "Standard", "Premium"])
    rows_data = [
        ("Pages", "Up to 5", "Up to 10", "Unlimited*"),
        ("WordPress CMS", "✓", "✓", "✓"),
        ("Custom / Premium UI", "Custom", "Premium UI/UX", "Fully Custom"),
        ("E-Commerce", "—", "WooCommerce", "Advanced WooCommerce"),
        ("Blog", "—", "✓", "✓"),
        ("Payment Gateway", "—", "✓", "✓"),
        ("CRM / Booking / API", "—", "—", "✓"),
        ("SEO Level", "Basic", "Advanced", "Advanced + Schema"),
        ("Analytics & Pixels", "GA", "GA + GSC + Pixel", "Full Stack"),
        ("Support Period", "30 Days", "60 Days", "90 Days"),
        ("Investment (PKR)", "60,000", "120,000", "250,000"),
    ]
    for row_data in rows_data:
        row = comp.add_row().cells
        for i, val in enumerate(row_data):
            bold = i == 0 or row_data[0] == "Investment (PKR)"
            format_table_cell(row[i], val, bold=bold)
    add_body(doc, "*Unlimited pages subject to agreed sitemap and content provision.", 4)


def technology_standards(doc):
    page_break(doc)
    section_header_bar(doc, "Technology & Quality Standards", "Built to perform, secure, and scale")
    add_body(
        doc,
        "Every engagement adheres to documented technical standards — regardless of package tier. Our "
        "WordPress implementations follow hardened baseline configurations, role-based access, and "
        "backup recommendations suitable for SMB and mid-market operators.",
    )
    add_heading(doc, "Development Stack", 2)
    add_bullets(
        doc,
        [
            "WordPress CMS with child-theme or block-theme best practices",
            "WooCommerce for e-commerce (Standard & Premium)",
            "SSL/TLS, security plugins, and login protection",
            "CDN-ready asset optimization and caching strategy",
            "Google Analytics 4, Search Console, and Meta Pixel where scoped",
        ],
    )
    add_heading(doc, "Quality Assurance", 2)
    add_bullets(
        doc,
        [
            "Responsive breakpoints: mobile, tablet, desktop",
            "Form submission and notification testing",
            "Checkout and payment sandbox validation (e-commerce tiers)",
            "Performance review prior to launch sign-off",
        ],
    )
    add_callout(
        doc,
        "Accessibility & Brand Consistency",
        "We apply semantic HTML, readable typography, and contrast-aware palettes aligned to your brand "
        "guidelines. Optional accessibility audits are available as an add-on.",
    )


def commercial_terms(doc):
    page_break(doc)
    section_header_bar(doc, "Commercial Terms & Assumptions", "Governance for a successful partnership")
    terms = [
        (
            "Payment Schedule",
            "Website projects: typically 50% upon SOW signature, 30% at design approval, 20% prior to "
            "go-live. Monthly retainers (social & marketing): invoiced in advance on the 1st of each month.",
        ),
        (
            "Client Responsibilities",
            "Timely provision of brand assets, copy, product data, and feedback within agreed review windows. "
            "Delays may shift timelines without penalty to CoLab Space Point.",
        ),
        (
            "Scope Changes",
            "Work outside the signed SOW is estimated separately and executed only after written approval.",
        ),
        (
            "Third-Party Costs",
            "Domain, hosting, premium plugins, stock photography, and advertising media are excluded unless "
            "explicitly bundled in writing.",
        ),
        (
            "Intellectual Property",
            "Upon full settlement, client receives agreed deliverables and licenses. CoLab Space Point may "
            "showcase non-confidential work in portfolio unless NDA restricts.",
        ),
        (
            "Confidentiality",
            "Both parties treat non-public business information as confidential for the term of engagement "
            "and twelve (12) months thereafter.",
        ),
    ]
    for title, body in terms:
        add_heading(doc, title, 3)
        add_body(doc, body, 10)
    add_callout(
        doc,
        "Currency",
        "All investments in this proposal are quoted in Pakistani Rupees (PKR) unless otherwise stated.",
    )


def social_media(doc):
    page_break(doc)
    section_header_bar(doc, "Section 2 — Social Media Management", "Monthly retainers — strategy, content & community")
    add_body(
        doc,
        "Social media is your always-on brand layer. We combine editorial planning, platform-native creative "
        "formats, and community response protocols so your audience sees consistency — not sporadic posting. "
        "Packages scale by channel breadth, content volume, and reporting cadence.",
    )
    add_callout(
        doc,
        "Strategy First",
        "Before publishing, we align on tone of voice, content pillars, and monthly themes tied to your "
        "commercial calendar (launches, promotions, hiring, events).",
    )
    packs = [
        ("Basic", "PKR 25,000 / month", ["Facebook", "Instagram", "12 posts per month", "Captions & copy", "Hashtag research", "Monthly performance report"]),
        ("Standard", "PKR 45,000 / month", ["Facebook, Instagram & LinkedIn", "20 posts per month", "Stories", "Reels planning", "Community management", "Analytics dashboard"]),
        ("Premium", "PKR 75,000 / month", ["Facebook, Instagram, LinkedIn & TikTok", "30+ posts per month", "Daily stories", "Reels strategy & direction", "Community management", "Weekly reports"]),
    ]
    table = doc.add_table(rows=1, cols=4)
    pricing_table_header(table, ["", "Basic", "Standard", "Premium"])
    features = [
        ("Monthly Fee (PKR)", "25,000", "45,000", "75,000"),
        ("Facebook", "✓", "✓", "✓"),
        ("Instagram", "✓", "✓", "✓"),
        ("LinkedIn", "—", "✓", "✓"),
        ("TikTok", "—", "—", "✓"),
        ("Posts / Month", "12", "20", "30+"),
        ("Stories", "—", "✓", "Daily"),
        ("Reels", "—", "Planning", "Full Strategy"),
        ("Community Mgmt", "—", "✓", "✓"),
        ("Reporting", "Monthly", "Analytics", "Weekly"),
    ]
    for fd in features:
        row = table.add_row().cells
        for i, v in enumerate(fd):
            format_table_cell(row[i], v, bold=(i == 0))
    doc.add_paragraph()
    for name, price, items in packs:
        add_heading(doc, f"{name} Package — {price}", 3)
        add_bullets(doc, items)


def digital_marketing(doc):
    page_break(doc)
    section_header_bar(doc, "Section 3 — Digital Marketing", "Paid media & growth — monthly management")
    add_body(
        doc,
        "Digital marketing at CoLab Space Point is engineered for accountability. Campaigns are structured "
        "around measurable objectives — cost per lead, return on ad spend, and qualified pipeline — with "
        "creative and landing experiences that reinforce your website investment.",
    )
    add_body(
        doc,
        "Premium engagements include SEO coordination so organic and paid channels compound rather than "
        "compete. Weekly strategy meetings ensure leadership visibility into spend, tests, and next actions.",
    )
    table = doc.add_table(rows=1, cols=4)
    pricing_table_header(table, ["", "Basic", "Standard", "Premium"])
    rows = [
        ("Management Fee (PKR/mo)", "35,000", "65,000", "120,000"),
        ("Meta Ads", "✓", "✓", "✓"),
        ("Google Ads", "—", "✓", "✓"),
        ("SEO", "—", "—", "✓"),
        ("Audience Targeting", "✓", "✓", "✓"),
        ("Lead Generation", "—", "✓", "✓"),
        ("Remarketing", "—", "—", "✓"),
        ("Conversion Tracking", "—", "✓", "✓"),
        ("Landing Page Guidance", "—", "✓", "✓"),
        ("CRO / Optimization", "Campaign", "Conversion", "Full Funnel"),
        ("Strategy Sessions", "—", "—", "Weekly"),
        ("Reporting", "Monthly", "Monthly", "Detailed + Weekly"),
    ]
    for r in rows:
        row = table.add_row().cells
        for i, v in enumerate(r):
            format_table_cell(row[i], v, bold=(i == 0))

    details = [
        ("Basic — PKR 35,000", "Meta Ads management with audience targeting, ongoing campaign optimization, and monthly reporting. Ad spend billed separately."),
        ("Standard — PKR 65,000", "Google Ads + Meta Ads, lead-generation funnels, conversion tracking, landing page recommendations, and monthly reports."),
        ("Premium — PKR 120,000", "Full-stack growth: Google & Meta, SEO alignment, remarketing, conversion optimization, marketing strategy, weekly meetings, and executive-ready reporting."),
    ]
    doc.add_paragraph()
    for t, b in details:
        add_heading(doc, t, 3)
        add_body(doc, b, 6)


def addons(doc):
    page_break(doc)
    section_header_bar(doc, "Add-On Services", "Optional enhancements — quoted per scope")
    addons_data = [
        ("Logo Design", "From PKR 15,000"),
        ("Brand Identity Package", "From PKR 45,000"),
        ("Landing Page Design", "From PKR 35,000"),
        ("Website Maintenance", "From PKR 8,000 / month"),
        ("Content Writing (per page)", "From PKR 3,500"),
        ("Graphic Design (per asset)", "From PKR 2,500"),
        ("Video Editing (per minute)", "From PKR 5,000"),
        ("SEO Audit", "From PKR 25,000"),
        ("Product Upload (bulk)", "From PKR 500 / product"),
        ("Business Email Setup", "From PKR 5,000"),
        ("Domain & Hosting Assistance", "At cost + PKR 3,000 setup"),
    ]
    t = doc.add_table(rows=1, cols=2)
    pricing_table_header(t, ["Service", "Indicative Investment"])
    for svc, price in addons_data:
        row = t.add_row().cells
        format_table_cell(row[0], svc)
        format_table_cell(row[1], price)
    add_body(
        doc,
        "Final pricing depends on complexity, timelines, and integration requirements. All add-ons are "
        "documented in a written change order before work begins.",
        12,
    )


def process_timeline(doc):
    page_break(doc)
    section_header_bar(doc, "Our Process", "A proven delivery framework")
    steps = [
        ("1. Discovery", "Stakeholder interviews, goals, audience, competitors, and technical requirements."),
        ("2. Planning", "Sitemap, wireframes, project plan, milestones, and success metrics."),
        ("3. Design", "Visual concepts, UI kit, and client approval gates."),
        ("4. Development", "Agile build, integrations, content population, and QA cycles."),
        ("5. Testing", "Cross-device testing, performance, security, and UAT sign-off."),
        ("6. Launch", "DNS, SSL, go-live checklist, and analytics verification."),
        ("7. Support", "Training, hypercare window, and ongoing optimization roadmap."),
    ]
    t = doc.add_table(rows=len(steps), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (phase, desc) in enumerate(steps):
        format_table_cell(t.rows[i].cells[0], phase, bold=True, header=(i % 2 == 0))
        format_table_cell(t.rows[i].cells[1], desc)
        if i % 2 == 0:
            set_cell_shading(t.rows[i].cells[0], "D6E4F0")
    add_callout(
        doc,
        "Governance",
        "You receive a dedicated project coordinator, shared timeline, and status updates at every phase.",
    )


def trust_section(doc):
    page_break(doc)
    section_header_bar(doc, "Why Clients Trust Us", "Reputation built in Gujrat — delivered globally")
    add_body(
        doc,
        f"{COMPANY} operates from CoLab Point's established coworking and technology hub in the heart of "
        "Gujrat — a space trusted by freelancers, startups, and enterprises since 2021. That foundation "
        "means we understand real business pressure: deadlines, budgets, and the need for partners who "
        "show up.",
    )
    add_body(
        doc,
        "Our digital practice inherits the same standards that earned CoLab Point 5-star community reviews: "
        "professional environments, high-speed infrastructure, and a collaborative culture. We translate "
        "that into deliverables you can defend in boardrooms and pitch decks.",
    )
    pillars = [
        "Documented scopes and change-control — no surprise invoices.",
        "Quality assurance on every launch — devices, browsers, and performance.",
        "Ethical marketing — compliant ads, honest SEO, and brand-safe creative.",
        "Local presence, global standards — Gujrat roots with international-grade output.",
        "Training and handover — you own your assets, credentials, and data.",
    ]
    add_bullets(doc, pillars, "✓")
    add_callout(
        doc,
        "Client Success Mindset",
        "We measure our success by your retention, referrals, and revenue impact — not by shipping pages alone.",
    )


def faq(doc):
    page_break(doc)
    section_header_bar(doc, "Frequently Asked Questions", "Clear answers before you commit")
    faqs = [
        (
            "How long does a typical website project take?",
            "Basic sites often complete in 3–4 weeks; Standard and Premium timelines depend on content readiness "
            "and integrations — typically 6–12 weeks. We provide a fixed schedule after Discovery.",
        ),
        (
            "Do you provide hosting and domain services?",
            "We assist with domain registration and hosting setup (add-on). You may also use your existing provider; "
            "we configure DNS, SSL, and deployments accordingly.",
        ),
        (
            "Are ad budgets included in digital marketing fees?",
            "Management fees cover strategy, setup, and optimization. Media spend (Google/Meta) is paid directly "
            "to platforms or reimbursed per agreed policy.",
        ),
        (
            "Can we upgrade from Basic to Standard later?",
            "Yes. We scope incremental work and migrate your site without unnecessary rebuilds where possible.",
        ),
        (
            "What support is included after launch?",
            "Each website package includes a defined support window (30/60/90 days). Extended maintenance is "
            "available via our Website Maintenance add-on.",
        ),
        (
            "Who owns the website and content?",
            "You retain full ownership of domain, hosting, designs (upon final payment), and content assets.",
        ),
        (
            "Do you work with clients outside Gujrat?",
            "Absolutely. We serve clients across Pakistan and internationally via structured remote collaboration.",
        ),
    ]
    for q, a in faqs:
        add_heading(doc, q, 3)
        add_body(doc, a, 10)


def contact_page(doc):
    page_break(doc)
    section_header_bar(doc, "Contact & Next Steps", "Let's build your next chapter")
    add_body(
        doc,
        "Thank you for considering CoLab Space Point. To proceed, we recommend a complimentary discovery "
        "call to align on objectives, package selection, and timeline. We will then issue a formal "
        "statement of work and project schedule.",
    )
    t = doc.add_table(rows=6, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = [
        ("Company", COMPANY),
        ("Website", WEB),
        ("Phone", f"{PHONE1}  |  {PHONE2}"),
        ("Email", f"{EMAIL1}  |  {EMAIL2}"),
        ("Office Address", ADDRESS),
        ("Parent Brand", "CoLab Point — Coworking & Innovation Hub, Est. 2021"),
    ]
    for i, (lab, val) in enumerate(labels):
        format_table_cell(t.rows[i].cells[0], lab, bold=True, header=True)
        format_table_cell(t.rows[i].cells[1], val)
    doc.add_paragraph()
    add_callout(
        doc,
        "Proposal Validity",
        "Pricing and scope in this document are valid for 30 days from the date of issue unless otherwise agreed in writing.",
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    r = p.add_run(f"© {COMPANY}  |  {WEB}")
    style_run(r, 9, False, GRAY)


def setup_document(doc):
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(f"{COMPANY}  ·  {WEB}  ·  Confidential")
    style_run(r, 8, False, GRAY)


def main():
    out = "/workspace/CoLab_Space_Point_Digital_Agency_Proposal.docx"
    doc = Document()
    setup_document(doc)
    cover_page(doc)
    executive_summary(doc)
    table_of_contents(doc)
    company_profile(doc)
    why_choose(doc)
    website_packages(doc)
    social_media(doc)
    digital_marketing(doc)
    addons(doc)
    technology_standards(doc)
    process_timeline(doc)
    trust_section(doc)
    commercial_terms(doc)
    faq(doc)
    contact_page(doc)
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
