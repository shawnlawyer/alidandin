from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from textwrap import wrap

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
DOCS = ASSETS / "docs"
IMAGES = ASSETS / "images"


NAVY = colors.HexColor("#1E2A3A")
SAND = colors.HexColor("#C2A57A")
TERRACOTTA = colors.HexColor("#B96939")
OLIVE = colors.HexColor("#617147")
CREAM = colors.HexColor("#FFFAF3")
INK = colors.HexColor("#1F2A36")
MUTED = colors.HexColor("#65707A")


@dataclass(frozen=True)
class Section:
    title: str
    summary: str
    bullets: list[str]


BUSINESS_PLAN_SECTIONS = [
    Section(
        "Executive Summary",
        "Ali Dandin is a premium textile brand built around provenance, material intelligence, and destination-led collections. "
        "The business combines premium product curation, editorial storytelling, and direct-to-consumer commerce into a system "
        "that scales from curated drops into signature lines, private label, and selective partnerships.",
        [
            "Category focus: premium textiles rather than broad travel goods.",
            "Hero launch: Kyoto Indigo, a collection that unifies material, process, and origin.",
            "Core edge: Ali's textile background supports sourcing judgment, supplier fluency, and product credibility.",
        ],
    ),
    Section(
        "Company Overview",
        "The company is designed as a modern textile house with direct-to-consumer economics and strong editorial identity. "
        "Its mission is to bring culturally rooted textile products to market through a premium brand system that combines "
        "sourcing, design judgment, and storytelling.",
        [
            "Direct-to-consumer first, with later retail and partnership expansion.",
            "Destination-led collections rather than a commodity catalog.",
            "Editorial voice anchored in material literacy and provenance.",
        ],
    ),
    Section(
        "Market Opportunity",
        "The market opportunity sits at the overlap of premium home textiles, artisanal product commerce, and content-led "
        "direct-to-consumer brands. Buyers in premium home and lifestyle categories respond to provenance, texture, design taste, "
        "and perceived authenticity.",
        [
            "Consumers increasingly want products with origin and taste embedded in them.",
            "Textiles support healthy pricing, repeat purchasing, and private label expansion.",
            "There is whitespace for a globally literate, editorially premium textile brand.",
        ],
    ),
    Section(
        "Brand Positioning",
        "Ali Dandin is positioned as a specialized authority in premium textiles. The brand tone is precise, restrained, "
        "and materially intelligent. It reads as a textile house with global reach rather than a souvenir shop or generic lifestyle label.",
        [
            "Primary message: premium textiles shaped by place, process, and curation.",
            "Visual system: midnight blue, sand, cream, with terracotta and olive accents.",
            "Symbols: primary mark, AD monogram, passport stamp, and textile globe.",
        ],
    ),
    Section(
        "Product Strategy",
        "Kyoto Indigo serves as the opening collection because it combines material clarity, process depth, and strong cultural "
        "association in a way that translates cleanly into product, imagery, and story. The launch assortment stays narrow to maintain clarity and control.",
        [
            "Launch products: Indigo Linen Throw, Indigo Scarf, Indigo Pillow, Indigo Textile Set.",
            "Every product ties to process, place, and use.",
            "Expansion path: throws, scarves, pillows, table linens, bedding, and signature textile accessories.",
        ],
    ),
    Section(
        "Business Model",
        "The company monetizes product first, then layers in recurring and higher-margin revenue streams. Product revenue establishes the brand. "
        "Membership deepens customer ownership. Private label increases margin. Partnerships and licensing extend reach without diluting the brand.",
        [
            "Primary revenue: direct-to-consumer textile collections.",
            "Secondary revenue: membership, collaborations, affiliate income, selective licensing.",
            "Long-term revenue: private label lines, prestige retail, and distribution leverage.",
        ],
    ),
    Section(
        "Go-To-Market Strategy",
        "The go-to-market model uses content and commerce as a single system. The customer acquisition model begins with story: "
        "the destination, the material, the maker, and the product. That story is translated into an editorial storefront and a tightly curated release calendar.",
        [
            "Primary channels: Shopify, Instagram, short-form video, email capture, editorial content.",
            "Content builds trust; trust lifts conversion; launches create momentum.",
            "Brand-building content is prioritized over discount-led performance marketing.",
        ],
    ),
    Section(
        "Operations and Sourcing",
        "The operating model begins with direct sourcing and disciplined assortment. The brand identifies regions with strong textile credibility, "
        "documents the process behind the product, evaluates shipping and durability, and scales only the products that hold both brand value and operational feasibility.",
        [
            "Core regions: Japan, India, Turkey, Italy, and later Peru.",
            "Sourcing documentation covers maker, process, region, and material details.",
            "Small-batch collection discipline protects quality and working capital.",
        ],
    ),
    Section(
        "Commerce Platform",
        "The storefront presents one clear featured collection, one supporting artisan or process story, and one disciplined product grid. "
        "That hierarchy protects the premium position and prevents the site from becoming visually noisy or operationally scattered.",
        [
            "Homepage anchored by a featured collection.",
            "Collection pages organized by destination and product family.",
            "Product pages structured around material, process, and origin.",
        ],
    ),
    Section(
        "Financial Plan",
        "The planning model assumes premium price architecture and disciplined collection growth. The base case grows from collection revenue into a broader mix "
        "of product and recurring revenue streams while preserving premium gross margins.",
        [
            "Year 1 revenue: $826K.",
            "Year 2 revenue: $3.6M.",
            "Year 3 revenue: $9.352M.",
        ],
    ),
    Section(
        "Growth Roadmap",
        "The business scales in three phases: discovery, authority, and lifestyle expansion. The long-term upside is not simply product margin. "
        "It is ownership of a trusted point of view around taste, provenance, and material literacy.",
        [
            "Phase 1: destination-led launches and clear brand imagery.",
            "Phase 2: repeat collections, stronger supplier systems, and private label development.",
            "Phase 3: bedding, home textiles, licensing, partnerships, and prestige retail.",
        ],
    ),
    Section(
        "Risk Management",
        "The strongest early risks are brand concentration, supplier variability, and over-expansion. The mitigation path is to protect the textile-first thesis, "
        "grow through repeatable supplier systems, and keep category expansion disciplined.",
        [
            "Protect the textile-first positioning.",
            "Use small-batch launches to validate demand.",
            "Maintain disciplined category expansion.",
        ],
    ),
    Section(
        "Capital Strategy",
        "The company supports a staged capital path that preserves control early and invites growth capital later. The initial capital range supports product development, "
        "travel, sourcing, content production, ecommerce operations, and early team build-out.",
        [
            "Illustrative initial capital need: $35K-$70K.",
            "Later outside capital supports inventory scale, team build-out, and supplier management.",
            "The strongest financing story follows product proof and early audience traction.",
        ],
    ),
]


DECK_SLIDES = [
    {
        "title": "Ali Dandin premium textiles",
        "subtitle": "A destination-led textile brand built around provenance, material intelligence, and premium direct-to-consumer commerce.",
        "bullets": [
            "Premium textile authority with global sourcing roots",
            "Hero launch: Kyoto Indigo",
            "Scales from curated drops to signature lines and partnerships",
        ],
        "image": None,
    },
    {
        "title": "Problem",
        "subtitle": "Consumers want products with origin and taste, but most marketplaces flatten both into commodity product.",
        "bullets": [
            "Selection breadth often replaces curation",
            "Material literacy is rarely a true brand moat",
            "Trust is fragmented across generic marketplaces and broad lifestyle brands",
        ],
        "image": None,
    },
    {
        "title": "Solution",
        "subtitle": "Ali Dandin operates as a textile-led discovery brand centered on provenance, material intelligence, and destination-led collections.",
        "bullets": [
            "One destination-led collection at a time",
            "One strong material story at a time",
            "One premium editorial storefront that translates story into commerce",
        ],
        "image": IMAGES / "logo_suite.png",
    },
    {
        "title": "Collection Architecture",
        "subtitle": "Kyoto Indigo establishes the opening collection and the brand's commercial logic.",
        "bullets": [
            "Launch products: throw, scarf, pillow, textile set",
            "Narrow assortment improves clarity and control",
            "Material, process, and origin form one coherent product story",
        ],
        "image": IMAGES / "kyoto_indigo_collection.png",
    },
    {
        "title": "Commerce System",
        "subtitle": "The storefront is editorial, premium, and collection-led rather than broad catalog commerce.",
        "bullets": [
            "Homepage anchored by a featured collection",
            "Collection pages organized by destination and product family",
            "Product pages structured around material, process, and origin",
        ],
        "image": IMAGES / "shopify_homepage.png",
    },
    {
        "title": "Sourcing Network",
        "subtitle": "Regional sourcing is a product strategy and a brand strategy at the same time.",
        "bullets": [
            "Japan: indigo and heritage process",
            "India: block print, handloom, flexible expansion",
            "Turkey and Italy: scalable production and luxury finishing",
        ],
        "image": IMAGES / "global_sourcing_overview.png",
    },
    {
        "title": "Business Model",
        "subtitle": "Product revenue establishes the brand, then higher-leverage revenue streams extend it.",
        "bullets": [
            "Primary revenue: direct-to-consumer product sales",
            "Secondary revenue: membership, collaborations, affiliate income",
            "Expansion revenue: private label lines, licensing, prestige retail",
        ],
        "image": None,
    },
    {
        "title": "Financial Profile",
        "subtitle": "The economic model is built on premium ASPs, disciplined assortment, and content-led acquisition.",
        "bullets": [
            "Year 1 revenue: $826K",
            "Year 2 revenue: $3.6M",
            "Year 3 revenue: $9.352M",
        ],
        "image": None,
    },
    {
        "title": "Growth Roadmap",
        "subtitle": "The company scales in three phases: discovery, authority, and lifestyle expansion.",
        "bullets": [
            "Phase 1: destination-led launches and narrow assortment",
            "Phase 2: repeat collections and private label development",
            "Phase 3: bedding, home textiles, partnerships, and licensing",
        ],
        "image": None,
    },
    {
        "title": "Capital Strategy",
        "subtitle": "The brand supports a staged capital path that preserves control early and attracts outside capital after proof.",
        "bullets": [
            "Initial capital range: $35K-$70K",
            "Use of funds: product, sourcing, content, ecommerce operations",
            "Later growth capital supports team build-out and inventory scale",
        ],
        "image": None,
    },
]


def ensure_dirs() -> None:
    DOCS.mkdir(parents=True, exist_ok=True)


def build_pdf() -> None:
    output = DOCS / "investor_business_plan.pdf"
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="PlanTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=26,
            leading=32,
            alignment=TA_CENTER,
            textColor=NAVY,
            spaceAfter=18,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionTitle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=NAVY,
            spaceBefore=6,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="BodyCopy",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.5,
            leading=15,
            textColor=INK,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="BulletCopy",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.3,
            leading=14,
            leftIndent=14,
            bulletIndent=2,
            textColor=INK,
            spaceAfter=4,
        )
    )

    story = []
    story.append(Spacer(1, 0.35 * inch))
    story.append(Paragraph("Ali Dandin Business Plan", styles["PlanTitle"]))
    story.append(
        Paragraph(
            "Premium textile brand built around provenance, material intelligence, destination-led collections, and editorial commerce.",
            ParagraphStyle(
                "Subtitle",
                parent=styles["BodyText"],
                fontName="Helvetica",
                fontSize=12,
                leading=17,
                alignment=TA_CENTER,
                textColor=MUTED,
                spaceAfter=18,
            ),
        )
    )

    cover_image = IMAGES / "logo_suite.png"
    if cover_image.exists():
        story.append(Image(str(cover_image), width=6.8 * inch, height=4.0 * inch))
        story.append(Spacer(1, 0.18 * inch))

    summary_table = Table(
        [
            ["Category", "Premium textiles"],
            ["Launch collection", "Kyoto Indigo"],
            ["Business model", "Collections, membership, private label, partnerships"],
            ["Planning horizon", "3 years with long-term expansion roadmap"],
        ],
        colWidths=[1.8 * inch, 4.5 * inch],
    )
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7D0C5")),
                ("TEXTCOLOR", (0, 0), (0, -1), NAVY),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("LEADING", (0, 0), (-1, -1), 13),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2EBDF")),
                ("ROWBACKGROUNDS", (1, 0), (1, -1), [colors.white, colors.HexColor("#FBF7F1")]),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(summary_table)
    story.append(PageBreak())

    section_images = {
        "Brand Positioning": IMAGES / "logo_suite.png",
        "Product Strategy": IMAGES / "kyoto_indigo_collection.png",
        "Operations and Sourcing": IMAGES / "global_sourcing_overview.png",
        "Commerce Platform": IMAGES / "shopify_homepage.png",
    }

    for idx, section in enumerate(BUSINESS_PLAN_SECTIONS):
        story.append(Paragraph(section.title, styles["SectionTitle"]))
        story.append(Paragraph(section.summary, styles["BodyCopy"]))
        image_path = section_images.get(section.title)
        if image_path and image_path.exists():
            story.append(Image(str(image_path), width=6.5 * inch, height=3.8 * inch))
            story.append(Spacer(1, 0.12 * inch))
        for bullet in section.bullets:
            story.append(Paragraph(bullet, styles["BulletCopy"], bulletText="•"))
        story.append(Spacer(1, 0.08 * inch))
        if idx in {2, 5, 8, 10}:
            story.append(PageBreak())

    revenue_table = Table(
        [
            ["Metric", "Year 1", "Year 2", "Year 3"],
            ["Audience size", "50,000", "250,000", "1,000,000"],
            ["Collections launched", "4", "8", "12"],
            ["Average selling price", "$65", "$70", "$72"],
            ["Product revenue", "$650,000", "$2,800,000", "$6,912,000"],
            ["Membership + partnerships", "$176,000", "$800,000", "$2,440,000"],
            ["Total revenue", "$826,000", "$3,600,000", "$9,352,000"],
        ],
        colWidths=[2.2 * inch, 1.35 * inch, 1.35 * inch, 1.35 * inch],
    )
    revenue_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7D0C5")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FBF7F1")]),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(PageBreak())
    story.append(Paragraph("Financial Snapshot", styles["SectionTitle"]))
    story.append(
        Paragraph(
            "The planning case assumes disciplined assortment growth, premium pricing, and a revenue mix that broadens over time through membership, partnerships, and private label.",
            styles["BodyCopy"],
        )
    )
    story.append(revenue_table)

    doc = SimpleDocTemplate(
        str(output),
        pagesize=letter,
        rightMargin=0.65 * inch,
        leftMargin=0.65 * inch,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
        title="Ali Dandin Business Plan",
        author="OpenAI Codex",
    )
    doc.build(story)


def add_textbox(slide, left, top, width, height, text, font_size, color, bold=False, name=None):
    tx = slide.shapes.add_textbox(left, top, width, height)
    if name:
        tx.name = name
    frame = tx.text_frame
    frame.clear()
    p = frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    return tx


def build_pptx() -> None:
    output = DOCS / "investor_deck.pptx"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    navy_rgb = RGBColor(0x1E, 0x2A, 0x3A)
    sand_rgb = RGBColor(0xC2, 0xA5, 0x7A)
    cream_rgb = RGBColor(0xFF, 0xFA, 0xF3)
    muted_rgb = RGBColor(0x65, 0x70, 0x7A)

    blank = prs.slide_layouts[6]

    for i, spec in enumerate(DECK_SLIDES):
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = cream_rgb if i else navy_rgb

        if i == 0:
            add_textbox(slide, Inches(0.8), Inches(0.6), Inches(6.4), Inches(0.35), "ALI DANDIN INVESTOR DECK", 11, sand_rgb, True)
            add_textbox(slide, Inches(0.8), Inches(1.15), Inches(6.6), Inches(1.6), spec["title"], 28, cream_rgb, True)
            add_textbox(slide, Inches(0.8), Inches(2.45), Inches(6.2), Inches(1.2), spec["subtitle"], 15, cream_rgb)
            box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(7.3), Inches(1.05), Inches(5.1), Inches(4.45))
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            box.fill.transparency = 0.9
            box.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            tf = box.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = "Investment thesis"
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = navy_rgb
            for bullet in spec["bullets"]:
                p = tf.add_paragraph()
                p.text = bullet
                p.level = 0
                p.font.size = Pt(15)
                p.font.color.rgb = navy_rgb
            continue

        add_textbox(slide, Inches(0.7), Inches(0.55), Inches(1.5), Inches(0.3), f"{i+1:02d}", 11, sand_rgb, True)
        add_textbox(slide, Inches(0.7), Inches(0.95), Inches(5.9), Inches(0.9), spec["title"], 26, navy_rgb, True)
        add_textbox(slide, Inches(0.7), Inches(1.75), Inches(5.8), Inches(1.15), spec["subtitle"], 14, muted_rgb)

        if spec["image"] and Path(spec["image"]).exists():
            slide.shapes.add_picture(str(spec["image"]), Inches(7.1), Inches(1.0), width=Inches(5.5), height=Inches(3.9))
            bullet_box_left = Inches(0.7)
            bullet_box_width = Inches(5.8)
        else:
            bullet_box_left = Inches(0.7)
            bullet_box_width = Inches(12.0)

        bullet_box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, bullet_box_left, Inches(3.05), bullet_box_width, Inches(3.5))
        bullet_box.fill.solid()
        bullet_box.fill.fore_color.rgb = RGBColor(0xFA, 0xF4, 0xEA)
        bullet_box.line.color.rgb = RGBColor(0xE3, 0xD8, 0xC8)
        tf = bullet_box.text_frame
        tf.clear()
        for idx, bullet in enumerate(spec["bullets"]):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.size = Pt(18 if idx == 0 else 16)
            p.font.bold = idx == 0
            p.font.color.rgb = navy_rgb
            p.space_after = Pt(10)

    prs.core_properties.title = "Ali Dandin Investor Deck"
    prs.core_properties.author = "OpenAI Codex"
    prs.save(str(output))


def build_xlsx() -> None:
    output = DOCS / "financial_model.xlsx"
    wb = Workbook()
    wb.remove(wb.active)

    navy_fill = PatternFill("solid", fgColor="1E2A3A")
    sand_fill = PatternFill("solid", fgColor="F2E7D6")
    green_fill = PatternFill("solid", fgColor="EEF4E7")
    border = Border(
        left=Side(style="thin", color="D7D0C5"),
        right=Side(style="thin", color="D7D0C5"),
        top=Side(style="thin", color="D7D0C5"),
        bottom=Side(style="thin", color="D7D0C5"),
    )

    def style_header(cell, dark=True):
        cell.font = Font(bold=True, color="FFFFFF" if dark else "1E2A3A")
        cell.fill = navy_fill if dark else sand_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def style_cell(cell, bold=False, fill=None, number_format=None):
        cell.font = Font(bold=bold, color="1F2A36")
        if fill:
            cell.fill = fill
        cell.border = border
        if number_format:
            cell.number_format = number_format
        cell.alignment = Alignment(vertical="center")

    ws = wb.create_sheet("Assumptions")
    ws.append(["Assumption", "Value", "Notes"])
    assumptions = [
        ("Audience Year 1", 50000, "Combined social and owned audience"),
        ("Audience Year 2", 250000, "Scaled through content and launches"),
        ("Audience Year 3", 1000000, "Brand authority and partnerships"),
        ("Collections Year 1", 4, "Quarterly launch cadence"),
        ("Collections Year 2", 8, "Bi-monthly cadence"),
        ("Collections Year 3", 12, "Monthly launch cadence"),
        ("ASP Year 1", 65, "Premium entry assortment"),
        ("ASP Year 2", 70, "Expanded premium mix"),
        ("ASP Year 3", 72, "Private label and richer assortment"),
        ("Gross Margin Target Low", 0.55, "Conservative blended margin"),
        ("Gross Margin Target High", 0.65, "Private label upside"),
        ("Membership Price", 8, "Passport Club monthly price"),
    ]
    for row in assumptions:
        ws.append(list(row))
    for cell in ws[1]:
        style_header(cell)
    for row in ws.iter_rows(min_row=2):
        for idx, cell in enumerate(row):
            style_cell(cell)
            if idx == 1 and isinstance(cell.value, (int, float)):
                cell.number_format = "$#,##0.00" if "ASP" in row[0].value or "Price" in row[0].value else "0.00%"
                if "Audience" in row[0].value or "Collections" in row[0].value:
                    cell.number_format = "#,##0"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 36

    ws = wb.create_sheet("Revenue")
    revenue_rows = [
        ["Metric", "Year 1", "Year 2", "Year 3"],
        ["Collections launched", 4, 8, 12],
        ["Total units sold", 10000, 40000, 96000],
        ["Average selling price", 65, 70, 72],
        ["Product revenue", 650000, 2800000, 6912000],
        ["Membership revenue", 96000, 480000, 1440000],
        ["Partnership + affiliate", 80000, 320000, 1000000],
        ["Total revenue", 826000, 3600000, 9352000],
    ]
    for row in revenue_rows:
        ws.append(row)
    for cell in ws[1]:
        style_header(cell)
    for row in ws.iter_rows(min_row=2):
        label = row[0].value
        for idx, cell in enumerate(row):
            style_cell(cell, bold=label == "Total revenue", fill=green_fill if label == "Total revenue" else None)
            if idx > 0:
                if "revenue" in str(label).lower() or "price" in str(label).lower() or "affiliate" in str(label).lower():
                    cell.number_format = "$#,##0"
                else:
                    cell.number_format = "#,##0"
    for col, width in zip("ABCD", [30, 16, 16, 16]):
        ws.column_dimensions[col].width = width

    ws = wb.create_sheet("COGS_Opex")
    rows = [
        ["Metric", "Year 1", "Year 2", "Year 3"],
        ["COGS", 280000, 1120000, 2688000],
        ["Travel and sourcing", 60000, 85000, 120000],
        ["Team", 150000, 450000, 900000],
        ["Marketing and growth", 82600, 432000, 1028720],
        ["General operations", 75000, 180000, 350000],
        ["Total operating costs", 367600, 1147000, 2398720],
    ]
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        style_header(cell)
    for row in ws.iter_rows(min_row=2):
        label = row[0].value
        for idx, cell in enumerate(row):
            style_cell(cell, bold=label == "Total operating costs", fill=green_fill if label == "Total operating costs" else None)
            if idx > 0:
                cell.number_format = "$#,##0"
    for col, width in zip("ABCD", [30, 16, 16, 16]):
        ws.column_dimensions[col].width = width

    ws = wb.create_sheet("Profitability")
    rows = [
        ["Metric", "Year 1", "Year 2", "Year 3"],
        ["Total revenue", 826000, 3600000, 9352000],
        ["COGS", 280000, 1120000, 2688000],
        ["Gross profit", 546000, 2480000, 6664000],
        ["Total operating costs", 367600, 1147000, 2398720],
        ["Indicative operating profit", 178400, 1333000, 4265280],
    ]
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        style_header(cell)
    for row in ws.iter_rows(min_row=2):
        label = row[0].value
        for idx, cell in enumerate(row):
            style_cell(cell, bold=label in {"Gross profit", "Indicative operating profit"}, fill=green_fill if label in {"Gross profit", "Indicative operating profit"} else None)
            if idx > 0:
                cell.number_format = "$#,##0"
    for col, width in zip("ABCD", [30, 16, 16, 16]):
        ws.column_dimensions[col].width = width

    ws = wb.create_sheet("Scenario_View")
    rows = [
        ["Scenario", "Year 3 Revenue", "Commentary"],
        ["Base case", 9352000, "Current planning case"],
        ["Upside case", 12000000, "Faster repeat purchase and earlier private label expansion"],
        ["Downside case", 6500000, "Slower audience growth and lower sell-through"],
    ]
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        style_header(cell)
    for row in ws.iter_rows(min_row=2):
        for idx, cell in enumerate(row):
            style_cell(cell)
            if idx == 1:
                cell.number_format = "$#,##0"
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 56

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

    wb.save(str(output))


def main() -> None:
    ensure_dirs()
    build_pdf()
    build_pptx()
    build_xlsx()


if __name__ == "__main__":
    main()
