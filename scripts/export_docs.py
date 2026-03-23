from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image as PILImage
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.shapes import Circle, Drawing, Line, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
DOCS = ASSETS / "docs"
IMAGES = ASSETS / "images"


NAVY = colors.HexColor("#1E2A3A")
SAND = colors.HexColor("#C2A57A")
SAND_SOFT = colors.HexColor("#EBDCC6")
TERRACOTTA = colors.HexColor("#B96939")
OLIVE = colors.HexColor("#617147")
CREAM = colors.HexColor("#FFFAF3")
INK = colors.HexColor("#1F2A36")
MUTED = colors.HexColor("#65707A")
LINE = colors.HexColor("#D8CFC2")
WHITE = colors.white


YEAR_LABELS = ["Year 1", "Year 2", "Year 3"]
REVENUE = [826000, 3600000, 9352000]
GROSS_PROFIT = [546000, 2480000, 6664000]
OPERATING_PROFIT = [178400, 1333000, 4265280]


@dataclass(frozen=True)
class ProductCard:
    name: str
    note: str


@dataclass(frozen=True)
class RegionCard:
    name: str
    note: str


PRODUCTS = [
    ProductCard("Indigo Linen Throw", "Hero product with premium texture, gifting appeal, and strong editorial presence."),
    ProductCard("Indigo Scarf", "Portable entry point that supports travel, styling, and repeat purchase."),
    ProductCard("Indigo Pillow", "Home category anchor that translates textile story into everyday living."),
    ProductCard("Indigo Textile Set", "Study set that turns process, swatch, and provenance into a collectible object."),
]

REGIONS = [
    RegionCard("Japan", "Indigo process, heritage finishing, and the Kyoto Indigo narrative foundation."),
    RegionCard("India", "Block print, handloom depth, flexible development capacity, and future assortment breadth."),
    RegionCard("Turkey", "Cotton scale, towels, throws, and production fluency for broader textile programs."),
    RegionCard("Italy", "Luxury finishing, linen credibility, and eventual prestige textile manufacturing."),
    RegionCard("Peru", "Future expansion into alpaca and rich fiber stories for premium soft goods."),
]


def money_short(value: float) -> str:
    if value >= 1_000_000:
        return f"${value / 1_000_000:.1f}M"
    if value >= 1_000:
        return f"${value / 1_000:.0f}K"
    return f"${value:,.0f}"


def ensure_dirs() -> None:
    DOCS.mkdir(parents=True, exist_ok=True)


def image_dimensions(path: Path) -> tuple[int, int]:
    with PILImage.open(path) as img:
        return img.size


def fitted_image(path: Path, max_width: float, max_height: float, h_align: str = "CENTER") -> Image:
    width_px, height_px = ImageReader(str(path)).getSize()
    scale = min(max_width / width_px, max_height / height_px)
    flowable = Image(str(path), width=width_px * scale, height=height_px * scale)
    flowable.hAlign = h_align
    return flowable


def build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CoverKicker",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10.5,
            leading=12,
            textColor=SAND,
            spaceAfter=10,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CoverTitle",
            parent=styles["Title"],
            fontName="Times-Bold",
            fontSize=30,
            leading=34,
            textColor=NAVY,
            alignment=TA_CENTER,
            spaceAfter=14,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CoverBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=12.5,
            leading=18,
            textColor=MUTED,
            alignment=TA_CENTER,
            spaceAfter=14,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionKicker",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9.5,
            leading=11,
            textColor=TERRACOTTA,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionTitle",
            parent=styles["Heading1"],
            fontName="Times-Bold",
            fontSize=22,
            leading=26,
            textColor=NAVY,
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="BodyCopy",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.7,
            leading=15,
            textColor=INK,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="BodyCopyMuted",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.2,
            leading=14,
            textColor=MUTED,
            spaceAfter=4,
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
    styles.add(
        ParagraphStyle(
            name="CardTitle",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=11.3,
            leading=13,
            textColor=NAVY,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CardBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.7,
            leading=13.2,
            textColor=INK,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="StatValue",
            parent=styles["BodyText"],
            fontName="Times-Bold",
            fontSize=21,
            leading=23,
            textColor=NAVY,
            alignment=TA_CENTER,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="StatLabel",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.8,
            leading=10.5,
            textColor=MUTED,
            alignment=TA_CENTER,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="MiniLabel",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=10,
            textColor=MUTED,
            alignment=TA_CENTER,
        )
    )
    return styles


def card(content, *, background=WHITE, border_color=LINE, padding=12, valign="TOP") -> Table:
    table = Table([[content]])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), background),
                ("BOX", (0, 0), (-1, -1), 0.8, border_color),
                ("INNERPADDING", (0, 0), (-1, -1), padding),
                ("VALIGN", (0, 0), (-1, -1), valign),
                ("ROUNDEDCORNERS", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return table


def stat_card(value: str, label: str, styles) -> Table:
    return card(
        [
            Paragraph(value, styles["StatValue"]),
            Paragraph(label, styles["StatLabel"]),
        ],
        background=WHITE,
        padding=10,
        valign="MIDDLE",
    )


def narrative_card(title: str, body: str, styles, *, background=WHITE) -> Table:
    return card(
        [
            Paragraph(title, styles["CardTitle"]),
            Paragraph(body, styles["CardBody"]),
        ],
        background=background,
        padding=12,
    )


def bullets(items: list[str], styles) -> list[Paragraph]:
    return [Paragraph(item, styles["BulletCopy"], bulletText="•") for item in items]


def section_heading(kicker: str, title: str, body: str, styles) -> list:
    return [
        Paragraph(kicker, styles["SectionKicker"]),
        Paragraph(title, styles["SectionTitle"]),
        Paragraph(body, styles["BodyCopy"]),
    ]


def revenue_chart(width: float = 6.7 * inch, height: float = 2.8 * inch) -> Drawing:
    drawing = Drawing(width, height)
    drawing.add(Rect(0, 0, width, height, fillColor=WHITE, strokeColor=LINE, rx=12, ry=12))

    chart = VerticalBarChart()
    chart.x = 34
    chart.y = 34
    chart.width = width - 68
    chart.height = height - 74
    chart.data = [[value / 1_000_000 for value in REVENUE], [value / 1_000_000 for value in GROSS_PROFIT]]
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = 10
    chart.valueAxis.valueStep = 2
    chart.valueAxis.labels.fontName = "Helvetica"
    chart.valueAxis.labels.fontSize = 7
    chart.valueAxis.labels.fillColor = MUTED
    chart.categoryAxis.categoryNames = YEAR_LABELS
    chart.categoryAxis.labels.fontName = "Helvetica-Bold"
    chart.categoryAxis.labels.fontSize = 8
    chart.categoryAxis.labels.fillColor = NAVY
    chart.categoryAxis.strokeColor = colors.transparent
    chart.valueAxis.strokeColor = colors.transparent
    chart.bars[0].fillColor = NAVY
    chart.bars[0].strokeColor = NAVY
    chart.bars[1].fillColor = TERRACOTTA
    chart.bars[1].strokeColor = TERRACOTTA
    chart.barWidth = 18
    chart.groupSpacing = 16
    chart.barSpacing = 6
    chart.categoryAxis.tickDown = 0
    drawing.add(chart)

    drawing.add(String(18, height - 18, "Revenue and gross profit ($M)", fontName="Helvetica-Bold", fontSize=10, fillColor=NAVY))
    drawing.add(Rect(width - 150, height - 22, 8, 8, fillColor=NAVY, strokeColor=NAVY))
    drawing.add(String(width - 136, height - 20, "Revenue", fontName="Helvetica", fontSize=8, fillColor=MUTED))
    drawing.add(Rect(width - 82, height - 22, 8, 8, fillColor=TERRACOTTA, strokeColor=TERRACOTTA))
    drawing.add(String(width - 68, height - 20, "Gross profit", fontName="Helvetica", fontSize=8, fillColor=MUTED))
    return drawing


def roadmap_chart(width: float = 6.7 * inch, height: float = 2.05 * inch) -> Drawing:
    drawing = Drawing(width, height)
    drawing.add(Rect(0, 0, width, height, fillColor=WHITE, strokeColor=LINE, rx=12, ry=12))
    y = height / 2
    x_positions = [width * 0.17, width * 0.5, width * 0.83]
    drawing.add(Line(46, y, width - 46, y, strokeColor=SAND, strokeWidth=3))
    phases = [
        ("Discovery", "Years 1-2", "Destination-led launches and clear brand imagery"),
        ("Authority", "Years 3-4", "Repeat collections, supplier depth, private label"),
        ("Expansion", "Years 5+", "Home textiles, prestige retail, licensing"),
    ]
    for x, (title, period, note) in zip(x_positions, phases):
        drawing.add(Circle(x, y, 13, fillColor=NAVY, strokeColor=NAVY))
        drawing.add(Circle(x, y, 5, fillColor=SAND, strokeColor=SAND))
        drawing.add(String(x - 34, y + 22, title, fontName="Helvetica-Bold", fontSize=10, fillColor=NAVY))
        drawing.add(String(x - 34, y + 10, period, fontName="Helvetica", fontSize=8, fillColor=TERRACOTTA))
        drawing.add(String(x - 64, y - 30, note, fontName="Helvetica", fontSize=7.3, fillColor=MUTED))
    return drawing


def financial_table(styles) -> Table:
    rows = [
        ["Metric", "Year 1", "Year 2", "Year 3"],
        ["Total revenue", money_short(REVENUE[0]), money_short(REVENUE[1]), money_short(REVENUE[2])],
        ["Gross profit", money_short(GROSS_PROFIT[0]), money_short(GROSS_PROFIT[1]), money_short(GROSS_PROFIT[2])],
        ["Operating profit", money_short(OPERATING_PROFIT[0]), money_short(OPERATING_PROFIT[1]), money_short(OPERATING_PROFIT[2])],
        ["Collections launched", "4", "8", "12"],
    ]
    table = Table(rows, colWidths=[2.0 * inch, 1.4 * inch, 1.4 * inch, 1.4 * inch])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 1), (-1, -1), WHITE),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#FBF6EF")]),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("TEXTCOLOR", (0, 1), (-1, -1), INK),
                ("GRID", (0, 0), (-1, -1), 0.6, LINE),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ]
        )
    )
    return table


def unit_economics_table() -> Table:
    rows = [
        ["Unit economics", "Value"],
        ["Average selling price", "$65-$72"],
        ["Gross margin target", "55%-65%"],
        ["Membership price", "$8 / month"],
        ["Initial capital range", "$35K-$70K"],
    ]
    table = Table(rows, colWidths=[2.5 * inch, 1.9 * inch])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), SAND_SOFT),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 0), (-1, 0), NAVY),
                ("GRID", (0, 0), (-1, -1), 0.6, LINE),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#FBF6EF")]),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return table


def draw_first_page(canvas, doc) -> None:
    canvas.saveState()
    width, height = letter
    canvas.setFillColor(CREAM)
    canvas.rect(0, 0, width, height, stroke=0, fill=1)
    canvas.setFillColor(NAVY)
    canvas.rect(0, height - 1.6 * inch, width, 1.6 * inch, stroke=0, fill=1)
    canvas.setFillColor(SAND)
    canvas.circle(width - 0.95 * inch, height - 0.78 * inch, 0.22 * inch, stroke=0, fill=1)
    canvas.setFillColor(TERRACOTTA)
    canvas.circle(0.95 * inch, 0.9 * inch, 0.12 * inch, stroke=0, fill=1)
    canvas.restoreState()


def draw_later_pages(canvas, doc) -> None:
    canvas.saveState()
    width, height = letter
    canvas.setStrokeColor(LINE)
    canvas.line(doc.leftMargin, height - 0.52 * inch, width - doc.rightMargin, height - 0.52 * inch)
    canvas.setFillColor(NAVY)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawString(doc.leftMargin, height - 0.38 * inch, "ALI DANDIN PREMIUM TEXTILES")
    canvas.setFillColor(MUTED)
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(width - doc.rightMargin, 0.42 * inch, f"{canvas.getPageNumber()}")
    canvas.restoreState()


def build_pdf() -> None:
    output = DOCS / "investor_business_plan.pdf"
    styles = build_styles()
    story: list = []

    cover_left = [
        Paragraph("Premium textile brand", styles["CoverKicker"]),
        Paragraph("Ali Dandin Business Plan", styles["CoverTitle"]),
        Paragraph(
            "Ali Dandin defines a focused position in global textile commerce. The brand centers premium textiles, destination-led collections, and disciplined sourcing inside an editorial direct-to-consumer model.",
            styles["CoverBody"],
        ),
        Spacer(1, 0.1 * inch),
        Paragraph(
            "Kyoto Indigo establishes the opening collection, while the broader company scales through signature lines, private label, and selective partnerships.",
            styles["BodyCopy"],
        ),
    ]
    cover_image = fitted_image(IMAGES / "kyoto_indigo_collection.png", 2.7 * inch, 4.8 * inch)
    cover_table = Table([[cover_left, card([cover_image], background=colors.HexColor("#F7F0E5"), padding=10)]], colWidths=[4.45 * inch, 2.15 * inch])
    cover_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))

    story.append(Spacer(1, 0.55 * inch))
    story.append(cover_table)
    story.append(Spacer(1, 0.22 * inch))

    metrics = Table(
        [[
            stat_card("Kyoto Indigo", "Opening collection", styles),
            stat_card("$826K", "Year 1 revenue", styles),
            stat_card("$9.4M", "Year 3 revenue", styles),
            stat_card("55%-65%", "Gross margin target", styles),
        ]],
        colWidths=[1.68 * inch, 1.4 * inch, 1.4 * inch, 1.4 * inch],
    )
    metrics.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story.append(metrics)
    story.append(PageBreak())

    story.extend(section_heading("Investment case", "A premium textile authority with clear commercial logic", "Ali Dandin is positioned as a specialized authority in premium textiles. The thesis is simple: provenance, material literacy, and destination-led curation create a cleaner premium brand than broad travel commerce or generic artisan marketplaces.", styles))
    thesis_cards = Table(
        [[
            narrative_card("Category edge", "Premium textiles combine tactile differentiation, repeat purchase, healthy price architecture, and eventual private label expansion.", styles),
            narrative_card("Brand edge", "Ali's background in textiles brings sourcing fluency, material judgment, and commercial credibility that most travel-led brands do not have.", styles),
            narrative_card("Growth edge", "A tightly edited collection structure makes the brand legible to customers, retailers, and future capital partners.", styles),
        ]],
        colWidths=[2.12 * inch, 2.12 * inch, 2.12 * inch],
    )
    thesis_cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(thesis_cards)
    story.append(Spacer(1, 0.16 * inch))
    story.append(revenue_chart())
    story.append(Spacer(1, 0.14 * inch))
    story.extend(
        bullets(
            [
                "Product revenue establishes the brand; recurring revenue and private label extend it.",
                "Destination-led collections build memory structure and keep the assortment disciplined.",
                "The value of the company grows as the brand moves from curated discovery into a repeatable textile system.",
            ],
            styles,
        )
    )
    story.append(PageBreak())

    story.extend(section_heading("Brand system", "Identity, tone, and visual authority", "The identity system includes a primary mark, monogram, passport stamp, and textile globe. The design language is restrained and editorial: midnight blue, sand, cream, terracotta, and olive support a premium textile house rather than a generic lifestyle label.", styles))
    brand_layout = Table(
        [[
            [
                Paragraph("Positioning", styles["CardTitle"]),
                Paragraph("Ali Dandin premium textiles", styles["BodyCopy"]),
                Paragraph("Voice", styles["CardTitle"]),
                Paragraph("Measured, exact, and materially intelligent. The brand speaks from judgment rather than hype.", styles["CardBody"]),
                Paragraph("System", styles["CardTitle"]),
                Paragraph("Primary mark, AD monogram, passport stamp, and textile globe anchor the visual architecture.", styles["CardBody"]),
            ],
            card([fitted_image(IMAGES / "logo_suite.png", 3.0 * inch, 3.6 * inch)], background=colors.HexColor("#F7F0E5"), padding=10),
        ]],
        colWidths=[3.45 * inch, 3.15 * inch],
    )
    brand_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    palette_row = Table(
        [[
            stat_card("Midnight Blue", "#1E2A3A", styles),
            stat_card("Sand", "#C2A57A", styles),
            stat_card("Cream", "#FFFAF3", styles),
            stat_card("Terracotta / Olive", "Accent palette", styles),
        ]],
        colWidths=[1.65 * inch, 1.45 * inch, 1.45 * inch, 1.95 * inch],
    )
    palette_row.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(brand_layout)
    story.append(Spacer(1, 0.16 * inch))
    story.append(palette_row)
    story.append(PageBreak())

    story.extend(section_heading("Collection architecture", "Kyoto Indigo as the opening system", "Kyoto Indigo establishes the opening collection, combining material, dye process, and cultural origin into a coherent commercial system. The assortment is narrow by design so the brand remains legible and premium from day one.", styles))
    collection_layout = Table(
        [[
            card([fitted_image(IMAGES / "kyoto_indigo_collection.png", 2.7 * inch, 4.0 * inch)], background=colors.HexColor("#F7F0E5"), padding=10),
            [
                Paragraph("Collection logic", styles["CardTitle"]),
                Paragraph("One destination. One material story. One restrained assortment that can be photographed, merchandised, and repeated without dilution.", styles["CardBody"]),
                Spacer(1, 0.06 * inch),
                Paragraph("Opening assortment", styles["CardTitle"]),
            ]
        ]],
        colWidths=[2.9 * inch, 3.7 * inch],
    )
    collection_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    product_grid = Table(
        [[
            narrative_card(product.name, product.note, styles, background=WHITE) for product in PRODUCTS[:2]
        ], [
            narrative_card(product.name, product.note, styles, background=WHITE) for product in PRODUCTS[2:]
        ]],
        colWidths=[3.18 * inch, 3.18 * inch],
    )
    product_grid.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(collection_layout)
    story.append(Spacer(1, 0.14 * inch))
    story.append(product_grid)
    story.append(PageBreak())

    story.extend(section_heading("Commerce platform", "Editorial storefront, disciplined hierarchy", "The storefront is collection-led rather than catalog-led. It presents one featured collection, one supporting story about process or maker, and one tightly edited product grid. That structure keeps the brand premium and intelligible.", styles))
    commerce_layout = Table(
        [[
            [
                Paragraph("Navigation", styles["CardTitle"]),
                Paragraph("Home, Collections, Journal, About, and Passport Club form the core public structure.", styles["CardBody"]),
                Paragraph("Merchandising", styles["CardTitle"]),
                Paragraph("The homepage leads with the collection hero, then reinforces the process story before moving into product cards.", styles["CardBody"]),
                Paragraph("Conversion", styles["CardTitle"]),
                Paragraph("Product pages elevate material, process, and origin so the brand sells judgment, not discounting.", styles["CardBody"]),
            ],
            card([fitted_image(IMAGES / "shopify_homepage.png", 2.9 * inch, 4.1 * inch)], background=colors.HexColor("#F7F0E5"), padding=10),
        ]],
        colWidths=[3.45 * inch, 3.15 * inch],
    )
    commerce_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    commerce_cards = Table(
        [[
            narrative_card("Editorial engine", "Journal, short-form video, and launch stories create awareness without reducing the brand to performance marketing alone.", styles),
            narrative_card("Membership layer", "Passport Club creates recurring revenue and deepens early access, storytelling, and collection ownership.", styles),
            narrative_card("Private label path", "Once repeat demand is proven, signature textile lines support higher margin and broader distribution.", styles),
        ]],
        colWidths=[2.12 * inch, 2.12 * inch, 2.12 * inch],
    )
    commerce_cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(commerce_layout)
    story.append(Spacer(1, 0.14 * inch))
    story.append(commerce_cards)
    story.append(PageBreak())

    story.extend(section_heading("Sourcing model", "Regional depth and disciplined supplier logic", "Sourcing is a brand expression as much as an operating function. Japan establishes the opening story, while India, Turkey, Italy, and later Peru create a broader network across handcraft, scale, and luxury finishing.", styles))
    story.append(card([fitted_image(IMAGES / "global_sourcing_overview.png", 6.4 * inch, 3.6 * inch)], background=colors.HexColor("#F7F0E5"), padding=10))
    story.append(Spacer(1, 0.16 * inch))
    region_rows = []
    for idx in range(0, len(REGIONS), 2):
        pair = REGIONS[idx: idx + 2]
        region_rows.append([narrative_card(region.name, region.note, styles) for region in pair] + ([Spacer(0, 0)] if len(pair) == 1 else []))
    region_grid = Table(region_rows, colWidths=[3.18 * inch, 3.18 * inch])
    region_grid.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(region_grid)
    story.append(PageBreak())

    story.extend(section_heading("Financial architecture", "Premium pricing, disciplined assortment, layered revenue", "The base case grows from product-led revenue into a broader mix of membership, collaborations, and eventually private label. The model stays strongest when the assortment remains curated and the positioning remains textile-first.", styles))
    finance_layout = Table(
        [[financial_table(styles), unit_economics_table()]],
        colWidths=[4.0 * inch, 2.6 * inch],
    )
    finance_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(finance_layout)
    story.append(Spacer(1, 0.18 * inch))
    finance_cards = Table(
        [[
            stat_card(money_short(GROSS_PROFIT[2]), "Year 3 gross profit", styles),
            stat_card(money_short(OPERATING_PROFIT[2]), "Year 3 operating profit", styles),
            stat_card("$8 / mo", "Membership pricing", styles),
        ]],
        colWidths=[2.1 * inch, 2.1 * inch, 2.1 * inch],
    )
    finance_cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(finance_cards)
    story.append(PageBreak())

    story.extend(section_heading("Scale path", "From destination-led discovery into a modern textile house", "The company scales from discovery into authority and into systems that support long-term growth. The expansion path moves through repeat collections, private label, prestige partnerships, and eventual home textile breadth.", styles))
    story.append(roadmap_chart())
    story.append(Spacer(1, 0.18 * inch))
    scale_cards = Table(
        [[
            narrative_card("Phase 1", "Kyoto Indigo and subsequent destination-led collections establish the public point of view and commercial grammar.", styles),
            narrative_card("Phase 2", "Repeat collections, stronger supplier systems, and signature product lines create a brand with operating depth.", styles),
            narrative_card("Phase 3", "Bedding, home textiles, partnerships, licensing, and prestige retail extend the company into a lifestyle-scale brand.", styles),
        ]],
        colWidths=[2.12 * inch, 2.12 * inch, 2.12 * inch],
    )
    scale_cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(scale_cards)
    story.append(Spacer(1, 0.18 * inch))
    story.extend(section_heading("Capital strategy", "Tight early capital, stronger later leverage", "Initial capital supports product development, sourcing travel, content production, and ecommerce operations. Later capital supports inventory scale, team build-out, and broader supplier management after the brand has already established proof.", styles))
    capital_cards = Table(
        [[
            narrative_card("Use of funds", "Product development, sourcing, inventory, content, and ecommerce operations.", styles),
            narrative_card("Capital range", "$35K-$70K for the initial operating phase, followed by growth capital after traction.", styles),
            narrative_card("Long-term upside", "A premium textile authority with room to scale into signature home collections, licensing, and selective retail.", styles),
        ]],
        colWidths=[2.12 * inch, 2.12 * inch, 2.12 * inch],
    )
    capital_cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(capital_cards)

    doc = SimpleDocTemplate(
        str(output),
        pagesize=letter,
        rightMargin=0.58 * inch,
        leftMargin=0.58 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.55 * inch,
        title="Ali Dandin Business Plan",
        author="OpenAI Codex",
    )
    doc.build(story, onFirstPage=draw_first_page, onLaterPages=draw_later_pages)


def set_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(slide, left, top, width, height, text, size, color, *, bold=False, font_name="Aptos", align=PP_ALIGN.LEFT):
    tx = slide.shapes.add_textbox(left, top, width, height)
    frame = tx.text_frame
    frame.clear()
    p = frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = align
    return tx


def add_panel(slide, left, top, width, height, *, fill_rgb, line_rgb, radius=True):
    shape_type = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    shape.line.color.rgb = line_rgb
    return shape


def add_picture_contain(slide, path: Path, left, top, width, height):
    image_w, image_h = image_dimensions(path)
    frame_ratio = width / height
    image_ratio = image_w / image_h
    if image_ratio > frame_ratio:
        draw_width = width
        draw_height = width / image_ratio
        draw_left = left
        draw_top = top + (height - draw_height) / 2
    else:
        draw_height = height
        draw_width = height * image_ratio
        draw_left = left + (width - draw_width) / 2
        draw_top = top
    slide.shapes.add_picture(str(path), draw_left, draw_top, width=draw_width, height=draw_height)


def add_bullet_card(slide, left, top, width, height, title, bullets, *, fill_rgb, line_rgb, title_rgb, body_rgb):
    panel = add_panel(slide, left, top, width, height, fill_rgb=fill_rgb, line_rgb=line_rgb)
    frame = panel.text_frame
    frame.clear()
    p = frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = title_rgb
    p.font.name = "Georgia"
    for bullet in bullets:
        para = frame.add_paragraph()
        para.text = bullet
        para.level = 0
        para.font.size = Pt(11.5)
        para.font.color.rgb = body_rgb
        para.font.name = "Aptos"
        para.space_after = Pt(5)
        para.bullet = True
    return panel


def add_stat_box(slide, left, top, width, height, value, label, *, fill_rgb, line_rgb, value_rgb, label_rgb):
    panel = add_panel(slide, left, top, width, height, fill_rgb=fill_rgb, line_rgb=line_rgb)
    frame = panel.text_frame
    frame.clear()
    p = frame.paragraphs[0]
    p.text = value
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = value_rgb
    p.font.name = "Georgia"
    p.alignment = PP_ALIGN.CENTER
    p = frame.add_paragraph()
    p.text = label
    p.font.size = Pt(11)
    p.font.color.rgb = label_rgb
    p.font.name = "Aptos"
    p.alignment = PP_ALIGN.CENTER
    return panel


def add_revenue_chart_slide(slide, left, top, width, height, *, bg_rgb, axis_rgb, title_rgb):
    add_panel(slide, left, top, width, height, fill_rgb=bg_rgb, line_rgb=RGBColor(0xD8, 0xCF, 0xC2))
    add_text(slide, left + Inches(0.18), top + Inches(0.08), width - Inches(0.36), Inches(0.25), "Revenue ramp", 13, title_rgb, bold=True, font_name="Aptos")
    max_value = max(REVENUE)
    chart_top = top + Inches(0.48)
    chart_bottom = top + height - Inches(0.5)
    chart_height = chart_bottom - chart_top
    bar_width = Inches(0.68)
    gap = Inches(0.56)
    base_left = left + Inches(0.55)
    line_y = chart_bottom
    axis_line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left + Inches(0.42), line_y, width - Inches(0.84), Inches(0.02))
    axis_line.fill.solid()
    axis_line.fill.fore_color.rgb = axis_rgb
    axis_line.line.color.rgb = axis_rgb
    for idx, (label, value) in enumerate(zip(YEAR_LABELS, REVENUE)):
        bar_height = chart_height * (value / max_value)
        x = base_left + idx * (bar_width + gap)
        y = chart_bottom - bar_height
        bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, bar_width, bar_height)
        bar.fill.solid()
        bar.fill.fore_color.rgb = NAVY_RGB
        bar.line.color.rgb = NAVY_RGB
        add_text(slide, x - Inches(0.05), y - Inches(0.28), bar_width + Inches(0.1), Inches(0.2), money_short(value), 10, title_rgb, bold=True, font_name="Aptos", align=PP_ALIGN.CENTER)
        add_text(slide, x - Inches(0.1), chart_bottom + Inches(0.05), bar_width + Inches(0.2), Inches(0.22), label, 10, axis_rgb, font_name="Aptos", align=PP_ALIGN.CENTER)


NAVY_RGB = RGBColor(0x1E, 0x2A, 0x3A)
SAND_RGB = RGBColor(0xC2, 0xA5, 0x7A)
CREAM_RGB = RGBColor(0xFF, 0xFA, 0xF3)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)
SOFT_RGB = RGBColor(0xF7, 0xF0, 0xE5)
LINE_RGB = RGBColor(0xD8, 0xCF, 0xC2)
MUTED_RGB = RGBColor(0x65, 0x70, 0x7A)
TERRACOTTA_RGB = RGBColor(0xB9, 0x69, 0x39)
OLIVE_RGB = RGBColor(0x61, 0x71, 0x47)


def add_slide_header(slide, page_number: int, title: str, subtitle: str, *, dark: bool = False):
    kicker_color = SAND_RGB if dark else TERRACOTTA_RGB
    title_color = CREAM_RGB if dark else NAVY_RGB
    subtitle_color = CREAM_RGB if dark else MUTED_RGB
    add_text(slide, Inches(0.7), Inches(0.42), Inches(1.0), Inches(0.2), f"{page_number:02d}", 10, kicker_color, bold=True)
    add_text(slide, Inches(0.7), Inches(0.78), Inches(5.9), Inches(0.55), title, 27 if dark else 25, title_color, bold=True, font_name="Georgia")
    add_text(slide, Inches(0.7), Inches(1.48), Inches(5.95), Inches(0.72), subtitle, 13.5, subtitle_color, font_name="Aptos")


def build_pptx() -> None:
    output = DOCS / "investor_deck.pptx"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Cover
    slide = prs.slides.add_slide(blank)
    set_bg(slide, NAVY_RGB)
    add_panel(slide, Inches(7.72), Inches(0.46), Inches(4.88), Inches(6.56), fill_rgb=RGBColor(0x2A, 0x39, 0x4C), line_rgb=RGBColor(0x2A, 0x39, 0x4C))
    add_picture_contain(slide, IMAGES / "kyoto_indigo_collection.png", Inches(7.86), Inches(0.58), Inches(4.6), Inches(6.32))
    add_text(slide, Inches(0.75), Inches(0.64), Inches(2.4), Inches(0.22), "ALI DANDIN", 11, SAND_RGB, bold=True)
    add_text(slide, Inches(0.75), Inches(1.15), Inches(5.7), Inches(1.2), "Ali Dandin premium textiles", 31, CREAM_RGB, bold=True, font_name="Georgia")
    add_text(slide, Inches(0.75), Inches(2.3), Inches(5.7), Inches(1.0), "A destination-led textile brand built around provenance, material intelligence, and premium direct-to-consumer commerce.", 15, CREAM_RGB, font_name="Aptos")
    add_stat_box(slide, Inches(0.75), Inches(4.45), Inches(1.7), Inches(1.0), "Kyoto Indigo", "Opening collection", fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x3F, 0x52, 0x6B), value_rgb=CREAM_RGB, label_rgb=RGBColor(0xDA, 0xD4, 0xCA))
    add_stat_box(slide, Inches(2.62), Inches(4.45), Inches(1.7), Inches(1.0), "$826K", "Year 1 revenue", fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x3F, 0x52, 0x6B), value_rgb=CREAM_RGB, label_rgb=RGBColor(0xDA, 0xD4, 0xCA))
    add_stat_box(slide, Inches(4.49), Inches(4.45), Inches(1.7), Inches(1.0), "55%-65%", "Gross margin", fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x3F, 0x52, 0x6B), value_rgb=CREAM_RGB, label_rgb=RGBColor(0xDA, 0xD4, 0xCA))
    add_text(slide, Inches(0.75), Inches(6.18), Inches(5.9), Inches(0.55), "Premium textiles provide a strong foundation for margin, repeatability, and expansion into private label.", 12.5, RGBColor(0xD8, 0xD1, 0xC7), font_name="Aptos")

    # Thesis
    slide = prs.slides.add_slide(blank)
    set_bg(slide, CREAM_RGB)
    add_slide_header(slide, 2, "Focused market position", "Ali Dandin operates as a textile-led discovery brand centered on provenance, material intelligence, and destination-led collections.")
    add_bullet_card(slide, Inches(0.72), Inches(2.15), Inches(3.75), Inches(4.52), "Why textiles", [
        "Textiles hold tactile differentiation and premium price perception.",
        "The category supports repeat purchase more cleanly than broad travel goods.",
        "Private label and prestige retail expansion remain credible long-term paths.",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_bullet_card(slide, Inches(4.62), Inches(2.15), Inches(3.75), Inches(4.52), "Why Ali", [
        "Material judgment is the moat, not generic sourcing access.",
        "A textile background improves product quality, supplier fluency, and trust.",
        "The brand speaks with authority rather than marketplace breadth.",
    ], fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_bullet_card(slide, Inches(8.52), Inches(2.15), Inches(4.08), Inches(4.52), "What the customer buys", [
        "A clear point of view on textiles.",
        "Collections that translate place and process into product.",
        "A premium editorial storefront that turns story into commerce.",
    ], fill_rgb=RGBColor(0x1E, 0x2A, 0x3A), line_rgb=RGBColor(0x32, 0x43, 0x59), title_rgb=CREAM_RGB, body_rgb=RGBColor(0xDD, 0xD6, 0xCC))

    # Brand system
    slide = prs.slides.add_slide(blank)
    set_bg(slide, CREAM_RGB)
    add_slide_header(slide, 3, "Brand system", "Ali Dandin defines a precise premium position in global textile commerce.")
    add_panel(slide, Inches(0.72), Inches(2.1), Inches(5.6), Inches(4.7), fill_rgb=WHITE_RGB, line_rgb=LINE_RGB)
    add_text(slide, Inches(0.98), Inches(2.38), Inches(4.7), Inches(0.35), "Identity system", 19, NAVY_RGB, bold=True, font_name="Georgia")
    add_text(slide, Inches(0.98), Inches(2.82), Inches(4.9), Inches(0.95), "The identity system includes a primary mark, monogram, passport stamp, and textile globe. Midnight blue, sand, and cream establish a textile house rather than a generic lifestyle brand.", 13, MUTED_RGB)
    add_text(slide, Inches(0.98), Inches(4.04), Inches(1.4), Inches(0.22), "Voice", 11.5, TERRACOTTA_RGB, bold=True)
    add_text(slide, Inches(0.98), Inches(4.3), Inches(4.9), Inches(0.54), "Measured, materially intelligent, and anchored in provenance.", 12, NAVY_RGB)
    add_text(slide, Inches(0.98), Inches(5.0), Inches(1.6), Inches(0.22), "Palette", 11.5, TERRACOTTA_RGB, bold=True)
    add_text(slide, Inches(0.98), Inches(5.26), Inches(4.9), Inches(0.6), "Midnight Blue · Sand · Cream · Terracotta · Olive", 12, NAVY_RGB)
    add_panel(slide, Inches(6.58), Inches(2.1), Inches(6.02), Inches(4.7), fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=LINE_RGB)
    add_picture_contain(slide, IMAGES / "logo_suite.png", Inches(6.72), Inches(2.24), Inches(5.74), Inches(4.42))

    # Collection architecture
    slide = prs.slides.add_slide(blank)
    set_bg(slide, NAVY_RGB)
    add_slide_header(slide, 4, "Kyoto Indigo", "Kyoto Indigo establishes the opening collection, combining material, dye process, and cultural origin into a cohesive product system.", dark=True)
    add_panel(slide, Inches(0.72), Inches(2.08), Inches(4.15), Inches(4.76), fill_rgb=RGBColor(0x27, 0x36, 0x49), line_rgb=RGBColor(0x35, 0x47, 0x5E))
    add_picture_contain(slide, IMAGES / "kyoto_indigo_collection.png", Inches(0.86), Inches(2.22), Inches(3.87), Inches(4.48))
    add_bullet_card(slide, Inches(5.1), Inches(2.08), Inches(3.45), Inches(2.22), "Opening assortment", [
        "Indigo Linen Throw",
        "Indigo Scarf",
        "Indigo Pillow",
        "Indigo Textile Set",
    ], fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x35, 0x47, 0x5E), title_rgb=CREAM_RGB, body_rgb=RGBColor(0xDD, 0xD6, 0xCC))
    add_bullet_card(slide, Inches(8.8), Inches(2.08), Inches(3.8), Inches(2.22), "Collection logic", [
        "One destination, one material story, one disciplined assortment.",
        "Clear photography, merchandising, and memory structure.",
        "Strong foundation for future destination-led releases.",
    ], fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x35, 0x47, 0x5E), title_rgb=CREAM_RGB, body_rgb=RGBColor(0xDD, 0xD6, 0xCC))
    add_stat_box(slide, Inches(5.1), Inches(4.62), Inches(2.28), Inches(1.12), "4", "Opening SKUs", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(7.62), Inches(4.62), Inches(2.28), Inches(1.12), "Japan", "Origin story", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(10.14), Inches(4.62), Inches(2.42), Inches(1.12), "Textile-led", "Category position", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)

    # Commerce platform
    slide = prs.slides.add_slide(blank)
    set_bg(slide, CREAM_RGB)
    add_slide_header(slide, 5, "Commerce platform", "The storefront is editorial, collection-led, and designed to convert on taste and provenance.")
    add_panel(slide, Inches(0.72), Inches(2.08), Inches(4.4), Inches(4.86), fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=LINE_RGB)
    add_picture_contain(slide, IMAGES / "shopify_homepage.png", Inches(0.86), Inches(2.22), Inches(4.12), Inches(4.58))
    add_bullet_card(slide, Inches(5.38), Inches(2.08), Inches(3.4), Inches(4.86), "Store logic", [
        "Homepage anchored by a featured collection.",
        "Journal and process story reinforce the premium position.",
        "Product pages sell material, process, and origin rather than discounting.",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_bullet_card(slide, Inches(8.98), Inches(2.08), Inches(3.58), Inches(4.86), "Revenue layers", [
        "Product sales",
        "Passport Club membership",
        "Private label expansion",
        "Selective collaborations and licensing",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)

    # Sourcing network
    slide = prs.slides.add_slide(blank)
    set_bg(slide, NAVY_RGB)
    add_slide_header(slide, 6, "Sourcing network", "Regional sourcing is both a product strategy and a brand strategy.", dark=True)
    add_panel(slide, Inches(7.02), Inches(1.95), Inches(5.56), Inches(4.9), fill_rgb=RGBColor(0x27, 0x36, 0x49), line_rgb=RGBColor(0x35, 0x47, 0x5E))
    add_picture_contain(slide, IMAGES / "global_sourcing_overview.png", Inches(7.16), Inches(2.08), Inches(5.28), Inches(4.62))
    add_bullet_card(slide, Inches(0.72), Inches(1.95), Inches(5.96), Inches(1.72), "Regional priorities", [
        "Japan: heritage process and indigo authority",
        "India: handloom and block print expansion",
        "Turkey and Italy: scale and premium finishing",
    ], fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x35, 0x47, 0x5E), title_rgb=CREAM_RGB, body_rgb=RGBColor(0xDD, 0xD6, 0xCC))
    add_stat_box(slide, Inches(0.72), Inches(4.04), Inches(1.8), Inches(1.05), "Japan", "Process anchor", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(2.66), Inches(4.04), Inches(1.8), Inches(1.05), "India", "Breadth", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(4.6), Inches(4.04), Inches(1.8), Inches(1.05), "Turkey", "Scale", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(0.72), Inches(5.36), Inches(1.8), Inches(1.05), "Italy", "Finishing", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(2.66), Inches(5.36), Inches(1.8), Inches(1.05), "Peru", "Future fiber", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_text(slide, Inches(4.88), Inches(5.43), Inches(1.75), Inches(0.82), "Supplier depth follows the same rule as the assortment: narrow, intentional, and quality-led.", 12, RGBColor(0xD8, 0xD1, 0xC7), font_name="Aptos")

    # Business model
    slide = prs.slides.add_slide(blank)
    set_bg(slide, CREAM_RGB)
    add_slide_header(slide, 7, "Business model", "Product revenue establishes the brand, then higher-leverage revenue streams extend it.")
    add_bullet_card(slide, Inches(0.72), Inches(2.08), Inches(3.0), Inches(4.92), "Revenue stack", [
        "Direct-to-consumer textile collections",
        "Passport Club membership",
        "Collaborations and affiliate income",
        "Private label and licensing",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_panel(slide, Inches(4.02), Inches(2.08), Inches(4.0), Inches(4.92), fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=LINE_RGB)
    add_text(slide, Inches(4.3), Inches(2.36), Inches(3.4), Inches(0.35), "Value ladder", 18, NAVY_RGB, bold=True, font_name="Georgia")
    steps = [
        ("Product", "Destination-led textile drops"),
        ("Membership", "Early access and editorial depth"),
        ("Private label", "Signature lines with stronger margin"),
        ("Partnerships", "Selective prestige distribution"),
    ]
    step_top = Inches(2.88)
    for idx, (name, desc) in enumerate(steps):
        width = Inches(2.25 + idx * 0.36)
        left = Inches(4.32) + Inches(0.18 * (3 - idx))
        panel = add_panel(slide, left, step_top + Inches(idx * 0.63), width, Inches(0.5), fill_rgb=[RGBColor(0x1E, 0x2A, 0x3A), RGBColor(0x3A, 0x49, 0x5D), RGBColor(0x6D, 0x53, 0x40), RGBColor(0x61, 0x71, 0x47)][idx], line_rgb=[RGBColor(0x1E, 0x2A, 0x3A), RGBColor(0x3A, 0x49, 0x5D), RGBColor(0x6D, 0x53, 0x40), RGBColor(0x61, 0x71, 0x47)][idx])
        frame = panel.text_frame
        frame.clear()
        p = frame.paragraphs[0]
        p.text = name
        p.font.name = "Aptos"
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = CREAM_RGB
        p = frame.add_paragraph()
        p.text = desc
        p.font.name = "Aptos"
        p.font.size = Pt(9.2)
        p.font.color.rgb = CREAM_RGB
    add_bullet_card(slide, Inches(8.32), Inches(2.08), Inches(4.24), Inches(4.92), "Why it scales", [
        "Clear category ownership supports premium pricing.",
        "Membership improves retention and launch depth.",
        "Private label expands margin without breaking the brand.",
        "Selective partnerships extend reach while preserving taste.",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)

    # Financial profile
    slide = prs.slides.add_slide(blank)
    set_bg(slide, NAVY_RGB)
    add_slide_header(slide, 8, "Financial profile", "The economic model is built on premium ASPs, disciplined assortment, and content-led acquisition.", dark=True)
    add_revenue_chart_slide(slide, Inches(0.72), Inches(2.02), Inches(5.76), Inches(4.9), bg_rgb=CREAM_RGB, axis_rgb=MUTED_RGB, title_rgb=NAVY_RGB)
    add_stat_box(slide, Inches(6.8), Inches(2.12), Inches(1.9), Inches(1.1), money_short(GROSS_PROFIT[2]), "Year 3 gross profit", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(8.92), Inches(2.12), Inches(1.9), Inches(1.1), money_short(OPERATING_PROFIT[2]), "Year 3 operating profit", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_stat_box(slide, Inches(11.04), Inches(2.12), Inches(1.5), Inches(1.1), "$8", "Membership", fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), value_rgb=NAVY_RGB, label_rgb=MUTED_RGB)
    add_bullet_card(slide, Inches(6.8), Inches(3.6), Inches(5.74), Inches(3.32), "Model characteristics", [
        "Year 1 revenue: $826K",
        "Year 2 revenue: $3.6M",
        "Year 3 revenue: $9.352M",
        "Margin expansion follows assortment quality and private label depth",
    ], fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x35, 0x47, 0x5E), title_rgb=CREAM_RGB, body_rgb=RGBColor(0xDD, 0xD6, 0xCC))

    # Roadmap
    slide = prs.slides.add_slide(blank)
    set_bg(slide, CREAM_RGB)
    add_slide_header(slide, 9, "Growth roadmap", "The brand scales from discovery into authority and then into a broader textile lifestyle platform.")
    add_panel(slide, Inches(0.72), Inches(2.08), Inches(11.86), Inches(2.28), fill_rgb=WHITE_RGB, line_rgb=LINE_RGB)
    line_y = Inches(3.1)
    timeline_line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(1.35), line_y, Inches(10.56), Inches(0.03))
    timeline_line.fill.solid()
    timeline_line.fill.fore_color.rgb = SAND_RGB
    timeline_line.line.color.rgb = SAND_RGB
    nodes = [
        (Inches(2.1), "Discovery", "Years 1-2", "Destination-led launches"),
        (Inches(6.05), "Authority", "Years 3-4", "Repeat collections and private label"),
        (Inches(10.0), "Expansion", "Years 5+", "Home textiles, partnerships, licensing"),
    ]
    for x, title, period, note in nodes:
        circ = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, x, line_y - Inches(0.17), Inches(0.34), Inches(0.34))
        circ.fill.solid()
        circ.fill.fore_color.rgb = NAVY_RGB
        circ.line.color.rgb = NAVY_RGB
        add_text(slide, x - Inches(0.28), Inches(2.38), Inches(1.05), Inches(0.22), title, 12, NAVY_RGB, bold=True, font_name="Aptos", align=PP_ALIGN.CENTER)
        add_text(slide, x - Inches(0.28), Inches(2.62), Inches(1.05), Inches(0.2), period, 10, TERRACOTTA_RGB, font_name="Aptos", align=PP_ALIGN.CENTER)
        add_text(slide, x - Inches(0.7), Inches(3.45), Inches(1.9), Inches(0.45), note, 10.5, MUTED_RGB, font_name="Aptos", align=PP_ALIGN.CENTER)
    add_bullet_card(slide, Inches(0.72), Inches(4.7), Inches(3.74), Inches(2.0), "Phase 1", [
        "Clear imagery",
        "Strong assortment discipline",
        "Focused destination story",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_bullet_card(slide, Inches(4.78), Inches(4.7), Inches(3.74), Inches(2.0), "Phase 2", [
        "Supplier depth",
        "Signature lines",
        "Private label confidence",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_bullet_card(slide, Inches(8.84), Inches(4.7), Inches(3.74), Inches(2.0), "Phase 3", [
        "Prestige retail",
        "Licensing",
        "Lifestyle-scale brand equity",
    ], fill_rgb=WHITE_RGB, line_rgb=LINE_RGB, title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)

    # Capital strategy / close
    slide = prs.slides.add_slide(blank)
    set_bg(slide, NAVY_RGB)
    add_slide_header(slide, 10, "Capital strategy", "The financing story is strongest when product proof, brand clarity, and supplier discipline reinforce one another.", dark=True)
    add_bullet_card(slide, Inches(0.72), Inches(2.08), Inches(5.86), Inches(4.72), "Use of funds", [
        "Product development and initial inventory",
        "Sourcing travel and supplier onboarding",
        "Content production and ecommerce operations",
        "Early team support and working capital",
    ], fill_rgb=RGBColor(0x2B, 0x3A, 0x4C), line_rgb=RGBColor(0x35, 0x47, 0x5E), title_rgb=CREAM_RGB, body_rgb=RGBColor(0xDD, 0xD6, 0xCC))
    add_bullet_card(slide, Inches(6.78), Inches(2.08), Inches(5.8), Inches(4.72), "Why this wins", [
        "Clear premium category position",
        "Textile fluency as a brand and sourcing moat",
        "Tightly curated collections rather than broad marketplace sprawl",
        "A path from editorial discovery into lasting brand equity",
    ], fill_rgb=RGBColor(0xF7, 0xF0, 0xE5), line_rgb=RGBColor(0x35, 0x47, 0x5E), title_rgb=NAVY_RGB, body_rgb=MUTED_RGB)
    add_text(slide, Inches(0.72), Inches(6.92), Inches(11.86), Inches(0.2), "Ali Dandin premium textiles", 15, SAND_RGB, bold=True, font_name="Georgia", align=PP_ALIGN.CENTER)

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
            style_cell(
                cell,
                bold=label in {"Gross profit", "Indicative operating profit"},
                fill=green_fill if label in {"Gross profit", "Indicative operating profit"} else None,
            )
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


def build_zip() -> None:
    output = DOCS / "alidandin_investor_package.zip"
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for filename in ["investor_business_plan.pdf", "investor_deck.pptx", "financial_model.xlsx"]:
            archive.write(DOCS / filename, arcname=filename)


def main() -> None:
    ensure_dirs()
    build_pdf()
    build_pptx()
    build_xlsx()
    build_zip()


if __name__ == "__main__":
    main()
