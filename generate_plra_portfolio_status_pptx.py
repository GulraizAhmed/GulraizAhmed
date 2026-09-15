#!/usr/bin/env python3
"""
Generate a concise PLRA Assigned Project Portfolio status briefing (5 slides).
Formal green/gold theme matching PLRA official presentation style.
"""

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


# Formal PLRA palette (from portfolio reference)
GREEN = RGBColor(0x00, 0x68, 0x37)
GREEN_DARK = RGBColor(0x0B, 0x4F, 0x2C)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
GOLD_DARK = RGBColor(0xA8, 0x86, 0x3E)
GRAY = RGBColor(0x4A, 0x4A, 0x4A)
GRAY_LIGHT = RGBColor(0x8A, 0x8A, 0x8A)
BLACK = RGBColor(0x22, 0x22, 0x22)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xF7, 0xF9, 0xF7)
CARD_LINE = RGBColor(0xE6, 0xEB, 0xE6)


def set_run(run, size=14, bold=False, color=BLACK, italic=False):
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_text(shape, lines, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    try:
        tf._txBody.bodyPr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(valign, "t"),
        )
    except Exception:
        pass
    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, size, bold, color, italic = item, 14, False, BLACK, False
        else:
            text, size, bold, color = item[0], item[1], item[2], item[3]
            italic = item[4] if len(item) > 4 else False
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(1)
        run = p.add_run()
        run.text = text
        set_run(run, size=size, bold=bold, color=color, italic=italic)


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def fill_bg(slide, prs, color=WHITE):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    return sh


def rounded(slide, left, top, width, height, fill=WHITE, line=CARD_LINE, line_width=1.0, radius=0.12):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_width)
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    return sh


def accent_divider(slide, left, top, width=1.55):
    """Gold + green dual accent line under subtitle (matches reference)."""
    g = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width * 0.55), Inches(0.035))
    g.fill.solid()
    g.fill.fore_color.rgb = GOLD
    g.line.fill.background()
    gr = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left + width * 0.55),
        Inches(top),
        Inches(width * 0.45),
        Inches(0.035),
    )
    gr.fill.solid()
    gr.fill.fore_color.rgb = GREEN
    gr.line.fill.background()


def slide_header(slide, number, title, subtitle, y=0.28):
    num = slide.shapes.add_textbox(Inches(0.45), Inches(y), Inches(0.7), Inches(0.55))
    add_text(num, [(f"{number:02d}", 28, True, GREEN)])
    t = slide.shapes.add_textbox(Inches(1.15), Inches(y + 0.02), Inches(11.5), Inches(0.45))
    add_text(t, [(title, 26, True, GREEN)])
    s = slide.shapes.add_textbox(Inches(1.15), Inches(y + 0.42), Inches(11.5), Inches(0.28))
    add_text(s, [(subtitle, 12, False, GRAY_LIGHT)])
    accent_divider(slide, 1.15, y + 0.72, width=1.7)


def footer(slide, prs):
    box = slide.shapes.add_textbox(Inches(0.45), Inches(7.15), Inches(12.4), Inches(0.25))
    add_text(
        box,
        [("Punjab Land Records Authority (PLRA)  |  Assigned Project Portfolio — Status Briefing", 9, False, GRAY_LIGHT)],
        align=PP_ALIGN.CENTER,
    )


def set_fade_transition(slide):
    sld = slide._element
    for child in list(sld):
        if etree.QName(child).localname == "transition":
            sld.remove(child)
    tr = etree.Element(qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))
    # insert after clrMapOvr / cSld
    insert_at = 0
    for idx, child in enumerate(list(sld)):
        if etree.QName(child).localname in ("cSld", "clrMapOvr"):
            insert_at = idx + 1
    sld.insert(insert_at, tr)


# -------------------- Slide 1: Portfolio Index --------------------

PORTFOLIO = [
    "Arazi Moawin",
    "CLRMIS",
    "E-Registration",
    "CIMS",
    "E-Girdawari",
    "Punjab Zameen App",
    "Irrigation",
    "CMS",
    "RCMS",
    "E-stamp Punjab",
    "E-stamp ICT",
    "DC Valuation",
    "MNP HUB",
    "Payment Gateway",
    "LMS",
    "Roznamcha",
]


def slide_portfolio_index(prs):
    s = blank(prs)
    fill_bg(s, prs, WHITE)
    slide_header(s, 1, "Assigned Project Portfolio", "Project-wise status briefing — current status & key updates")

    cols, rows = 4, 4
    left0, top0 = 0.45, 1.35
    gap_x, gap_y = 0.18, 0.16
    card_w = (12.4 - gap_x * (cols - 1)) / cols
    card_h = 1.15

    for i, name in enumerate(PORTFOLIO):
        r, c = divmod(i, cols)
        x = left0 + c * (card_w + gap_x)
        y = top0 + r * (card_h + gap_y)
        rounded(s, Inches(x), Inches(y), Inches(card_w), Inches(card_h), fill=WHITE, line=CARD_LINE, radius=0.14)

        n = s.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.12), Inches(0.55), Inches(0.28))
        add_text(n, [(f"{i + 1:02d}", 11, True, GREEN)])
        t = s.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.38), Inches(card_w - 0.3), Inches(0.55))
        add_text(t, [(name, 13, True, BLACK)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    footer(s, prs)
    return s


# -------------------- Compact system cards --------------------

def system_card(slide, left, top, width, height, title, status, bullets):
    rounded(slide, Inches(left), Inches(top), Inches(width), Inches(height), fill=WHITE, line=CARD_LINE, radius=0.1)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(0.07), Inches(height))
    bar.fill.solid()
    bar.fill.fore_color.rgb = GREEN
    bar.line.fill.background()

    ht = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.08), Inches(width - 0.3), Inches(0.28))
    add_text(ht, [(title, 13, True, GREEN)])

    st = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.34), Inches(width - 0.3), Inches(0.42))
    add_text(st, [("Status: " + status, 10, False, GRAY)])

    lines = [(f"•  {b}", 10, False, BLACK) for b in bullets]
    bt = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.78), Inches(width - 0.3), Inches(height - 0.9))
    add_text(bt, lines)


def two_col_systems(slide, systems, top=1.2, card_h=2.55):
    """systems: list of (title, status, bullets) — laid out in 2 columns."""
    gap = 0.18
    width = (12.4 - gap) / 2
    for i, (title, status, bullets) in enumerate(systems):
        c = i % 2
        r = i // 2
        x = 0.45 + c * (width + gap)
        y = top + r * (card_h + 0.14)
        system_card(slide, x, y, width, card_h, title, status, bullets)


# -------------------- Slides 2–5 --------------------

def slide_core_land(prs):
    s = blank(prs)
    fill_bg(s, prs, SOFT)
    slide_header(s, 2, "Core Land & Identity Systems", "Arazi Moawin · CLRMIS · E-Registration · CIMS")
    two_col_systems(
        s,
        [
            (
                "Arazi Moawin",
                "On track — UAT deployment planned; current sprint completed for UAT",
                [
                    "Multi-Tehsil functionality",
                    "Territory-based admin access",
                    "Production / UAT issue resolution",
                    "e-Registry integration",
                ],
            ),
            (
                "CLRMIS",
                "Release 12 in progress",
                [
                    "GPC enhancement",
                    "Fard Survey task",
                    "KMI",
                    "Security fixes",
                ],
            ),
            (
                "E-Registration",
                "Release 0",
                [
                    "FBR integration enhancement",
                    "Arazi Moawin integration & feedback",
                    "Audit logging",
                ],
            ),
            (
                "CIMS",
                "Sprint 14 — integration work in progress",
                [
                    "HSMS integration",
                    "CMS integration",
                    "Arazi Moawin integration",
                ],
            ),
        ],
        top=1.2,
        card_h=2.6,
    )
    footer(s, prs)
    return s


def slide_apps_ops(prs):
    s = blank(prs)
    fill_bg(s, prs, WHITE)
    slide_header(s, 3, "Citizen Apps & Operations", "E-Girdawari · Punjab Zameen · Irrigation · CMS · RCMS")

    # Top row: 3 cards
    gap = 0.14
    w3 = (12.4 - gap * 2) / 3
    top_systems = [
        (
            "E-Girdawari",
            "Sprint 0 — Requirement gathering & Figma",
            [
                "Girdawari Figma completed",
                "Irrigation design to be incorporated",
                "Requirement gathering ongoing",
                "CMS issues under resolution",
            ],
        ),
        (
            "Punjab Zameen App",
            "Release 2 — merging code for production",
            [
                "CIMS, CMS, HSMS integrated on UAT",
                "Nearest Arazi Moawin & RCMS on UAT",
                "All four integrations ready to go live",
            ],
        ),
        (
            "Irrigation",
            "Design phase — to be incorporated in Girdawari App",
            [
                "Irrigation design activity",
                "Figma / design coordination with Girdawari",
            ],
        ),
    ]
    for i, item in enumerate(top_systems):
        system_card(s, 0.45 + i * (w3 + gap), 1.2, w3, 2.75, *item)

    # Bottom row: 2 wider cards
    w2 = (12.4 - gap) / 2
    bottom = [
        (
            "CMS",
            "Planning",
            [
                "Department-wise & pendency reports",
                "Overdue / escalation workflow enhancements",
                "Mobile-friendly UI",
                "CIMS integration: backend / UAT / sign-up / update APIs",
            ],
        ),
        (
            "RCMS",
            "On track — UAT / production activities in current sprint",
            [
                "Stay Order implementation in CLRMIS",
                "SMBR Court Account updation",
                "Punjab Zameen / Proclamation linkage",
            ],
        ),
    ]
    for i, item in enumerate(bottom):
        system_card(s, 0.45 + i * (w2 + gap), 4.1, w2, 2.7, *item)

    footer(s, prs)
    return s


def slide_platforms(prs):
    s = blank(prs)
    fill_bg(s, prs, SOFT)
    slide_header(s, 4, "Stamping, Valuation & Payments", "E-stamp · DC Valuation · MNP HUB · Payment Gateway")

    gap = 0.12
    w = (12.4 - gap * 2) / 3
    systems = [
        (
            "E-stamp Punjab",
            "Oct 2024 takeover — in-house development & support",
            [
                "Delivered: Central Hub SMS/MNP, GPC challan & mapping, eRegistry APIs, stamp reprint, duty revision",
                "In progress: Approved-map (sale deed), Sec. 9 exemption, Arazi Fard/Survey, PMD NO-call, eStamp GB",
                "Next: Infra 2022, Court Copy API, Arazi Moawin vendor link, GIS DC calculator",
            ],
        ),
        (
            "E-stamp ICT",
            "Jan 2026 launched — in-house development & support",
            [
                "Delivered: Ufone SMS/OTP, adhesive stamp + QR, NITB payment API, email/agent download, DB 2022, DC import 2026–27",
                "In progress: Stamp reprint & support, refund process (testing), adhesive verification (web)",
                "Next: M&P for adhesive challan, delisting, refund training, Exchange Deed",
            ],
        ),
        (
            "DC Valuation",
            "Apr 2025 launched — digital valuation platform",
            [
                "Delivered: PERA & Excise APIs, eStamp PUSH, GIS view/entry 2026–27, QR PDF, CLRMIS territory mapping",
                "In progress: Rural/Urban corrigendum, FBR & CIMS integration, E-Biz API (PULSE)",
                "In progress: Urban property-area GIS mapping; CLRMIS territory 2026–27",
            ],
        ),
    ]
    for i, item in enumerate(systems):
        system_card(s, 0.45 + i * (w + gap), 1.15, w, 3.15, *item)

    # Bottom two
    w2 = (12.4 - gap) / 2
    bottom = [
        (
            "MNP HUB",
            "Jan 2026 launched — in-house development & support",
            [
                "Centralized CNIC–MSISDN pairing & authentication",
                "OTP-based mobile verification via centralized SMS",
                "Reusable APIs for multiple PLRA applications",
            ],
        ),
        (
            "Payment Gateway",
            "31 Jul launched — in-house development & support",
            [
                "Delivered: Secure webhook, Payment Ledger, CSV export",
                "In progress: Admin filters/analytics (district-wise), BOP reconciliation, new client onboarding (IP whitelist)",
            ],
        ),
    ]
    for i, item in enumerate(bottom):
        system_card(s, 0.45 + i * (w2 + gap), 4.45, w2, 2.35, *item)

    footer(s, prs)
    return s


def slide_takeovers(prs):
    s = blank(prs)
    fill_bg(s, prs, WHITE)
    slide_header(s, 5, "LMS & Roznamcha Takeovers", "September 2026 — in-house development & support")

    gap = 0.18
    w = (12.4 - gap) / 2

    # LMS
    system_card(
        s,
        0.45,
        1.2,
        w,
        5.5,
        "LMS (Learning Management System)",
        "Sep 2026 takeover — in-house development & support",
        [
            "Standardize training across departments & regions",
            "Digitize assessments, certifications & learner records",
            "Improve progress / compliance visibility",
            "Automate admin effort via self-service workflows",
            "Future: Digital Exam Portal with HRMS elevation exams",
            "Biometric verification, secure Q-bank, CBT & auto evaluation",
            "HR/committee workflow; MPDD post-elevation training",
        ],
    )

    # Roznamcha
    system_card(
        s,
        0.45 + w + gap,
        1.2,
        w,
        5.5,
        "Roznamcha",
        "Sep 2026 takeover — in-house development & support",
        [
            "Digitize Roznamcha — capture واقعاتی entries by territory",
            "Nambardar Registry — centralized نمبردار master records",
            "Official Messaging — circulate ہدایات / پیغامات with attachments",
            "Configurable Forms — add fields without full redevelopment",
            "Leadership Visibility — dashboards & territory filters",
            "Official Reports — printable Urdu PDF records",
        ],
    )

    footer(s, prs)
    return s


def validate_pptx(path):
    import zipfile
    from pptx import Presentation as P

    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            raise RuntimeError(f"Corrupt zip entry: {bad}")
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                etree.fromstring(z.read(name))
    return len(P(path).slides)


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_portfolio_index(prs)
    slide_core_land(prs)
    slide_apps_ops(prs)
    slide_platforms(prs)
    slide_takeovers(prs)

    for slide in prs.slides:
        set_fade_transition(slide)

    out = Path("/workspace/PLRA_Assigned_Project_Portfolio_Status.pptx")
    tmp = Path("/workspace/PLRA_Assigned_Project_Portfolio_Status.tmp.pptx")
    prs.save(str(tmp))
    n = validate_pptx(str(tmp))
    tmp.replace(out)
    print(f"Saved {out} with {n} slides (validated)")
    return str(out)


if __name__ == "__main__":
    build()
