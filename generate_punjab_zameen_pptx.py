#!/usr/bin/env python3
"""Generate Punjab Zameen Application PowerPoint presentation."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


# PLRA / Punjab Zameen brand colors (from attached designs)
GREEN = RGBColor(0x1B, 0x6B, 0x4A)
GREEN_DARK = RGBColor(0x0F, 0x4A, 0x32)
GREEN_MID = RGBColor(0x2E, 0x8B, 0x63)
GREEN_SOFT = RGBColor(0xE8, 0xF5, 0xEE)
GREEN_BORDER = RGBColor(0x2F, 0x9E, 0x6B)
TITLE_BLUE = RGBColor(0x1A, 0x2B, 0x4A)
GRAY = RGBColor(0x55, 0x55, 0x55)
GRAY_LIGHT = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)


def set_run(run, size=14, bold=False, color=BLACK, italic=False, font="Calibri"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_text(shape, lines, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    """lines: list of (text, size, bold, color[, italic]) or plain strings."""
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, size, bold, color, italic = item, 14, False, BLACK, False
        else:
            text = item[0]
            size = item[1]
            bold = item[2]
            color = item[3]
            italic = item[4] if len(item) > 4 else False
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = text
        set_run(run, size=size, bold=bold, color=color, italic=italic)


def add_footer(slide, prs):
    box = slide.shapes.add_textbox(
        Inches(0.5), prs.slide_height - Inches(0.45),
        prs.slide_width - Inches(1.0), Inches(0.3),
    )
    add_text(
        box,
        [("PUNJAB LAND RECORDS AUTHORITY (PLRA)  |  GOVT OF THE PUNJAB", 10, False, GRAY_LIGHT)],
        align=PP_ALIGN.CENTER,
    )


def add_eyebrow(slide, text="CORE APPLICATION MODULES"):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.28), Inches(6), Inches(0.3))
    add_text(box, [(text, 11, True, GREEN)])


def add_title_block(slide, title, subtitle=None, y=0.45):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(12.3), Inches(0.5))
    add_text(box, [(title, 28, True, TITLE_BLUE)])
    if subtitle:
        box2 = slide.shapes.add_textbox(Inches(0.5), Inches(y + 0.45), Inches(12.3), Inches(0.35))
        add_text(box2, [(subtitle, 14, False, GRAY)])


def add_rounded_rect(slide, left, top, width, height, fill=WHITE, line=GREEN_BORDER, line_width=1.5):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(line_width)
    # Soften corner radius
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def add_circle(slide, left, top, size, fill=GREEN, text=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if text is not None:
        add_text(shape, [(str(text), 14, True, WHITE)], align=PP_ALIGN.CENTER)
        shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        for p in shape.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
    return shape


def add_bullet_card(slide, left, top, width, height, title, bullets, icon_label=None):
    card = add_rounded_rect(slide, left, top, width, height, fill=WHITE, line=GREEN_BORDER)
    # accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = GREEN
    bar.line.fill.background()

    t = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.15), width - Inches(0.4), Inches(0.35))
    add_text(t, [(title, 15, True, GREEN_DARK)])

    body = slide.shapes.add_textbox(
        left + Inches(0.25), top + Inches(0.5), width - Inches(0.4), height - Inches(0.65)
    )
    lines = []
    for b in bullets:
        lines.append((f"•  {b}", 12, False, GRAY))
    add_text(body, lines)
    return card


def blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)


# ---------- Slides ----------

def slide_title(prs):
    s = blank_slide(prs)
    # green header band
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(3.2))
    band.fill.solid()
    band.fill.fore_color.rgb = GREEN
    band.line.fill.background()

    soft = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.2), prs.slide_width, Inches(0.15))
    soft.fill.solid()
    soft.fill.fore_color.rgb = GREEN_MID
    soft.line.fill.background()

    box = s.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.5), Inches(0.4))
    add_text(box, [("PUNJAB LAND RECORDS AUTHORITY (PLRA)", 14, True, WHITE)], align=PP_ALIGN.CENTER)

    box = s.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.8))
    add_text(box, [("Punjab Zameen Application", 40, True, WHITE)], align=PP_ALIGN.CENTER)

    box = s.shapes.add_textbox(Inches(0.8), Inches(2.3), Inches(11.5), Inches(0.5))
    add_text(
        box,
        [("Citizen-Facing Digital Land Services Platform", 18, False, RGBColor(0xD8, 0xF0, 0xE4))],
        align=PP_ALIGN.CENTER,
    )

    # key points
    points = [
        ("11+ Integrated Systems", "Enterprise interoperability across provincial departments"),
        ("End-to-End Digital Services", "GPC, Record Copy, Payments, Complaints & more"),
        ("Citizen Empowerment", "Verified holdings, online payments, and status tracking"),
    ]
    x = Inches(0.7)
    for title, desc in points:
        card = add_rounded_rect(s, x, Inches(3.8), Inches(3.9), Inches(1.6), fill=GREEN_SOFT, line=GREEN_BORDER)
        t = s.shapes.add_textbox(x + Inches(0.2), Inches(4.0), Inches(3.5), Inches(0.4))
        add_text(t, [(title, 15, True, GREEN_DARK)])
        d = s.shapes.add_textbox(x + Inches(0.2), Inches(4.45), Inches(3.5), Inches(0.7))
        add_text(d, [(desc, 12, False, GRAY)])
        x += Inches(4.1)

    add_footer(s, prs)


def slide_agenda(prs):
    s = blank_slide(prs)
    add_eyebrow(s, "OVERVIEW")
    add_title_block(s, "Agenda", "Module-by-module walkthrough of Punjab Zameen")

    left_items = [
        "1. Why Punjab Zameen Remodeling?",
        "2. Integrated Enterprise Ecosystem",
        "3. My Properties & GPC Workflow",
        "4. Approved Housing Society (HSMS)",
        "5. Get Record Copy",
        "6. Update Personal Records",
        "7. Application for Partition",
        "8. Complaints Module (CMS)",
    ]
    right_items = [
        "9. Payment Gateway",
        "10. Revenue Court Cases (RCMS)",
        "11. Arazi Muawin",
        "12. e-Registration",
        "13. e-Stamp",
        "14. Property Tax (PTS)",
        "15. e-Leasing",
        "16. Future Roadmap",
    ]

    for i, items in enumerate([left_items, right_items]):
        x = Inches(0.6) if i == 0 else Inches(6.8)
        card = add_rounded_rect(s, x, Inches(1.5), Inches(5.8), Inches(5.0), fill=WHITE, line=GREEN_BORDER)
        box = s.shapes.add_textbox(x + Inches(0.35), Inches(1.7), Inches(5.1), Inches(4.5))
        add_text(box, [(item, 15, False, BLACK) for item in items])

    add_footer(s, prs)


def slide_why_remodel(prs):
    s = blank_slide(prs)
    add_eyebrow(s, "STRATEGIC RATIONALE")
    add_title_block(s, "Why Punjab Zameen Remodeling?", "From fragmented legacy access to a unified citizen platform")

    cards = [
        ("Citizen Empowerment", [
            "Direct access to verified property holdings",
            "Initiate GPC requests without middlemen",
            "Pay taxes and fees digitally",
        ]),
        ("Enterprise Interoperability", [
            "Standardized RESTful APIs",
            "Connecting 11+ provincial systems",
            "Single platform for PLRA services",
        ]),
        ("Financial Transparency", [
            "Automated fee calculation",
            "PSID tracking & audit trail",
            "Duplicate-payment safeguards",
        ]),
        ("Access to PLRA Services", [
            "Multiple services from one app",
            "Reduced citizen load on offices",
            "Live status & digital vault storage",
        ]),
    ]

    positions = [
        (Inches(0.5), Inches(1.5)),
        (Inches(6.7), Inches(1.5)),
        (Inches(0.5), Inches(4.0)),
        (Inches(6.7), Inches(4.0)),
    ]
    for (x, y), (title, bullets) in zip(positions, cards):
        add_bullet_card(s, x, y, Inches(5.9), Inches(2.2), title, bullets)

    add_footer(s, prs)


def slide_ecosystem(prs):
    s = blank_slide(prs)
    add_eyebrow(s, "INTEGRATED ENTERPRISE ECOSYSTEM")
    add_title_block(
        s,
        "Seamless Interoperability Across 10+ Government Systems",
        "Punjab Zameen as the citizen front-end to PLRA and partner systems",
    )

    systems = [
        ("CIMS", "Central Identity\nManagement System"),
        ("CLRMIS", "Computerized Land Records\nManagement Information System"),
        ("Payment Gateway", "Online challan fee\ncollection & reconciliation"),
        ("HSMS", "Housing Society\nManagement System"),
        ("RCMS", "Revenue Court\nManagement System"),
        ("Arazi Muawin", "Franchise centers &\nland facilitation"),
        ("CMS", "Complaint\nManagement System"),
        ("e-Registration", "Digital deed &\nregistration services"),
        ("e-Stamping", "Stamp instrument\nfee & papers"),
        ("PTS", "Property Tax\nSystem"),
    ]

    start_x, start_y = Inches(0.45), Inches(1.55)
    card_w, card_h = Inches(2.4), Inches(1.55)
    gap_x, gap_y = Inches(0.2), Inches(0.2)

    for i, (code, desc) in enumerate(systems):
        row, col = divmod(i, 5)
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        card = add_rounded_rect(s, x, y, card_w, card_h, fill=GREEN_SOFT, line=GREEN_BORDER)
        t = s.shapes.add_textbox(x + Inches(0.1), y + Inches(0.25), card_w - Inches(0.2), Inches(0.4))
        add_text(t, [(code, 14, True, GREEN_DARK)], align=PP_ALIGN.CENTER)
        d = s.shapes.add_textbox(x + Inches(0.1), y + Inches(0.7), card_w - Inches(0.2), Inches(0.7))
        add_text(d, [(desc, 11, False, GRAY)], align=PP_ALIGN.CENTER)

    # bottom note
    note = s.shapes.add_textbox(Inches(0.5), Inches(5.15), Inches(12.3), Inches(0.4))
    add_text(
        note,
        [("My Properties also surfaces GPC (Green Property Certificate) as a complete actionable service hub.", 12, True, GREEN)],
        align=PP_ALIGN.CENTER,
    )
    add_footer(s, prs)


def slide_home_cards(prs):
    s = blank_slide(prs)
    add_eyebrow(s, "APPLICATION HOME SCREEN")
    add_title_block(s, "Service Cards on Punjab Zameen", "Each card opens a complete integrated system or module")

    cards = [
        ("My Properties", "GPC & actionable property hub"),
        ("Approved Housing Society", "HSMS — GIS & plot verification"),
        ("Get Record Copy", "Fard, Registry, Mutation, Girdawari"),
        ("Update Personal Records", "Citizen profile & record updates"),
        ("Application for Partition", "Land partition applications"),
        ("Complaints", "CMS — register & track grievances"),
        ("Payments", "Payment Gateway for challans"),
        ("Revenue Court Cases", "RCMS — case status & tracking"),
        ("Arazi Muawin", "Franchise locator & applications"),
        ("e-Registration", "Digital registration services"),
        ("e-Stamp", "Stamp instruments & fees"),
        ("Property Tax", "PTS — property tax services"),
        ("e-Leasing", "Digital leasing workflows"),
    ]

    start_x, start_y = Inches(0.45), Inches(1.45)
    card_w, card_h = Inches(3.05), Inches(0.95)
    gap_x, gap_y = Inches(0.18), Inches(0.15)

    for i, (title, desc) in enumerate(cards):
        row, col = divmod(i, 4)
        if row == 3 and col > 0:
            # last row center remaining - actually 13 cards: 4+4+4+1
            pass
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        if i == 12:
            x = start_x + 1.5 * (card_w + gap_x)
        card = add_rounded_rect(s, x, y, card_w, card_h, fill=WHITE, line=GREEN_BORDER)
        icon = add_rounded_rect(
            s, x + Inches(0.12), y + Inches(0.2), Inches(0.55), Inches(0.55),
            fill=GREEN_SOFT, line=GREEN_SOFT,
        )
        t = s.shapes.add_textbox(x + Inches(0.8), y + Inches(0.15), card_w - Inches(0.95), Inches(0.35))
        add_text(t, [(title, 12, True, BLACK)])
        d = s.shapes.add_textbox(x + Inches(0.8), y + Inches(0.5), card_w - Inches(0.95), Inches(0.35))
        add_text(d, [(desc, 10, False, GRAY)])

    add_footer(s, prs)


def slide_my_properties(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "My Properties Module",
        "Actionable property hub — from static list to service execution",
    )

    # AS-IS / TO-BE
    add_bullet_card(
        s, Inches(0.5), Inches(1.45), Inches(5.9), Inches(2.0),
        "AS-IS — Static List",
        [
            "Legacy app showed a plain static list of properties",
            "Zero functions or services could be performed",
            "No direct actions from property cards",
        ],
    )
    add_bullet_card(
        s, Inches(6.7), Inches(1.45), Inches(5.9), Inches(2.0),
        "TO-BE — Actionable Hub",
        [
            "Initiate Green Property Certificates (GPC)",
            "Request Fard / Registry copies",
            "Check court cases & transfer ownership",
        ],
    )

    # Tabs
    tabs = s.shapes.add_textbox(Inches(0.5), Inches(3.65), Inches(12.3), Inches(0.35))
    add_text(tabs, [("Property Categories (with total counts)", 14, True, GREEN_DARK)])

    for i, (name, desc) in enumerate([
        ("All", "Complete inventory of properties linked to the logged-in citizen"),
        ("Verified", "Properties confirmed against CLRMIS / ownership checks"),
        ("Unverified", "Properties awaiting verification or ownership confirmation"),
    ]):
        x = Inches(0.5) + i * Inches(4.15)
        card = add_rounded_rect(s, x, Inches(4.1), Inches(3.95), Inches(1.7), fill=GREEN_SOFT, line=GREEN_BORDER)
        add_circle(s, x + Inches(0.2), Inches(4.3), Inches(0.4), text=str(i + 1))
        t = s.shapes.add_textbox(x + Inches(0.75), Inches(4.35), Inches(3.0), Inches(0.35))
        add_text(t, [(name, 16, True, GREEN_DARK)])
        d = s.shapes.add_textbox(x + Inches(0.2), Inches(4.85), Inches(3.55), Inches(0.75))
        add_text(d, [(desc, 12, False, GRAY)])

    add_footer(s, prs)


def slide_workflow(prs, eyebrow, title, subtitle, steps, footer=True):
    """Horizontal numbered workflow cards matching attached GPC/CMS style."""
    s = blank_slide(prs)
    add_eyebrow(s, eyebrow)
    add_title_block(s, title, subtitle)

    n = len(steps)
    margin = Inches(0.35)
    gap = Inches(0.12)
    usable = prs.slide_width - 2 * margin - (n - 1) * gap
    card_w = usable / n
    card_h = Inches(4.35)
    top = Inches(1.45)

    for i, (step_title, desc, system) in enumerate(steps):
        x = margin + i * (card_w + gap)
        card = add_rounded_rect(s, x, top, card_w, card_h, fill=WHITE, line=GREEN_BORDER, line_width=1.75)

        # number circle centered near top
        circle_size = Inches(0.42)
        cx = x + (card_w - circle_size) / 2
        add_circle(s, cx, top + Inches(0.25), circle_size, fill=GREEN, text=str(i + 1))

        # title
        t = s.shapes.add_textbox(x + Inches(0.1), top + Inches(0.85), card_w - Inches(0.2), Inches(0.95))
        add_text(t, [(step_title, 12, True, BLACK)], align=PP_ALIGN.CENTER)

        # description
        d = s.shapes.add_textbox(x + Inches(0.12), top + Inches(1.9), card_w - Inches(0.24), Inches(1.5))
        add_text(d, [(desc, 11, False, GRAY)], align=PP_ALIGN.CENTER)

        # system label
        sys_box = s.shapes.add_textbox(x + Inches(0.08), top + card_h - Inches(0.7), card_w - Inches(0.16), Inches(0.5))
        add_text(sys_box, [(f"System: {system}", 11, True, GREEN)], align=PP_ALIGN.CENTER)

    if footer:
        add_footer(s, prs)
    return s


def slide_gpc_workflow(prs):
    steps = [
        ("100% Ownership Check", "System verifies 100% ownership share in CLRMIS", "CLRMIS"),
        ("Data Capture & English Name", "Auto-retrieves Urdu details; captures applicant English name", "PZA App"),
        ("Mobile OTP Verification", "Validates 6-digit OTP sent to registered mobile", "SMS Gateway"),
        ("Survey Representative", "Nominate optional representative for physical survey", "PZA App"),
        ("SCO/ARC Verification & PSID", "Officer verifies record; generates PSID fee challan", "LAMP"),
        ("Physical Survey & ADLR Approval", "Survey team completes survey; ADLR grants final approval", "LAMP"),
        ("GPC Issued & Digital Vault", "Digital Certificate generated & filed in Digital Vault", "Digital Vault"),
    ]
    slide_workflow(
        prs,
        "CORE APPLICATION MODULES",
        "Green Property Certificate (GPC) Workflow",
        "End-to-End Digital Certification Pipeline for 100% Owned Land",
        steps,
    )


def slide_hsms(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Approved Housing Society (HSMS)",
        "GIS satellite visualization of approved societies, block boundaries, and plot legal verification",
    )

    features = [
        ("GIS Satellite Interface", [
            "Interactive map with zoom and pan",
            "Full-screen polygon rendering",
            "Block and society boundary overlays",
        ]),
        ("Administrative Filters", [
            "Search by District and Tehsil",
            "Filter Authority-Owned societies",
            "Filter Private Housing Societies",
        ]),
        ("Color-Coded Plot Statuses", [
            "Mortgage — Blue",
            "Legal Case — Orange",
            "Available — Green / Sold — Grey",
        ]),
        ("Block & Plot Cards", [
            "Plot size and attributes",
            "Last transfer date",
            "Current legal / sale status",
        ]),
    ]
    positions = [
        (Inches(0.5), Inches(1.5)),
        (Inches(6.7), Inches(1.5)),
        (Inches(0.5), Inches(4.0)),
        (Inches(6.7), Inches(4.0)),
    ]
    for (x, y), (title, bullets) in zip(positions, features):
        add_bullet_card(s, x, y, Inches(5.9), Inches(2.2), title, bullets)
    add_footer(s, prs)


def slide_record_copy(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Get Record Copy Module",
        "Digitizing Fard, Registry & Mutation copies with integrated payments",
    )

    add_bullet_card(
        s, Inches(0.5), Inches(1.45), Inches(12.2), Inches(1.35),
        "Expanded Service Suite",
        [
            "Get Fard Copy  •  Get Registry Copy  •  Get Mutation (Intiqal) Copy  •  Crop Inspection Reports (Girdawari)",
            "Certified official copies with in-app fee settlement via Payment Gateway",
        ],
    )

    features = [
        ("Multi-Criteria Search", [
            "Lookup by e-Stamp Number",
            "Mutation Number / Registry Number",
            "Search by CNIC",
        ]),
        ("In-App Fee Settlement", [
            "Dedicated Payment Gateway",
            "Challan fee via Debit/Credit Card",
            "Instant receipt & vault storage",
        ]),
        ("Source Systems", [
            "CLRMIS integration for land records",
            "e-Stamp linkage where applicable",
            "Secure digital delivery to citizen",
        ]),
    ]
    for i, (title, bullets) in enumerate(features):
        x = Inches(0.5) + i * Inches(4.15)
        add_bullet_card(s, x, Inches(3.1), Inches(3.95), Inches(2.7), title, bullets)
    add_footer(s, prs)


def slide_update_records(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Update Personal Records",
        "Keep citizen identity and contact data synchronized for land services",
    )
    cards = [
        ("Profile Accuracy", [
            "Update personal information linked to land services",
            "Ensure CNIC and contact data remain current",
            "Reduce failed OTP / delivery issues",
        ]),
        ("Service Continuity", [
            "Accurate records required for GPC and payments",
            "Supports complaint and court case notifications",
            "Aligned with CIMS identity services",
        ]),
        ("Citizen Convenience", [
            "Self-service updates from Punjab Zameen",
            "Fewer office visits for demographic corrections",
            "Faster processing of downstream applications",
        ]),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = Inches(0.5) + i * Inches(4.15)
        add_bullet_card(s, x, Inches(1.6), Inches(3.95), Inches(4.0), title, bullets)
    add_footer(s, prs)


def slide_partition(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Application for Partition",
        "Digital intake for land partition requests initiated from Punjab Zameen",
    )
    cards = [
        ("Purpose", [
            "Enable citizens to apply for partition of land holdings",
            "Reduce manual paper-based submission",
            "Connect applications to backend PLRA workflows",
        ]),
        ("Citizen Journey", [
            "Select property from My Properties hub",
            "Submit partition application details",
            "Track progress within the application",
        ]),
        ("Value", [
            "Transparent application status",
            "Less dependency on intermediaries",
            "Aligned with digital land governance",
        ]),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = Inches(0.5) + i * Inches(4.15)
        add_bullet_card(s, x, Inches(1.6), Inches(3.95), Inches(4.0), title, bullets)
    add_footer(s, prs)


def slide_cms(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Complaints Module (CMS)",
        "Citizen grievance redressal integrated with CMS Helpdesk workflows",
    )

    # Two options
    add_bullet_card(
        s, Inches(0.5), Inches(1.45), Inches(5.9), Inches(2.35),
        "Option 1 — Track My Complaint",
        [
            "Check complaint status in real time",
            "Statuses: Pending, Resolved, Rejected",
            "View timeline and official updates",
            "Push notifications for progress changes",
        ],
    )
    add_bullet_card(
        s, Inches(6.7), Inches(1.45), Inches(5.9), Inches(2.35),
        "Option 2 — Register a Complaint",
        [
            "Select complaint type / category",
            "Fill required workflow form fields",
            "Review → Edit (if needed) → Submit",
            "Receive CMS reference number on success",
        ],
    )

    add_bullet_card(
        s, Inches(0.5), Inches(4.05), Inches(12.2), Inches(1.75),
        "Citizen Portal Capabilities",
        [
            "Register / log in, file complaints, monitor progress, and communicate with authorities",
            "Synchronized category & sub-category taxonomy from CMS",
            "Attachments supported (documents/photos, up to 5MB) during complaint lodging",
        ],
    )
    add_footer(s, prs)


def slide_cms_workflow(prs):
    steps = [
        ("Category Selection", "Synchronized category & sub-category taxonomy from CMS", "CMS"),
        ("Form & Attachment Upload", "Enter complaint details; attach documents/photos (up to 5MB)", "PZA App"),
        ("CMS Registration & Ref #", "Payload submitted via secure API; returns unique CMS Ref Number", "CMS"),
        ("Helpdesk Assignment", "CMS routes complaint to designated officer or district desk", "CMS Helpdesk"),
        ("Live Progress Tracking", "Real-time timeline updates delivered via push notifications", "PZA App"),
        ("Resolution & Feedback", "Citizen views official resolution text; rates resolution satisfaction", "CMS"),
    ]
    slide_workflow(
        prs,
        "CORE APPLICATION MODULES",
        "Complaints Module",
        "Citizen Grievance Redressal Integrated with CMS Helpdesk Workflows",
        steps,
    )


def slide_payments(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Payment Gateway Module",
        "Online payment of challans for Fard, e-Stamp, copy fee, and more",
    )

    add_bullet_card(
        s, Inches(0.5), Inches(1.4), Inches(5.9), Inches(2.4),
        "AS-IS — Manual Payment Bottlenecks",
        [
            "Only basic Get Record Copy & static history",
            "No in-app payment gateway",
            "Citizens visited bank booths for challan fee",
            "Limited modules and fragmented experience",
        ],
    )
    add_bullet_card(
        s, Inches(6.7), Inches(1.4), Inches(5.9), Inches(2.4),
        "TO-BE — Integrated Payment Gateway",
        [
            "CLRMIS + e-Stamp integrated fee collection",
            "Debit/Credit card payments in-app & web",
            "Automated Digital Vault for receipts",
            "Community module for guidance & discussion",
        ],
    )

    # Three components
    note = s.shapes.add_textbox(Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.3))
    add_text(note, [("Three Integrated Components", 14, True, GREEN_DARK)])

    comps = [
        ("1. Citizen Web Portal", "Enter challan number, fetch details, and pay fee online"),
        ("2. Punjab Zameen App", "Payment card on home + Record Copy tab integration"),
        ("3. Admin Portal", "Monitor, verify, reconcile, and audit all transactions"),
    ]
    for i, (title, desc) in enumerate(comps):
        x = Inches(0.5) + i * Inches(4.15)
        card = add_rounded_rect(s, x, Inches(4.4), Inches(3.95), Inches(1.4), fill=GREEN_SOFT, line=GREEN_BORDER)
        t = s.shapes.add_textbox(x + Inches(0.2), Inches(4.55), Inches(3.55), Inches(0.35))
        add_text(t, [(title, 13, True, GREEN_DARK)])
        d = s.shapes.add_textbox(x + Inches(0.2), Inches(4.95), Inches(3.55), Inches(0.65))
        add_text(d, [(desc, 12, False, GRAY)])

    add_footer(s, prs)


def slide_payment_sources(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Payment Gateway — Supported Fees & Future Scope",
        "Current collections and planned expansion across PLRA systems",
    )

    add_bullet_card(
        s, Inches(0.5), Inches(1.45), Inches(6.0), Inches(3.3),
        "Currently Supported — CLRMIS",
        [
            "Fard Fee (Copy Fee of Fard)",
            "Mutation Fee",
            "Crop Inspection Fee",
            "Registry Fee",
        ],
    )
    add_bullet_card(
        s, Inches(6.8), Inches(1.45), Inches(5.9), Inches(3.3),
        "Currently Supported — E-Stamp",
        [
            "Fee of all E-Stamp instruments / papers",
            "Online settlement without bank booth visits",
            "Receipt available for download & vault",
            "Audit-ready transaction logs",
        ],
    )

    future = add_rounded_rect(s, Inches(0.5), Inches(5.0), Inches(12.2), Inches(1.1), fill=GREEN_SOFT, line=GREEN_BORDER)
    t = s.shapes.add_textbox(Inches(0.75), Inches(5.15), Inches(11.7), Inches(0.35))
    add_text(t, [("Future Payment Gateway Integrations", 14, True, GREEN_DARK)])
    d = s.shapes.add_textbox(Inches(0.75), Inches(5.55), Inches(11.7), Inches(0.4))
    add_text(
        d,
        [("RCMS  •  HSMS  •  Arazi Muawin  •  Additional PLRA systems via web portal & Punjab Zameen app", 13, False, GRAY)],
    )
    add_footer(s, prs)


def slide_payment_steps(prs):
    steps = [
        ("Enter Challan Number", "Citizen enters PSID / challan number on portal or app", "Payment Gateway"),
        ("Fetch & Review Details", "System retrieves and displays challan amount & particulars", "Payment Gateway"),
        ("Click Pay Now", "Proceed to the bank payment page securely", "Bank Page"),
        ("Enter Card Details", "Provide Debit / Credit card information", "Bank Page"),
        ("Click Pay", "Authorize and complete the payment transaction", "Bank Page"),
        ("Download Receipt", "Obtain payment receipt; store in Digital Vault", "Digital Vault"),
    ]
    slide_workflow(
        prs,
        "CORE APPLICATION MODULES",
        "Online Payment Steps",
        "Citizen journey for challan fee payment via web portal & Punjab Zameen",
        steps,
    )


def slide_rcms(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Revenue Court Cases (RCMS)",
        "Revenue Court Management System — case visibility from Punjab Zameen",
    )
    cards = [
        ("What Citizens Get", [
            "Access revenue court case information",
            "Track case status from the mobile app",
            "Reduce dependency on physical court desks",
        ]),
        ("System Integration", [
            "Powered by RCMS backend",
            "Linked from home screen card",
            "Fits enterprise interoperability model",
        ]),
        ("Future Enablement", [
            "Planned Payment Gateway linkage",
            "Online fee settlement for RCMS challans",
            "Unified citizen experience with other PLRA systems",
        ]),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = Inches(0.5) + i * Inches(4.15)
        add_bullet_card(s, x, Inches(1.6), Inches(3.95), Inches(4.0), title, bullets)
    add_footer(s, prs)


def slide_arazi(prs):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(
        s,
        "Arazi Muawin Module",
        "Locate authorized franchise centers and submit new franchise applications",
    )
    features = [
        ("GPS Proximity Engine", [
            "Nearest Arazi Muawin office on interactive map",
            "GPS proximity enhancements planned",
            "Citizen-friendly discovery of centers",
        ]),
        ("Dual View Mode", [
            "Interactive GIS Map View",
            "Filtered List View",
            "Seamless toggle between modes",
        ]),
        ("Direct Device Actions", [
            "One-tap Call opens phone dialer",
            "Navigate launches Google Maps",
            "Faster real-world assistance",
        ]),
        ("Franchise Application Portal", [
            "Capture personal data & CNIC",
            "Office GPS & infrastructure details",
            "Tenancy and application particulars",
        ]),
    ]
    positions = [
        (Inches(0.5), Inches(1.5)),
        (Inches(6.7), Inches(1.5)),
        (Inches(0.5), Inches(4.0)),
        (Inches(6.7), Inches(4.0)),
    ]
    for (x, y), (title, bullets) in zip(positions, features):
        add_bullet_card(s, x, y, Inches(5.9), Inches(2.2), title, bullets)
    add_footer(s, prs)


def slide_simple_system(prs, title, subtitle, cards):
    s = blank_slide(prs)
    add_eyebrow(s)
    add_title_block(s, title, subtitle)
    n = len(cards)
    width = Inches(12.2 / n - 0.15) if n else Inches(3.9)
    for i, (ctitle, bullets) in enumerate(cards):
        x = Inches(0.5) + i * (width + Inches(0.2))
        add_bullet_card(s, x, Inches(1.6), width, Inches(4.0), ctitle, bullets)
    add_footer(s, prs)


def slide_ereg(prs):
    slide_simple_system(
        prs,
        "e-Registration",
        "Digital registration services integrated into Punjab Zameen",
        [
            ("System Role", [
                "Dedicated e-Registration system card on home screen",
                "Citizen access to registration-related services",
                "Part of the 10+ system enterprise ecosystem",
            ]),
            ("Citizen Value", [
                "Reduced need for multiple portals",
                "Consistent identity via CIMS where applicable",
                "Aligned digital land transaction journey",
            ]),
            ("Platform Fit", [
                "Accessible from Punjab Zameen menu",
                "Complements e-Stamp and Record Copy",
                "Supports end-to-end property lifecycle",
            ]),
        ],
    )


def slide_estamp(prs):
    slide_simple_system(
        prs,
        "e-Stamp",
        "Stamp instruments and online fee collection through Payment Gateway",
        [
            ("Capabilities", [
                "Access e-Stamp services from the app card",
                "Fee of e-Stamp instruments / papers supported",
                "Integrated with Payment Gateway",
            ]),
            ("Payment Linkage", [
                "Enter challan and pay online",
                "Debit / Credit card settlement",
                "Downloadable payment receipt",
            ]),
            ("Citizen Outcome", [
                "No physical bank booth for stamp fees",
                "Faster completion of registration journeys",
                "Transparent fee tracking & audit trail",
            ]),
        ],
    )


def slide_pts(prs):
    slide_simple_system(
        prs,
        "Property Tax (PTS)",
        "Property Tax System services available from Punjab Zameen",
        [
            ("System Access", [
                "Dedicated Property Tax card on home screen",
                "PTS as an integrated government system",
                "Citizen self-service entry point",
            ]),
            ("Citizen Empowerment", [
                "Tax-related actions without middlemen",
                "Visibility into property tax obligations",
                "Consistent with remodeling objectives",
            ]),
            ("Ecosystem", [
                "Works alongside My Properties & Payments",
                "Supports financial transparency goals",
                "Extends PLRA digital service coverage",
            ]),
        ],
    )


def slide_eleasing(prs):
    slide_simple_system(
        prs,
        "e-Leasing",
        "Digital leasing workflows surfaced through Punjab Zameen",
        [
            ("Module Purpose", [
                "Dedicated e-Leasing system card",
                "Citizen access to leasing-related services",
                "Complements ownership and registration modules",
            ]),
            ("Digital Journey", [
                "Initiate leasing processes from the app",
                "Reduce paper-based leasing coordination",
                "Trackable digital interactions",
            ]),
            ("Platform Benefit", [
                "Single-app access to PLRA services",
                "Interoperable enterprise architecture",
                "Lower citizen load on physical offices",
            ]),
        ],
    )


def slide_future(prs):
    s = blank_slide(prs)
    add_eyebrow(s, "ROADMAP")
    add_title_block(
        s,
        "Future Integrations",
        "Expanding Payment Gateway and system reach across web portal & mobile app",
    )
    cards = [
        ("Payment Gateway Expansion", [
            "RCMS challan fee online",
            "HSMS payment enablement",
            "Arazi Muawin fee collection",
            "More PLRA systems over time",
        ]),
        ("Channels", [
            "Punjab Zameen mobile application",
            "Citizen Payment Gateway web portal",
            "Admin portal for reconciliation",
            "Digital Vault for receipts & certificates",
        ]),
        ("Outcomes", [
            "Fewer bank booth visits",
            "Faster service completion",
            "Stronger audit & transparency",
            "Unified citizen land services hub",
        ]),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = Inches(0.5) + i * Inches(4.15)
        add_bullet_card(s, x, Inches(1.6), Inches(3.95), Inches(4.0), title, bullets)
    add_footer(s, prs)


def slide_closing(prs):
    s = blank_slide(prs)
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.0), prs.slide_width, Inches(3.2))
    band.fill.solid()
    band.fill.fore_color.rgb = GREEN
    band.line.fill.background()

    t = s.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.5), Inches(0.7))
    add_text(t, [("Thank You", 40, True, WHITE)], align=PP_ALIGN.CENTER)

    d = s.shapes.add_textbox(Inches(0.8), Inches(3.3), Inches(11.5), Inches(0.5))
    add_text(
        d,
        [("Punjab Zameen — Empowering Citizens Through Digital Land Services", 16, False, RGBColor(0xD8, 0xF0, 0xE4))],
        align=PP_ALIGN.CENTER,
    )

    f = s.shapes.add_textbox(Inches(0.8), Inches(4.0), Inches(11.5), Inches(0.4))
    add_text(
        f,
        [("PUNJAB LAND RECORDS AUTHORITY (PLRA)  |  GOVT OF THE PUNJAB", 12, True, WHITE)],
        align=PP_ALIGN.CENTER,
    )


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_title(prs)
    slide_agenda(prs)
    slide_why_remodel(prs)
    slide_ecosystem(prs)
    slide_home_cards(prs)
    slide_my_properties(prs)
    slide_gpc_workflow(prs)
    slide_hsms(prs)
    slide_record_copy(prs)
    slide_update_records(prs)
    slide_partition(prs)
    slide_cms(prs)
    slide_cms_workflow(prs)
    slide_payments(prs)
    slide_payment_sources(prs)
    slide_payment_steps(prs)
    slide_rcms(prs)
    slide_arazi(prs)
    slide_ereg(prs)
    slide_estamp(prs)
    slide_pts(prs)
    slide_eleasing(prs)
    slide_future(prs)
    slide_closing(prs)

    out = "/workspace/Punjab_Zameen_Application.pptx"
    prs.save(out)
    print(f"Saved {out} with {len(prs.slides)} slides")
    return out


if __name__ == "__main__":
    build()
