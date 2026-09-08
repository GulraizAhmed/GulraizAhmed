#!/usr/bin/env python3
"""Generate formal PLRA LMS PowerPoint presentation matching attached green/white design."""

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu
from pathlib import Path


GREEN = RGBColor(0x00, 0x68, 0x37)
GREEN_DARK = RGBColor(0x0B, 0x4F, 0x2C)
GREEN_MID = RGBColor(0x1B, 0x7A, 0x4A)
GREEN_SOFT = RGBColor(0xE8, 0xF5, 0xEE)
GREEN_BORDER = RGBColor(0x2F, 0x9E, 0x6B)
TITLE = RGBColor(0x00, 0x68, 0x37)
NAVY = RGBColor(0x1A, 0x2B, 0x4A)
GRAY = RGBColor(0x4A, 0x4A, 0x4A)
GRAY_LIGHT = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x22, 0x22, 0x22)
BG = RGBColor(0xFB, 0xFB, 0xFB)

ASSETS = Path("/workspace/lms_assets")


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
        p.space_after = Pt(3)
        run = p.add_run()
        run.text = text
        set_run(run, size=size, bold=bold, color=color, italic=italic)


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def fill_bg(slide, prs, color=BG):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    # send to back via z-order: recreate later content on top; pptx adds in order so this first is fine


def dots(slide, left, top, cols=8, rows=3, size=0.06, gap=0.14, color=GREEN_BORDER):
    for r in range(rows):
        for c in range(cols):
            s = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(left + c * gap),
                Inches(top + r * gap),
                Inches(size),
                Inches(size),
            )
            s.fill.solid()
            s.fill.fore_color.rgb = color
            s.line.fill.background()
            try:
                s.fill.fore_color.brightness = 0.35
            except Exception:
                pass


def corner_accents(slide, prs, style="angled"):
    # top-right green rounded block
    tr = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        prs.slide_width - Inches(2.4),
        Inches(-0.6),
        Inches(3.0),
        Inches(1.8),
    )
    tr.fill.solid()
    tr.fill.fore_color.rgb = GREEN
    tr.line.fill.background()
    try:
        tr.adjustments[0] = 0.25
    except Exception:
        pass

    # bottom wave / angled bars
    b1 = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_TRIANGLE,
        Inches(-0.2),
        prs.slide_height - Inches(1.35),
        Inches(5.2),
        Inches(1.5),
    )
    b1.fill.solid()
    b1.fill.fore_color.rgb = GREEN
    b1.line.fill.background()
    # flip triangle via rotation isn't simple; use parallelogram bars instead
    slide.shapes._spTree.remove(b1._element)

    bar1 = slide.shapes.add_shape(
        MSO_SHAPE.PARALLELOGRAM,
        Inches(-0.8),
        prs.slide_height - Inches(1.05),
        Inches(7.5),
        Inches(0.55),
    )
    bar1.fill.solid()
    bar1.fill.fore_color.rgb = GREEN
    bar1.line.fill.background()

    bar2 = slide.shapes.add_shape(
        MSO_SHAPE.PARALLELOGRAM,
        Inches(1.2),
        prs.slide_height - Inches(0.7),
        Inches(8.0),
        Inches(0.45),
    )
    bar2.fill.solid()
    bar2.fill.fore_color.rgb = GREEN_MID
    bar2.line.fill.background()

    br = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        prs.slide_width - Inches(1.5),
        prs.slide_height - Inches(1.3),
        Inches(2.2),
        Inches(1.8),
    )
    br.fill.solid()
    br.fill.fore_color.rgb = GREEN_DARK
    br.line.fill.background()
    try:
        br.adjustments[0] = 0.2
    except Exception:
        pass


def brand_header(slide, text="PLRA - LMS"):
    # logo circle proxy
    logo = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.45), Inches(0.28), Inches(0.42), Inches(0.42))
    logo.fill.solid()
    logo.fill.fore_color.rgb = GREEN
    logo.line.fill.background()
    add_text(logo, [("P", 12, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    box = slide.shapes.add_textbox(Inches(0.98), Inches(0.32), Inches(4.5), Inches(0.35))
    add_text(box, [(text, 16, True, GREEN)])


def slide_title_block(slide, title, y=0.85):
    box = slide.shapes.add_textbox(Inches(0.55), Inches(y), Inches(8.5), Inches(0.55))
    add_text(box, [(title.upper(), 30, True, GREEN)])


def add_check_bullets(slide, items, left=0.55, top=1.55, width=7.2, size=15):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(5.0))
    lines = []
    for it in items:
        lines.append((f"➤  {it}", size, False, GRAY))
    add_text(box, lines)


def add_arrow_bullets(slide, items, left=0.55, top=1.55, width=7.5, size=16):
    y = top
    for it in items:
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(y), Inches(0.32), Inches(0.32))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GREEN
        circ.line.fill.background()
        add_text(circ, [("›", 14, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        tb = slide.shapes.add_textbox(Inches(left + 0.45), Inches(y - 0.02), Inches(width), Inches(0.55))
        add_text(tb, [(it, size, False, GRAY)])
        y += 0.72


def set_transition(slide, kind="fade"):
    sld = slide._element
    for child in list(sld):
        if child.tag == qn("p:transition"):
            sld.remove(child)
    transition = etree.Element(qn("p:transition"))
    transition.set("spd", "med")
    transition.set("advClick", "1")
    if kind == "push":
        node = etree.SubElement(transition, qn("p:push"))
        node.set("dir", "l")
    elif kind == "wipe":
        node = etree.SubElement(transition, qn("p:wipe"))
        node.set("dir", "l")
    else:
        etree.SubElement(transition, qn("p:fade"))
    sld.append(transition)


def try_picture(slide, path, left, top, width, height):
    p = Path(path)
    if not p.exists():
        return None
    return slide.shapes.add_picture(str(p), Inches(left), Inches(top), Inches(width), Inches(height))


def image_panel(slide, prs, path=None, left=8.3, top=1.3, width=4.5, height=4.6):
    # green rounded frame
    frame = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    frame.fill.solid()
    frame.fill.fore_color.rgb = GREEN
    frame.line.fill.background()
    try:
        frame.adjustments[0] = 0.12
    except Exception:
        pass
    if path and Path(path).exists():
        # inset image
        slide.shapes.add_picture(
            str(path),
            Inches(left + 0.12),
            Inches(top + 0.12),
            Inches(width - 0.24),
            Inches(height - 0.24),
        )
    else:
        inner = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left + 0.15),
            Inches(top + 0.15),
            Inches(width - 0.3),
            Inches(height - 0.3),
        )
        inner.fill.solid()
        inner.fill.fore_color.rgb = GREEN_SOFT
        inner.line.fill.background()


# ---------------- slides ----------------

def slide_title(prs):
    s = blank(prs)
    fill_bg(s, prs, WHITE)
    corner_accents(s, prs)
    dots(s, 5.8, 0.35, cols=7, rows=2)

    # brand
    logo = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.55), Inches(0.4), Inches(0.5), Inches(0.5))
    logo.fill.solid()
    logo.fill.fore_color.rgb = GREEN
    logo.line.fill.background()
    add_text(logo, [("P", 14, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    box = s.shapes.add_textbox(Inches(1.15), Inches(0.48), Inches(3), Inches(0.35))
    add_text(box, [("PLRA", 18, True, GREEN)])

    box = s.shapes.add_textbox(Inches(0.55), Inches(2.0), Inches(7.5), Inches(1.0))
    add_text(box, [("LMS", 66, True, GREEN)])
    box = s.shapes.add_textbox(Inches(0.55), Inches(3.1), Inches(7.5), Inches(0.55))
    add_text(box, [("Learning Management System", 26, True, GREEN)])

    # website banner
    banner = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.35), Inches(6.6), Inches(0.7)
    )
    banner.fill.solid()
    banner.fill.fore_color.rgb = GREEN
    banner.line.fill.background()
    try:
        banner.adjustments[0] = 0.5
    except Exception:
        pass
    add_text(
        banner,
        [("🌐   Website   lms.punjab-zameen.gov.pk", 16, True, WHITE)],
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )

    # right decorative panel using reference if available
    image_panel(s, prs, ASSETS / "slide_title.png", left=8.55, top=1.55, width=4.2, height=4.3)


def slide_agenda(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.2, 0.3, cols=6, rows=2)
    slide_title_block(s, "Agenda")

    left = [
        "01  Project Background",
        "02  Key Objectives",
        "03  Purpose & Scope",
        "04  Types of Users",
        "05  Role Dashboards",
        "06  Core Platform Capabilities",
    ]
    right = [
        "07  Learning & Assessment",
        "08  Certification & Analytics",
        "09  Application Workflows",
        "10  Technology Stack",
        "11  Operating Environment",
        "12  Closing",
    ]
    for i, items in enumerate([left, right]):
        x = 0.55 if i == 0 else 7.0
        card = s.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.6), Inches(5.7), Inches(4.5)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        card.line.width = Pt(1.5)
        try:
            card.adjustments[0] = 0.08
        except Exception:
            pass
        tb = s.shapes.add_textbox(Inches(x + 0.35), Inches(1.85), Inches(5.1), Inches(4.0))
        add_text(tb, [(it, 16, False, GRAY) for it in items])


def slide_background(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Project Background")

    paras = [
        "LMS was developed to modernize and standardize learning across PLRA by providing a centralized digital platform for training, assessment, certification, and compliance management.",
        "Traditional training processes often rely on manual coordination, fragmented records, and limited visibility into learner progress. LMS addresses these challenges by delivering a secure, scalable, and data-driven learning ecosystem that supports administrators, trainers, and learners through a single platform.",
    ]
    tb = s.shapes.add_textbox(Inches(0.55), Inches(1.6), Inches(7.4), Inches(4.5))
    lines = []
    for p in paras:
        lines.append((p, 16, False, GRAY))
        lines.append(("", 8, False, GRAY))
    add_text(tb, lines)
    image_panel(s, prs, ASSETS / "slide_background.png", left=8.4, top=1.45, width=4.4, height=4.5)


def slide_objectives(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Key Objectives")
    items = [
        "Standardize training delivery across all departments and regions",
        "Digitize assessments, certifications, and learner records",
        "Improve visibility into training progress and compliance status",
        "Reduce administrative effort through automation and self-service workflows",
    ]
    add_arrow_bullets(s, items, left=0.55, top=1.7, width=7.4, size=17)
    image_panel(s, prs, ASSETS / "slide_objectives.png", left=8.5, top=1.4, width=4.3, height=4.4)


def slide_purpose_scope(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Purpose & Scope")

    # purpose card
    c1 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(12.3), Inches(1.7))
    c1.fill.solid()
    c1.fill.fore_color.rgb = GREEN_SOFT
    c1.line.color.rgb = GREEN_BORDER
    t = s.shapes.add_textbox(Inches(0.75), Inches(1.65), Inches(11.8), Inches(0.35))
    add_text(t, [("PURPOSE", 14, True, GREEN)])
    b = s.shapes.add_textbox(Inches(0.75), Inches(2.05), Inches(11.8), Inches(0.95))
    add_text(
        b,
        [
            (
                "PULSE LMS is a Laravel-based Learning Management System for PLRA that delivers CIMS authentication, "
                "role-based access, trainee onboarding, course delivery, quizzes, certificates, reports, notifications, "
                "support tickets, chat, forums, and protected media access — reflecting the software as implemented.",
                14,
                False,
                GRAY,
            )
        ],
    )

    # in / out scope
    for i, (title, bullets) in enumerate(
        [
            (
                "IN SCOPE",
                [
                    "CIMS login, RBAC, onboarding & approvals",
                    "Courses, modules, lessons, quizzes & media",
                    "Progress, certificates, badges & reports",
                    "Alerts, support tickets, chat & forums",
                ],
            ),
            (
                "OUT OF SCOPE",
                [
                    "Authoring of legal training content itself",
                    "Video / translation / trainer-guide production",
                    "Content creation remains a project deliverable",
                    "LMS stores and delivers materials once ready",
                ],
            ),
        ]
    ):
        x = 0.5 + i * 6.35
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(3.45), Inches(6.1), Inches(2.85))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        ht = s.shapes.add_textbox(Inches(x + 0.25), Inches(3.6), Inches(5.5), Inches(0.35))
        add_text(ht, [(title, 14, True, GREEN)])
        bt = s.shapes.add_textbox(Inches(x + 0.25), Inches(4.1), Inches(5.5), Inches(2.0))
        add_text(bt, [(f"•  {b}", 13, False, GRAY) for b in bullets])


def slide_users(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Types of Users")

    users = [
        ("01", "Super Admin", "Platform governance & control"),
        ("02", "Admin", "Learning operations & management"),
        ("03", "Trainer", "Course delivery & learner insights"),
        ("04", "Trainee", "Personalized learning journey"),
    ]
    for i, (num, name, desc) in enumerate(users):
        x = 0.55 + i * 3.15
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.1), Inches(2.95), Inches(3.4))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        card.line.width = Pt(1.75)
        try:
            card.adjustments[0] = 0.1
        except Exception:
            pass
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.95), Inches(2.45), Inches(1.0), Inches(1.0))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GREEN
        circ.line.fill.background()
        add_text(circ, [(num, 18, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        nt = s.shapes.add_textbox(Inches(x + 0.15), Inches(3.7), Inches(2.65), Inches(0.45))
        add_text(nt, [(name, 16, True, GREEN)], align=PP_ALIGN.CENTER)
        dt = s.shapes.add_textbox(Inches(x + 0.15), Inches(4.25), Inches(2.65), Inches(0.9))
        add_text(dt, [(desc, 12, False, GRAY)], align=PP_ALIGN.CENTER)


def role_slide(prs, title, subtitle, intro, responsibilities, image=None):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 5.8, 0.28, cols=6, rows=2)

    t = s.shapes.add_textbox(Inches(0.55), Inches(0.9), Inches(7.8), Inches(0.5))
    add_text(t, [(title.upper(), 28, True, GREEN)])
    st = s.shapes.add_textbox(Inches(0.55), Inches(1.4), Inches(7.8), Inches(0.35))
    add_text(st, [(subtitle, 15, True, GREEN_MID)])
    intro_box = s.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(7.6), Inches(1.15))
    add_text(intro_box, [(intro, 13, False, GRAY)])

    kr = s.shapes.add_textbox(Inches(0.55), Inches(3.1), Inches(7.5), Inches(0.3))
    add_text(kr, [("Key Responsibilities", 14, True, GREEN)])
    y = 3.5
    for item in responsibilities:
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.55), Inches(y), Inches(0.26), Inches(0.26))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GREEN
        circ.line.fill.background()
        add_text(circ, [("✓", 10, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        tb = s.shapes.add_textbox(Inches(0.95), Inches(y - 0.02), Inches(7.1), Inches(0.35))
        add_text(tb, [(item, 13, False, GRAY)])
        y += 0.42

    image_panel(s, prs, image, left=8.45, top=1.35, width=4.35, height=4.55)


def slide_capabilities_auth_users(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Core Capabilities — Access & Users")

    cards = [
        (
            "Authentication & Access",
            [
                "CIMS login with PKCE for trainees, admins & trainers",
                "Role mapping and server-side session context",
                "Auto-provision trainees; no auto-create for admins",
                "Pending approval, profile completion & logout",
            ],
        ),
        (
            "Dashboards",
            [
                "Role-specific dashboards for every user type",
                "Admin: stats, enrollments & performance links",
                "Trainer: course and trainee metrics",
                "Trainee: active courses, hours, badges & calendar",
            ],
        ),
        (
            "User & Master Data",
            [
                "Search, filter, activate/deactivate trainees",
                "Approve/reject registrations & assign ranks",
                "Manage admins, trainers, education & ranks",
                "District / tehsil data and rank-course mapping",
            ],
        ),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = 0.45 + i * 4.25
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.55), Inches(4.05), Inches(4.6))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.55), Inches(0.1), Inches(4.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = GREEN
        bar.line.fill.background()
        ht = s.shapes.add_textbox(Inches(x + 0.25), Inches(1.75), Inches(3.6), Inches(0.45))
        add_text(ht, [(title, 15, True, GREEN)])
        bt = s.shapes.add_textbox(Inches(x + 0.25), Inches(2.35), Inches(3.6), Inches(3.5))
        add_text(bt, [(f"•  {b}", 12, False, GRAY) for b in bullets])


def slide_capabilities_learning(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Learning Delivery & Assessment")

    cards = [
        (
            "Course Management",
            [
                "Create/edit bilingual courses, categories & topics",
                "Modules, lessons, quizzes, duration & outcomes",
                "Trainer & rank assignment; activate/deactivate",
                "Enrollment due dates and expired-course handling",
            ],
        ),
        (
            "Content & Media",
            [
                "Videos, PDFs, documents, subtitles & resources",
                "Media library with folders and S3/MinIO storage",
                "Protected streaming for authorized learners only",
                "Block unauthorized direct media navigation",
            ],
        ),
        (
            "Quizzes & Progress",
            [
                "MCQ / fill-in quizzes with CSV question upload",
                "Time limits, passing scores & auto-grading",
                "Lesson/course progress and study-hour tracking",
                "Completion when lessons done & quizzes passed",
            ],
        ),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = 0.45 + i * 4.25
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.55), Inches(4.05), Inches(4.6))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.55), Inches(0.1), Inches(4.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = GREEN
        bar.line.fill.background()
        ht = s.shapes.add_textbox(Inches(x + 0.25), Inches(1.75), Inches(3.6), Inches(0.45))
        add_text(ht, [(title, 15, True, GREEN)])
        bt = s.shapes.add_textbox(Inches(x + 0.25), Inches(2.35), Inches(3.6), Inches(3.5))
        add_text(bt, [(f"•  {b}", 12, False, GRAY) for b in bullets])


def slide_capabilities_cert_reports(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Certification, Analytics & Engagement")

    cards = [
        (
            "Certificates & Badges",
            [
                "Auto-issue certificates on eligibility",
                "Hash-based public certificate verification",
                "Template, signature & settings support",
                "Gamification badges for learning milestones",
            ],
        ),
        (
            "Reports & Analytics",
            [
                "Admin overview with charts & summaries",
                "Trainee/course progress with rich filters",
                "District, tehsil, rank & status dimensions",
                "CSV export for operational reporting",
            ],
        ),
        (
            "Alerts & Support",
            [
                "Admin alerts and deadline notifications",
                "Notification history with read status",
                "Support tickets, chat & discussion forums",
                "System, certificate & profile settings",
            ],
        ),
    ]
    for i, (title, bullets) in enumerate(cards):
        x = 0.45 + i * 4.25
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.55), Inches(4.05), Inches(4.6))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.55), Inches(0.1), Inches(4.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = GREEN
        bar.line.fill.background()
        ht = s.shapes.add_textbox(Inches(x + 0.25), Inches(1.75), Inches(3.6), Inches(0.45))
        add_text(ht, [(title, 15, True, GREEN)])
        bt = s.shapes.add_textbox(Inches(x + 0.25), Inches(2.35), Inches(3.6), Inches(3.5))
        add_text(bt, [(f"•  {b}", 12, False, GRAY) for b in bullets])


def workflow_cards_slide(prs, title, workflows):
    """workflows: list of (name, steps_list) — 4 per slide ideally."""
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, title)

    n = len(workflows)
    # 2x2 grid
    positions = [
        (0.45, 1.5),
        (6.85, 1.5),
        (0.45, 4.15),
        (6.85, 4.15),
    ]
    for i, (name, steps) in enumerate(workflows[:4]):
        x, y = positions[i]
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(6.05), Inches(2.4))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        # number pill
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.2), Inches(y + 0.2), Inches(0.38), Inches(0.38))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GREEN
        circ.line.fill.background()
        add_text(circ, [(f"{i+1:02d}", 11, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        ht = s.shapes.add_textbox(Inches(x + 0.7), Inches(y + 0.22), Inches(5.1), Inches(0.35))
        add_text(ht, [(name, 14, True, GREEN)])
        bt = s.shapes.add_textbox(Inches(x + 0.25), Inches(y + 0.7), Inches(5.55), Inches(1.55))
        add_text(bt, [(f"•  {st}", 11, False, GRAY) for st in steps])


def slide_tech_stack(prs):
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Technology Stack & Operating Environment")

    intro = s.shapes.add_textbox(Inches(0.55), Inches(1.45), Inches(12.2), Inches(0.7))
    add_text(
        intro,
        [
            (
                "LMS is a server-rendered Laravel web application organized by domains. It runs with PHP-FPM and Nginx, "
                "uses a relational database, integrates with CIMS for authentication, and stores media via S3-compatible storage (MinIO).",
                13,
                False,
                GRAY,
            )
        ],
    )

    stacks = [
        ("Application", ["Laravel (PHP 8.2+)", "Blade role-specific UI", "Composer dependencies", "Node / Vite frontend assets"]),
        ("Runtime", ["Docker containers", "Nginx reverse proxy", "PHP-FPM app runtime", "Queue & scheduler workers"]),
        ("Data & Storage", ["Relational database", "S3 / MinIO object storage", "Persistent volumes", "CSV reporting data"]),
        ("Integrations", ["CIMS auth (PKCE)", "SMTP / mail services", "HTTPS termination", "Modern browser clients"]),
    ]
    for i, (title, bullets) in enumerate(stacks):
        x = 0.45 + i * 3.2
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.35), Inches(3.05), Inches(3.7))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        head = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(2.35), Inches(3.05), Inches(0.55))
        head.fill.solid()
        head.fill.fore_color.rgb = GREEN
        head.line.fill.background()
        ht = s.shapes.add_textbox(Inches(x + 0.1), Inches(2.45), Inches(2.85), Inches(0.35))
        add_text(ht, [(title, 14, True, WHITE)], align=PP_ALIGN.CENTER)
        bt = s.shapes.add_textbox(Inches(x + 0.2), Inches(3.1), Inches(2.7), Inches(2.7))
        add_text(bt, [(f"•  {b}", 13, False, GRAY) for b in bullets])


def slide_closing(prs):
    s = blank(prs)
    fill_bg(s, prs, WHITE)
    # full green band
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.0), prs.slide_width, Inches(3.4))
    band.fill.solid()
    band.fill.fore_color.rgb = GREEN
    band.line.fill.background()
    dots(s, 5.5, 0.4, cols=8, rows=2)

    logo = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.25), Inches(2.25), Inches(0.7), Inches(0.7))
    logo.fill.solid()
    logo.fill.fore_color.rgb = WHITE
    logo.line.fill.background()
    add_text(logo, [("P", 18, True, GREEN)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    t = s.shapes.add_textbox(Inches(0.8), Inches(3.15), Inches(11.7), Inches(0.7))
    add_text(t, [("Thank You", 42, True, WHITE)], align=PP_ALIGN.CENTER)
    d = s.shapes.add_textbox(Inches(0.8), Inches(3.9), Inches(11.7), Inches(0.45))
    add_text(
        d,
        [("PLRA LMS — Modernizing Learning Across Punjab Land Records Authority", 15, False, RGBColor(0xD8, 0xF0, 0xE4))],
        align=PP_ALIGN.CENTER,
    )
    f = s.shapes.add_textbox(Inches(0.8), Inches(4.5), Inches(11.7), Inches(0.35))
    add_text(
        f,
        [("lms.punjab-zameen.gov.pk  |  Punjab Land Records Authority (PLRA)", 13, True, WHITE)],
        align=PP_ALIGN.CENTER,
    )


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_title(prs)
    slide_agenda(prs)
    slide_background(prs)
    slide_objectives(prs)
    slide_purpose_scope(prs)
    slide_users(prs)

    role_slide(
        prs,
        "Super Admin",
        "Complete Platform Governance & Control",
        "The Super Admin dashboard provides centralized oversight of the entire LMS ecosystem, enabling efficient management of administrators and trainers.",
        [
            "Manage Admin & Trainer Accounts",
            "Monitor Platform Health & System Resources",
            "Configure Roles, Permissions & Access Controls",
            "Oversee Alerts, Compliance & Governance",
        ],
        ASSETS / "slide_superadmin.png",
    )
    role_slide(
        prs,
        "Admin",
        "Centralized Learning Operations & Management",
        "The Admin dashboard serves as the operational hub of LMS, enabling efficient management of courses, trainers, trainees, approvals, and platform activities from a single interface.",
        [
            "Create & Manage Courses, Categories & Topics",
            "Review & Approve Trainee Registrations",
            "Manage Trainers, Ranks & Learning Assignments",
            "Publish Alerts & System Notifications",
            "Oversee Media Library & Learning Resources",
            "Monitor Support Tickets & User Activity",
        ],
        ASSETS / "slide_admin.png",
    )
    role_slide(
        prs,
        "Trainer",
        "Empowering Instructors Through Data & Insights",
        "The Trainer dashboard provides a centralized workspace for managing assigned courses, monitoring learner progress, and tracking performance outcomes to ensure successful course delivery.",
        [
            "Manage Assigned Courses & Learning Activities",
            "Monitor Student Progress & Course Completion",
            "Analyze Learner Performance & Assessment Results",
            "Identify At-Risk Learners Through Analytics",
            "Respond to Course-Related Support Tickets",
            "Improve Learning Outcomes Through Data-Driven Insights",
        ],
        ASSETS / "slide_objectives.png",
    )
    role_slide(
        prs,
        "Trainee",
        "Your Learning. Your Progress. Your Growth.",
        "The Trainee dashboard provides a personalized learning experience, enabling users to access courses, track progress, manage learning activities, and achieve certification goals through a single, intuitive platform.",
        [
            "Access Enrolled Courses & Learning Materials",
            "Track Lesson Progress & Course Completion",
            "Monitor Study Hours & Learning Performance",
            "Stay Updated with Alerts & Notifications",
            "Manage Calendar, Deadlines & Learning Tasks",
            "Create & Track Support Tickets",
            "Earn Certificates Upon Course Completion",
        ],
        ASSETS / "slide_background.png",
    )

    slide_capabilities_auth_users(prs)
    slide_capabilities_learning(prs)
    slide_capabilities_cert_reports(prs)

    workflow_cards_slide(
        prs,
        "Application Workflows — Access & Learning",
        [
            (
                "Login",
                [
                    "Select login mode and redirect to CIMS (PKCE)",
                    "Exchange code, resolve user/role, store session",
                    "Route to dashboard, profile, or pending approval",
                ],
            ),
            (
                "Trainee Onboarding",
                [
                    "Authenticate via CIMS and provision local user",
                    "Complete profile fields and documents",
                    "Admin approves; trainee gains course access",
                ],
            ),
            (
                "Course Creation",
                [
                    "Admin enters course structure and bilingual info",
                    "Attach media, modules, lessons and quizzes",
                    "Validate, save, assign and activate course",
                ],
            ),
            (
                "Trainee Learning",
                [
                    "Enroll / continue an active non-expired course",
                    "Complete lessons; system tracks progress & time",
                    "Pass quizzes; certificate issued on eligibility",
                ],
            ),
        ],
    )

    workflow_cards_slide(
        prs,
        "Application Workflows — Assessment & Operations",
        [
            (
                "Course Expiry",
                [
                    "Due date = enrollment date + duration (days)",
                    "Overdue incomplete courses marked expired",
                    "Removed from active lists; content access blocked",
                ],
            ),
            (
                "Quiz Attempt",
                [
                    "Start attempt with selected questions",
                    "Answer / skip / flag; finalize on submit or timeout",
                    "Auto-grade; show results; allow retry where enabled",
                ],
            ),
            (
                "Certificate",
                [
                    "Check lessons complete and quizzes passed",
                    "Issue certificate with hash and mark course done",
                    "Trainee views certificate; public verify by hash",
                ],
            ),
            (
                "Reports & Notifications",
                [
                    "Admin filters reports and exports CSV",
                    "Alerts and deadline notifications generated",
                    "Users view, act, and mark notifications read",
                ],
            ),
        ],
    )

    # Support workflow as compact extra card slide with 3 items + tech already separate
    s = blank(prs)
    fill_bg(s, prs)
    corner_accents(s, prs)
    brand_header(s)
    dots(s, 6.0, 0.28, cols=6, rows=2)
    slide_title_block(s, "Support & Communication Workflow")
    steps = [
        ("01", "Raise", "User creates a support ticket or starts chat / forum interaction."),
        ("02", "Track", "System stores messages, participants, attachments, and ticket status."),
        ("03", "Resolve", "Authorized users respond, update status, or moderate where permitted."),
    ]
    for i, (num, title, desc) in enumerate(steps):
        x = 0.7 + i * 4.15
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.2), Inches(3.9), Inches(3.5))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = GREEN_BORDER
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 1.35), Inches(2.55), Inches(1.1), Inches(1.1))
        circ.fill.solid()
        circ.fill.fore_color.rgb = GREEN
        circ.line.fill.background()
        add_text(circ, [(num, 18, True, WHITE)], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        nt = s.shapes.add_textbox(Inches(x + 0.2), Inches(3.9), Inches(3.5), Inches(0.4))
        add_text(nt, [(title, 18, True, GREEN)], align=PP_ALIGN.CENTER)
        dt = s.shapes.add_textbox(Inches(x + 0.25), Inches(4.45), Inches(3.4), Inches(1.0))
        add_text(dt, [(desc, 13, False, GRAY)], align=PP_ALIGN.CENTER)

    slide_tech_stack(prs)
    slide_closing(prs)

    for i, slide in enumerate(prs.slides):
        set_transition(slide, ["fade", "push", "wipe"][i % 3])

    out = "/workspace/PLRA_LMS_Presentation.pptx"
    prs.save(out)
    print(f"Saved {out} with {len(prs.slides)} slides")
    return out


if __name__ == "__main__":
    build()
