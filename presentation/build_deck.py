#!/usr/bin/env python3
"""
Generates the DineEase partner (B2B) presentation as a .pptx deck.

Audience : Restaurants / venue operators in Cyprus
Purpose  : Explain what DineEase offers, the subscription model, and how it
           differs from competitors operating in the Cyprus market.

Pricing is taken from dineeaseapp.com (annual billing = "save 20%", VAT incl.).
Competitor figures are from public market research (mid-2026, indicative).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ─── Brand palette (from index.html) ──────────────────────────────────────
NAVY      = RGBColor(0x0F, 0x2D, 0x40)
TEAL      = RGBColor(0x14, 0x43, 0x5E)
TEAL_MID  = RGBColor(0x1A, 0x55, 0x70)
ACCENT    = RGBColor(0x8B, 0x1A, 0x1A)
ACCENT2   = RGBColor(0xC0, 0x39, 0x2B)
GOLD      = RGBColor(0xC8, 0xA9, 0x6E)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE  = RGBColor(0xF4, 0xF1, 0xEC)
MUTED     = RGBColor(0xB0, 0xC4, 0xCC)
CARD      = RGBColor(0x16, 0x3A, 0x52)
GREEN     = RGBColor(0x2E, 0x7D, 0x5B)

SERIF = "Georgia"          # stands in for Playfair Display
SANS  = "Calibri"

HERE = os.path.dirname(os.path.abspath(__file__))
IMG  = os.path.join(HERE, "..", "images")

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ─── helpers ───────────────────────────────────────────────────────────────
def slide(bg=NAVY):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element)
    s.shapes._spTree.insert(2, r._element)
    return s


def box(s, x, y, w, h):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    return tb, tf


def para(tf, text, size, color, *, font=SANS, bold=False, align=PP_ALIGN.LEFT,
         space_after=6, space_before=0, first=False, line=None, italic=False):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    if line is not None:
        p.line_spacing = line
    r = p.add_run(); r.text = text
    f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    f.name = font; f.color.rgb = color
    return p


def rect(s, x, y, w, h, fill, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp


def kicker(s, text, x=Inches(0.9), y=Inches(0.65), color=GOLD):
    _, tf = box(s, x, y, Inches(8), Inches(0.4))
    para(tf, text.upper(), 13, color, bold=True, first=True, space_after=0)


def page_no(s, n):
    _, tf = box(s, Inches(12.3), Inches(7.0), Inches(0.8), Inches(0.4))
    para(tf, str(n), 10, MUTED, align=PP_ALIGN.RIGHT, first=True, space_after=0)
    _, tf2 = box(s, Inches(0.9), Inches(7.0), Inches(5), Inches(0.4))
    para(tf2, "DineEase  ·  Partner Overview", 10, MUTED, first=True, space_after=0)


def gold_rule(s, x, y, w=Inches(0.9)):
    rect(s, x, y, w, Pt(3), GOLD)


# ════════════════════════════════════════════════════════════════════════════
# 1 — TITLE
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
rect(s, 0, 0, SW, Inches(0.18), GOLD)
rect(s, 0, Inches(5.0), SW, Inches(0.05), TEAL_MID)

_, tf = box(s, Inches(0.9), Inches(0.7), Inches(8), Inches(0.5))
para(tf, "✦  DINEEASE  ·  FOR RESTAURANTS", 15, GOLD, bold=True, first=True)

_, tf = box(s, Inches(0.9), Inches(2.1), Inches(11.5), Inches(2.8))
para(tf, "Fill more tables.", 54, WHITE, font=SERIF, bold=True, first=True, space_after=2)
para(tf, "Pay one flat fee.", 54, GOLD, font=SERIF, bold=True, space_after=14)
para(tf, "The reservation platform built for Cyprus restaurants — plus a "
         "customer app that sends diners straight to your door.",
     20, MUTED, space_after=0, line=1.25)

_, tf = box(s, Inches(0.9), Inches(6.4), Inches(11), Inches(0.6))
para(tf, "Partner Presentation  ·  2026", 14, MUTED, first=True, space_after=0)

# ════════════════════════════════════════════════════════════════════════════
# 2 — THE PROBLEM
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
kicker(s, "The problem")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(1.2))
para(tf, "Taking bookings in Cyprus is harder than it should be", 34,
     WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

problems = [
    ("📞", "Bookings scattered everywhere",
     "Phone calls, Instagram DMs, WhatsApp and walk-ins — no single view of "
     "your night, and double-bookings happen."),
    ("🪑", "Empty tables, costly no-shows",
     "Without reminders or deposits, last-minute no-shows leave covers — and "
     "revenue — on the table."),
    ("💸", "Commission platforms eat margin",
     "Global booking sites charge a fee for every cover they send. The busier "
     "you get, the more you pay."),
    ("📵", "No direct line to diners",
     "Your guest data sits with someone else, so you can't bring customers "
     "back on your own terms."),
]
cx = Inches(0.9); cw = Inches(5.75); gap = Inches(0.35)
cy = Inches(2.45); ch = Inches(2.0); vgap = Inches(0.25)
for i, (ic, h, d) in enumerate(problems):
    col = i % 2; row = i // 2
    x = cx + col * (cw + gap)
    y = cy + row * (ch + vgap)
    rect(s, x, y, cw, ch, CARD, line=TEAL_MID, line_w=1)
    _, tf = box(s, x + Inches(0.3), y + Inches(0.25), cw - Inches(0.6), ch - Inches(0.4))
    para(tf, ic, 26, GOLD, first=True, space_after=4)
    para(tf, h, 18, WHITE, font=SERIF, bold=True, space_after=6)
    para(tf, d, 13.5, MUTED, line=1.2)
page_no(s, 2)

# ════════════════════════════════════════════════════════════════════════════
# 3 — WHAT DINEEASE IS  (two-sided)
# ════════════════════════════════════════════════════════════════════════════
s = slide(TEAL)
kicker(s, "What is DineEase")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(1.2))
para(tf, "One platform. Two sides. Both working for you.", 34,
     WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

# left card — restaurant
rect(s, Inches(0.9), Inches(2.5), Inches(5.6), Inches(4.1), NAVY, line=GOLD, line_w=1.5)
_, tf = box(s, Inches(1.25), Inches(2.85), Inches(4.95), Inches(3.5))
para(tf, "🍽  FOR YOUR RESTAURANT", 14, GOLD, bold=True, first=True, space_after=8)
para(tf, "A booking & floor manager", 22, WHITE, font=SERIF, bold=True, space_after=10)
for t in ["Reservations & waitlist in one dashboard",
          "Visual floorplan calendar of your night",
          "Automatic reminders to cut no-shows",
          "Manage one venue — or your whole group"]:
    para(tf, "✓  " + t, 15, OFFWHITE, space_after=8, line=1.15)

# right card — diner
rect(s, Inches(6.85), Inches(2.5), Inches(5.6), Inches(4.1), NAVY, line=GOLD, line_w=1.5)
_, tf = box(s, Inches(7.2), Inches(2.85), Inches(4.95), Inches(3.5))
para(tf, "📱  FOR YOUR CUSTOMERS", 14, GOLD, bold=True, first=True, space_after=8)
para(tf, "A discovery & booking app", 22, WHITE, font=SERIF, bold=True, space_after=10)
for t in ["Diners find you by cuisine, price & distance",
          "See real-time availability for your tables",
          "Book in seconds — no phone call needed",
          "Reviews & favourites keep them coming back"]:
    para(tf, "✓  " + t, 15, OFFWHITE, space_after=8, line=1.15)

# center connector
con = rect(s, Inches(6.15), Inches(4.05), Inches(1.05), Inches(1.05), GOLD,
           shape=MSO_SHAPE.OVAL)
tf = con.text_frame; tf.word_wrap = True
para(tf, "↔", 30, NAVY, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
page_no(s, 3)

# ════════════════════════════════════════════════════════════════════════════
# 4 — WHAT YOU GET (features)
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
kicker(s, "What you get")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(1.2))
para(tf, "Everything you need to run a busy service", 34,
     WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

feats = [
    ("📅", "Reservations & Waitlist",
     "Take, edit and track every booking in one place, with a smart waitlist for peak hours."),
    ("🗺", "Floorplan Calendar",
     "See your whole night at a glance and assign tables to match your real dining room."),
    ("🔔", "No-show Protection",
     "Automatic SMS / push reminders — and optional deposits — keep tables full."),
    ("📲", "Customer App Listing",
     "Your venue featured in the DineEase app, in front of diners searching nearby."),
    ("📊", "Insights & Analytics",
     "Understand covers, peak times and repeat guests to plan staffing and stock."),
    ("🏢", "Multi-location Ready",
     "Run two, three or five venues from a single account as you grow."),
]
cx = Inches(0.9); cw = Inches(3.75); gapx = Inches(0.32)
cy = Inches(2.4); ch = Inches(2.0); gapy = Inches(0.25)
for i, (ic, h, d) in enumerate(feats):
    col = i % 3; row = i // 3
    x = cx + col * (cw + gapx)
    y = cy + row * (ch + gapy)
    rect(s, x, y, cw, ch, CARD, line=TEAL_MID, line_w=1)
    _, tf = box(s, x + Inches(0.28), y + Inches(0.24), cw - Inches(0.56), ch - Inches(0.4))
    para(tf, ic, 24, GOLD, first=True, space_after=4)
    para(tf, h, 16.5, WHITE, font=SERIF, bold=True, space_after=5)
    para(tf, d, 12.5, MUTED, line=1.2)
page_no(s, 4)

# ════════════════════════════════════════════════════════════════════════════
# 5 — THE CUSTOMER APP (with screenshots)
# ════════════════════════════════════════════════════════════════════════════
s = slide(TEAL)
kicker(s, "The customer app")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(7.6), Inches(1.2))
para(tf, "We bring the diners to you", 34, WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

_, tf = box(s, Inches(0.9), Inches(2.35), Inches(5.4), Inches(4))
para(tf, "Unlike booking software you have to fill yourself, DineEase is also a "
         "consumer app people open when they're deciding where to eat.",
     16, OFFWHITE, first=True, line=1.3, space_after=14)
for t in ["Find. Book. Enjoy. — three taps to a table",
          "Search by cuisine, price and 5–20 km radius",
          "Real-time availability for your venue",
          "Secure in-app payments & confirmations",
          "Reviews and favourites build loyalty"]:
    para(tf, "✦  " + t, 15, WHITE, space_after=9, line=1.15)

# screenshots
shots = ["1.jpg", "2.jpg", "3.jpg", "4.jpg"]
labels = ["Welcome", "Find", "Book", "Enjoy"]
ph_w = Inches(1.45); ph_h = Inches(2.97)
start_x = Inches(6.7); base_y = Inches(2.7); step = Inches(1.55)
for i, (f, lab) in enumerate(zip(shots, labels)):
    p = os.path.join(IMG, f)
    if os.path.exists(p):
        x = start_x + i * step
        s.shapes.add_picture(p, x, base_y, width=ph_w, height=ph_h)
        _, tf = box(s, x - Inches(0.1), base_y + ph_h + Inches(0.1), ph_w + Inches(0.2), Inches(0.35))
        para(tf, lab.upper(), 11, GOLD, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
page_no(s, 5)

# ════════════════════════════════════════════════════════════════════════════
# 6 — HOW IT WORKS
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
kicker(s, "How it works")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(1.2))
para(tf, "Live in days, not months", 34, WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

steps = [
    ("1", "List your venue",
     "We set up your profile, floorplan and tables. Your listing goes live in the DineEase app."),
    ("2", "Manage bookings",
     "Reservations from the app, your website and phone all land in one simple dashboard."),
    ("3", "Fill your tables",
     "Diners discover and book you; reminders cut no-shows. You just run great service."),
]
cw = Inches(3.7); gapx = Inches(0.45)
total = cw * 3 + gapx * 2
cx = (SW - total) / 2
cy = Inches(2.7); ch = Inches(3.1)
for i, (n, h, d) in enumerate(steps):
    x = cx + i * (cw + gapx)
    rect(s, x, cy, cw, ch, CARD, line=TEAL_MID, line_w=1)
    circ = rect(s, x + cw/2 - Inches(0.55), cy - Inches(0.55), Inches(1.1), Inches(1.1),
                ACCENT2, shape=MSO_SHAPE.OVAL)
    tf = circ.text_frame
    para(tf, n, 30, WHITE, font=SERIF, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
    _, tf = box(s, x + Inches(0.35), cy + Inches(0.85), cw - Inches(0.7), ch - Inches(1.1))
    para(tf, h, 21, GOLD, font=SERIF, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=10)
    para(tf, d, 14.5, MUTED, align=PP_ALIGN.CENTER, line=1.3)
page_no(s, 6)

# ════════════════════════════════════════════════════════════════════════════
# 7 — SUBSCRIPTION MODEL (pricing)
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
kicker(s, "Subscription model")
_, tf = box(s, Inches(0.9), Inches(1.1), Inches(11.5), Inches(0.9))
para(tf, "Simple, flat pricing — no commission, ever", 32,
     WHITE, font=SERIF, bold=True, first=True)
_, tf = box(s, Inches(0.95), Inches(1.85), Inches(11.5), Inches(0.5))
para(tf, "Annual plans shown (save 20%) · VAT included · cancel anytime",
     14, GOLD, first=True, space_after=0)

tiers = [
    ("Essentials", "€60", "€719.90 / yr", "1 location", "or €74.99/mo monthly", False),
    ("Duo", "€104", "€1,247.90 / yr", "Up to 2 locations", "or €129.99/mo monthly", True),
    ("Group", "€144", "€1,727.90 / yr", "Up to 3 locations", "or €179.99/mo monthly", False),
    ("Portfolio", "€200", "€2,399.90 / yr", "4–5 locations", "or €249.99/mo monthly", False),
]
cw = Inches(2.85); gapx = Inches(0.3)
total = cw * 4 + gapx * 3
cx = (SW - total) / 2
cy = Inches(2.55); ch = Inches(4.05)
for i, (name, price, yr, loc, mo, pop) in enumerate(tiers):
    x = cx + i * (cw + gapx)
    fill = TEAL if pop else CARD
    ln = GOLD if pop else TEAL_MID
    rect(s, x, cy, cw, ch, fill, line=ln, line_w=2 if pop else 1)
    if pop:
        badge = rect(s, x + cw/2 - Inches(0.95), cy - Inches(0.22), Inches(1.9), Inches(0.44),
                     ACCENT2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = badge.text_frame
        para(tf, "MOST POPULAR", 10.5, WHITE, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
    _, tf = box(s, x + Inches(0.25), cy + Inches(0.45), cw - Inches(0.5), ch - Inches(0.7))
    para(tf, name, 20, WHITE, font=SERIF, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=8)
    p = tf.add_paragraph(); p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(0)
    r = p.add_run(); r.text = price; r.font.size = Pt(38); r.font.bold = True
    r.font.name = SERIF; r.font.color.rgb = GOLD
    r2 = p.add_run(); r2.text = " /mo"; r2.font.size = Pt(14); r2.font.color.rgb = MUTED; r2.font.name = SANS
    para(tf, yr, 12, MUTED, align=PP_ALIGN.CENTER, space_after=10)
    para(tf, loc, 14.5, WHITE, bold=True, align=PP_ALIGN.CENTER, space_after=4)
    para(tf, mo, 11.5, MUTED, align=PP_ALIGN.CENTER, space_after=8, italic=True)
    para(tf, "All DineEase features", 12, OFFWHITE, align=PP_ALIGN.CENTER, space_after=2)
    para(tf, "Reservations & waitlist", 12, OFFWHITE, align=PP_ALIGN.CENTER, space_after=2)
    para(tf, "Floorplan calendar", 12, OFFWHITE, align=PP_ALIGN.CENTER, space_after=2)
    para(tf, "Customer app listing", 12, OFFWHITE, align=PP_ALIGN.CENTER, space_after=2)
page_no(s, 7)

# ════════════════════════════════════════════════════════════════════════════
# 8 — HOW WE'RE DIFFERENT (comparison table)
# ════════════════════════════════════════════════════════════════════════════
s = slide(TEAL)
kicker(s, "How we compare")
_, tf = box(s, Inches(0.9), Inches(1.05), Inches(11.5), Inches(0.9))
para(tf, "Built for Cyprus restaurants — not the other way round", 30,
     WHITE, font=SERIF, bold=True, first=True)

rows = [
    ["", "DineEase", "Eat App", "OpenTable", "TheFork / Quandoo", "Klisto (CY)"],
    ["Pricing model", "Flat subscription", "Flat subscription", "Subscription + fees", "Commission per cover", "Marketplace"],
    ["Per-cover commission", "None", "None", "$1–1.50 / cover*", "~€2.60 / guest*", "Varies"],
    ["Restaurant-specialised", "Yes", "Yes", "Yes", "Yes", "Multi-vertical"],
    ["Own customer app", "Yes", "Limited", "Yes", "Yes", "Yes"],
    ["Cyprus-focused & EUR", "Yes — VAT incl.", "No (USD)", "No (USD)", "Limited in CY", "Yes"],
    ["Cost as you grow", "Stays flat", "Stays flat", "Rises per cover", "Rises per cover", "Varies"],
]
nrows = len(rows); ncols = 6
tx = Inches(0.7); ty = Inches(2.05)
tw = Inches(11.93); th = Inches(4.5)
tbl_shape = s.shapes.add_table(nrows, ncols, tx, ty, tw, th)
table = tbl_shape.table
table.columns[0].width = Inches(2.63)
for c in range(1, 6):
    table.columns[c].width = Inches(1.86)
table.first_row = False
table.horz_banding = False

for ri, row in enumerate(rows):
    table.rows[ri].height = Inches(0.64)
    for ci, val in enumerate(row):
        cell = table.cell(ri, ci)
        cell.margin_left = Inches(0.1); cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        # cell fill
        if ri == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY if ci != 1 else ACCENT
        elif ci == 1:
            cell.fill.solid(); cell.fill.fore_color.rgb = GOLD
        elif ci == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb = CARD
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
        r = p.add_run(); r.text = val
        f = r.font
        if ri == 0:
            f.size = Pt(12); f.bold = True; f.name = SERIF
            f.color.rgb = NAVY if ci == 1 else WHITE
        elif ci == 0:
            f.size = Pt(11); f.bold = True; f.name = SANS; f.color.rgb = WHITE
        elif ci == 1:
            f.size = Pt(11); f.bold = True; f.name = SANS; f.color.rgb = NAVY
        else:
            f.size = Pt(10.5); f.name = SANS; f.color.rgb = MUTED

_, tf = box(s, Inches(0.7), Inches(6.75), Inches(11.9), Inches(0.4))
para(tf, "*Competitor pricing from public sources, mid-2026, indicative. OpenTable / "
         "TheFork charge per seated guest, so cost rises with every booking.",
     10.5, MUTED, first=True, space_after=0, italic=True)
page_no(s, 8)

# ════════════════════════════════════════════════════════════════════════════
# 9 — DINEEASE vs EAT APP (closest competitor, side by side)
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
kicker(s, "Our closest competitor")
_, tf = box(s, Inches(0.9), Inches(1.05), Inches(11.5), Inches(0.9))
para(tf, "DineEase vs Eat App — same idea, built for Cyprus", 30,
     WHITE, font=SERIF, bold=True, first=True)
_, tf = box(s, Inches(0.95), Inches(1.8), Inches(11.5), Inches(0.5))
para(tf, "Both are flat-fee and commission-free. The difference is how you're "
         "priced — and who fills your tables.", 14, GOLD, first=True, space_after=0)

# DineEase card
rect(s, Inches(0.7), Inches(2.45), Inches(6.0), Inches(3.45), TEAL, line=GOLD, line_w=2)
_, tf = box(s, Inches(1.05), Inches(2.7), Inches(5.4), Inches(3.1))
para(tf, "DINEEASE", 16, GOLD, bold=True, first=True, space_after=2)
para(tf, "Priced per location · covers never capped · EUR, VAT incl.", 12.5,
     OFFWHITE, italic=True, space_after=10)
for name, price, detail in [
    ("Essentials", "€60/mo", "1 location"),
    ("Duo", "€104/mo", "up to 2 locations"),
    ("Group", "€144/mo", "up to 3 locations"),
    ("Portfolio", "€200/mo", "4–5 locations"),
]:
    p = tf.add_paragraph(); p.space_after = Pt(7); p.line_spacing = 1.0
    r = p.add_run(); r.text = f"{name}  "; r.font.size = Pt(15); r.font.bold = True
    r.font.name = SANS; r.font.color.rgb = WHITE
    r = p.add_run(); r.text = f"{price}"; r.font.size = Pt(15); r.font.bold = True
    r.font.name = SERIF; r.font.color.rgb = GOLD
    r = p.add_run(); r.text = f"   ·  {detail}"; r.font.size = Pt(12.5)
    r.font.name = SANS; r.font.color.rgb = MUTED
para(tf, "+ your venue in the DineEase consumer app for Cyprus diners", 12.5,
     GOLD, space_before=4, space_after=0)

# Eat App card
rect(s, Inches(6.95), Inches(2.45), Inches(5.7), Inches(3.45), CARD, line=TEAL_MID, line_w=1)
_, tf = box(s, Inches(7.3), Inches(2.7), Inches(5.1), Inches(3.1))
para(tf, "EAT APP", 16, MUTED, bold=True, first=True, space_after=2)
para(tf, "Priced by covers & add-ons · USD · no local diner audience", 12.5,
     MUTED, italic=True, space_after=10)
for name, price, detail in [
    ("Free", "$0/mo", "up to 100 covers/mo"),
    ("Starter", "$48/mo", "up to 300 covers/mo"),
    ("Essential", "$111/mo", "unlimited covers"),
    ("Pro", "$209/mo", "unlimited covers"),
]:
    p = tf.add_paragraph(); p.space_after = Pt(7); p.line_spacing = 1.0
    r = p.add_run(); r.text = f"{name}  "; r.font.size = Pt(15); r.font.bold = True
    r.font.name = SANS; r.font.color.rgb = WHITE
    r = p.add_run(); r.text = f"{price}"; r.font.size = Pt(15); r.font.bold = True
    r.font.name = SERIF; r.font.color.rgb = OFFWHITE
    r = p.add_run(); r.text = f"   ·  {detail}"; r.font.size = Pt(12.5)
    r.font.name = SANS; r.font.color.rgb = MUTED
para(tf, "Reservation software only — you fill the tables yourself", 12.5,
     MUTED, space_before=4, space_after=0)

_, tf = box(s, Inches(0.7), Inches(6.05), Inches(11.95), Inches(0.95))
para(tf, "Takeaway:  Eat App caps covers on cheaper plans, prices in USD, and bills "
         "$19–$39/mo add-ons for waitlist, CRM & reports — even on its recommended "
         "plan. DineEase includes every feature in each tier, priced per location in "
         "EUR (VAT incl.), and sends Cyprus diners your way.",
     13, OFFWHITE, first=True, align=PP_ALIGN.CENTER, line=1.2, space_after=0)
_, tf = box(s, Inches(0.7), Inches(7.02), Inches(11.95), Inches(0.35))
para(tf, "Eat App prices: yearly billing (30% off), per restaurant.eatapp.co, 2026.",
     9.5, MUTED, first=True, align=PP_ALIGN.CENTER, italic=True, space_after=0)
page_no(s, 9)

# ════════════════════════════════════════════════════════════════════════════
# 10 — THE COST ADVANTAGE (math)
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
kicker(s, "The cost advantage")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(1.2))
para(tf, "Flat beats commission as you get busy", 34,
     WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

_, tf = box(s, Inches(0.9), Inches(2.25), Inches(11.5), Inches(0.6))
para(tf, "Example: a restaurant seating 1,000 booked covers in a month",
     16, GOLD, first=True, space_after=0)

cards = [
    ("Commission platform", "≈ €2,600", "/month", "at ~€2.60 per guest, every month — and it climbs as you grow", ACCENT2),
    ("OpenTable-style", "€140 + fees", "/month", "subscription plus $1–1.50 for every network cover", TEAL_MID),
    ("DineEase Duo", "€104", "/month", "flat — VAT included, no per-cover fee, no matter how busy", GREEN),
]
cw = Inches(3.7); gapx = Inches(0.45)
total = cw * 3 + gapx * 2
cx = (SW - total) / 2
cy = Inches(3.1); ch = Inches(2.9)
for i, (h, big, unit, d, accent) in enumerate(cards):
    x = cx + i * (cw + gapx)
    rect(s, x, cy, cw, ch, CARD, line=accent, line_w=2)
    rect(s, x, cy, cw, Inches(0.12), accent)
    _, tf = box(s, x + Inches(0.3), cy + Inches(0.4), cw - Inches(0.6), ch - Inches(0.6))
    para(tf, h, 16, WHITE, font=SERIF, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=10)
    para(tf, big, 36, GOLD if i == 2 else WHITE, font=SERIF, bold=True,
         align=PP_ALIGN.CENTER, space_after=0)
    para(tf, unit, 13, MUTED, align=PP_ALIGN.CENTER, space_after=10)
    para(tf, d, 13, MUTED, align=PP_ALIGN.CENTER, line=1.25)

_, tf = box(s, Inches(0.9), Inches(6.35), Inches(11.5), Inches(0.6))
para(tf, "With DineEase your cost is predictable. The more diners we send you, the "
         "more you save versus commission models.", 14, OFFWHITE, first=True,
     align=PP_ALIGN.CENTER, space_after=0)
page_no(s, 10)

# ════════════════════════════════════════════════════════════════════════════
# 11 — WHY DINEEASE FOR CYPRUS
# ════════════════════════════════════════════════════════════════════════════
s = slide(TEAL)
kicker(s, "Why DineEase")
_, tf = box(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(1.2))
para(tf, "The local advantage", 34, WHITE, font=SERIF, bold=True, first=True)
gold_rule(s, Inches(0.95), Inches(2.05))

points = [
    ("🇨🇾", "Made for Cyprus", "Local team, EUR pricing with VAT included, and support that knows your market."),
    ("🚫", "Zero commission", "A flat monthly fee — we never take a cut of your covers or your revenue."),
    ("🤝", "You own your guests", "Your customer data and relationships stay yours, to bring diners back."),
    ("📈", "Scales with you", "From a single bistro to a five-venue group, on one account and one bill."),
]
cx = Inches(0.9); cw = Inches(5.75); gap = Inches(0.35)
cy = Inches(2.45); ch = Inches(1.95); vgap = Inches(0.25)
for i, (ic, h, d) in enumerate(points):
    col = i % 2; row = i // 2
    x = cx + col * (cw + gap)
    y = cy + row * (ch + vgap)
    rect(s, x, y, cw, ch, NAVY, line=GOLD, line_w=1)
    _, tf = box(s, x + Inches(0.35), y + Inches(0.3), cw - Inches(0.7), ch - Inches(0.5))
    para(tf, ic + "  " + h, 19, GOLD, font=SERIF, bold=True, first=True, space_after=8)
    para(tf, d, 14.5, OFFWHITE, line=1.25)
page_no(s, 11)

# ════════════════════════════════════════════════════════════════════════════
# 11 — CALL TO ACTION
# ════════════════════════════════════════════════════════════════════════════
s = slide(NAVY)
rect(s, 0, 0, SW, Inches(0.18), GOLD)
# glow panel
rect(s, Inches(1.4), Inches(1.6), Inches(10.5), Inches(4.4), TEAL, line=GOLD, line_w=1.5,
     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
_, tf = box(s, Inches(2.0), Inches(2.1), Inches(9.3), Inches(1.4))
para(tf, "Ready to fill more tables?", 40, WHITE, font=SERIF, bold=True,
     align=PP_ALIGN.CENTER, first=True, space_after=10)
_, tf = box(s, Inches(2.4), Inches(3.35), Inches(8.5), Inches(1.0))
para(tf, "Start with Essentials from €60/mo, or pick the plan that fits your group. "
         "No commission, no surprises — VAT included.",
     17, OFFWHITE, align=PP_ALIGN.CENTER, first=True, line=1.3, space_after=0)

# CTA button
btn = rect(s, SW/2 - Inches(1.6), Inches(4.55), Inches(3.2), Inches(0.75), GOLD,
           shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tf = btn.text_frame
para(tf, "Get started today", 18, NAVY, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)

_, tf = box(s, Inches(2.0), Inches(5.5), Inches(9.3), Inches(0.5))
para(tf, "dineeaseapp.com   ·   Find. Book. Enjoy.", 15, GOLD,
     align=PP_ALIGN.CENTER, first=True, space_after=0)

out = os.path.join(HERE, "DineEase-Partner-Presentation.pptx")
prs.save(out)
print("Saved:", out, "—", len(prs.slides._sldIdLst), "slides")
