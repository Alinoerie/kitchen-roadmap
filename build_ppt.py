#!/usr/bin/env python3
"""
Genereert de PPT proposal voor Keukengerij launch.
Output: /home/alisionary/.hermes/profiles/hermes_ecom_research/projects/kitchen-roadmap/kitchen-launch-proposal.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# =============================================================
# KLEUREN
# =============================================================
COL_PRIMARY = RGBColor(0x2d, 0x5a, 0x3d)    # dark green
COL_ACCENT = RGBColor(0xc8, 0x98, 0x60)     # warm amber
COL_DARK = RGBColor(0x1f, 0x3d, 0x2d)        # darker green
COL_TEXT = RGBColor(0x1a, 0x1a, 0x1a)        # near-black
COL_MUTED = RGBColor(0x6a, 0x6a, 0x6a)      # gray
COL_BG = RGBColor(0xfa, 0xfa, 0xf7)          # cream
COL_WARN = RGBColor(0x8a, 0x45, 0x20)        # warning orange

# Owner colors
OWNER_COLORS = {
    'ali': RGBColor(0xc8, 0x98, 0x60),
    'kheibar': RGBColor(0x6a, 0x8a, 0xaa),
    'farshad': RGBColor(0x8a, 0x6a, 0xaa),
    'designer': RGBColor(0xaa, 0x6a, 0x8a),
    'hermes': RGBColor(0x4a, 0x8a, 0x6a),
}

# =============================================================
# PREST SETUP
# =============================================================
prs = Presentation()
prs.slide_width = Inches(13.333)  # 16:9 widescreen
prs.slide_height = Inches(7.5)

SW = prs.slide_width
SH = prs.slide_height

# =============================================================
# HELPERS
# =============================================================

def add_blank_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid(); bg.fill.fore_color.rgb = COL_BG
    bg.line.fill.background()
    return slide

def add_text(slide, x, y, w, h, text, size=18, bold=False, color=COL_TEXT, align=PP_ALIGN.LEFT, font='Calibri'):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font
    return tb

def add_rect(slide, x, y, w, h, color=COL_PRIMARY, line_color=None):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    rect.fill.solid(); rect.fill.fore_color.rgb = color
    if line_color:
        rect.line.color.rgb = line_color
    else:
        rect.line.fill.background()
    return rect

def add_bullets(slide, x, y, w, h, bullets, size=14, color=COL_TEXT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = "• " + b
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = 'Calibri'
    return tb

def add_phase_strip(slide, current_idx):
    """Bottom strip met alle fases als pills, current is amber."""
    phases = ['F0', 'F1', 'F2', 'F3', 'F4', 'F5']
    titles = ['Foundations', 'Brand+Products', 'Listings+Wave 1', 'Wave 2+3+Validation', 'Scale Winners', 'Lock+Optimize']
    strip_y = SH - Inches(0.45)
    pill_w = Inches(2.0)
    total_w = pill_w * len(phases) + Inches(0.05) * (len(phases) - 1)
    start_x = (SW - total_w) / 2

    for i, (code, title) in enumerate(zip(phases, titles)):
        x = start_x + i * (pill_w + Inches(0.05))
        if i == current_idx:
            bg = COL_ACCENT
            txt = COL_DARK
        elif i < current_idx:
            bg = COL_PRIMARY
            txt = COL_BG
        else:
            bg = COL_BG
            txt = COL_MUTED
        add_rect(slide, x, strip_y, pill_w, Inches(0.35), color=bg, line_color=COL_MUTED)
        add_text(slide, x, strip_y + Inches(0.04), pill_w, Inches(0.28),
                 f"{code}: {title}", size=9, bold=(i == current_idx), color=txt, align=PP_ALIGN.CENTER)

def add_header(slide, phase_code, phase_title, days, goal):
    """Standaard header voor milestone slides."""
    # Phase badge
    badge = add_rect(slide, Inches(0.4), Inches(0.3), Inches(0.7), Inches(0.5), color=COL_PRIMARY)
    add_text(slide, Inches(0.4), Inches(0.32), Inches(0.7), Inches(0.45),
             phase_code, size=18, bold=True, color=COL_BG, align=PP_ALIGN.CENTER)
    # Phase title
    add_text(slide, Inches(1.3), Inches(0.3), Inches(8), Inches(0.45),
             phase_title, size=22, bold=True, color=COL_DARK)
    add_text(slide, Inches(1.3), Inches(0.75), Inches(8), Inches(0.3),
             f"{days} · doel: {goal}", size=11, color=COL_MUTED, font='Calibri')

def add_owner_legend(slide, x, y):
    owners = [('Ali', 'ali'), ('Kheibar', 'kheibar'), ('Farshad', 'farshad'),
              ('Designer (F)', 'designer'), ('Hermes', 'hermes')]
    for i, (label, key) in enumerate(owners):
        ox = x + i * Inches(1.2)
        # Dot
        dot = add_rect(slide, ox, y + Inches(0.08), Inches(0.15), Inches(0.15),
                        color=OWNER_COLORS[key])
        add_text(slide, ox + Inches(0.2), y, Inches(1.0), Inches(0.3),
                 label, size=9, color=COL_MUTED)

# =============================================================
# SLIDE 1 — COVER
# =============================================================
slide = add_blank_slide()
# Big primary band
add_rect(slide, 0, 0, SW, SH, color=COL_DARK)
# Accent stripe
add_rect(slide, 0, SH - Inches(0.3), SW, Inches(0.3), color=COL_ACCENT)

# Title
add_text(slide, Inches(0.6), Inches(2.0), Inches(12), Inches(1.5),
         "Keukengerij Launch", size=64, bold=True, color=COL_BG, font='Cambria')
add_text(slide, Inches(0.6), Inches(3.3), Inches(12), Inches(0.6),
         "0 → €10k+/maand · 90 dagen · EU10 · Dropship",
         size=22, color=COL_ACCENT, font='Cambria')

# Sub
add_text(slide, Inches(0.6), Inches(4.5), Inches(12), Inches(0.4),
         "Ali · Kheibar · Farshad", size=18, color=COL_BG, font='Calibri')
add_text(slide, Inches(0.6), Inches(4.9), Inches(12), Inches(0.4),
         "v1.0 · 2026-09-09", size=12, color=COL_ACCENT, font='Calibri')

# Bottom band
add_text(slide, Inches(0.6), Inches(6.8), Inches(12), Inches(0.4),
         "Strategic Roadmap + Execution Plan", size=16, color=COL_BG, font='Cambria')

# =============================================================
# SLIDE 2 — EXECUTIVE SUMMARY
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "Executive Summary", size=32, bold=True, color=COL_DARK, font='Cambria')

# 3 cards
cards = [
    ("Wat", "Keukengerij e-commerce launch in EU10. 10 producten, gefaseerde wave-based introductie. Shopify als enige kanaal. Dropship-only fulfilment."),
    ("Waarom", "HK-only seller structuur (bewust gekozen) voor tax-efficiency. Eigen merk opbouwen voor lange-termijn exit-potentieel. AI-first aanpak via Hermes."),
    ("Hoe", "Team van 3 founders (33% elk). €5-10k budget + €1000 eerste weken ad spend. 30% ACoS target. 6 fasen, 90 dagen.")
]
for i, (title, body) in enumerate(cards):
    x = Inches(0.6 + i * 4.1)
    card = add_rect(slide, x, Inches(1.3), Inches(3.9), Inches(2.2), color=COL_BG, line_color=COL_PRIMARY)
    add_text(slide, x + Inches(0.2), Inches(1.45), Inches(3.5), Inches(0.4),
             title, size=20, bold=True, color=COL_PRIMARY, font='Cambria')
    add_text(slide, x + Inches(0.2), Inches(1.95), Inches(3.5), Inches(1.4),
             body, size=12, color=COL_TEXT)

# KPI summary
add_text(slide, Inches(0.6), Inches(4.0), Inches(12), Inches(0.4),
         "Targets (eerste 90 dagen)", size=20, bold=True, color=COL_DARK)
kpis = [
    ("Doel", "Maximaliseren (geen vast target-getal)"),
    ("ACoS", "30% (break-even ROAS 3.33)"),
    ("Retour", "14 dgn, wij betalen retourlabel"),
    ("Ad budget eerste weken", "€1000 (~€35/dag)"),
    ("Customer service", "24h response, 7 talen, AI-chatbot"),
    ("Brand positioning", "Mid-range met mix budget/premium"),
]
for i, (k, v) in enumerate(kpis):
    row = i // 2; col = i % 2
    y = Inches(4.5) + row * Inches(0.5)
    x = Inches(0.6) + col * Inches(6.2)
    add_rect(slide, x, y, Inches(1.8), Inches(0.4), color=COL_PRIMARY)
    add_text(slide, x, y, Inches(1.8), Inches(0.4),
             k, size=12, bold=True, color=COL_BG, align=PP_ALIGN.CENTER)
    add_rect(slide, x + Inches(1.8), y, Inches(4.4), Inches(0.4), color=COL_BG, line_color=COL_MUTED)
    add_text(slide, x + Inches(1.95), y + Inches(0.07), Inches(4.2), Inches(0.3),
             v, size=12, color=COL_TEXT)

add_phase_strip(slide, 0)

# =============================================================
# SLIDE 3 — TEAM + OWNERSHIP
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "Team · 33% per persoon", size=32, bold=True, color=COL_DARK, font='Cambria')

team = [
    ("Ali", COL_ACCENT, "Strategie · Marketing · Content · AI-agents",
     "33%", "Stigma, marketing, content generation, Hermes orchestration"),
    ("Kheibar", COL_PRIMARY, "Finance · Banking · Payouts",
     "33%", "HK bank, WISE, Currency Ways, payment infrastructure"),
    ("Farshad", COL_DARK, "E-com operator · Supplier netwerk",
     "33%", "Dropship supplier in China, quality check, designer (zijn team)"),
]
for i, (name, color, role, owner_pct, brings) in enumerate(team):
    x = Inches(0.6 + i * 4.1)
    # Color band top
    add_rect(slide, x, Inches(1.3), Inches(3.9), Inches(0.5), color=color)
    add_text(slide, x, Inches(1.32), Inches(3.9), Inches(0.45),
             name, size=24, bold=True, color=COL_BG, align=PP_ALIGN.CENTER)
    # Body card
    add_rect(slide, x, Inches(1.8), Inches(3.9), Inches(4.5), color=COL_BG, line_color=color)
    add_text(slide, x + Inches(0.2), Inches(2.0), Inches(3.5), Inches(0.6),
             role, size=14, bold=True, color=color, font='Cambria')
    add_text(slide, x + Inches(0.2), Inches(2.7), Inches(3.5), Inches(0.4),
             f"Ownership: {owner_pct}", size=12, color=COL_MUTED)
    add_text(slide, x + Inches(0.2), Inches(3.2), Inches(3.5), Inches(3.0),
             brings, size=12, color=COL_TEXT)

add_phase_strip(slide, 0)

# =============================================================
# SLIDE 4 — PRODUCT STRATEGY
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "Product Strategy · 10 keukengerij SKU's",
         size=32, bold=True, color=COL_DARK, font='Cambria')

# Wave diagram
add_text(slide, Inches(0.6), Inches(1.3), Inches(12), Inches(0.4),
         "Wave-based launch · 3-4 producten per wave · 14-dgn test cyclus",
         size=18, bold=True, color=COL_PRIMARY)

# 3 waves
waves = [
    ("Wave 1 · D20-D35", "3-4 SKUs", "Eerste validatie, kill losers", COL_ACCENT),
    ("Wave 2 · D40-D55", "3-4 nieuwe SKUs", "Anders segment dan W1", COL_PRIMARY),
    ("Wave 3 · D51-D65", "3-4 final SKUs", "Portfolio compleet", COL_DARK),
]
for i, (title, count, desc, color) in enumerate(waves):
    x = Inches(0.6 + i * 4.1)
    add_rect(slide, x, Inches(2.0), Inches(3.9), Inches(2.5), color=COL_BG, line_color=color)
    # Wave title bar
    add_rect(slide, x, Inches(2.0), Inches(3.9), Inches(0.5), color=color)
    add_text(slide, x, Inches(2.05), Inches(3.9), Inches(0.4),
             title, size=14, bold=True, color=COL_BG, align=PP_ALIGN.CENTER)
    add_text(slide, x + Inches(0.2), Inches(2.7), Inches(3.5), Inches(0.5),
             count, size=24, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_text(slide, x + Inches(0.2), Inches(3.4), Inches(3.5), Inches(1.0),
             desc, size=12, color=COL_TEXT, align=PP_ALIGN.CENTER)

# Pricing strategy
add_text(slide, Inches(0.6), Inches(4.8), Inches(12), Inches(0.4),
         "Pricing strategie · Mid-range met mix budget/premium",
         size=18, bold=True, color=COL_PRIMARY)
prices = [
    ("Budget", "€15-25", "Gadgets, tools, messensets <€30"),
    ("Mid", "€30-50", "Pannen, snijplanken, keukensets"),
    ("Premium", "€50-80", "High-end messen, speciale pannen"),
]
for i, (label, range_, examples) in enumerate(prices):
    x = Inches(0.6 + i * 4.1)
    add_rect(slide, x, Inches(5.4), Inches(3.9), Inches(1.5), color=COL_BG, line_color=COL_MUTED)
    add_text(slide, x + Inches(0.2), Inches(5.5), Inches(3.5), Inches(0.4),
             label, size=16, bold=True, color=COL_DARK)
    add_text(slide, x + Inches(0.2), Inches(5.95), Inches(3.5), Inches(0.4),
             range_, size=22, bold=True, color=COL_ACCENT)
    add_text(slide, x + Inches(0.2), Inches(6.45), Inches(3.5), Inches(0.4),
             examples, size=11, color=COL_MUTED)

add_phase_strip(slide, 0)

# =============================================================
# SLIDE 5 — CHANNELS & GEO
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "Channels · EU10 · Multi-language",
         size=32, bold=True, color=COL_DARK, font='Cambria')

# Big channel focus box
add_rect(slide, Inches(0.6), Inches(1.3), Inches(7), Inches(2.5), color=COL_PRIMARY)
add_text(slide, Inches(0.8), Inches(1.5), Inches(6.6), Inches(0.5),
         "Hoofdkanaal (90 dagen)", size=14, bold=True, color=COL_ACCENT)
add_text(slide, Inches(0.8), Inches(1.95), Inches(6.6), Inches(1.0),
         "Shopify webshop", size=36, bold=True, color=COL_BG, font='Cambria')
add_text(slide, Inches(0.8), Inches(2.9), Inches(6.6), Inches(0.8),
         "Multi-currency · Multi-language · Dropship-only",
         size=14, color=COL_BG)

# Excluded box
add_rect(slide, Inches(8), Inches(1.3), Inches(4.7), Inches(2.5), color=COL_BG, line_color=COL_MUTED)
add_text(slide, Inches(8.2), Inches(1.5), Inches(4.3), Inches(0.5),
         "UITGESLOTEN in 90 dagen", size=14, bold=True, color=COL_MUTED)
excluded = ["TikTok Shop", "Marktplaats", "Amazon EU", "bol.com", "Etsy"]
for i, ch in enumerate(excluded):
    y = Inches(2.05) + i * Inches(0.35)
    add_rect(slide, Inches(8.2), y, Inches(0.2), Inches(0.2), color=COL_WARN)
    add_text(slide, Inches(8.5), y - Inches(0.05), Inches(4.0), Inches(0.3),
             ch, size=14, color=COL_TEXT)

# EU10
add_text(slide, Inches(0.6), Inches(4.1), Inches(12), Inches(0.4),
         "EU10 markten (Amazon EU set)", size=18, bold=True, color=COL_PRIMARY)
eu10 = ['NL', 'BE', 'DE', 'FR', 'IT', 'ES', 'PL', 'SE', 'IE', 'CZ']
for i, country in enumerate(eu10):
    x = Inches(0.6) + (i % 5) * Inches(2.45)
    y = Inches(4.7) + (i // 5) * Inches(0.7)
    add_rect(slide, x, y, Inches(2.3), Inches(0.6), color=COL_PRIMARY)
    add_text(slide, x, y, Inches(2.3), Inches(0.6),
             country, size=22, bold=True, color=COL_BG, align=PP_ALIGN.CENTER)

# Languages
add_text(slide, Inches(0.6), Inches(6.3), Inches(12), Inches(0.4),
         "Talen: NL · EN · DE · FR · IT · ES · PL (+ meer via AI vertaling)",
         size=14, color=COL_MUTED)

add_phase_strip(slide, 0)

# =============================================================
# SLIDE 6 — COMPLIANCE DISCLOSURE (HK-only)
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "⚠️ Compliance Disclosure · HK-only seller",
         size=28, bold=True, color=COL_WARN, font='Cambria')

add_text(slide, Inches(0.6), Inches(1.1), Inches(12), Inches(0.5),
         "Jullie hebben expliciet gekozen voor een HK-bedrijf als enige seller.",
         size=14, color=COL_TEXT)
add_text(slide, Inches(0.6), Inches(1.5), Inches(12), Inches(0.5),
         "Dit brengt bekende risico's met zich mee. Mitigatie in de roadmap.",
         size=14, color=COL_TEXT)

risks = [
    ("GPSR Responsible Person", "Extern service (~€50-150/SKU/jaar)", COL_WARN),
    ("Douane / Btw", "Klanten kunnen btw-verrassingen krijgen bij import (geen IOSS)", COL_WARN),
    ("NL-belastingdienst", "DAC7 reporting via Shopify/Meta/WISE — wereldwijd inkomen belastbaar omdat Ali fiscaal inwoner NL is", COL_WARN),
    ("Productaansprakelijkheid", "Blijft gelden ongeacht seller-land", COL_WARN),
    ("AVB verzekering", "Aanbevolen voor keuken-producten (food contact)", COL_PRIMARY),
]
for i, (title, desc, color) in enumerate(risks):
    y = Inches(2.3) + i * Inches(0.85)
    add_rect(slide, Inches(0.6), y, Inches(0.3), Inches(0.7), color=color)
    add_text(slide, Inches(1.0), y, Inches(4.0), Inches(0.4),
             title, size=14, bold=True, color=COL_DARK)
    add_text(slide, Inches(5.0), y, Inches(7.5), Inches(0.7),
             desc, size=12, color=COL_TEXT)

# Recommendation box
add_rect(slide, Inches(0.6), Inches(6.8), Inches(12.2), Inches(0.4), color=COL_ACCENT)
add_text(slide, Inches(0.7), Inches(6.85), Inches(12.0), Inches(0.3),
         "Aanbeveling: NL-belastingadviseur voor DAC7 compliance monitoring (jaarlijks, niet in 90-dgn roadmap)",
         size=12, bold=True, color=COL_DARK)

# =============================================================
# SLIDE 7 — F0 MILESTONE
# =============================================================
slide = add_blank_slide()
add_header(slide, "F0", "Foundations", "D-7 → D0", "HK live, Shopify opgezet, design klaar, alle infra")

# Sub-tasks
tasks = [
    ("HK bank account aanmaken", "Kheibar", "vrijdag"),
    ("WISE + Currency Ways accounts", "Kheibar", "vrijdag"),
    ("HK bedrijf oprichten", "Kheibar", "vrijdag"),
    ("Shopify master + Markets setup", "Farshad", "D-7"),
    ("Custom theme design", "Designer", "D-5 → D-1"),
    ("GPSR Responsible Person service", "Ali", "D-5"),
    ("Privacy + Cookie + Terms + Returns policies", "Hermes", "D-3"),
    ("Shopify Inbox + chatbot widget", "Ali", "D-2"),
    ("GA4 + Meta Pixel", "Hermes", "D-2"),
    ("Introductie Farshad's China supplier", "Farshad", "ma"),
]
y = Inches(1.5)
for name, owner, day in tasks:
    color = OWNER_COLORS.get(owner.lower().replace(" (f)", "designer").split()[0], COL_MUTED)
    # Owner dot
    add_rect(slide, Inches(0.6), y + Inches(0.1), Inches(0.15), Inches(0.15), color=color)
    add_text(slide, Inches(0.9), y, Inches(7.0), Inches(0.3),
             name, size=12, color=COL_TEXT)
    add_text(slide, Inches(8.0), y, Inches(1.5), Inches(0.3),
             owner, size=11, color=color, bold=True)
    add_text(slide, Inches(9.7), y, Inches(3.0), Inches(0.3),
             day, size=11, color=COL_MUTED)
    y += Inches(0.42)

add_phase_strip(slide, 0)

# =============================================================
# SLIDE 8 — F1 MILESTONE
# =============================================================
slide = add_blank_slide()
add_header(slide, "F1", "Brand + Product Setup", "D0 → D14", "Brand naam, logo, 10 producten, listings in voorbereiding")

tasks = [
    ("Brand voice document (Ali)", "Ali", "D0"),
    ("Brand naam brainstorm (60 namen)", "Ali", "D0-D1"),
    ("EUIPO + WIPO trademark search", "Ali", "D1"),
    ("Brand naam lock-in", "Ali", "D2"),
    ("Logo design", "Designer", "D2-D5"),
    ("Product onderzoek (10 producten)", "Ali", "vrijdag"),
    ("Pain/Diff/Ad-ability score × 10", "Ali", "D2"),
    ("Amazon scrape (per segment)", "Hermes", "D2"),
    ("Supplier shortlist per SKU", "Farshad", "D3-D5"),
    ("Sample aanvragen × 10 SKUs", "Farshad", "D5"),
    ("Sample quality check", "Farshad", "D10-D14"),
    ("Per-product compliance check", "Ali", "D7"),
    ("GPSR × 10 SKUs", "Ali", "D10"),
    ("Product page copy × 10", "Ali", "D10-D14"),
]
y = Inches(1.4)
for name, owner, day in tasks:
    owner_key = owner.lower().replace(" (f)", "designer").split()[0]
    color = OWNER_COLORS.get(owner_key, COL_MUTED)
    add_rect(slide, Inches(0.6), y + Inches(0.08), Inches(0.15), Inches(0.15), color=color)
    add_text(slide, Inches(0.9), y, Inches(7.0), Inches(0.3),
             name, size=11, color=COL_TEXT)
    add_text(slide, Inches(8.0), y, Inches(1.5), Inches(0.3),
             owner, size=10, color=color, bold=True)
    add_text(slide, Inches(9.7), y, Inches(3.0), Inches(0.3),
             day, size=10, color=COL_MUTED)
    y += Inches(0.36)

add_phase_strip(slide, 1)

# =============================================================
# SLIDE 9 — F2 MILESTONE
# =============================================================
slide = add_blank_slide()
add_header(slide, "F2", "Listings + Wave 1 Launch", "D14 → D35", "10 listings live, Wave 1 eerste ads")

tasks = [
    ("Product photos × 10 SKUs (5-8/product)", "Designer", "D14-D18"),
    ("Product videos × 10 SKUs", "Designer", "D15-D20"),
    ("Listings live × 10", "Farshad", "D18"),
    ("Pricing per SKU", "Ali", "D18"),
    ("FAQ × 10 SKUs", "Hermes", "D19"),
    ("Wave 1 selectie (3-4 SKUs)", "Ali", "D20"),
    ("UGC + AI Higgsfield × 10/SKU", "Ali", "D20-D24"),
    ("Meta Ads Wave 1 launch", "Ali", "D25"),
    ("Google Ads Wave 1 launch", "Ali", "D27"),
    ("Email flows live", "Ali", "D24"),
    ("Daily metrics dashboard", "Hermes", "D26"),
    ("Wave 1 review (D+14)", "Ali", "D39"),
]
y = Inches(1.4)
for name, owner, day in tasks:
    owner_key = owner.lower().replace(" (f)", "designer").split()[0]
    color = OWNER_COLORS.get(owner_key, COL_MUTED)
    add_rect(slide, Inches(0.6), y + Inches(0.08), Inches(0.15), Inches(0.15), color=color)
    add_text(slide, Inches(0.9), y, Inches(7.0), Inches(0.3),
             name, size=11, color=COL_TEXT)
    add_text(slide, Inches(8.0), y, Inches(1.5), Inches(0.3),
             owner, size=10, color=color, bold=True)
    add_text(slide, Inches(9.7), y, Inches(3.0), Inches(0.3),
             day, size=10, color=COL_MUTED)
    y += Inches(0.4)

add_phase_strip(slide, 2)

# =============================================================
# SLIDE 10 — F3 MILESTONE
# =============================================================
slide = add_blank_slide()
add_header(slide, "F3", "Wave 2 + Wave 3 + Validation", "D35 → D65", "Alle 3 waves getest, winners geïdentificeerd")

tasks = [
    ("Wave 2 selectie + listings + video's", "Ali/Farshad", "D40-D44"),
    ("Wave 2 ads launch", "Ali", "D45"),
    ("Wave 2 review (D+14)", "Ali", "D59"),
    ("Wave 3 selectie + listings + video's", "Ali/Farshad", "D51-D55"),
    ("Wave 3 ads launch", "Ali", "D56"),
    ("Wave 3 review (D+14)", "Ali", "D70"),
    ("Kill rules automation script", "Hermes", "D36"),
    ("Per-SKU dashboard", "Hermes", "D38"),
    ("Cohort analysis tooling", "Hermes", "D50"),
    ("Customer survey NPS", "Ali", "D55"),
    ("Weekly review meetings", "Ali", "D36+"),
    ("Decision log per SKU", "Ali", "D36+"),
]
y = Inches(1.4)
for name, owner, day in tasks:
    add_text(slide, Inches(0.6), y, Inches(8.5), Inches(0.3),
             name, size=11, color=COL_TEXT)
    add_text(slide, Inches(9.2), y, Inches(1.5), Inches(0.3),
             owner, size=10, color=COL_MUTED, bold=True)
    add_text(slide, Inches(10.8), y, Inches(2.0), Inches(0.3),
             day, size=10, color=COL_MUTED)
    y += Inches(0.4)

add_phase_strip(slide, 3)

# =============================================================
# SLIDE 11 — F4 MILESTONE
# =============================================================
slide = add_blank_slide()
add_header(slide, "F4", "Scale Winners", "D65 → D85", "Budget omhoog op winners, repeat halo start")

tasks = [
    ("Winners definitief (3 SKUs)", "Ali", "D65"),
    ("Scale budget €100/dag/winner", "Ali", "D67"),
    ("Lookalikes op kopers", "Ali", "D68"),
    ("Creative refresh elke 14 dgn", "Ali", "D70+"),
    ("Google PMax campaigns", "Ali", "D72"),
    ("Email flow post-purchase upsell", "Ali", "D75"),
    ("Email flow win-back (60d)", "Ali", "D78"),
    ("Bundle pricing tests", "Ali", "D75"),
    ("Cross-SKU recommendations", "Hermes", "D80"),
    ("LTV:CAC dashboard per channel", "Hermes", "D78"),
    ("Repeat-buyer cohort analyse", "Hermes", "D82"),
    ("Cash forecast 4-weken runway", "Kheibar", "D85"),
]
y = Inches(1.4)
for name, owner, day in tasks:
    owner_key = owner.lower().split()[0]
    color = OWNER_COLORS.get(owner_key, COL_MUTED)
    add_rect(slide, Inches(0.6), y + Inches(0.08), Inches(0.15), Inches(0.15), color=color)
    add_text(slide, Inches(0.9), y, Inches(7.0), Inches(0.3),
             name, size=11, color=COL_TEXT)
    add_text(slide, Inches(8.0), y, Inches(1.5), Inches(0.3),
             owner, size=10, color=color, bold=True)
    add_text(slide, Inches(9.7), y, Inches(3.0), Inches(0.3),
             day, size=10, color=COL_MUTED)
    y += Inches(0.4)

add_phase_strip(slide, 4)

# =============================================================
# SLIDE 12 — F5 MILESTONE
# =============================================================
slide = add_blank_slide()
add_header(slide, "F5", "Lock + Optimize", "D85 → D90", "Optimalisaties, documentatie, exit-potentieel assessment")

tasks = [
    ("Eindstand metrics review", "Ali", "D87"),
    ("Lessons learned doc", "Ali", "D88"),
    ("Updated roadmap v2.0", "Hermes", "D89"),
    ("Supplier scorecard", "Farshad", "D88"),
    ("Customer feedback analysis", "Ali", "D89"),
    ("Compliance audit", "Ali", "D89"),
    ("Financial close-out (90 dgn P&L)", "Kheibar", "D90"),
    ("Exit-potentieel assessment", "Ali", "D90"),
    ("NL-belastingadvies call", "Ali", "D90+"),
    ("Influencer strategy (post-90)", "Ali", "D90+"),
]
y = Inches(1.4)
for name, owner, day in tasks:
    owner_key = owner.lower().split()[0]
    color = OWNER_COLORS.get(owner_key, COL_MUTED)
    add_rect(slide, Inches(0.6), y + Inches(0.08), Inches(0.15), Inches(0.15), color=color)
    add_text(slide, Inches(0.9), y, Inches(7.0), Inches(0.3),
             name, size=11, color=COL_TEXT)
    add_text(slide, Inches(8.0), y, Inches(1.5), Inches(0.3),
             owner, size=10, color=color, bold=True)
    add_text(slide, Inches(9.7), y, Inches(3.0), Inches(0.3),
             day, size=10, color=COL_MUTED)
    y += Inches(0.5)

add_phase_strip(slide, 5)

# =============================================================
# SLIDE 13 — RECURRING OPS
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "Recurring Operations · Dagelijks ritme",
         size=28, bold=True, color=COL_DARK, font='Cambria')

ops = [
    ("Dagelijks 30 min", "Morning metrics (Hermes cron)", "Sys", "Telegram digest 06:30"),
    ("09:00", "Customer service sweep", "Ali/Kheibar/Farshad", "Shopify Inbox"),
    ("12:00", "Trend-spotting + content scroll", "Ali", "TikTok/IG"),
    ("18:00", "Content post (1 video/brand)", "Ali", "Multi-language"),
    ("21:00", "EOD notes deposit", "Ali", "Hermes eod"),
    ("22:00", "Review queue (5/brand)", "Ali", "Shopify"),
]
y = Inches(1.3)
for time, what, who, tool in ops:
    add_text(slide, Inches(0.6), y, Inches(1.5), Inches(0.3),
             time, size=12, bold=True, color=COL_PRIMARY)
    add_text(slide, Inches(2.2), y, Inches(4.5), Inches(0.3),
             what, size=12, color=COL_TEXT)
    add_text(slide, Inches(7.0), y, Inches(3.0), Inches(0.3),
             who, size=12, color=COL_MUTED)
    add_text(slide, Inches(10.0), y, Inches(3.0), Inches(0.3),
             tool, size=12, color=COL_MUTED)
    y += Inches(0.5)

# Weekly
add_text(slide, Inches(0.6), Inches(4.6), Inches(12), Inches(0.4),
         "Wekelijks", size=18, bold=True, color=COL_PRIMARY)
weekly = [
    ("Ma 08:00", "Auto-sourcing bottom-5 vervangen", "Sys"),
    ("Di 14:00", "Weekly review (Ali + Kheibar + Farshad)", "Ali"),
    ("Wo 10:00", "Ad-budget + bid adjustments", "Ali"),
    ("Wo 16:00", "Supplier check-in", "Farshad"),
    ("Vr 14:00", "Bulk order decisions", "Farshad"),
    ("Vr 16:00", "Cash forecast update", "Kheibar"),
    ("Zo 21:00", "Weekly Telegram digest (auto)", "Sys"),
]
y = Inches(5.1)
for time, what, who in weekly:
    add_text(slide, Inches(0.6), y, Inches(1.5), Inches(0.3),
             time, size=11, color=COL_PRIMARY)
    add_text(slide, Inches(2.2), y, Inches(7.0), Inches(0.3),
             what, size=11, color=COL_TEXT)
    add_text(slide, Inches(9.5), y, Inches(3.0), Inches(0.3),
             who, size=11, color=COL_MUTED, bold=True)
    y += Inches(0.27)

add_phase_strip(slide, 5)

# =============================================================
# SLIDE 14 — NEXT STEPS
# =============================================================
slide = add_blank_slide()
add_text(slide, Inches(0.6), Inches(0.4), Inches(12), Inches(0.6),
         "Next Steps · Eerste 14 dagen",
         size=32, bold=True, color=COL_DARK, font='Cambria')

next_steps = [
    ("vrijdag 12 sept", "HK entity + bank + WISE + Currency Ways", "Kheibar"),
    ("vrijdag 12 sept", "Product research (10 producten)", "Ali"),
    ("vrijdag 12 sept", "Brand naam input delen met Hermes", "Ali"),
    ("ma 15 sept", "Introductie Farshad's China supplier", "Farshad"),
    ("D-7 (zondag)", "Shopify master + Markets setup start", "Farshad"),
    ("D-5", "GPSR Responsible Person service kiezen", "Ali"),
    ("D-5", "Trademark search (vóór naam lock-in)", "Ali"),
    ("D-2", "Compliance disclosure met Kheibar + Farshad reviewen", "Allen"),
    ("D0", "GO/NO-GO 🚀", "Allen"),
]
y = Inches(1.3)
for when, what, who in next_steps:
    add_text(slide, Inches(0.6), y, Inches(2.0), Inches(0.4),
             when, size=14, bold=True, color=COL_ACCENT)
    add_text(slide, Inches(2.8), y, Inches(7.5), Inches(0.4),
             what, size=14, color=COL_TEXT)
    add_text(slide, Inches(10.5), y, Inches(2.5), Inches(0.4),
             who, size=12, color=COL_PRIMARY, bold=True)
    y += Inches(0.55)

# Companion webapp reference
add_rect(slide, Inches(0.6), Inches(6.5), Inches(12.2), Inches(0.7), color=COL_PRIMARY)
add_text(slide, Inches(0.7), Inches(6.55), Inches(12.0), Inches(0.3),
         "📱 Companion webapp: Vercel-deployed voor live tracking per owner + status",
         size=14, bold=True, color=COL_BG)
add_text(slide, Inches(0.7), Inches(6.85), Inches(12.0), Inches(0.3),
         "Filters per Ali/Kheibar/Farshad/Designer/Hermes · Status checkboxes · Comments per taak",
         size=11, color=COL_BG)

# =============================================================
# SAVE
# =============================================================
output_path = '/home/alisionary/.hermes/profiles/hermes_ecom_research/projects/kitchen-roadmap/kitchen-launch-proposal.pptx'
prs.save(output_path)
print(f"PPTX saved: {output_path}")
print(f"Total slides: {len(prs.slides)}")