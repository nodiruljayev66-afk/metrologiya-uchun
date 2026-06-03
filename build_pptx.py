"""
PPTX Generator — stdlib only (zipfile + XML)
10 ta professional slide:
  Noaniqlik hisoblash | Toksik moddalar | Metrologik ishlar
ISO/IEC 17025 | GUM | ILAC-G8 | SANTE
"""
import zipfile, textwrap, re

# ── EMU helpers ───────────────────────────────────────────────────
W  = 9144000   # slide width  (25.4 cm)
H  = 5143500   # slide height (14.29 cm) — widescreen 16:9
CM = 360000    # 1 cm in EMU
PT = 12700     # 1 pt in EMU

def emu(cm): return int(cm * CM)
def pt(p):   return int(p  * PT)

# ── Colour palette ────────────────────────────────────────────────
NAVY   = "1F3864"
BLUE   = "2E75B6"
LBLUE  = "D6E4F0"
GREEN  = "375623"
LGREEN = "E2EFDA"
RED    = "C00000"
ORANGE = "F4B942"
YELLOW = "FFD966"
WHITE  = "FFFFFF"
LGRAY  = "F2F2F2"
DGRAY  = "595959"
TEAL   = "00B0F0"
PURPLE = "7030A0"

# ═══════════════════════════════════════════════════════════════════
#  LOW-LEVEL XML PRIMITIVES
# ═══════════════════════════════════════════════════════════════════
def xml_esc(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def solid_fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

def sp_pr(x, y, cx, cy, fill=None, border_color=None, rounding=None):
    """Shape properties (position + size + optional fill + optional border)."""
    ln = ""
    if border_color:
        ln = f'<a:ln><a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill></a:ln>'
    rnd = f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {rounding}"/></a:avLst></a:prstGeom>' \
          if rounding else '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    fill_xml = solid_fill(fill) if fill else "<a:noFill/>"
    return (f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
            f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'{rnd}{fill_xml}{ln}</p:spPr>')

def run(text, bold=False, sz=18, color="000000", italic=False):
    b  = "<a:b/>" if bold else ""
    i  = "<a:i/>" if italic else ""
    return (f'<a:r><a:rPr lang="uz-UZ" sz="{sz*100}" dirty="0" b="{1 if bold else 0}" '
            f'i="{1 if italic else 0}">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr>'
            f'<a:t>{xml_esc(text)}</a:t></a:r>')

def para(runs_xml, align="l", space_before=0, space_after=0, line_spacing=None):
    ls = f'<a:lnSpc><a:spcPts val="{line_spacing*100}"/></a:lnSpc>' if line_spacing else ""
    spc = (f'<a:spcBef><a:spcPts val="{space_before*100}"/></a:spcBef>'
           f'<a:spcAft><a:spcPts val="{space_after*100}"/></a:spcAft>')
    return (f'<a:p><a:pPr algn="{align}" indent="0">{spc}{ls}</a:pPr>'
            f'{"".join(runs_xml)}</a:p>')

def txBody(paras_xml, anchor="t", inset_l=91440, inset_t=45720):
    return (f'<p:txBody><a:bodyPr anchor="{anchor}" lIns="{inset_l}" '
            f'tIns="{inset_t}" rIns="{inset_l}" bIns="{inset_t}" wrap="square"/>'
            f'<a:lstStyle/>'
            f'{"".join(paras_xml)}</p:txBody>')

def shape(sp_pr_xml, txbody_xml, sp_id):
    return (f'<p:sp><p:nvSpPr>'
            f'<p:cNvPr id="{sp_id}" name="sp{sp_id}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr/></p:nvSpPr>'
            f'{sp_pr_xml}{txbody_xml}</p:sp>')


# ═══════════════════════════════════════════════════════════════════
#  HIGH-LEVEL SLIDE BUILDERS
# ═══════════════════════════════════════════════════════════════════
_shapes = []   # accumulator per slide
_sid    = [1]  # shape id counter

def reset():
    _shapes.clear()
    _sid[0] = 1

def nid():
    v = _sid[0]; _sid[0] += 1; return v

# ── Coloured rectangle with text ─────────────────────────────────
def box(x, y, cx, cy, text, fill, text_color=WHITE,
        bold=True, sz=14, align="ctr", anchor="ctr",
        italic=False, rounding=None, border=None):
    sp = sp_pr(x, y, cx, cy, fill=fill, border_color=border, rounding=rounding)
    lines = text.split("\n")
    ps = []
    for li, line in enumerate(lines):
        ps.append(para([run(line, bold=bold, sz=sz, color=text_color, italic=italic)],
                       align=align, space_before=0 if li else 0))
    tx = txBody(ps, anchor=anchor)
    _shapes.append(shape(sp, tx, nid()))

# ── Plain text box (transparent bg) ──────────────────────────────
def txt(x, y, cx, cy, text, bold=False, sz=13, color="000000",
        align="l", anchor="t", italic=False):
    sp = sp_pr(x, y, cx, cy)
    lines = text.split("\n")
    ps = [para([run(l, bold=bold, sz=sz, color=color, italic=italic)],
               align=align, space_before=0) for l in lines]
    tx = txBody(ps, anchor=anchor)
    _shapes.append(shape(sp, tx, nid()))

# ── Bullet list ───────────────────────────────────────────────────
def bullets(x, y, cx, cy, items, sz=11, color="000000",
            bullet_color=BLUE, bg=None, bold_first=False):
    sp = sp_pr(x, y, cx, cy, fill=bg, border_color="CCCCCC" if bg else None)
    ps = []
    for i, item in enumerate(items):
        bullet = "• "
        b = bold_first and i == 0
        ps.append(para(
            [run(bullet, bold=True, sz=sz, color=bullet_color),
             run(item,   bold=b,    sz=sz, color=color)],
            align="l", space_before=3, space_after=3))
    tx = txBody(ps, anchor="t", inset_l=emu(0.3), inset_t=emu(0.2))
    _shapes.append(shape(sp, tx, nid()))

# ── Horizontal divider line ───────────────────────────────────────
def hline(y, color=BLUE):
    _shapes.append(
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="{nid()}" name="line{nid()}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{emu(0.5)}" y="{y}"/>'
        f'<a:ext cx="{emu(24.4)}" cy="{pt(2)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'{solid_fill(color)}</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')

# ── Table ─────────────────────────────────────────────────────────
def table(x, y, cx, rows_data, col_widths,
          header_fill=NAVY, header_fc=WHITE,
          alt_fill=LGRAY, border_color="CCCCCC"):
    """rows_data: list of lists. First row = header."""
    row_h = pt(22)
    total_h = row_h * len(rows_data)
    tbl_id = nid()

    col_xml = "".join(f'<a:gridCol w="{w}"/>' for w in col_widths)
    rows_xml = ""
    for ri, row in enumerate(rows_data):
        fill = header_fill if ri == 0 else (alt_fill if ri % 2 == 1 else WHITE)
        fc   = header_fc   if ri == 0 else "000000"
        bld  = (ri == 0)
        cells = ""
        for ci, cell in enumerate(row):
            border_xml = (f'<a:lnL w="12700"><a:solidFill><a:srgbClr val="{border_color}"/>'
                          f'</a:solidFill></a:lnL>'
                          f'<a:lnR w="12700"><a:solidFill><a:srgbClr val="{border_color}"/>'
                          f'</a:solidFill></a:lnR>'
                          f'<a:lnT w="12700"><a:solidFill><a:srgbClr val="{border_color}"/>'
                          f'</a:solidFill></a:lnT>'
                          f'<a:lnB w="12700"><a:solidFill><a:srgbClr val="{border_color}"/>'
                          f'</a:solidFill></a:lnB>')
            cells += (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>'
                      f'<a:p><a:pPr algn="ctr"/>'
                      f'{run(str(cell), bold=bld, sz=10, color=fc)}'
                      f'</a:p></a:txBody>'
                      f'<a:tcPr>{border_xml}'
                      f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
                      f'</a:tcPr></a:tc>')
        rows_xml += f'<a:tr h="{row_h}">{cells}</a:tr>'

    tbl_xml = (f'<p:graphicFrame><p:nvGraphicFramePr>'
               f'<p:cNvPr id="{tbl_id}" name="tbl{tbl_id}"/>'
               f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
               f'<p:nvPr/></p:nvGraphicFramePr>'
               f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{total_h}"/></p:xfrm>'
               f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
               f'<a:tbl><a:tblPr firstRow="1" bandRow="1">'
               f'<a:tableStyleId>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId>'
               f'</a:tblPr>'
               f'<a:tblGrid>{col_xml}</a:tblGrid>'
               f'{rows_xml}'
               f'</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')
    _shapes.append(tbl_xml)

# ── Slide header banner ───────────────────────────────────────────
def header_banner(title, subtitle=None, slide_num=None):
    # Dark navy top bar
    box(0, 0, W, emu(2.4), "", fill=NAVY)
    # Slide number badge
    if slide_num:
        box(emu(0.3), emu(0.25), emu(1.2), emu(1.0),
            str(slide_num), fill=TEAL, sz=22, bold=True, anchor="ctr", align="ctr")
    # Title
    tx_x = emu(1.8) if slide_num else emu(0.5)
    txt(tx_x, emu(0.2), W - tx_x - emu(0.5), emu(1.2),
        title, bold=True, sz=22, color=WHITE, align="l", anchor="ctr")
    # Subtitle bar
    if subtitle:
        box(0, emu(2.4), W, emu(0.7), subtitle,
            fill=BLUE, sz=10, bold=False, align="ctr", anchor="ctr")

# ── Footer ────────────────────────────────────────────────────────
def footer(ref_text):
    box(0, H - emu(0.6), W, emu(0.6),
        ref_text, fill=NAVY, sz=8, bold=False, align="ctr", anchor="ctr")

# ── Assemble slide XML ────────────────────────────────────────────
def slide_xml():
    ns = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    body = "\n".join(str(s) for s in _shapes)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:sld {ns}><p:cSld><p:spTree>'
            f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
            f'<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>'
            f'{body}'
            f'</p:spTree></p:cSld></p:sld>')


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 1 — TITLE SLIDE
# ═══════════════════════════════════════════════════════════════════
def slide1():
    reset()
    # Full background
    box(0, 0, W, H, "", fill=NAVY)
    # Accent top stripe
    box(0, 0, W, emu(0.5), "", fill=TEAL)
    # Accent bottom stripe
    box(0, H - emu(0.5), W, emu(0.5), "", fill=TEAL)

    # Main title
    txt(emu(0.8), emu(1.2), W - emu(1.6), emu(2.0),
        "O'LCHOV NOANIQLIGINI HISOBLASH\nVA QAROR QABUL QILISH QOIDALARI",
        bold=True, sz=28, color=WHITE, align="ctr", anchor="ctr")

    # Subtitle
    txt(emu(1.0), emu(3.4), W - emu(2.0), emu(0.8),
        "Mahsulotlarda Toksik Moddalar Aniqlanishi va Umumiy Metrologik Amaliyot",
        bold=False, sz=15, color=TEAL, align="ctr", anchor="ctr")

    hline(emu(4.4), TEAL)

    # Standards bar
    box(emu(0.5), emu(4.6), W - emu(1.0), emu(0.65),
        "ISO/IEC 17025:2017  |  GUM JCGM 100:2008  |  ILAC-G8:2019  |  SANTE/11813  |  EA-4/02",
        fill="17375E", sz=11, bold=False, align="ctr")

    # Bottom info
    txt(emu(0.8), emu(5.4), W - emu(1.6), emu(0.5),
        "Laboratoriya Sifat Menejeri  •  Metrologik Nazorat Tizimi  •  2025",
        bold=False, sz=10, color="8DB4E2", align="ctr")

    return slide_xml()

# ═══════════════════════════════════════════════════════════════════
#  SLIDE 2 — NOANIQLIK NIMA VA NIMA UCHUN KERAK?
# ═══════════════════════════════════════════════════════════════════
def slide2():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("NOANIQLIK — ASOSIY TUSHUNCHALAR", "GUM JCGM 100:2008 | ISO/IEC 17025:2017 Cl.7.6", 2)
    footer("GUM (2008): 'Measurement uncertainty — parameter characterizing the dispersion of values reasonably attributed to the measurand'")

    # Definition box
    box(emu(0.4), emu(3.15), emu(14.5), emu(1.6),
        "Ta'rif (GUM §2.2.3):\n«O'lchov noaniqlik — o'lchanayotgan kattalikka asosli ravishda\nnisblash mumkin bo'lgan qiymatlar tarqalishini tavsifluvchi parametr»",
        fill=LBLUE, text_color=NAVY, sz=11, bold=False, align="l", anchor="t",
        border="2E75B6")

    # 3 concept boxes
    concepts = [
        (BLUE,   "XATO ≠ NOANIQLIK",
                 "Xato — aniq ma'lum farq\nNoaniqlik — mumkin bo'lgan\nfarqlar oralig'i"),
        (GREEN,  "NIMA UCHUN KERAK?",
                 "• Natija to'g'riligini isbotlash\n• Qarorni asoslash\n• ISO 17025 talabi"),
        (PURPLE, "FORMULA:",
                 "U = k × u_c\nu_c = √Σu²(xᵢ)\nk=2 → ~95% ishonch"),
    ]
    box_w = emu(7.8)
    for i, (col, ttl, desc) in enumerate(concepts):
        col_x = [emu(0.4), emu(8.6), emu(0.4)][i]
        col_y = [emu(3.15), emu(3.15), emu(6.5)][i]
        if i < 2:
            box(col_x, col_y, box_w, emu(2.9), ttl, fill=col, sz=13, bold=True, anchor="t", align="ctr")
            txt(col_x + emu(0.2), col_y + emu(0.8), box_w - emu(0.4), emu(2.0),
                desc, sz=11, color=WHITE, align="l")

    # Formula panel
    box(emu(0.4), emu(6.3), W - emu(0.8), emu(1.35),
        "U = k × u_c    →    Natija = x̄ ± U [birlik]   (k=2, ~95% ishonch oralig'i)",
        fill=NAVY, sz=15, bold=True, align="ctr", anchor="ctr")

    # Key points
    bullets(emu(8.6), emu(3.15), box_w, emu(2.9),
            ["U hisoblash majburiy (ISO 17025)",
             "Hisobotda: x̄ ± U (k=2) ko'rinishida",
             "U > 1/3×tolerans → metod yaroqsiz",
             "Kuzatilish zanjiri talab qilinadi"],
            sz=10, color=WHITE, bullet_color=YELLOW, bg=BLUE)

    return slide_xml()

# ═══════════════════════════════════════════════════════════════════
#  SLIDE 3 — NOANIQLIK MANBALARI (ISHIKAWA)
# ═══════════════════════════════════════════════════════════════════
def slide3():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("NOANIQLIK MANBALARI — BALIQ SUYAGI DIAGRAMMASI",
                  "EURACHEM/CITAC CG4 | Bottom-up yondashuv | ISO/IEC 17025 Cl.7.6.2", 3)
    footer("Har bir manba u(xi) sifatida miqdorlanadi va kvadrat yig'indisi orqali uc ga qo'shiladi")

    # Central arrow
    box(emu(3.5), emu(5.8), emu(18.2), emu(0.55), "➤  UMUMIY NOANIQLIK  uc", fill=RED, sz=12, bold=True, align="ctr")

    # 6 source boxes — 3 top, 3 bottom
    sources_top = [
        (BLUE,   "1. STANDARTLAR\n& CRM",
                 "• Sertifikat sof.(%)\n• Lot-to-lot o'zgarish\n• Saqlash sharoiti"),
        (PURPLE, "2. TAROZILASH\n& HAJM",
                 "• Tarozilay og'ishi\n• Pipetka/kolba sertif.\n• Harorat ta'siri"),
        (GREEN,  "3. EKSTRAKSIYA\n& TAYYORLASH",
                 "• Recovery o'zgarishi\n• SPE/LLE unumi\n• Matritsa ta'siri"),
    ]
    sources_bot = [
        ("C55A11", "4. ASBOB\nJAVOBI",
                   "• Kalibrasiya aniqligi\n• Detektor barqarorligi\n• MS ion suppression"),
        (NAVY,    "5. ATROF-MUHIT",
                  "• Harorat (±1°C)\n• Namlik ta'siri\n• Titrasyon xatosi"),
        (RED,     "6. OPERATORGA\nBOG'LIQ",
                  "• Operator farqi\n• Kunlararo farq\n• IQC intermediate SD"),
    ]
    bw = emu(7.5)
    bh = emu(2.4)
    xs = [emu(0.3), emu(8.5), emu(16.7)]
    for i, (col, ttl, desc) in enumerate(sources_top):
        box(xs[i], emu(3.1), bw, bh, ttl, fill=col, sz=11, bold=True, anchor="t", align="ctr")
        txt(xs[i]+emu(0.2), emu(3.85), bw-emu(0.4), emu(1.6), desc, sz=9, color=WHITE)

    for i, (col, ttl, desc) in enumerate(sources_bot):
        box(xs[i], emu(6.45), bw, bh, ttl, fill=col, sz=11, bold=True, anchor="t", align="ctr")
        txt(xs[i]+emu(0.2), emu(7.2), bw-emu(0.4), emu(1.6), desc, sz=9, color=WHITE)

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 4 — NOANIQLIK BYUDJETI (JADVAL)
# ═══════════════════════════════════════════════════════════════════
def slide4():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("NOANIQLIK BYUDJETI — HISOBLASH JADVALI",
                  "GUM §5 | EURACHEM/CITAC CG4 — Pestitsid qoldig'i (LC-MS/MS, 50 µg/kg) misoli", 4)
    footer("u_c = √(0.029²+0.017²+0.023²+0.031²+0.012²+0.008²) = 0.053 → U = 2×0.053 = 0.106 → 10.6 µg/kg (k=2)")

    # Budget table
    rows = [
        ["Noaniqlik manbai", "Qiymat (a)", "Taqsimot", "Divisor", "u(xᵢ)", "u²(xᵢ)", "Hissa (%)"],
        ["CRM sofligi (99.2±0.4%)", "0.4%", "To'g'ri burchak (R)", "√3=1.732", "0.231%", "0.053%²", "10.7"],
        ["Tarozilay (±0.1 mg)",     "0.1 mg","R",                  "√3=1.732", "0.058 mg","0.003 mg²","1.1"],
        ["Ekstraksiya Recovery",    "SD=2.3%","Normal (N)",         "1",        "2.3%",    "5.3%²",    "23.8"],
        ["Kalibrasiya egri chizig'i","SD=1.7%","N",                 "1",        "1.7%",    "2.9%²",    "13.0"],
        ["Matritsa ta'siri (IS CV)","SD=3.1%","N",                 "1",        "3.1%",    "9.6%²",    "43.0"],
        ["Hajmiy idishlar",         "0.08%", "R",                  "√3",       "0.046%",  "0.002%²",  "0.9"],
        ["Harorat ta'siri",         "0.5%",  "R",                  "√3",       "0.289%",  "0.083%²",  "7.5"],
        ["JAMI  u_c (rel.)",        "—",     "—",                  "—",        "5.3%",    "28.0%²",   "100"],
    ]
    col_w = [emu(6.8), emu(3.5), emu(3.5), emu(3.2), emu(2.8), emu(2.8), emu(2.4)]
    table(emu(0.3), emu(3.15), W - emu(0.6), rows, col_w,
          header_fill=NAVY, alt_fill=LBLUE)

    # Result box
    box(emu(0.3), emu(7.95), W - emu(0.6), emu(0.75),
        "NATIJA:  50 µg/kg ± 10.6 µg/kg   (U, k=2, ~95% ishonch)   |   U_rel = 10.6 / 50 × 100 = 21.2%",
        fill=NAVY, sz=13, bold=True, align="ctr", anchor="ctr")

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 5 — TOP-DOWN va BOTTOM-UP TAQQOSLASH
# ═══════════════════════════════════════════════════════════════════
def slide5():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("NOANIQLIK HISOBLASH YONDASHUVLARI",
                  "Top-down (empirik) vs Bottom-up (GUM) — Foydalanish holatlari va afzalliklar", 5)
    footer("EURACHEM/CITAC CG4 (2012): Routine laboratoriyalar uchun Top-down, yangi metodlar uchun Bottom-up tavsiya etiladi")

    # Top-down box
    box(emu(0.3), emu(3.2), emu(12.0), emu(0.65),
        "⬆  TOP-DOWN YONDASHUV (EURACHEM CG4)", fill=BLUE, sz=13, bold=True, align="ctr")
    bullets(emu(0.3), emu(3.85), emu(12.0), emu(3.1),
            ["IQC ma'lumotlaridan hisoblash (≥6 oy, ≥60 ta natija)",
             "u_c = √[ u²(repeat.) + u²(inter-day SD) + u²(bias) ]",
             "Bias = |(x̄ − μ_ref)| / μ_ref  (CRM yoki PT dan)",
             "Kamchiligi: Sistemik xatolarni ko'rsatmasligi mumkin",
             "Qachon: Ko'p IQC tarixi bor kundalik testlar uchun"],
            sz=10, color="000000", bullet_color=BLUE, bg=LBLUE)

    # Bottom-up box
    box(emu(13.0), emu(3.2), emu(11.5), emu(0.65),
        "⬇  BOTTOM-UP YONDASHUV (GUM §5)", fill=GREEN, sz=13, bold=True, align="ctr")
    bullets(emu(13.0), emu(3.85), emu(11.5), emu(3.1),
            ["Har bir manba alohida identifikatsiya va miqdorlanadi",
             "u(xi) = a / divisor  (taqsimotga qarab)",
             "Kragten elektron jadvali yoki qisman hosilalar",
             "Afzalligi: Sistemik komponentlarni aniqlaydi",
             "Qachon: Yangi metod, validatsiya, akkreditatsiya"],
            sz=10, color="000000", bullet_color=GREEN, bg=LGREEN)

    # Combined formula
    box(emu(0.3), emu(7.15), W - emu(0.6), emu(0.65),
        "Ikkala yondashuv uchun ham:   U = k × u_c   (k=2 → ~95%)   |   Hisobotda: x̄ ± U [birlik]; k=2; CRM kuzatilishi",
        fill=NAVY, sz=12, bold=True, align="ctr")

    # Distribution types
    box(emu(0.3), emu(7.9), emu(8.0), emu(0.65),
        "NORMAL taqsimot → divisor = 1  (IQC SD)", fill=PURPLE, sz=11, align="ctr")
    box(emu(8.6), emu(7.9), emu(7.0), emu(0.65),
        "TO'G'RI BURCHAK → divisor = √3  (sertifikat ±)", fill="C55A11", sz=11, align="ctr")
    box(emu(16.0), emu(7.9), emu(8.5), emu(0.65),
        "UCHBURCHAK → divisor = √6  (ikkala chegara)", fill=RED, sz=11, align="ctr")

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 6 — TOKSIK MODDALAR: NOANIQLIK MISOLLARI
# ═══════════════════════════════════════════════════════════════════
def slide6():
    reset()
    box(0, 0, W, H, "", fill="FFF8F0")
    header_banner("TOKSIK MODDALAR — NOANIQLIK MISOLLARI",
                  "Pestitsidlar | Og'ir metallar | Mikotoksinlar | PAH | Veterinar dori qoldiqlari", 6)
    footer("SANTE/11813/2017: Recovery 70–120%, RSD ≤20%. EC 333/2007: Pb, Cd, Hg uchun MU ≤50% at MRL. COMMISSION REG (EC) 401/2006: Aflatoksin U≤50%")

    rows = [
        ["Analitik / Guruh", "Metod", "Matriks", "MRL / Cutoff", "Maqbul U_rel", "Qaror qoidasi"],
        ["Pestitsid qoldig'i\n(masalan, imidakloprid)", "LC-MS/MS\nSANTE/11813", "Meva/sabzavot", "0.01 mg/kg", "≤50% MRL da\n(~25-30% maqbul)", "Qat'iy Guard Band\nw = U"],
        ["Aflatoksin B1", "HPLC-FLD\nISO 16050", "Don, yong'oq", "2 µg/kg (EU)", "≤50% MRL", "Stringent GB\nILAC-G8"],
        ["Qo'rg'oshin (Pb)", "ICP-MS\nEN 15763", "Oziq-ovqat", "0.1 mg/kg", "≤50% (Reg.333/2007)", "Shared Risk\nyoki Stringent"],
        ["Simob (Hg)", "CV-AAS / ICP-MS\nISO 17852", "Baliq", "0.5 mg/kg", "≤50%", "Stringent GB"],
        ["Okratoksin A", "HPLC-FLD\nEN 14132", "Donli ekinlar", "3 µg/kg", "≤50%", "Guard Band w=U"],
        ["Veterinar dori\n(enrofloksatsin)", "LC-MS/MS\n2002/657/EC", "Go'sht, tuxum", "100 µg/kg", "≤20% (CCα)", "CCα/CCβ qaror"],
        ["PAH (B[a]P)", "GC-MS\nEN 16619", "Tutun, yog'", "2 µg/kg", "≤50%", "Guard Band"],
    ]
    col_w = [emu(5.5), emu(4.2), emu(3.5), emu(3.5), emu(4.5), emu(4.2)]
    table(emu(0.3), emu(3.15), W - emu(0.6), rows, col_w,
          header_fill=RED, alt_fill="FFF0E8")

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 7 — QAROR QABUL QILISH QOIDALARI (ILAC-G8)
# ═══════════════════════════════════════════════════════════════════
def slide7():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("QAROR QABUL QILISH QOIDALARI",
                  "ILAC-G8:09/2019 | ISO/IEC 17025:2017 Cl.7.8.6 | ASME B89.7.3.1", 7)
    footer("ILAC-G8: 'The decision rule shall be documented in the test report when a statement of conformity is made'")

    # Number line visual (manual boxes)
    y_nl = emu(3.3)
    # zones
    box(emu(0.3),  y_nl, emu(6.5), emu(1.1), "✅  MANFIY ZONA\nx < TL − U", fill="375623", sz=11, bold=True)
    box(emu(6.8),  y_nl, emu(4.0), emu(1.1), "🟡  NOANIQLIK\nZONASI", fill=ORANGE, sz=11, bold=True)
    box(emu(10.8), y_nl, emu(5.0), emu(1.1), "⚠ CHEGARA\nHUDUDI ±U", fill=YELLOW, text_color=NAVY, sz=11, bold=True)
    box(emu(15.8), y_nl, emu(8.7), emu(1.1), "🔴  MUSBAT ZONA\nx > TL + U", fill=RED, sz=11, bold=True)

    # Labels
    txt(emu(6.5), y_nl - emu(0.6), emu(2.0), emu(0.5), "TL − U", bold=True, sz=10, color=RED, align="ctr")
    txt(emu(15.5), y_nl - emu(0.6), emu(1.5), emu(0.5), "TL", bold=True, sz=11, color=NAVY, align="ctr")
    txt(emu(21.5), y_nl - emu(0.6), emu(2.0), emu(0.5), "TL + U", bold=True, sz=10, color=RED, align="ctr")

    # 4 rules comparison
    rules = [
        (BLUE,    "1. ODDIY QABUL",
                  "• Noaniqlik hisobga olinmaydi\n• x > TL → MUSBAT\n• Qonuniy jihatdan zaif!\n• Faqat mijoz rozi bo'lsa"),
        (GREEN,   "2. TAQSIMLANGAN RISK",
                  "• x > TL → Musbat\n• Yolg'on +/- risk teng\n• Har ikki tomon baham ko'radi\n• O'rta risk holatlarda"),
        ("C55A11","3. QAT'IY GUARD BAND\n(w = U) — TAVSIYA",
                  "• x ≥ TL+U → Musbat ✓\n• x ≤ TL−U → Manfiy ✓\n• Oradagilar: noaniq\n• Forensik/huquqiy uchun"),
        (PURPLE,  "4. YUMSHOQ GUARD BAND\n(klinik uchun)",
                  "• x ≥ TL−U → Musbat\n• Yolg'on manfiy minimal\n• Klinik favqulodda hollarda\n• Xavfli moddalar uchun"),
    ]
    bw = emu(6.0)
    for i, (col, ttl, desc) in enumerate(rules):
        bx = emu(0.3) + i * emu(6.2)
        box(bx, emu(4.65), bw, emu(0.7), ttl, fill=col, sz=9, bold=True, align="ctr")
        bullets(bx, emu(5.35), bw, emu(2.4),
                desc.strip().split("\n"),
                sz=9, color="000000", bullet_color=col, bg=LGRAY)

    # TUR box
    box(emu(0.3), emu(7.95), W - emu(0.6), emu(0.65),
        "TUR (Test Uncertainty Ratio) = Tolerans / (2×U)   →   TUR ≥ 4: Ishonchli qaror  |  TUR < 2: Qaror ishonchsiz",
        fill=NAVY, sz=12, bold=True, align="ctr")

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 8 — KUZATILISH ZANJIRI VA CRM
# ═══════════════════════════════════════════════════════════════════
def slide8():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("METROLOGIK KUZATILISH ZANJIRI VA CRM",
                  "ISO/IEC 17025 Cl.6.6 | ILAC-P10:2022 | VIM §2.41 | EA-4/02", 8)
    footer("VIM §2.41: Traceability — property of a measurement result relating it to a reference through a documented unbroken chain")

    # Pyramid levels
    levels = [
        (NAVY,    emu(9.5),  emu(1.2),  "SI BIRLIKLARI (mol, kg, m, K)",                             "BIPM / NMI"),
        (RED,     emu(6.5),  emu(1.2),  "BIRLAMCHI ETALON & CRM (NIST, LGC, BAM, IRMM)",             "U_ref ±δ%"),
        (BLUE,    emu(5.0),  emu(1.2),  "IKKILAMCHI STANDART / IN-HOUSE CRM",                         "Sertifikat + U"),
        (GREEN,   emu(4.0),  emu(1.2),  "ISHCHI KALIBRASIYA STANDARTLARI",                            "Kundalik tayyorlanish"),
        ("C55A11",emu(3.0),  emu(1.2),  "IQC / QC NAMUNALARI",                                        "Westgard nazorati"),
        (PURPLE,  emu(2.0),  emu(1.2),  "LABORATORIYA NATIJASI",                                       "x ± U (k=2)"),
    ]
    total_h = sum(l[2] for l in levels)
    start_y = emu(3.0)
    center_x = W // 2
    for i, (col, width, height, label_txt, right_txt) in enumerate(levels):
        x_pos = center_x - width // 2
        box(x_pos, start_y + i * emu(1.25), width, height,
            label_txt, fill=col, sz=10, bold=True, align="ctr")
        txt(x_pos + width + emu(0.2), start_y + i * emu(1.25) + emu(0.2),
            emu(5.0), emu(0.8), right_txt, sz=9, color=DGRAY, italic=True)
        # Arrow between levels
        if i < len(levels) - 1:
            txt(center_x - emu(0.3), start_y + i * emu(1.25) + height,
                emu(0.8), emu(0.25), "▼", sz=9, color=DGRAY, align="ctr")

    # CRM selection criteria
    bullets(emu(17.0), emu(3.2), emu(7.5), emu(5.0),
            ["Matrix-matched CRM afzal (LC, GC, ICP)",
             "Sertifikat: U va k ko'rsatilgan bo'lishi kerak",
             "NMI kuzatilishi: NIST/LGC/BAM/IRMM",
             "Amal qilish muddati tekshirilsin",
             "Saqlash: sertifikat haroratida",
             "Qayta kalibrlash: ma'lumotga asosida\n(R-chart, Nelson qoidasi)"],
            sz=9, color="000000", bullet_color=BLUE, bg=LBLUE)
    txt(emu(17.0), emu(3.0), emu(7.5), emu(0.5),
        "CRM TANLASH MEZONLARI", bold=True, sz=11, color=NAVY)

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 9 — KALIBRASIYA VA LOD/LOQ
# ═══════════════════════════════════════════════════════════════════
def slide9():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("KALIBRASIYA EGRI CHIZIG'I VA LOD/LOQ",
                  "ISO/IEC 17025 Cl.7.3 | FDA BMV 2018 | SANTE/11813 | ISO 11843", 9)
    footer("LOD = 3.3×Sy/b  |  LOQ = 10×Sy/b  |  R²≥0.999 (LC-MS/MS)  |  Back-calc. dev.≤15% at LOQ, ≤10% iboshqa nuqtalarda")

    # Left: calibration types
    box(emu(0.3), emu(3.2), emu(11.5), emu(0.65),
        "KALIBRASIYA MODELI TANLASH", fill=BLUE, sz=12, bold=True, align="ctr")
    cal_rows = [
        ["Mezon", "Linear (1/x²)", "Non-linear 4PL", "Weighted Lin."],
        ["Qachon", "LC-MS/MS\nGC-MS (tor diapazon)", "Immunoassay\nELISA, CLIA", "Geteroscedastic\nkeng diapazon"],
        ["R² talabi", "≥ 0.999", "≥ 0.995", "≥ 0.999"],
        ["Residual", "≤±10% (≤15% LOQ)", "≤±20% (~LOQ)", "≤±15%"],
        ["Og'irlik", "1/x² (asosiy)", "1/y² tavsiya", "1/x yoki 1/x²"],
        ["Tekshirish", "Back-calc farq%", "4PL parametrlar", "SSRR minimum"],
    ]
    col_w4 = [emu(3.0), emu(2.8), emu(2.8), emu(3.0)]
    table(emu(0.3), emu(3.85), emu(11.5), cal_rows, col_w4,
          header_fill=NAVY, alt_fill=LBLUE)

    # Right: LOD/LOQ methods
    box(emu(12.2), emu(3.2), emu(12.4), emu(0.65),
        "LOD / LOQ HISOBLASH METODLARI", fill=RED, sz=12, bold=True, align="ctr")
    lod_rows = [
        ["Metod", "Formula", "Qachon qo'llanadi"],
        ["Kalibrasiya\negri chizig'i", "LOD=3.3×Sy/b\nLOQ=10×Sy/b", "LC-MS/MS, GC-MS\n(asosiy metod)"],
        ["Signal/Shovqin\n(S/N)", "LOD: S/N=3\nLOQ: S/N=10", "Xromatografiya\n(amaliy)"],
        ["Blank spike\n(ISO 11843)", "LOD=x̄blank\n+3×s_blank", "Biologik matritsa\n(eng qat'iy)"],
        ["Aniqlik profili\n(CV%)", "LOQ: CV≤20%\n+Recovery 80-120%", "Validatsiya uchun\n(tasdiqlovchi)"],
    ]
    col_w5 = [emu(3.5), emu(4.2), emu(4.7)]
    table(emu(12.2), emu(3.85), emu(12.4), lod_rows, col_w5,
          header_fill=RED, alt_fill="FFF0E8")

    # Bottom acceptance criteria
    box(emu(0.3), emu(7.95), W - emu(0.6), emu(0.65),
        "Kalibrasiya qabul mezonlari:  R²≥0.999 (MS)  |  ≥2/3 kalibr nuqtalari ±10% ichida  |  "
        "LOQ ≤ 50%×MRL  |  IS CV < 15%",
        fill=NAVY, sz=11, bold=True, align="ctr")

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  SLIDE 10 — INTEGRATSIYA: TO'LIQ METROLOGIK AMALIYOT
# ═══════════════════════════════════════════════════════════════════
def slide10():
    reset()
    box(0, 0, W, H, "", fill="F7FBFF")
    header_banner("TO'LIQ METROLOGIK SIFAT TIZIMI — INTEGRATSIYA",
                  "ISO/IEC 17025:2017 | ILAC-G8 | GUM | SANTE — Kundalik amaliyot jadvali", 10)
    footer("Muvaffaqiyatli akkreditatsiya = Kuzatilish + Noaniqlik + Qaror qoidasi + IQC + Hujjatlashtirish — barchasi bir vaqtda")

    # Workflow table
    rows = [
        ["Bosqich", "Amal", "Standart", "KPI / Qabul mezoni", "Hujjat"],
        ["1. Kalibrasiya",
         "Egri chiziq qurish\nLOD/LOQ aniqlash\nResidual tahlil",
         "ISO 17025\nFDA BMV 2018",
         "R²≥0.999\nCV≤15% (IS)\nBack-calc ≤10%",
         "Kalibrasiya jurnali\nEgri chiziq fayli"],
        ["2. IQC",
         "Kunlik QC o'lchash\nLevey-Jennings\nWestgard 6 qoida",
         "ISO 17025 Cl.7.7\nILAC-G24",
         "CV≤10% (LC-MS)\nCV≤15% (ELISA)\nWestgard alert",
         "IQC jadvali\nWestgard chart"],
        ["3. Noaniqlik",
         "Top-down IQC dan\nhisoblash, Budget\nu_c va U (k=2)",
         "GUM JCGM 100\nEURACHEM CG4",
         "U≤50%×MRL\nTUR≥4\nYillik qayta baholash",
         "MU budget fayli\nHisobot"],
        ["4. Kuzatilish",
         "CRM tekshirish\nPT ishtirok\nZanjir hujjatlashtirish",
         "ISO 17025 Cl.6.6\nILAC-P10\nISO 13528",
         "z-score |z|≤2\nEn≤1.0\nCRM: ±2U ichida",
         "CRM sertifikati\nPT hisoboti"],
        ["5. Qaror qabul",
         "Guard Band w=U\nILAC-G8 qoidasi\nTUR hisoblash",
         "ILAC-G8:2019\nISO 17025 Cl.7.8.6",
         "TUR≥4\nNoaniqlik zonasi\nhisobotda ko'rsatilsin",
         "Qaror qoidasi\nhujjati (SOP)"],
        ["6. Hujjatlashtirish",
         "Barcha qadam\nSOP ga yozish\nAudit trail",
         "ISO 17025 Cl.7.5\nILAC-P9",
         "0 ta kuzatilish\nbog'liq kamchilik\n(akkreditatsiya audit)",
         "LIMS / Lab. daftar\nSOP to'plami"],
    ]
    col_w = [emu(3.5), emu(5.5), emu(4.0), emu(5.5), emu(6.0)]
    table(emu(0.3), emu(3.15), W - emu(0.6), rows, col_w,
          header_fill=NAVY, alt_fill=LBLUE)

    return slide_xml()


# ═══════════════════════════════════════════════════════════════════
#  PPTX ASSEMBLER
# ═══════════════════════════════════════════════════════════════════
SLIDE_BUILDERS = [slide1, slide2, slide3, slide4, slide5,
                  slide6, slide7, slide8, slide9, slide10]

PRES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  saveSubsetFonts="1">
  <p:sldMasterIdLst/>
  <p:sldSz cx="9144000" cy="5143500" type="custom"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:sldIdLst>
SLIDE_ID_LIST
  </p:sldIdLst>
</p:presentation>"""

PRES_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
SLIDE_RELS
</Relationships>"""

SLIDE_REL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

CT_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
SLIDE_CT
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>"""


def build_pptx(path):
    n = len(SLIDE_BUILDERS)

    # Build slide XMLs
    slides_xml = [fn() for fn in SLIDE_BUILDERS]

    # Presentation XML
    sid_list = "\n".join(
        f'    <p:sldId id="{256+i}" r:id="rId{i+1}"/>' for i in range(n))
    pres = PRES_XML.replace("SLIDE_ID_LIST", sid_list)

    # Presentation rels
    slide_rels_lines = "\n".join(
        f'  <Relationship Id="rId{i+1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i+1}.xml"/>' for i in range(n))
    pres_rels = PRES_RELS.replace("SLIDE_RELS", slide_rels_lines)

    # Content types
    slide_ct = "\n".join(
        f'  <Override PartName="/ppt/slides/slide{i+1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(n))
    ct = CT_TEMPLATE.replace("SLIDE_CT", slide_ct)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("ppt/presentation.xml", pres)
        zf.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        for i, sx in enumerate(slides_xml):
            zf.writestr(f"ppt/slides/slide{i+1}.xml", sx)
            zf.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", SLIDE_REL)

    import os
    size = os.path.getsize(path)
    print(f"✅ PPTX yaratildi: {path}")
    print(f"   Slaydlar: {n} ta")
    print(f"   Hajmi: {size:,} bayt ({size/1024:.1f} KB)")


if __name__ == "__main__":
    out = "/projects/sandbox/metrologiya-uchun/Metrologik_Noaniqlik_Qaror.pptx"
    build_pptx(out)
