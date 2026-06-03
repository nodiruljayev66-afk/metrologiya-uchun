"""
Metrologik QA Tizimi - Excel Generator
ISO/IEC 17025 | ILAC-G8 | GUM | SANTE
"""
import zipfile, io, os, math
from xml.etree.ElementTree import Element, SubElement, tostring

# ─── COLOUR PALETTE ───────────────────────────────────────────────
C_HEADER   = "1F3864"   # dark navy
C_SUBHDR   = "2E75B6"   # blue
C_ACCENT   = "00B0F0"   # light blue
C_GREEN    = "70AD47"
C_YELLOW   = "FFD966"
C_RED      = "FF0000"
C_ORANGE   = "F4B942"
C_WHITE    = "FFFFFF"
C_LGRAY    = "F2F2F2"
C_DGRAY    = "D9D9D9"
C_LABEL    = "D6E4F0"

# ─── SHARED STRINGS REGISTRY ──────────────────────────────────────
_sst = []
_sst_map = {}

def si(text):
    t = str(text)
    if t not in _sst_map:
        _sst_map[t] = len(_sst)
        _sst.append(t)
    return _sst_map[t]

# ─── STYLE REGISTRY ───────────────────────────────────────────────
_fonts = []
_fills = []
_borders = []
_numfmts = []
_xfs = []

def reg_font(bold=False, sz=11, color="000000", name="Calibri", italic=False):
    f = (bold, sz, color, name, italic)
    if f not in _fonts:
        _fonts.append(f)
    return _fonts.index(f)

def reg_fill(fgColor=None):
    if fgColor is None:
        f = ("none", None)
    else:
        f = ("solid", fgColor)
    if f not in _fills:
        _fills.append(f)
    return _fills.index(f)

def reg_border(style="thin"):
    if style not in _borders:
        _borders.append(style)
    return _borders.index(style)

def reg_numfmt(fmt_str):
    if fmt_str not in _numfmts:
        _numfmts.append(fmt_str)
    return _numfmts.index(fmt_str)

def reg_xf(font_id=0, fill_id=0, border_id=0, numfmt_id=0,
           halign="general", valign="bottom", wrap=False):
    x = (font_id, fill_id, border_id, numfmt_id, halign, valign, wrap)
    if x not in _xfs:
        _xfs.append(x)
    return _xfs.index(x)


# ─── PRE-REGISTER BASE STYLES ─────────────────────────────────────
reg_font()                                          # 0: default
reg_fill()                                          # 0: none
reg_fill("FFFFFF")                                  # 1: white
reg_border("none")                                  # 0: none
reg_border("thin")                                  # 1: thin

# style IDs (computed after full registration at build time)
# We use helper below to get XF index with all params

def make_style(bold=False, sz=11, color="000000", bg=None,
               halign="general", wrap=False, numfmt="", italic=False):
    fi = reg_font(bold=bold, sz=sz, color=color, italic=italic)
    fli = reg_fill(bg) if bg else 0
    bi = 1  # thin border
    ni = 0
    if numfmt:
        ni = reg_numfmt(numfmt)
    return reg_xf(font_id=fi, fill_id=fli, border_id=bi,
                  numfmt_id=ni, halign=halign, wrap=wrap)

# ─── WORKSHEET BUILDER ────────────────────────────────────────────
class Sheet:
    def __init__(self, name):
        self.name = name
        self.rows = {}          # {row: {col: (value, style, type)}}
        self.merges = []        # list of "A1:D1" strings
        self.col_widths = {}    # {col_idx: width}
        self.row_heights = {}   # {row_idx: height}
        self.freeze = None      # "B2" etc

    def write(self, row, col, value, style=0, typ="s"):
        if row not in self.rows:
            self.rows[row] = {}
        self.rows[row][col] = (value, style, typ)

    def write_str(self, row, col, value, style=0):
        self.write(row, col, si(value), style, "s")

    def write_num(self, row, col, value, style=0):
        self.write(row, col, value, style, "n")

    def write_formula(self, row, col, formula, style=0):
        self.write(row, col, formula, style, "f")

    def merge(self, r1, c1, r2, c2):
        def col_letter(n):
            s = ""
            while n >= 0:
                s = chr(65 + n % 26) + s
                n = n // 26 - 1
            return s
        self.merges.append(
            f"{col_letter(c1)}{r1+1}:{col_letter(c2)}{r2+1}"
        )

    def set_col_width(self, col, width):
        self.col_widths[col] = width

    def set_row_height(self, row, height):
        self.row_heights[row] = height


# ─── XML GENERATORS ───────────────────────────────────────────────
def col_letter(n):
    s = ""
    while n >= 0:
        s = chr(65 + n % 26) + s
        n = n // 26 - 1
    return s

def sheet_xml(sheet):
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

    if sheet.freeze:
        lines.append('<sheetViews><sheetView tabSelected="1" workbookViewId="0">')
        lines.append(f'<pane ySplit="1" xSplit="0" topLeftCell="{sheet.freeze}" activePane="bottomLeft" state="frozen"/>')
        lines.append('</sheetView></sheetViews>')

    lines.append('<sheetFormatPr defaultRowHeight="15"/>')

    if sheet.col_widths:
        lines.append('<cols>')
        for ci, w in sorted(sheet.col_widths.items()):
            lines.append(f'<col min="{ci+1}" max="{ci+1}" width="{w}" customWidth="1"/>')
        lines.append('</cols>')

    lines.append('<sheetData>')
    for ri in sorted(sheet.rows.keys()):
        rh = sheet.row_heights.get(ri, "")
        rh_attr = f' ht="{rh}" customHeight="1"' if rh else ""
        lines.append(f'<row r="{ri+1}"{rh_attr}>')
        for ci in sorted(sheet.rows[ri].keys()):
            val, sty, typ = sheet.rows[ri][ci]
            ref = f"{col_letter(ci)}{ri+1}"
            if typ == "s":
                lines.append(f'<c r="{ref}" t="s" s="{sty}"><v>{val}</v></c>')
            elif typ == "n":
                lines.append(f'<c r="{ref}" t="n" s="{sty}"><v>{val}</v></c>')
            elif typ == "f":
                lines.append(f'<c r="{ref}" s="{sty}"><f>{val}</f></c>')
        lines.append('</row>')
    lines.append('</sheetData>')

    if sheet.merges:
        lines.append('<mergeCells>')
        for m in sheet.merges:
            lines.append(f'<mergeCell ref="{m}"/>')
        lines.append('</mergeCells>')

    lines.append('</worksheet>')
    return "\n".join(lines)


def styles_xml():
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

    # numFmts
    custom_fmts = [(i+164, f) for i,f in enumerate(_numfmts)]
    if custom_fmts:
        lines.append(f'<numFmts count="{len(custom_fmts)}">')
        for fid, fs in custom_fmts:
            fs_esc = fs.replace('"', '&quot;')
            lines.append(f'<numFmt numFmtId="{fid}" formatCode="{fs_esc}"/>')
        lines.append('</numFmts>')

    # fonts
    lines.append(f'<fonts count="{len(_fonts)}">')
    for (bold, sz, color, name, italic) in _fonts:
        lines.append('<font>')
        if bold: lines.append('<b/>')
        if italic: lines.append('<i/>')
        lines.append(f'<sz val="{sz}"/>')
        lines.append(f'<color rgb="FF{color}"/>')
        lines.append(f'<name val="{name}"/>')
        lines.append('</font>')
    lines.append('</fonts>')

    # fills (first two must be none/gray per spec)
    all_fills = [("none", None), ("gray125", None)] + _fills
    lines.append(f'<fills count="{len(all_fills)}">')
    for (pt, fg) in all_fills:
        lines.append('<fill>')
        if fg:
            lines.append(f'<patternFill patternType="solid"><fgColor rgb="FF{fg}"/></patternFill>')
        else:
            lines.append(f'<patternFill patternType="{pt}"/>')
        lines.append('</fill>')
    lines.append('</fills>')

    # borders
    lines.append(f'<borders count="{len(_borders)}">')
    for b in _borders:
        if b == "none":
            lines.append('<border><left/><right/><top/><bottom/><diagonal/></border>')
        else:
            lines.append(f'<border><left style="{b}"><color auto="1"/></left>'
                         f'<right style="{b}"><color auto="1"/></right>'
                         f'<top style="{b}"><color auto="1"/></top>'
                         f'<bottom style="{b}"><color auto="1"/></bottom>'
                         f'<diagonal/></border>')
    lines.append('</borders>')

    # cellStyleXfs
    lines.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')

    # cellXfs
    lines.append(f'<cellXfs count="{len(_xfs)}">')
    for (fi, fli, bi, ni, ha, va, wrap) in _xfs:
        fid = ni + 164 if ni > 0 else 0
        fill_id = fli + 2  # offset for the 2 required fills
        wa = ' wrapText="1"' if wrap else ""
        lines.append(f'<xf numFmtId="{fid}" fontId="{fi}" fillId="{fill_id}" borderId="{bi}" xfId="0"'
                     f' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
                     f'<alignment horizontal="{ha}" vertical="{va}"{wa}/></xf>')
    lines.append('</cellXfs>')
    lines.append('</styleSheet>')
    return "\n".join(lines)

def sst_xml():
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append(f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 f' count="{len(_sst)}" uniqueCount="{len(_sst)}">')
    for s in _sst:
        esc = str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        lines.append(f'<si><t xml:space="preserve">{esc}</t></si>')
    lines.append('</sst>')
    return "\n".join(lines)


# ─── WORKBOOK XML ─────────────────────────────────────────────────
def workbook_xml(sheets):
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    lines.append('<sheets>')
    for i, s in enumerate(sheets):
        lines.append(f'<sheet name="{s.name}" sheetId="{i+1}" r:id="rId{i+1}"/>')
    lines.append('</sheets>')
    lines.append('</workbook>')
    return "\n".join(lines)

def workbook_rels(sheets):
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for i, s in enumerate(sheets):
        lines.append(f'<Relationship Id="rId{i+1}" '
                     f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                     f'Target="worksheets/sheet{i+1}.xml"/>')
    lines.append(f'<Relationship Id="rId{len(sheets)+1}" '
                 f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                 f'Target="sharedStrings.xml"/>')
    lines.append(f'<Relationship Id="rId{len(sheets)+2}" '
                 f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                 f'Target="styles.xml"/>')
    lines.append('</Relationships>')
    return "\n".join(lines)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
SHEET_OVERRIDES
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


# ═══════════════════════════════════════════════════════════════════
#  STYLE SHORTCUTS  (built lazily so indices are stable at call time)
# ═══════════════════════════════════════════════════════════════════
def S(bold=False, sz=11, color="000000", bg=None,
      halign="general", wrap=False, numfmt="", italic=False):
    return make_style(bold=bold, sz=sz, color=color, bg=bg,
                      halign=halign, wrap=wrap, numfmt=numfmt, italic=italic)

# ─── HELPER: section header row ───────────────────────────────────
def sec_header(sh, row, col, text, ncols, bg=C_HEADER, fc=C_WHITE, sz=12):
    sty = S(bold=True, sz=sz, color=fc, bg=bg, halign="center")
    sh.write_str(row, col, text, sty)
    sh.merge(row, col, row, col+ncols-1)
    sh.set_row_height(row, 20)

def sub_header(sh, row, col, text, ncols, bg=C_SUBHDR, fc=C_WHITE, sz=11):
    sty = S(bold=True, sz=sz, color=fc, bg=bg, halign="center")
    sh.write_str(row, col, text, sty)
    sh.merge(row, col, row, col+ncols-1)
    sh.set_row_height(row, 18)

def label(sh, row, col, text, bg=C_LABEL):
    sh.write_str(row, col, text, S(bold=True, bg=bg, halign="left"))

def note(sh, row, col, text):
    sh.write_str(row, col, text, S(italic=True, color="595959", sz=9))

# ═══════════════════════════════════════════════════════════════════
#  SHEET 1: DASHBOARD
# ═══════════════════════════════════════════════════════════════════
def build_dashboard():
    sh = Sheet("📊 DASHBOARD")
    sh.set_col_width(0, 3)
    sh.set_col_width(1, 28)
    sh.set_col_width(2, 18)
    sh.set_col_width(3, 18)
    sh.set_col_width(4, 18)
    sh.set_col_width(5, 22)

    # Title block
    sec_header(sh, 0, 1, "METROLOGIK SIFAT NAZORATI TIZIMI", 5, bg=C_HEADER, sz=16)
    sec_header(sh, 1, 1, "ISO/IEC 17025 | ILAC-G8 | GUM | SANTE | EA-4/02", 5,
               bg=C_SUBHDR, sz=11)
    sh.set_row_height(0, 30)

    # Lab info
    sh.set_row_height(3, 18)
    sub_header(sh, 3, 1, "LABORATORIYA MA'LUMOTLARI", 5, bg="17375E")
    fields = [
        ("Laboratoriya nomi:", ""),
        ("Akkreditatsiya №:", ""),
        ("Mas'ul shaxs:", ""),
        ("Sana:", ""),
        ("Uskunalar:", "IC | GC | GC-MS | LC | HPLC-MS | ICP-MS"),
    ]
    for i, (lbl, val) in enumerate(fields):
        r = 4 + i
        sh.write_str(r, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_str(r, 2, val, S(bg=C_LGRAY))
        sh.merge(r, 2, r, 5)

    # Module navigator
    r = 10
    sec_header(sh, r, 1, "MODULLAR — QISQACHA NAVIGATOR", 5, bg="243F60")
    r += 1

    modules = [
        ("📈 M1: Kalibrasiya",     "Linear/Non-linear, R², LOD/LOQ, Residual, 1/x² weighting", "M1_KALIB"),
        ("📉 M2: Statistika",       "SD, CV%, Bias, Recovery, Grubbs, Dixon, Shapiro-Wilk",     "M2_STAT"),
        ("🔍 M3: IQC Monitoring",  "Levey-Jennings, Westgard qoidalari (6 qoida), Control limits","M3_IQC"),
        ("📐 M4: MU / Noaniqlik",  "Top-down, Uncertainty budget, Expanded U (k=2), GUM",       "M4_MU"),
        ("⚖️ M5: Conformity",      "Guard band, ILAC-G8, TUR, False conformance risk",          "M5_CONF"),
        ("🌐 M6: Interlaboratory", "z-score, En number, Robust statistics (median, MAD)",       "M6_INTER"),
    ]
    colors = [C_GREEN, C_SUBHDR, "7030A0", "C55A11", C_RED, "00B0F0"]
    sh.write_str(r, 1, "Modul", S(bold=True, bg=C_HEADER, color=C_WHITE, halign="center"))
    sh.write_str(r, 2, "Funksiya", S(bold=True, bg=C_HEADER, color=C_WHITE, halign="center"))
    sh.merge(r, 2, r, 4)
    sh.write_str(r, 5, "Sheet nomi", S(bold=True, bg=C_HEADER, color=C_WHITE, halign="center"))
    r += 1

    for i, (mod, desc, sname) in enumerate(modules):
        c = colors[i]
        sh.write_str(r+i, 1, mod,   S(bold=True, bg=c, color=C_WHITE))
        sh.write_str(r+i, 2, desc,  S(bg=C_LGRAY, wrap=True))
        sh.merge(r+i, 2, r+i, 4)
        sh.write_str(r+i, 5, sname, S(color="1F78B4", halign="center"))
        sh.set_row_height(r+i, 20)

    # Standards reference
    r2 = r + len(modules) + 2
    sec_header(sh, r2, 1, "QOʻLLANILGAN STANDARTLAR VA HUJJATLAR", 5, bg="1F3864")
    standards = [
        ("ISO/IEC 17025:2017", "Sinov va kalibrlash laboratoriyalari kompetentsiyasi"),
        ("ILAC-G8:09/2019",    "Muvofiqlik bayonotlari uchun qaror qoidalari"),
        ("GUM (JCGM 100:2008)","O'lchov noaniqligini ifodalash qo'llanmasi"),
        ("EURACHEM/CITAC CG4", "Analitik kimyoda o'lchov noaniqligini aniqlash"),
        ("ISO 5725-2:2019",    "Aniqlik va haqiqiylik — takrorlanish metodlari"),
        ("SANTE/11813/2017",   "Pestitsid qoldiqlarini tahlil qilish sifat nazorati"),
        ("EA-4/02 M:2022",     "Kalibrlash laboratoriyalarida noaniqlikni ifodalash"),
        ("FDA BMV 2018",       "Bioanalitik metod validatsiyasi — aniqlik talablari"),
    ]
    r2 += 1
    sh.write_str(r2, 1, "Standart", S(bold=True, bg=C_SUBHDR, color=C_WHITE))
    sh.write_str(r2, 2, "Tavsif",   S(bold=True, bg=C_SUBHDR, color=C_WHITE))
    sh.merge(r2, 2, r2, 5)
    for i, (std, desc) in enumerate(standards):
        bg = C_LGRAY if i%2==0 else C_WHITE
        sh.write_str(r2+1+i, 1, std,  S(bold=True, bg=bg))
        sh.write_str(r2+1+i, 2, desc, S(bg=bg))
        sh.merge(r2+1+i, 2, r2+1+i, 5)

    return sh


# ═══════════════════════════════════════════════════════════════════
#  SHEET 2: M1 — KALIBRASIYA
# ═══════════════════════════════════════════════════════════════════
def build_calibration():
    sh = Sheet("M1_KALIB")
    for c,w in enumerate([3,22,16,16,16,16,16,20]):
        sh.set_col_width(c, w)

    sec_header(sh, 0, 1, "MODUL 1 — KALIBRASIYA TAHLILI", 6, sz=14)
    sub_header(sh, 1, 1, "ISO/IEC 17025 Cl.7.3 | SANTE/11813 | FDA BMV 2018", 6, bg="1F4E79")

    # ── Input table ──────────────────────────────────────────
    r = 3
    sub_header(sh, r, 1, "KIRISH MA'LUMOTLARI — KALIBRASIYA NUQTALARI", 6, bg=C_SUBHDR)
    r += 1
    hdrs = ["№", "Konsentrasiya (x)\n[ng/mL]", "Signal / Area (y₁)",
            "Signal / Area (y₂)", "Signal / Area (y₃)", "O'rtacha ȳ", "RSD (%)"]
    for ci, h in enumerate(hdrs):
        sh.write_str(r, ci+1, h, S(bold=True, bg="17375E", color=C_WHITE,
                                   halign="center", wrap=True, sz=10))
        sh.set_row_height(r, 30)

    concs = [1, 5, 10, 25, 50, 100, 250, 500]
    for i, c in enumerate(concs):
        rr = r+1+i
        sh.write_num(rr, 1, i+1,  S(halign="center", bg=C_LGRAY if i%2==0 else C_WHITE))
        sh.write_num(rr, 2, c,    S(halign="center", bg=C_LGRAY if i%2==0 else C_WHITE, numfmt="0.000"))
        for ci in [3,4,5]:
            sh.write_str(rr, ci, "", S(bg="EBF3FB"))
        # Average formula
        sh.write_formula(rr, 6, f"AVERAGE(D{rr+1}:F{rr+1})",
                         S(numfmt="0.0000", bg=C_YELLOW, halign="center"))
        # RSD%
        sh.write_formula(rr, 7,
            f'IF(AVERAGE(D{rr+1}:F{rr+1})=0,"—",STDEV(D{rr+1}:F{rr+1})/AVERAGE(D{rr+1}:F{rr+1})*100)',
            S(numfmt="0.00", bg=C_LGRAY, halign="center"))

    # ── Results block ─────────────────────────────────────────
    r2 = r + len(concs) + 2
    sub_header(sh, r2, 1, "REGRESSION NATIJALARI (Formulalar avtomatik)", 6, bg="375623")
    r2 += 1
    results = [
        ("Slope (b):",        f"=SLOPE(G{r+2}:G{r+1+len(concs)},C{r+2}:C{r+1+len(concs)})", "0.00000E+00"),
        ("Intercept (a):",    f"=INTERCEPT(G{r+2}:G{r+1+len(concs)},C{r+2}:C{r+1+len(concs)})", "0.00000E+00"),
        ("R² (Pearson²):",    f"=RSQ(G{r+2}:G{r+1+len(concs)},C{r+2}:C{r+1+len(concs)})", "0.000000"),
        ("R (Pearson):",      f"=CORREL(G{r+2}:G{r+1+len(concs)},C{r+2}:C{r+1+len(concs)})", "0.000000"),
        ("n (nuqtalar):",     f"=COUNT(C{r+2}:C{r+1+len(concs)})", "0"),
        ("Sy (residual SD):", f"=STEYX(G{r+2}:G{r+1+len(concs)},C{r+2}:C{r+1+len(concs)})", "0.00000E+00"),
    ]
    for i, (lbl, frm, fmt) in enumerate(results):
        sh.write_str(r2+i, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_formula(r2+i, 2, frm, S(numfmt=fmt, bg=C_YELLOW, halign="center"))
        sh.merge(r2+i, 2, r2+i, 3)

    # LOD / LOQ
    r3 = r2 + len(results) + 1
    sub_header(sh, r3, 1, "LOD / LOQ HISOBLASH  (Kalibrasiya egri chizig'i metodi)", 6, bg="833C00")
    r3 += 1
    sh.write_str(r3,   1, "LOD = 3.3 × Sy / b",  S(bold=True, bg=C_LABEL))
    sh.write_formula(r3, 2, f"=3.3*{col_letter(2)}{r2+6+1}/{col_letter(2)}{r2+1+1}",
                     S(numfmt="0.000", bg=C_GREEN, color=C_WHITE, halign="center"))
    sh.write_str(r3,   3, "ng/mL", S(italic=True))
    sh.write_str(r3+1, 1, "LOQ = 10 × Sy / b",   S(bold=True, bg=C_LABEL))
    sh.write_formula(r3+1, 2, f"=10*{col_letter(2)}{r2+6+1}/{col_letter(2)}{r2+1+1}",
                     S(numfmt="0.000", bg=C_ORANGE, halign="center"))
    sh.write_str(r3+1, 3, "ng/mL", S(italic=True))

    # Acceptance criteria note
    r4 = r3 + 3
    note(sh, r4, 1, "✔ R² ≥ 0.999 talab qilinadi (LC-MS/MS, GC-MS). "
                    "R² ≥ 0.995 qabul qilinadi (IC, HPLC-UV). "
                    "Back-calculated deviation ≤ 15% at LOQ, ≤ 10% at all other points (FDA BMV 2018).")
    sh.merge(r4, 1, r4, 7)
    note(sh, r4+1, 1, "✔ Agar fan-shaped residual aniqlansa — 1/x yoki 1/x² weighting qo'llang. "
                      "SSRR (sum of squared relative residuals) minimumini tanlang.")
    sh.merge(r4+1, 1, r4+1, 7)

    # Residual analysis table
    r5 = r4 + 3
    sub_header(sh, r5, 1, "RESIDUAL TAHLIL — Back-calculated qiymatlar", 6, bg="2E4057")
    r5 += 1
    res_hdrs = ["№", "x (nom.)", "ȳ (o'lchangan)", "x_back = (ȳ−a)/b", "Residual %", "Status"]
    for ci, h in enumerate(res_hdrs):
        sh.write_str(r5, ci+1, h, S(bold=True, bg=C_SUBHDR, color=C_WHITE, halign="center"))
    for i in range(len(concs)):
        rr = r5+1+i
        bg = C_LGRAY if i%2==0 else C_WHITE
        sh.write_num(rr, 1, i+1, S(halign="center", bg=bg))
        sh.write_formula(rr, 2, f"=C{r+2+i}", S(numfmt="0.000", bg=bg, halign="center"))
        sh.write_formula(rr, 3, f"=G{r+2+i}", S(numfmt="0.0000", bg=bg, halign="center"))
        sh.write_formula(rr, 4,
            f"=IF(C{r2+1+1}=0,\"\",({col_letter(3)}{rr+1}-{col_letter(2)}{r2+2+1})/{col_letter(2)}{r2+1+1})",
            S(numfmt="0.000", bg=C_YELLOW, halign="center"))
        sh.write_formula(rr, 5,
            f"=IF(C{r+2+i}=0,\"\",({col_letter(4)}{rr+1}-C{r+2+i})/C{r+2+i}*100)",
            S(numfmt="0.00\"%\"", bg=C_LGRAY, halign="center"))
        sh.write_formula(rr, 6,
            f'=IF(F{rr+1}="","—",IF(ABS(F{rr+1})<=10,"✔ OK",IF(ABS(F{rr+1})<=15,"⚠ LOQ?","✘ REJECT")))',
            S(halign="center", bg=bg))

    sh.freeze = "B2"
    return sh


# ═══════════════════════════════════════════════════════════════════
#  SHEET 3: M2 — STATISTIKA
# ═══════════════════════════════════════════════════════════════════
def build_statistics():
    sh = Sheet("M2_STAT")
    for c,w in enumerate([3,24,14,14,14,14,14,18,14]):
        sh.set_col_width(c, w)

    sec_header(sh, 0, 1, "MODUL 2 — STATISTIK TAHLIL", 7, sz=14)
    sub_header(sh, 1, 1, "ISO 5725-2:2019 | EURACHEM CG4 | FDA BMV 2018 | Dixon & Grubbs", 7, bg="1F4E79")

    # ── Input ─────────────────────────────────────────────────
    r = 3
    sub_header(sh, r, 1, "NATIJALAR KIRISH (maksimal 20 ta o'lchov)", 7, bg=C_SUBHDR)
    r += 1
    hdrs = ["№","O'lchov natijasi\n[birlik]","","","","","","",""]
    sh.write_str(r, 1, "№", S(bold=True, bg="17375E", color=C_WHITE, halign="center"))
    sh.write_str(r, 2, "O'lchov natijasi [birlik / ng·mL⁻¹ / mg·kg⁻¹]",
                 S(bold=True, bg="17375E", color=C_WHITE, halign="center"))
    sh.merge(r, 2, r, 8)
    sh.set_row_height(r, 22)

    # 20 input rows in 2 columns of 10
    for i in range(10):
        rr = r+1+i
        sh.write_num(rr, 1, i+1, S(halign="center", bg=C_LGRAY))
        sh.write_str(rr, 2, "", S(bg="EBF3FB"))
        sh.write_num(rr, 4, i+11, S(halign="center", bg=C_LGRAY))
        sh.write_str(rr, 5, "", S(bg="EBF3FB"))

    # ── Statistics ────────────────────────────────────────────
    r2 = r + 12
    sub_header(sh, r2, 1, "ASOSIY STATISTIKA (Avtomatik hisoblash)", 7, bg="375623")
    r2 += 1

    data_range = f"C{r+2}:C{r+11},F{r+2}:F{r+11}"
    flat = f"C{r+2}:C{r+11}"   # first column only for some formulas
    full = f"(C{r+2}:C{r+11},F{r+2}:F{r+11})"

    stats = [
        ("n (namunalar soni):",           f"=COUNT(C{r+2}:C{r+11})+COUNT(F{r+2}:F{r+11})", "0"),
        ("Xo'rtacha (x̄):",               f"=IFERROR(AVERAGE(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.0000"),
        ("Median:",                        f"=IFERROR(MEDIAN(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.0000"),
        ("SD (standart og'ish, s):",       f"=IFERROR(STDEV(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.0000"),
        ("Variance (s²):",                 f"=IFERROR(VAR(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.000000"),
        ("CV% (nisbiy og'ish):",          f"=IFERROR(STDEV(C{r+2}:C{r+11},F{r+2}:F{r+11})/AVERAGE(C{r+2}:C{r+11},F{r+2}:F{r+11})*100,\"\")", "0.00"),
        ("Min:", f"=IFERROR(MIN(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.0000"),
        ("Max:", f"=IFERROR(MAX(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.0000"),
        ("Range (Max−Min):", f"=IFERROR(MAX(C{r+2}:C{r+11},F{r+2}:F{r+11})-MIN(C{r+2}:C{r+11},F{r+2}:F{r+11}),\"\")", "0.0000"),
    ]
    for i, (lbl, frm, fmt) in enumerate(stats):
        sh.write_str(r2+i, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_formula(r2+i, 2, frm, S(numfmt=fmt, bg=C_YELLOW, halign="center"))
        sh.merge(r2+i, 2, r2+i, 3)

    # ── Bias & Recovery ───────────────────────────────────────
    r3 = r2 + len(stats) + 1
    sub_header(sh, r3, 1, "BIAS & RECOVERY HISOBLASH", 7, bg="833C00")
    r3 += 1
    sh.write_str(r3,   1, "Sertifikat / nominal qiymat (μref):", S(bold=True, bg=C_LABEL))
    sh.write_str(r3,   2, "", S(bg="EBF3FB", numfmt="0.0000"))
    sh.merge(r3, 2, r3, 3)
    sh.write_str(r3+1, 1, "Bias (%) = (x̄ − μref)/μref × 100:", S(bold=True, bg=C_LABEL))
    sh.write_formula(r3+1, 2,
        f'=IFERROR(({col_letter(2)}{r2+2+1}-C{r3+1+1})/C{r3+1+1}*100,"")',
        S(numfmt="0.00", bg=C_YELLOW, halign="center"))
    sh.merge(r3+1, 2, r3+1, 3)
    sh.write_str(r3+2, 1, "Recovery (%) = x̄/μref × 100:", S(bold=True, bg=C_LABEL))
    sh.write_formula(r3+2, 2,
        f'=IFERROR({col_letter(2)}{r2+2+1}/C{r3+1+1}*100,"")',
        S(numfmt="0.00", bg=C_YELLOW, halign="center"))
    sh.merge(r3+2, 2, r3+2, 3)
    sh.write_str(r3+3, 1, "Status (Bias ≤ ±5%?):", S(bold=True, bg=C_LABEL))
    sh.write_formula(r3+3, 2,
        f'=IFERROR(IF(ABS(C{r3+2+1})<=5,"✔ QABUL (≤5%)","⚠ TEKSHIRISH KERAK (>5%)"),"Qiymat kiriting")',
        S(bg=C_LGRAY, halign="center"))
    sh.merge(r3+3, 2, r3+3, 3)

    # ── Grubbs Test ───────────────────────────────────────────
    r4 = r3 + 5
    sub_header(sh, r4, 1, "GRUBBS TESTI — Outlier aniqlash (ISO 5725-2)", 7, bg="2E4057")
    r4 += 1
    note(sh, r4, 1, "G_hisob = |x_shubhali − x̄| / s    →    G_hisob > G_krit (α=0.05) bo'lsa — outlier")
    sh.merge(r4, 1, r4, 8)
    r4 += 1
    sh.write_str(r4,   1, "Shubhali qiymat (x_shubhali):", S(bold=True, bg=C_LABEL))
    sh.write_str(r4,   2, "", S(bg="EBF3FB"))
    sh.write_str(r4+1, 1, "G_hisob:", S(bold=True, bg=C_LABEL))
    sh.write_formula(r4+1, 2,
        f'=IFERROR(ABS(C{r4+1+1}-{col_letter(2)}{r2+2+1})/{col_letter(2)}{r2+4+1},"")',
        S(numfmt="0.0000", bg=C_YELLOW, halign="center"))
    sh.write_str(r4+2, 1, "G_krit (n=10, α=0.05) = 2.290:", S(bold=True, bg=C_LABEL))
    sh.write_num(r4+2,  2, 2.290, S(numfmt="0.000", bg=C_LGRAY, halign="center"))
    sh.write_str(r4+3, 1, "Xulosa:", S(bold=True, bg=C_LABEL))
    sh.write_formula(r4+3, 2,
        f'=IFERROR(IF(C{r4+2+1}>C{r4+3+1},"⚠ OUTLIER — chiqarib tashlang","✔ Normal — outlier emas"),"Qiymat kiriting")',
        S(bg=C_LGRAY, halign="center"))
    sh.merge(r4+3, 2, r4+3, 4)
    note(sh, r4+5, 1,
         "Dixon Q-test (n≤7): Q = |x_shubhali − x_yaqin| / Range. Q_krit(n=7,α=0.05)=0.570. "
         "Agar ma'lumotlar normal taqsilanmagan bo'lsa — Hampel: median ± 3×MAD ishlating.")
    sh.merge(r4+5, 1, r4+5, 8)

    sh.freeze = "B2"
    return sh


# ═══════════════════════════════════════════════════════════════════
#  SHEET 4: M3 — IQC MONITORING (Levey-Jennings + Westgard)
# ═══════════════════════════════════════════════════════════════════
def build_iqc():
    sh = Sheet("M3_IQC")
    widths = [3,12,14,14,14,14,14,14,14,20]
    for c,w in enumerate(widths):
        sh.set_col_width(c, w)

    sec_header(sh, 0, 1, "MODUL 3 — IQC MONITORING (Levey-Jennings / Westgard)", 8, sz=14)
    sub_header(sh, 1, 1, "ISO/IEC 17025 Cl.7.7 | Westgard 6 qoidasi | ILAC-G24", 8, bg="1F4E79")

    # Control limits setup
    r = 3
    sub_header(sh, r, 1, "NAZORAT CHEGARALARI SOZLASH", 8, bg="17375E")
    r += 1
    sh.write_str(r,   1, "Analitik nomi:", S(bold=True, bg=C_LABEL))
    sh.write_str(r,   2, "", S(bg="EBF3FB"))
    sh.merge(r, 2, r, 4)
    sh.write_str(r+1, 1, "Birlik:", S(bold=True, bg=C_LABEL))
    sh.write_str(r+1, 2, "", S(bg="EBF3FB"))
    sh.write_str(r+2, 1, "X̄ (target mean):", S(bold=True, bg=C_LABEL))
    sh.write_str(r+2, 2, "", S(bg="EBF3FB", numfmt="0.0000"))
    sh.write_str(r+3, 1, "SD (nazorat SD):", S(bold=True, bg=C_LABEL))
    sh.write_str(r+3, 2, "", S(bg="EBF3FB", numfmt="0.0000"))

    # Limit formulas
    limits = [
        ("+3s UCL (Harakatlar chegarasi):", f"=C{r+3+1}+3*C{r+4+1}", C_RED),
        ("+2s UWL (Ogohlantirish):",        f"=C{r+3+1}+2*C{r+4+1}", C_ORANGE),
        ("+1s:",                             f"=C{r+3+1}+1*C{r+4+1}", C_YELLOW),
        ("Target (X̄):",                     f"=C{r+3+1}",             C_GREEN),
        ("−1s:",                             f"=C{r+3+1}-1*C{r+4+1}", C_YELLOW),
        ("−2s LWL (Ogohlantirish):",        f"=C{r+3+1}-2*C{r+4+1}", C_ORANGE),
        ("−3s LCL (Harakatlar chegarasi):", f"=C{r+3+1}-3*C{r+4+1}", C_RED),
    ]
    r5 = r + 5
    for i, (lbl, frm, col) in enumerate(limits):
        sh.write_str(r5+i, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_formula(r5+i, 2, frm, S(numfmt="0.0000", bg=col,
                                          color=C_WHITE if col in [C_RED,"1F4E79"] else "000000",
                                          halign="center"))

    # ── IQC Data table ────────────────────────────────────────
    r6 = r5 + len(limits) + 2
    sub_header(sh, r6, 1, "IQC NATIJALAR JADVALI (30 kun)", 8, bg=C_SUBHDR)
    r6 += 1
    hdrs2 = ["№","Sana","IQC qiymati","z-score\n(x−X̄)/s",
             "+3s","+2s","−2s","−3s","WESTGARD STATUS"]
    for ci, h in enumerate(hdrs2):
        sh.write_str(r6, ci+1, h, S(bold=True, bg="17375E", color=C_WHITE,
                                    halign="center", wrap=True))
        sh.set_row_height(r6, 28)

    mean_ref  = f"$C${r+3+1}"
    sd_ref    = f"$C${r+4+1}"
    ucl_ref   = f"$C${r5+1}"
    uwl_ref   = f"$C${r5+2}"
    lwl_ref   = f"$C${r5+6}"
    lcl_ref   = f"$C${r5+7}"

    for i in range(30):
        rr = r6+1+i
        bg = C_LGRAY if i%2==0 else C_WHITE
        sh.write_num(rr,  1, i+1, S(halign="center", bg=bg))
        sh.write_str(rr,  2, "",  S(bg="EBF3FB"))
        sh.write_str(rr,  3, "",  S(bg="EBF3FB", numfmt="0.0000"))
        # z-score
        sh.write_formula(rr, 4,
            f'=IFERROR((D{rr+1}-{mean_ref})/{sd_ref},"")',
            S(numfmt="0.00", bg=C_LGRAY, halign="center"))
        # Limit references
        sh.write_formula(rr, 5, f"={ucl_ref}", S(numfmt="0.0000", bg="FFE0E0", halign="center"))
        sh.write_formula(rr, 6, f"={uwl_ref}", S(numfmt="0.0000", bg="FFF0CC", halign="center"))
        sh.write_formula(rr, 7, f"={lwl_ref}", S(numfmt="0.0000", bg="FFF0CC", halign="center"))
        sh.write_formula(rr, 8, f"={lcl_ref}", S(numfmt="0.0000", bg="FFE0E0", halign="center"))
        # Westgard status
        sh.write_formula(rr, 9,
            f'=IFERROR(IF(D{rr+1}="","—",'
            f'IF(OR(D{rr+1}>={ucl_ref},D{rr+1}<={lcl_ref}),"🔴 13s HARAKATLAR",'
            f'IF(OR(D{rr+1}>={uwl_ref},D{rr+1}<={lwl_ref}),"🟡 12s OGOHLANTIRISH",'
            f'"✅ NAZORAT ICHIDA"))),"—")',
            S(halign="center", bg=bg))

    # Westgard rules explanation
    r7 = r6 + 32
    sub_header(sh, r7, 1, "WESTGARD QOIDALARI — QISQACHA", 8, bg="2E4057")
    r7 += 1
    rules = [
        ("1₂s",  "1 ta natija ±2s dan tashqarida — OGOHLANTIRISH (tekshiring)"),
        ("1₃s",  "1 ta natija ±3s dan tashqarida — HARAKATLAR (to'xtating)"),
        ("2₂s",  "Ketma-ket 2 ta natija bir tomonda ±2s — HARAKATLAR (sistemik xato)"),
        ("R₄s",  "1 partiya ichida max−min > 4s — HARAKATLAR (tasodifiy xato)"),
        ("4₁s",  "Ketma-ket 4 ta natija bir tomonda ±1s — OGOHLANTIRISH (trend)"),
        ("10ₓ",  "Ketma-ket 10 ta natija X̄ ning bir tomonida — HARAKATLAR (drift)"),
    ]
    sh.write_str(r7, 1, "Qoida", S(bold=True, bg=C_SUBHDR, color=C_WHITE, halign="center"))
    sh.write_str(r7, 2, "Ta'rif",S(bold=True, bg=C_SUBHDR, color=C_WHITE))
    sh.merge(r7, 2, r7, 9)
    for i, (rule, desc) in enumerate(rules):
        bg = C_LGRAY if i%2==0 else C_WHITE
        sh.write_str(r7+1+i, 1, rule, S(bold=True, bg=bg, halign="center"))
        sh.write_str(r7+1+i, 2, desc, S(bg=bg))
        sh.merge(r7+1+i, 2, r7+1+i, 9)

    sh.freeze = "B2"
    return sh


# ═══════════════════════════════════════════════════════════════════
#  SHEET 5: M4 — O'LCHOV NOANIQLIK (MU)
# ═══════════════════════════════════════════════════════════════════
def build_mu():
    sh = Sheet("M4_MU")
    for c,w in enumerate([3,28,16,16,16,14,14,18]):
        sh.set_col_width(c, w)

    sec_header(sh, 0, 1, "MODUL 4 — O'LCHOV NOANIQLIK (Measurement Uncertainty)", 6, sz=14)
    sub_header(sh, 1, 1, "GUM JCGM 100:2008 | EURACHEM/CITAC CG4 | EA-4/02 M:2022", 6, bg="1F4E79")

    # Method info
    r = 3
    sub_header(sh, r, 1, "METOD VA ANALITIK MA'LUMOTLAR", 6, bg="17375E")
    r += 1
    info_fields = [
        "Analitik / Analit nomi:",
        "Metod (ISO/EN/BS/SANTE):",
        "Matriks (siydik/qon/suv/o'simlik):",
        "Birlik (ng/mL, mg/kg, μg/L):",
        "Hisoblash metodi (Top-down / Bottom-up):",
    ]
    for i, f in enumerate(info_fields):
        sh.write_str(r+i, 1, f, S(bold=True, bg=C_LABEL))
        sh.write_str(r+i, 2, "", S(bg="EBF3FB"))
        sh.merge(r+i, 2, r+i, 4)

    # ── TOP-DOWN (IQC based) ──────────────────────────────────
    r2 = r + len(info_fields) + 2
    sub_header(sh, r2, 1, "TOP-DOWN YONDASHUV (IQC ma'lumotlaridan)", 6, bg="375623")
    r2 += 1
    note(sh, r2, 1, "u_combined = √[ u²(repeatability) + u²(reproducibility) + u²(bias) ]   "
                    "→   U = k × u_combined    (k=2, ~95% ishonch)")
    sh.merge(r2, 1, r2, 7)
    r2 += 1

    td_fields = [
        ("u_repeatability  (sr — kunlik tekrorlash SD):",        "0.0000",  C_LABEL),
        ("u_reproducibility (sI — kunlararo intermediate SD):",  "0.0000",  C_LABEL),
        ("u_bias  (|Bias%|/100 × x̄):",                          "0.0000",  C_LABEL),
        ("u_combined = √(u_r² + u_R² + u_b²):",                 "0.00000", C_YELLOW),
        ("Relative u_combined (%):",                             "0.00",    C_YELLOW),
        ("Expanded U  (k=2, ~95%):",                             "0.00000", C_GREEN),
        ("Relative U% (k=2):",                                   "0.00",    C_GREEN),
    ]
    row_refs = {}
    for i, (lbl, fmt, bg) in enumerate(td_fields):
        rr = r2+i
        row_refs[i] = rr
        sh.write_str(rr, 1, lbl, S(bold=True, bg=C_LABEL))
        if i < 3:
            sh.write_str(rr, 2, "", S(bg="EBF3FB", numfmt=fmt))
        elif i == 3:
            sh.write_formula(rr, 2,
                f"=IFERROR(SQRT(C{r2+1+1}^2+C{r2+2+1}^2+C{r2+3+1}^2),\"\")",
                S(numfmt=fmt, bg=bg, halign="center"))
        elif i == 4:
            sh.write_formula(rr, 2,
                f'=IFERROR(C{r2+4+1}/AVERAGE(C{r+1+1}:C{r+1+1})*100,"")',
                S(numfmt=fmt, bg=bg, halign="center"))
        elif i == 5:
            sh.write_formula(rr, 2,
                f"=IFERROR(2*C{r2+4+1},\"\")",
                S(numfmt=fmt, bg=bg, color=C_WHITE, halign="center"))
        elif i == 6:
            sh.write_formula(rr, 2,
                f"=IFERROR(2*C{r2+5+1},\"\")",
                S(numfmt=fmt, bg=bg, color=C_WHITE, halign="center"))
        sh.merge(rr, 2, rr, 3)

    # ── UNCERTAINTY BUDGET ────────────────────────────────────
    r3 = r2 + len(td_fields) + 2
    sub_header(sh, r3, 1, "BOTTOM-UP — NOANIQLIK BYUDJETI", 6, bg="833C00")
    r3 += 1
    bhdrs = ["Noaniqlik manbai","Qiymat (a)","Taqsimot\n(N/R/T)","Divisor","u(xi)","u(xi)² / uc² (%)","Izoh"]
    for ci, h in enumerate(bhdrs):
        sh.write_str(r3, ci+1, h, S(bold=True, bg="17375E", color=C_WHITE,
                                    halign="center", wrap=True))
        sh.set_row_height(r3, 28)

    sources = [
        ("Kalibrlash standartining sofligi","0.005","R","1.7321","","","CRM sertifikatidan"),
        ("Tarozilay tortish og'ishi","0.0001","R","1.7321","","","Tosh sertifikatidan"),
        ("Hajmiy idish (pipetka/kolba)","0.05","R","1.7321","","","Sertifikat/rul"),
        ("Ekstraksiya unumi (Recovery)","","N","1","","","Validatsiya SD"),
        ("Matritsa ta'siri (ion suppression)","","N","1","","","IS monitoring"),
        ("Asbob kalibratsiyasi","","N","1","","","IQC SD"),
        ("Harorat ta'siri","0.002","R","1.7321","","","Texnik ma'lumotnoma"),
    ]
    for i, row_data in enumerate(sources):
        rr = r3+1+i
        bg = C_LGRAY if i%2==0 else C_WHITE
        for ci, val in enumerate(row_data):
            if ci == 4:  # u(xi) = a/divisor
                sh.write_formula(rr, ci+1,
                    f'=IFERROR(IF(C{rr+1}="","",C{rr+1}/E{rr+1}),"")',
                    S(numfmt="0.000000", bg=C_YELLOW, halign="center"))
            elif ci == 5:  # contribution %
                sh.write_formula(rr, ci+1, '""', S(bg=C_LGRAY, halign="center"))
            else:
                sh.write_str(rr, ci+1, val, S(bg=bg))

    # Combined from budget
    r_comb = r3 + len(sources) + 1
    sh.write_str(r_comb, 1, "uc (combined) = √Σu(xi)²:", S(bold=True, bg=C_LABEL))
    sh.write_formula(r_comb, 2,
        f"=IFERROR(SQRT(SUMPRODUCT(F{r3+2}:F{r3+1+len(sources)}^2)),\"\")",
        S(numfmt="0.000000", bg=C_GREEN, color=C_WHITE, halign="center"))
    sh.write_str(r_comb+1, 1, "U = 2 × uc (k=2, ~95%):", S(bold=True, bg=C_LABEL))
    sh.write_formula(r_comb+1, 2,
        f"=IFERROR(2*C{r_comb+1},\"\")",
        S(numfmt="0.000000", bg=C_GREEN, color=C_WHITE, halign="center"))

    # Result reporting format
    r4 = r_comb + 3
    sub_header(sh, r4, 1, "NATIJA TAQDIMOT FORMATI (ISO/IEC 17025 Cl.7.8.3)", 6, bg="2E4057")
    r4 += 1
    note(sh, r4, 1,
         "Tavsiya etilgan format: \"[Analitik]: [x̄] [birlik] ± [U] [birlik]; "
         "U k=2 da hisoblangan (~95% ishonch oralig'i); "
         "NIST/LGC/BAM CRM ga nisbatan kuzatilishi ta'minlangan.\"")
    sh.merge(r4, 1, r4, 7)
    note(sh, r4+1, 1,
         "Muvofiqlik shartidir: U ≤ (1/3) × tolerans. "
         "Agar U > 1/3 × tolerans — metod maqsad uchun yaroqsiz, optimallashtirish zarur.")
    sh.merge(r4+1, 1, r4+1, 7)

    sh.freeze = "B2"
    return sh


# ═══════════════════════════════════════════════════════════════════
#  SHEET 6: M5 — CONFORMITY ASSESSMENT (Guard Band / ILAC-G8)
# ═══════════════════════════════════════════════════════════════════
def build_conformity():
    sh = Sheet("M5_CONF")
    for c,w in enumerate([3,26,16,16,16,16,18]):
        sh.set_col_width(c, w)

    sec_header(sh, 0, 1, "MODUL 5 — MUVOFIQLIK BAHOLASH VA QAROR QOIDALARI", 5, sz=14)
    sub_header(sh, 1, 1, "ILAC-G8:09/2019 | ISO/IEC 17025:2017 Cl.7.8.6 | ASME B89.7.3.1", 5, bg="1F4E79")

    # Inputs
    r = 3
    sub_header(sh, r, 1, "KIRISH PARAMETRLARI", 5, bg="17375E")
    r += 1

    inputs = [
        ("O'lchangan natija (x):",           "0.0000", "EBF3FB"),
        ("Kengaytirilgan noaniqlik U (k=2):", "0.0000", "EBF3FB"),
        ("Maqbullik chegarasi / Cutoff (TL):","0.0000", "EBF3FB"),
        ("Tolerans kengligi (±δ):",           "0.0000", "EBF3FB"),
    ]
    for i, (lbl, fmt, bg) in enumerate(inputs):
        sh.write_str(r+i, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_str(r+i, 2, "",  S(bg=bg, numfmt=fmt))
        sh.merge(r+i, 2, r+i, 3)

    x_ref  = f"$C${r+1}"
    U_ref  = f"$C${r+2}"
    TL_ref = f"$C${r+3}"
    tol_ref= f"$C${r+4}"

    # ── Guard Band calculation ────────────────────────────────
    r2 = r + len(inputs) + 2
    sub_header(sh, r2, 1, "GUARD BAND HISOBLASH (ILAC-G8)", 5, bg="375623")
    r2 += 1

    gb_rows = [
        ("Guard Band w = U:",                   f"={U_ref}",                         "0.0000", C_YELLOW),
        ("Qabul qilish zonasi yuqori chegarasi:",f"={TL_ref}+{U_ref}",               "0.0000", C_ORANGE),
        ("Qabul qilish zonasi quyi chegarasi:",  f"={TL_ref}-{U_ref}",               "0.0000", C_ORANGE),
        ("Qat'iy MUSBAT zona (≥ TL+U):",        f"={TL_ref}+{U_ref}",               "0.0000", C_RED),
        ("Noaniqlik zonasi [TL−U ; TL+U]:",     f"=\"[\"&TEXT({TL_ref}-{U_ref},\"0.00\")&\" ; \"&TEXT({TL_ref}+{U_ref},\"0.00\")&\"]\"",
                                                                                      "@",      C_YELLOW),
        ("Qat'iy MANFIY zona (≤ TL−U):",        f"={TL_ref}-{U_ref}",               "0.0000", C_GREEN),
    ]
    for i, (lbl, frm, fmt, bg) in enumerate(gb_rows):
        sh.write_str(r2+i, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_formula(r2+i, 2, frm, S(numfmt=fmt, bg=bg, halign="center"))
        sh.merge(r2+i, 2, r2+i, 3)

    # ── TUR ───────────────────────────────────────────────────
    r3 = r2 + len(gb_rows) + 2
    sub_header(sh, r3, 1, "TUR — Test Uncertainty Ratio", 5, bg="833C00")
    r3 += 1
    sh.write_str(r3,   1, "TUR = δ / U  (ASME B89.7.3.1):", S(bold=True, bg=C_LABEL))
    sh.write_formula(r3, 2,
        f'=IFERROR({tol_ref}/{U_ref},"")',
        S(numfmt="0.00", bg=C_YELLOW, halign="center"))
    sh.merge(r3, 2, r3, 3)
    sh.write_str(r3+1, 1, "TUR baholash:", S(bold=True, bg=C_LABEL))
    sh.write_formula(r3+1, 2,
        f'=IFERROR(IF(C{r3+1}>=4,"✔ TUR≥4: Ishonchli qaror",'
        f'IF(C{r3+1}>=2,"⚠ TUR 2-4: Ehtiyotkor bo\'ling",'
        f'"🔴 TUR<2: Qaror ishonchsiz — metodni yaxshilang")),"")',
        S(bg=C_LGRAY, halign="center"))
    sh.merge(r3+1, 2, r3+1, 5)

    # ── Conformity Decision ───────────────────────────────────
    r4 = r3 + 3
    sub_header(sh, r4, 1, "AVTOMATIK QAROR — ILAC-G8 Qat'iy Guard Band", 5, bg="2E4057")
    r4 += 1
    sh.write_str(r4, 1, "QAROR:", S(bold=True, sz=13, bg=C_LABEL))
    sh.write_formula(r4, 2,
        f'=IFERROR(IF({x_ref}="","Qiymat kiriting",'
        f'IF({x_ref}>={TL_ref}+{U_ref},"🔴 MUSBAT — Chegaradan yuqori (U hisobga olingan)",'
        f'IF({x_ref}<={TL_ref}-{U_ref},"✅ MANFIY — Chegaradan past (U hisobga olingan)",'
        f'"🟡 NOANIQLIK ZONASI — Qo\'shimcha test zarur"))),"—")',
        S(bold=True, sz=12, bg=C_LGRAY, halign="center"))
    sh.merge(r4, 2, r4, 6)
    sh.set_row_height(r4, 26)

    # ── Decision rules comparison ─────────────────────────────
    r5 = r4 + 3
    sub_header(sh, r5, 1, "QAROR QOIDALARI TAQQOSLASH (ILAC-G8:2019)", 5, bg="1F3864")
    r5 += 1
    rules_tbl = [
        ("Oddiy qabul",             "Noaniqlik e'tiborga olinmaydi",           "Qonuniy jihatdan zaif"),
        ("Taqsimlangan risk",        "x > TL → Musbat (U bo'linmaydi)",        "Yolg'on ijobiy/salbiy ehtimoli teng"),
        ("Qat'iy Guard Band (w=U)", "x ≥ TL+U → Musbat; x ≤ TL−U → Manfiy", "Yolg'on ijobiy minimal — TAVSIYA"),
        ("Yumshoq Guard Band",       "x ≥ TL−U → Musbat (klinik uchun)",      "Yolg'on manfiy minimal"),
    ]
    sh.write_str(r5, 1, "Qoida", S(bold=True, bg=C_SUBHDR, color=C_WHITE))
    sh.write_str(r5, 2, "Qo'llash",S(bold=True, bg=C_SUBHDR, color=C_WHITE))
    sh.write_str(r5, 3, "Risk",    S(bold=True, bg=C_SUBHDR, color=C_WHITE))
    sh.merge(r5, 3, r5, 5)
    for i, (rule, appl, risk) in enumerate(rules_tbl):
        bg = C_LGRAY if i%2==0 else C_WHITE
        sh.write_str(r5+1+i, 1, rule, S(bold=True, bg=bg))
        sh.write_str(r5+1+i, 2, appl, S(bg=bg))
        sh.write_str(r5+1+i, 3, risk, S(bg=bg, italic=True))
        sh.merge(r5+1+i, 3, r5+1+i, 5)

    sh.freeze = "B2"
    return sh


# ═══════════════════════════════════════════════════════════════════
#  SHEET 7: M6 — INTERLABORATORY COMPARISON
# ═══════════════════════════════════════════════════════════════════
def build_interlaboratory():
    sh = Sheet("M6_INTER")
    for c,w in enumerate([3,24,14,14,14,14,14,14,16]):
        sh.set_col_width(c, w)

    sec_header(sh, 0, 1, "MODUL 6 — LABORATORIYALARARO TAQQOSLASH", 7, sz=14)
    sub_header(sh, 1, 1, "ISO 13528:2022 | ILAC-P9 | ISO/IEC 17043:2023 | En number", 7, bg="1F4E79")

    # ── z-score section ───────────────────────────────────────
    r = 3
    sub_header(sh, r, 1, "Z-SCORE HISOBLASH (Laboratoriyamiz natijasi)", 7, bg="17375E")
    r += 1
    note(sh, r, 1,
         "z = (x_lab − x_ref) / σ_p     "
         "|z| ≤ 2.0 = Qoniqarli  |  2.0 < |z| ≤ 3.0 = Ogohlantirish  |  |z| > 3.0 = Muvaffaqiyatsiz")
    sh.merge(r, 1, r, 8)
    r += 1

    z_inputs = [
        ("Laboratoriyamiz natijasi (x_lab):",    "EBF3FB", "0.0000"),
        ("Mos yozuvlar qiymati (x_ref):",        "EBF3FB", "0.0000"),
        ("σ_p (PT o'rnatilgan SD):",             "EBF3FB", "0.0000"),
        ("z-score = (x_lab − x_ref) / σ_p:",    C_YELLOW,  "0.00"),
        ("|z| baholash:",                         C_LGRAY,   "@"),
    ]
    for i, (lbl, bg, fmt) in enumerate(z_inputs):
        sh.write_str(r+i, 1, lbl, S(bold=True, bg=C_LABEL))
        if i < 3:
            sh.write_str(r+i, 2, "", S(bg=bg, numfmt=fmt))
        elif i == 3:
            sh.write_formula(r+i, 2,
                f'=IFERROR((C{r+1+1}-C{r+2+1})/C{r+3+1},"")',
                S(numfmt=fmt, bg=bg, halign="center"))
        else:
            sh.write_formula(r+i, 2,
                f'=IFERROR(IF(ABS(C{r+4+1})<=2,"✅ QONIQARLI (|z|≤2)",'
                f'IF(ABS(C{r+4+1})<=3,"⚠ OGOHLANTIRISH (2<|z|≤3)","🔴 MUVAFFAQIYATSIZ (|z|>3)")),"—")',
                S(bg=bg, halign="center"))
        sh.merge(r+i, 2, r+i, 4)

    # ── En number ────────────────────────────────────────────
    r2 = r + len(z_inputs) + 2
    sub_header(sh, r2, 1, "En SONI HISOBLASH (Kalibrlash laboratoriyalari — EA-4/02)", 7, bg="375623")
    r2 += 1
    note(sh, r2, 1,
         "En = (x_lab − x_ref) / √(U_lab² + U_ref²)     "
         "|En| ≤ 1.0 = Qoniqarli  |  |En| > 1.0 = Muvaffaqiyatsiz")
    sh.merge(r2, 1, r2, 8)
    r2 += 1

    en_inputs = [
        ("x_lab (laboratoriyamiz natijasi):", "EBF3FB", "0.0000"),
        ("U_lab (kengaytirilgan noaniqlik):", "EBF3FB", "0.0000"),
        ("x_ref (mos yozuvlar natijasi):",    "EBF3FB", "0.0000"),
        ("U_ref (mos yozuvlar noaniqlik):",   "EBF3FB", "0.0000"),
        ("En = (x_lab−x_ref)/√(U_lab²+U_ref²):", C_YELLOW, "0.0000"),
        ("|En| baholash:",                     C_LGRAY, "@"),
    ]
    for i, (lbl, bg, fmt) in enumerate(en_inputs):
        sh.write_str(r2+i, 1, lbl, S(bold=True, bg=C_LABEL))
        if i < 4:
            sh.write_str(r2+i, 2, "", S(bg=bg, numfmt=fmt))
        elif i == 4:
            sh.write_formula(r2+i, 2,
                f'=IFERROR((C{r2+1+1}-C{r2+3+1})/SQRT(C{r2+2+1}^2+C{r2+4+1}^2),"")',
                S(numfmt=fmt, bg=bg, halign="center"))
        else:
            sh.write_formula(r2+i, 2,
                f'=IFERROR(IF(ABS(C{r2+5+1})<=1,"✅ QONIQARLI (|En|≤1)","🔴 MUVAFFAQIYATSIZ (|En|>1)"),"—")',
                S(bg=bg, halign="center"))
        sh.merge(r2+i, 2, r2+i, 4)

    # ── Multi-lab comparison table ────────────────────────────
    r3 = r2 + len(en_inputs) + 2
    sub_header(sh, r3, 1, "KO'P LABORATORIYA TAQQOSLASH JADVALI", 7, bg="2E4057")
    r3 += 1
    ml_hdrs = ["№","Laboratoriya","Natija","Noaniqlik U","z-score","En soni","Xulosa"]
    for ci, h in enumerate(ml_hdrs):
        sh.write_str(r3, ci+1, h, S(bold=True, bg="17375E", color=C_WHITE, halign="center"))

    # Robust statistics header
    note(sh, r3+1, 1, "Quyidagi 15 qator — har bir laboratoriya uchun natija kiriting:")
    sh.merge(r3+1, 1, r3+1, 8)

    for i in range(15):
        rr = r3+2+i
        bg = C_LGRAY if i%2==0 else C_WHITE
        sh.write_num(rr, 1, i+1, S(halign="center", bg=bg))
        sh.write_str(rr, 2, "", S(bg="EBF3FB"))
        sh.write_str(rr, 3, "", S(bg="EBF3FB", numfmt="0.0000"))
        sh.write_str(rr, 4, "", S(bg="EBF3FB", numfmt="0.0000"))
        # z-score for each lab
        sh.write_formula(rr, 5,
            f'=IFERROR((D{rr+1}-MEDIAN($D${r3+3}:$D${r3+17}))/MAD_placeholder,"—")',
            S(numfmt="0.00", bg=C_LGRAY, halign="center"))
        sh.write_str(rr, 5, "=(D"+str(rr+1)+"-xref)/σp — kiriting", S(bg=C_LGRAY, halign="center"))
        sh.write_str(rr, 6, "", S(bg="EBF3FB", numfmt="0.0000"))
        sh.write_formula(rr, 7,
            f'=IF(E{rr+1}="","—",IF(ABS(E{rr+1})<=2,"✅","IF(ABS(E)<=3,\"⚠\",\"🔴\")"))',
            S(bg=bg, halign="center"))

    # Robust statistics block
    r4 = r3 + 18
    sub_header(sh, r4, 1, "ROBUST STATISTIKA (ISO 13528 — Algorithm A)", 7, bg="1F3864")
    r4 += 1
    rob = [
        ("Median (robust o'rtacha):",  f"=IFERROR(MEDIAN(D{r3+3}:D{r3+17}),\"\")", "0.0000"),
        ("MAD = Median|xi−Median|:",   '""', "0.0000"),
        ("Robust SD (s*) = 1.4826×MAD:", f'=IFERROR(1.4826*C{r4+2+1},"")', "0.0000"),
        ("n (ishtirokchilar):",          f"=COUNTA(B{r3+3}:B{r3+17})", "0"),
    ]
    for i, (lbl, frm, fmt) in enumerate(rob):
        sh.write_str(r4+i, 1, lbl, S(bold=True, bg=C_LABEL))
        sh.write_formula(r4+i, 2, frm, S(numfmt=fmt, bg=C_YELLOW, halign="center"))
        sh.merge(r4+i, 2, r4+i, 4)

    note(sh, r4+len(rob)+1, 1,
         "ISO 13528 Algorithm A: Robust mean va SD iterativ hisoblash orqali outlierlardan himoyalangan. "
         "z-score uchun σ_p = robust SD yoki PT organizer tomonidan belgilangan qiymat ishlating.")
    sh.merge(r4+len(rob)+1, 1, r4+len(rob)+1, 8)

    sh.freeze = "B2"
    return sh


# ═══════════════════════════════════════════════════════════════════
#  MAIN BUILD — assemble all sheets into .xlsx
# ═══════════════════════════════════════════════════════════════════
def build_xlsx(path):
    sheets = [
        build_dashboard(),
        build_calibration(),
        build_statistics(),
        build_iqc(),
        build_mu(),
        build_conformity(),
        build_interlaboratory(),
    ]

    # Build all sheet XMLs first to populate SST and styles
    sheet_xmls = [sheet_xml(s) for s in sheets]

    # Content types with sheet overrides
    overrides = ""
    for i in range(len(sheets)):
        overrides += (f'  <Override PartName="/xl/worksheets/sheet{i+1}.xml" '
                      f'ContentType="application/vnd.openxmlformats-officedocument'
                      f'.spreadsheetml.worksheet+xml"/>\n')
    ct = CONTENT_TYPES.replace("SHEET_OVERRIDES", overrides)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("xl/workbook.xml", workbook_xml(sheets))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels(sheets))
        zf.writestr("xl/sharedStrings.xml", sst_xml())
        zf.writestr("xl/styles.xml", styles_xml())
        for i, (s, sx) in enumerate(zip(sheets, sheet_xmls)):
            zf.writestr(f"xl/worksheets/sheet{i+1}.xml", sx)

    print(f"✅ Excel fayl yaratildi: {path}")
    print(f"   Sahifalar: {[s.name for s in sheets]}")

if __name__ == "__main__":
    out = "/projects/sandbox/metrologiya-uchun/Metrologik_QA_Tizimi.xlsx"
    build_xlsx(out)
    import os
    size = os.path.getsize(out)
    print(f"   Hajmi: {size:,} bayt ({size/1024:.1f} KB)")
