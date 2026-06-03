"""
PPTX fixer — slideMaster + slideLayout qo'shadi
Shu ikkisi bo'lmasa PowerPoint ochishni rad etadi
"""
import zipfile, os, shutil

# ── Minimal slideMaster XML ────────────────────────────────────────
SLIDE_MASTER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgRef idx="1001">
    <a:schemeClr val="bg1"/>
  </p:bgRef></p:bg>
  <p:spTree>
    <p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm><a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="0" cy="0"/></a:xfrm>
    </p:grpSpPr>
  </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
    accent1="accent1" accent2="accent2" accent3="accent3"
    accent4="accent4" accent5="accent5" accent6="accent6"
    hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr><a:defRPr lang="uz-UZ" sz=2800 b="1">
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:latin typeface="Calibri"/>
      </a:defRPr></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr lang="uz-UZ" sz=1800>
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
        <a:latin typeface="Calibri"/>
      </a:defRPr></a:lvl1pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:defPPr><a:defRPr lang="uz-UZ">
        <a:latin typeface="Calibri"/>
      </a:defRPr></a:defPPr>
    </p:otherStyle>
  </p:txStyles>
</p:sldMaster>'''

SLIDE_MASTER_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''

# ── Minimal slideLayout XML ────────────────────────────────────────
SLIDE_LAYOUT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/>
        <a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>'''

SLIDE_LAYOUT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''

# ── theme XML ─────────────────────────────────────────────────────
THEME = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5A96C5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''

THEME_RELS_MASTER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    Target="../theme/theme1.xml"/>
</Relationships>'''


def fix_pptx(src, dst=None):
    if dst is None:
        dst = src
    tmp = src + ".tmp"

    with zipfile.ZipFile(src, "r") as zin:
        names = zin.namelist()

        # Read existing content
        files = {}
        for n in names:
            files[n] = zin.read(n)

    # ── Patch presentation.xml ────────────────────────────────────
    pres = files["ppt/presentation.xml"].decode("utf-8")

    # Replace empty sldMasterIdLst with one pointing to our master
    pres = pres.replace(
        "<p:sldMasterIdLst/>",
        '<p:sldMasterIdLst>'
        '<p:sldMasterId id="2147483648" r:id="rIdM1"/>'
        '</p:sldMasterIdLst>'
    )
    files["ppt/presentation.xml"] = pres.encode("utf-8")

    # ── Patch ppt/_rels/presentation.xml.rels ────────────────────
    pres_rels = files["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    n_slides = sum(1 for k in files if k.startswith("ppt/slides/slide") and k.endswith(".xml") and "_rels" not in k)

    master_rel = (
        f'  <Relationship Id="rIdM1" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        f'Target="slideMasters/slideMaster1.xml"/>\n'
    )
    pres_rels = pres_rels.replace("</Relationships>", master_rel + "</Relationships>")
    files["ppt/_rels/presentation.xml.rels"] = pres_rels.encode("utf-8")

    # ── Patch [Content_Types].xml ─────────────────────────────────
    ct = files["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>",
        '  <Override PartName="/ppt/slideMasters/slideMaster1.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\n'
        '  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>\n'
        '  <Override PartName="/ppt/theme/theme1.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n'
        '</Types>'
    )
    files["[Content_Types].xml"] = ct.encode("utf-8")

    # ── Patch each slide XML — add r:id reference to layout ──────
    for i in range(1, n_slides + 1):
        key = f"ppt/slides/slide{i}.xml"
        sld = files[key].decode("utf-8")
        # Add show attribute and layout relationship reference
        if 'r:id="rIdL1"' not in sld:
            sld = sld.replace(
                "<p:cSld>",
                '<p:cSld>',
                1
            )
        files[key] = sld.encode("utf-8")

        # Patch slide rels to reference layout
        rkey = f"ppt/slides/_rels/slide{i}.xml.rels"
        srel = files.get(rkey, b'').decode("utf-8")
        if "slideLayout" not in srel:
            srel = srel.replace("</Relationships>",
                '  <Relationship Id="rIdL1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
                'Target="../slideLayouts/slideLayout1.xml"/>\n'
                '</Relationships>'
            )
            # If empty rels file
            if "<Relationships" not in srel:
                srel = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
                    '  <Relationship Id="rIdL1" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
                    'Target="../slideLayouts/slideLayout1.xml"/>\n'
                    '</Relationships>'
                )
        files[rkey] = srel.encode("utf-8")

    # ── Add new files ─────────────────────────────────────────────
    files["ppt/slideMasters/slideMaster1.xml"]            = SLIDE_MASTER.encode("utf-8")
    files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = SLIDE_MASTER_RELS.encode("utf-8")
    files["ppt/slideLayouts/slideLayout1.xml"]            = SLIDE_LAYOUT.encode("utf-8")
    files["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = SLIDE_LAYOUT_RELS.encode("utf-8")
    files["ppt/theme/theme1.xml"]                         = THEME.encode("utf-8")
    # Add theme rels to master rels (already in SLIDE_MASTER_RELS above — but add theme path)
    files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = (
        SLIDE_MASTER_RELS.replace("</Relationships>",
            '  <Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
            'Target="../theme/theme1.xml"/>\n'
            '</Relationships>'
        )
    ).encode("utf-8")

    # ── Write output ──────────────────────────────────────────────
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    os.replace(tmp, dst)
    size = os.path.getsize(dst)
    print(f"✅ Tuzatildi: {dst}")
    print(f"   Fayl hajmi: {size:,} bayt ({size/1024:.1f} KB)")
    print(f"   Jami fayllar: {len(files)} ta")


if __name__ == "__main__":
    src = "/projects/sandbox/metrologiya-uchun/Metrologik_Noaniqlik_Qaror.pptx"
    fix_pptx(src)
