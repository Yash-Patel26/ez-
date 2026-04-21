#!/usr/bin/env python3
"""
Generate a custom MD2PPTX Slide Master template.

Color palette — modern dark-navy / sky-blue / violet:
  Primary bg  : #0D1B4B  (deep navy)
  Accent 1    : #00C2CB  (cyan-teal)
  Accent 2    : #7C3AED  (violet)
  Accent 3    : #10B981  (emerald)
  Accent 4    : #F59E0B  (amber)
  Accent 5    : #EF4444  (red)
  Accent 6    : #EC4899  (pink)

Run from md2pptx/ directory:
    python create_custom_template.py
Output: templates/Template_MD2PPTX_Custom_Presentation.pptx
"""

from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# ── Palette (hex without #) ──────────────────────────────────────
P = {
    "navy":       "0D1B4B",
    "white":      "FFFFFF",
    "dark_navy":  "1A2A5E",
    "light_bg":   "EEF2F7",
    "cyan":       "00C2CB",
    "violet":     "7C3AED",
    "emerald":    "10B981",
    "amber":      "F59E0B",
    "red":        "EF4444",
    "pink":       "EC4899",
    "mid_navy":   "243570",   # header bar on content slides
    "pale_cyan":  "E0F7FA",   # light accent bg
}

# ── Slide dimensions (13.333" × 7.5", standard widescreen) ───────
W  = Inches(13.333)
H  = Inches(7.5)
NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# ════════════════════════════════════════════════════════════════
# Theme XML
# ════════════════════════════════════════════════════════════════

THEME_XML = f'''\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="{NS}" name="MD2PPTX Navy">
  <a:themeElements>
    <a:clrScheme name="MD2PPTX Palette">
      <a:dk1><a:srgbClr val="{P['navy']}"/></a:dk1>
      <a:lt1><a:srgbClr val="{P['white']}"/></a:lt1>
      <a:dk2><a:srgbClr val="{P['dark_navy']}"/></a:dk2>
      <a:lt2><a:srgbClr val="{P['light_bg']}"/></a:lt2>
      <a:accent1><a:srgbClr val="{P['cyan']}"/></a:accent1>
      <a:accent2><a:srgbClr val="{P['violet']}"/></a:accent2>
      <a:accent3><a:srgbClr val="{P['emerald']}"/></a:accent3>
      <a:accent4><a:srgbClr val="{P['amber']}"/></a:accent4>
      <a:accent5><a:srgbClr val="{P['red']}"/></a:accent5>
      <a:accent6><a:srgbClr val="{P['pink']}"/></a:accent6>
      <a:hlink><a:srgbClr val="{P['cyan']}"/></a:hlink>
      <a:folHlink><a:srgbClr val="{P['violet']}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="MD2PPTX Fonts">
      <a:majorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/><a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/><a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''


# ════════════════════════════════════════════════════════════════
# XML helpers
# ════════════════════════════════════════════════════════════════

PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"


def rgb_to_hex(r, g, b):
    return f"{r:02X}{g:02X}{b:02X}"


def _solid_fill_xml(hex_color: str) -> str:
    return (
        f'<p:bg xmlns:p="{PML}" xmlns:a="{DML}">'
        f'  <p:bgPr>'
        f'    <a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
        f'    <a:effectLst/>'
        f'  </p:bgPr>'
        f'</p:bg>'
    )


def set_layout_name(layout, name: str):
    """Rename a slide layout via XML."""
    cSld = layout._element.find(qn("p:cSld"))
    if cSld is not None:
        cSld.set("name", name)


def set_layout_bg(layout, hex_color: str):
    """Set a solid background color on a slide layout."""
    elem = layout._element
    cSld = elem.find(qn("p:cSld"))
    if cSld is None:
        return

    # Remove existing bg if present
    existing_bg = cSld.find(qn("p:bg"))
    if existing_bg is not None:
        cSld.remove(existing_bg)

    # Build new bg element
    bg_xml = (
        f'<p:bg xmlns:p="{PML}" xmlns:a="{DML}">'
        f'  <p:bgPr>'
        f'    <a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
        f'    <a:effectLst/>'
        f'  </p:bgPr>'
        f'</p:bg>'
    )
    bg_elem = etree.fromstring(bg_xml)

    # Insert bg as first child of cSld (before spTree)
    sp_tree = cSld.find(qn("p:spTree"))
    if sp_tree is not None:
        cSld.insert(list(cSld).index(sp_tree), bg_elem)
    else:
        cSld.insert(0, bg_elem)


def add_rect_to_layout(layout, left, top, width, height, fill_hex, name="rect"):
    """Add a solid colored rectangle to a slide layout's shape tree."""
    cSld = layout._element.find(qn("p:cSld"))
    spTree = cSld.find(qn("p:spTree"))

    shape_xml = f'''\
<p:sp xmlns:p="{PML}" xmlns:a="{DML}"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="100" name="{name}"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{left}" y="{top}"/>
      <a:ext cx="{width}" cy="{height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p/>
  </p:txBody>
</p:sp>'''
    spTree.append(etree.fromstring(shape_xml))


def add_text_to_layout(layout, left, top, width, height,
                       text, font_size, bold, color_hex,
                       align="left", name="lbl"):
    """Add a styled text box to a slide layout."""
    cSld = layout._element.find(qn("p:cSld"))
    spTree = cSld.find(qn("p:spTree"))
    algn_map = {"left": "l", "center": "ctr", "right": "r"}
    algn = algn_map.get(align, "l")
    bold_str = "1" if bold else "0"

    shape_xml = f'''\
<p:sp xmlns:p="{PML}" xmlns:a="{DML}"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="101" name="{name}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{left}" y="{top}"/>
      <a:ext cx="{width}" cy="{height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="{algn}"/>
      <a:r>
        <a:rPr lang="en-US" sz="{font_size}" b="{bold_str}" dirty="0">
          <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
        </a:rPr>
        <a:t>{text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''
    spTree.append(etree.fromstring(shape_xml))


def add_placeholder_to_layout(layout, idx, ph_type, left, top, width, height,
                               font_size=2400, color_hex="FFFFFF", name="ph"):
    """
    Add a placeholder to a layout's spTree.
    ph_type: 1=TITLE, 2=BODY, 13=SLIDE_NUMBER
    """
    cSld = layout._element.find(qn("p:cSld"))
    spTree = cSld.find(qn("p:spTree"))

    ph_xml = f'''\
<p:sp xmlns:p="{PML}" xmlns:a="{DML}"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{200 + idx}" name="{name}"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="{"title" if ph_type == 1 else "body"}" idx="{idx}"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{left}" y="{top}"/>
      <a:ext cx="{width}" cy="{height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="{font_size}" dirty="0">
          <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
        </a:rPr>
        <a:t/>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''
    spTree.append(etree.fromstring(ph_xml))


def clear_layout_shapes(layout):
    """Remove all non-sp-lock shapes from a layout's spTree."""
    cSld = layout._element.find(qn("p:cSld"))
    spTree = cSld.find(qn("p:spTree"))
    for sp in list(spTree.findall(qn("p:sp"))):
        spTree.remove(sp)
    for pic in list(spTree.findall(qn("p:pic"))):
        spTree.remove(pic)
    for grp in list(spTree.findall(qn("p:grpSp"))):
        spTree.remove(grp)


# ════════════════════════════════════════════════════════════════
# EMU helpers (1 inch = 914400 EMU)
# ════════════════════════════════════════════════════════════════

def i(inches): return int(inches * 914400)
def p(pt):     return int(pt * 12700)   # 1 pt = 12700 EMU


# ════════════════════════════════════════════════════════════════
# Build template
# ════════════════════════════════════════════════════════════════

def build_template(out_path: str):
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    master = prs.slide_masters[0]

    # ── 1. Inject custom theme ───────────────────────────────────
    theme_part = None
    for rel in master.part.rels.values():
        if "theme" in rel.reltype.lower():
            theme_part = rel.target_part
            break

    if theme_part:
        theme_part._blob = THEME_XML.encode("utf-8")
    else:
        # Fallback: find any theme part in the package
        for part in prs.part.package.iter_parts():
            if part.partname and "theme" in str(part.partname).lower():
                part._blob = THEME_XML.encode("utf-8")
                break

    # ── 2. Set slide master background (white) ───────────────────
    master_cSld = master._element.find(qn("p:cSld"))
    existing_master_bg = master_cSld.find(qn("p:bg"))
    if existing_master_bg is not None:
        master_cSld.remove(existing_master_bg)
    master_bg_xml = (
        f'<p:bg xmlns:p="{PML}" xmlns:a="{DML}">'
        f'  <p:bgPr>'
        f'    <a:solidFill><a:srgbClr val="{P["white"]}"/></a:solidFill>'
        f'    <a:effectLst/>'
        f'  </p:bgPr>'
        f'</p:bg>'
    )
    master_bg = etree.fromstring(master_bg_xml)
    spTree_m = master_cSld.find(qn("p:spTree"))
    master_cSld.insert(list(master_cSld).index(spTree_m), master_bg)

    layouts = prs.slide_layouts

    # ── 3. Layout 0 — 1_Cover (dark navy, decorative) ────────────
    cover = layouts[0]
    clear_layout_shapes(cover)
    set_layout_name(cover, "1_Cover")
    set_layout_bg(cover, P["navy"])

    # Full-width cyan accent bar at top
    add_rect_to_layout(cover, i(0), i(0), i(13.333), i(0.08), P["cyan"], "top_bar")
    # Left decorative panel (dark navy shade)
    add_rect_to_layout(cover, i(0), i(0), i(0.5), i(7.5), P["mid_navy"], "left_panel")
    # Bottom accent strip
    add_rect_to_layout(cover, i(0), i(7.1), i(13.333), i(0.4), P["mid_navy"], "bottom_strip")
    # Cyan vertical accent line on left
    add_rect_to_layout(cover, i(0.5), i(2.2), i(0.06), i(2.8), P["cyan"], "accent_line")
    # Decorative circle (top-right)
    shape_xml = f'''\
<p:sp xmlns:p="{PML}" xmlns:a="{DML}"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="110" name="deco_circle"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{i(11.5)}" y="{i(-1.0)}"/><a:ext cx="{i(3.5)}" cy="{i(3.5)}"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{P['mid_navy']}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''
    cover._element.find(qn("p:cSld")).find(qn("p:spTree")).append(etree.fromstring(shape_xml))

    # Small cyan circle accent (bottom-right)
    shape_xml2 = f'''\
<p:sp xmlns:p="{PML}" xmlns:a="{DML}"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="111" name="deco_circle2"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{i(10.8)}" y="{i(5.6)}"/><a:ext cx="{i(1.2)}" cy="{i(1.2)}"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{P['cyan']}"><a:alpha val="30000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''
    cover._element.find(qn("p:cSld")).find(qn("p:spTree")).append(etree.fromstring(shape_xml2))

    # Placeholders: title (idx=10) and subtitle (idx=11)
    add_placeholder_to_layout(cover, 10, 2, i(0.7), i(2.5), i(11.8), i(1.2),
                               font_size=3600, color_hex=P["white"], name="Title Placeholder")
    add_placeholder_to_layout(cover, 11, 2, i(0.7), i(4.0), i(9.5), i(0.7),
                               font_size=1800, color_hex=P["cyan"], name="Subtitle Placeholder")

    # ── 4. Layout 1 — Divider (navy, centered label) ──────────────
    divider = layouts[1]
    clear_layout_shapes(divider)
    set_layout_name(divider, "Divider")
    set_layout_bg(divider, P["navy"])

    add_rect_to_layout(divider, i(0), i(0), i(13.333), i(0.08), P["cyan"], "top_bar")
    add_rect_to_layout(divider, i(0), i(7.1), i(13.333), i(0.4), P["mid_navy"], "bottom_bar")
    # Horizontal cyan line in the middle
    add_rect_to_layout(divider, i(2.5), i(3.7), i(8.33), i(0.04), P["cyan"], "mid_line")
    add_placeholder_to_layout(divider, 10, 2, i(1.5), i(2.8), i(10.33), i(1.0),
                               font_size=3200, color_hex=P["white"], name="Section Title")

    # ── 5. Layout 2 — Blank (white, slim navy header) ─────────────
    blank = layouts[2]
    clear_layout_shapes(blank)
    set_layout_name(blank, "Blank")
    set_layout_bg(blank, P["white"])

    add_rect_to_layout(blank, i(0), i(0), i(13.333), i(0.08), P["cyan"], "top_bar")
    add_rect_to_layout(blank, i(0), i(0), i(13.333), i(0.65), P["navy"], "header_bar")
    add_rect_to_layout(blank, i(0), i(7.1), i(13.333), i(0.08), P["mid_navy"], "footer_bar")

    # ── 6. Layout 3 — Title only (white, title placeholder) ───────
    title_only = layouts[3]
    clear_layout_shapes(title_only)
    set_layout_name(title_only, "Title only")
    set_layout_bg(title_only, P["white"])

    add_rect_to_layout(title_only, i(0), i(0), i(13.333), i(0.08), P["cyan"], "top_bar")
    add_rect_to_layout(title_only, i(0), i(0), i(13.333), i(0.65), P["navy"], "header_bar")
    add_rect_to_layout(title_only, i(0), i(7.1), i(13.333), i(0.08), P["mid_navy"], "footer_bar")
    add_placeholder_to_layout(title_only, 0, 1, i(0.375), i(0.08), i(12.58), i(0.57),
                               font_size=2000, color_hex=P["white"], name="Title 1")
    add_placeholder_to_layout(title_only, 11, 2, i(0.375), i(7.2), i(9.5), i(0.25),
                               font_size=1000, color_hex=P["light_bg"], name="Footer")

    # ── 7. Layout 4 — Thank You (navy, baked-in text) ─────────────
    thankyou = layouts[4]
    clear_layout_shapes(thankyou)
    set_layout_name(thankyou, "Thank You")
    set_layout_bg(thankyou, P["navy"])

    add_rect_to_layout(thankyou, i(0), i(0), i(13.333), i(0.08), P["cyan"], "top_bar")
    add_rect_to_layout(thankyou, i(0), i(7.1), i(13.333), i(0.4), P["mid_navy"], "bottom_bar")
    add_rect_to_layout(thankyou, i(3.5), i(3.6), i(6.33), i(0.04), P["cyan"], "mid_line")

    # Baked-in "Thank You" text (renderer detects this)
    add_text_to_layout(thankyou, i(1.5), i(2.3), i(10.33), i(1.4),
                       "Thank You", 5400, True, P["white"], "center", "thankyou_text")
    add_text_to_layout(thankyou, i(1.5), i(3.9), i(10.33), i(0.6),
                       "Questions &amp; Discussion", 2000, False, P["cyan"], "center", "thankyou_sub")

    # ── 8. Remaining layouts — set white bg + header bar ──────────
    for idx in range(5, len(layouts)):
        lay = layouts[idx]
        name = f"Content_{idx}"
        set_layout_name(lay, name)
        set_layout_bg(lay, P["white"])

    # ── 9. Save ───────────────────────────────────────────────────
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    print(f"Template saved -> {out_path}")


if __name__ == "__main__":
    out = "templates/Template_MD2PPTX_Custom_Presentation.pptx"
    build_template(out)
