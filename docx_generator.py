"""
Professional DOCX generator — Pre-Scoping Privacy Questionnaire
100% pure lxml XML. Zero python-docx table/document private APIs.
"""
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree


# ── Embedded Protiviti logo (base64 PNG, white on transparent) ────────────────
#import base64 as _b64
#_LOGO_B64 = "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"
#_LOGO_BYTES = _b64.b64decode(_LOGO_B64)

C_DARK_NAVY  = "1F3864"
C_MID_BLUE   = "2E75B6"
C_LIGHT_BLUE = "D6E4F0"
C_NEAR_WHITE = "F2F7FB"
C_WHITE      = "FFFFFF"
C_TEXT_DARK  = "1A1A2E"
C_TEXT_MID   = "4A5568"
C_GOLD       = "C8973A"
C_BORDER     = "B8CCE4"
FONT         = "Aptos"
FONT_SZ      = 10
W14          = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_SPC      = "{http://www.w3.org/XML/1998/namespace}space"
_CB          = [1000]

# ═══════════════════════════════════════════════════════════
# Pure lxml helpers — NO python-docx private methods used
# ═══════════════════════════════════════════════════════════

def _find_or_add(parent, tag):
    e = parent.find(qn(tag))
    if e is None:
        e = OxmlElement(tag)
        parent.append(e)
    return e

def _replace(parent, tag, new_elem):
    for old in parent.findall(qn(tag)):
        parent.remove(old)
    parent.append(new_elem)

# ─── tblPr from raw tbl lxml element ─────────────────────
def _tblPr_raw(tbl_lxml):
    pr = tbl_lxml.find(qn("w:tblPr"))
    if pr is None:
        pr = OxmlElement("w:tblPr")
        tbl_lxml.insert(0, pr)
    return pr

def _tbl_lxml(tbl):
    """Get raw lxml element from python-docx Table."""
    return tbl._tbl

# ─── table width (pure XML) ───────────────────────────────
def tbl_width(tbl, dxa):
    pr = _tblPr_raw(_tbl_lxml(tbl))
    for old in pr.findall(qn("w:tblW")):
        pr.remove(old)
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"), str(dxa)); w.set(qn("w:type"), "dxa")
    pr.append(w)

# ─── table alignment (pure XML, avoids .alignment attr) ───
def tbl_align_center(tbl):
    pr = _tblPr_raw(_tbl_lxml(tbl))
    for old in pr.findall(qn("w:jc")):
        pr.remove(old)
    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "center")
    pr.append(jc)

# ─── table borders ────────────────────────────────────────
def tbl_borders(tbl, color=C_BORDER):
    pr = _tblPr_raw(_tbl_lxml(tbl))
    for old in pr.findall(qn("w:tblBorders")):
        pr.remove(old)
    bdr = OxmlElement("w:tblBorders")
    for side in ["top","left","bottom","right","insideH","insideV"]:
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),"single"); b.set(qn("w:sz"),"4")
        b.set(qn("w:space"),"0");   b.set(qn("w:color"), color.lstrip("#"))
        bdr.append(b)
    pr.append(bdr)

def tbl_clear_style(tbl):
    """Remove table style + look overrides so cell shading is never overridden."""
    pr = _tblPr_raw(_tbl_lxml(tbl))
    for old in pr.findall(qn("w:tblStyle")): pr.remove(old)
    st = OxmlElement("w:tblStyle"); st.set(qn("w:val"), "TableNormal"); pr.insert(0, st)
    for old in pr.findall(qn("w:tblLook")): pr.remove(old)
    lk = OxmlElement("w:tblLook"); lk.set(qn("w:val"), "0000"); pr.append(lk)

# ─── cell helpers ─────────────────────────────────────────
def _tcPr(cell):
    tc = cell._tc
    pr = tc.find(qn("w:tcPr"))
    if pr is None:
        pr = OxmlElement("w:tcPr"); tc.insert(0, pr)
    return pr

def cell_shade(cell, fill):
    tcPr = _tcPr(cell)
    for old in tcPr.findall(qn("w:shd")): tcPr.remove(old)
    s = OxmlElement("w:shd")
    s.set(qn("w:val"),"clear"); s.set(qn("w:color"),"auto")
    s.set(qn("w:fill"), fill.lstrip("#")); tcPr.append(s)

def cell_valign(cell, val="top"):
    tcPr = _tcPr(cell)
    for old in tcPr.findall(qn("w:vAlign")): tcPr.remove(old)
    v = OxmlElement("w:vAlign"); v.set(qn("w:val"), val); tcPr.append(v)

def cell_w(cell, dxa):
    tcPr = _tcPr(cell)
    for old in tcPr.findall(qn("w:tcW")): tcPr.remove(old)
    w = OxmlElement("w:tcW")
    w.set(qn("w:w"), str(dxa)); w.set(qn("w:type"), "dxa"); tcPr.append(w)

def cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = _tcPr(cell)
    for old in tcPr.findall(qn("w:tcMar")): tcPr.remove(old)
    m = OxmlElement("w:tcMar")
    for side, val in [("top",top),("bottom",bottom),("left",left),("right",right)]:
        s = OxmlElement(f"w:{side}")
        s.set(qn("w:w"), str(val)); s.set(qn("w:type"), "dxa"); m.append(s)
    tcPr.append(m)

def cell_left_border(cell, color, sz="18"):
    tcPr = _tcPr(cell)
    tcBd = tcPr.find(qn("w:tcBorders"))
    if tcBd is None:
        tcBd = OxmlElement("w:tcBorders"); tcPr.append(tcBd)
    for old in tcBd.findall(qn("w:left")): tcBd.remove(old)
    lb = OxmlElement("w:left")
    lb.set(qn("w:val"),"single"); lb.set(qn("w:sz"), sz)
    lb.set(qn("w:space"),"0");   lb.set(qn("w:color"), color.lstrip("#"))
    tcBd.append(lb)

def cell_bottom_border(cell, color, sz="18"):
    tcPr = _tcPr(cell)
    tcBd = tcPr.find(qn("w:tcBorders"))
    if tcBd is None:
        tcBd = OxmlElement("w:tcBorders"); tcPr.append(tcBd)
    for old in tcBd.findall(qn("w:bottom")): tcBd.remove(old)
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"),"single"); b.set(qn("w:sz"), sz)
    b.set(qn("w:space"),"0");   b.set(qn("w:color"), color.lstrip("#"))
    tcBd.append(b)

# ─── row height ───────────────────────────────────────────
def row_h(row, pt):
    tr = row._tr
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr"); tr.insert(0, trPr)
    for old in trPr.findall(qn("w:trHeight")): trPr.remove(old)
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(pt*20))); h.set(qn("w:hRule"), "atLeast")
    trPr.append(h)

# ─── paragraph helpers ────────────────────────────────────
def _pPr(para):
    p = para._p
    pr = p.find(qn("w:pPr"))
    if pr is None:
        pr = OxmlElement("w:pPr"); p.insert(0, pr)
    return pr

def no_space(para):
    """Zero before/after spacing. Uses auto line so text is never clipped."""
    pPr = _pPr(para)
    for old in pPr.findall(qn("w:spacing")): pPr.remove(old)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),  "0")
    sp.set(qn("w:after"),   "0")
    sp.set(qn("w:line"),    "240")
    sp.set(qn("w:lineRule"),"auto")
    pPr.append(sp)

def tight_space(para):
    """Exact single-line spacing — use ONLY for checkbox option lines."""
    pPr = _pPr(para)
    for old in pPr.findall(qn("w:spacing")): pPr.remove(old)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),  "0")
    sp.set(qn("w:after"),   "0")
    sp.set(qn("w:line"),    "240")
    sp.set(qn("w:lineRule"),"exact")
    pPr.append(sp)

def _rPr(run):
    r = run._r
    pr = r.find(qn("w:rPr"))
    if pr is None:
        pr = OxmlElement("w:rPr"); r.insert(0, pr)
    return pr

def _set_font(rPr, font):
    for old in rPr.findall(qn("w:rFonts")): rPr.remove(old)
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"),font); rf.set(qn("w:hAnsi"),font)
    rf.set(qn("w:cs"),font);   rf.set(qn("w:eastAsia"),font)
    rPr.insert(0, rf)

def srun(para, text, bold=False, italic=False, size=None, color=None, font=None):
    run = para.add_run(text)
    run.bold=bold; run.italic=italic
    f=font or FONT; sz=size or FONT_SZ
    run.font.name=f; run.font.size=Pt(sz)
    if color: run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))
    _set_font(_rPr(run), f)
    return run

def cell_new_para(cell):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),  "0")
    sp.set(qn("w:after"),   "0")
    sp.set(qn("w:line"),    "240")
    sp.set(qn("w:lineRule"),"exact")
    pPr.append(sp); p.append(pPr)
    cell._tc.append(p)
    from docx.text.paragraph import Paragraph
    return Paragraph(p, cell)

def cell_new_para_auto(cell):
    """Paragraph with auto line height — for italic notes & cover text."""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),  "0")
    sp.set(qn("w:after"),   "0")
    sp.set(qn("w:line"),    "240")
    sp.set(qn("w:lineRule"),"auto")
    pPr.append(sp); p.append(pPr)
    cell._tc.append(p)
    from docx.text.paragraph import Paragraph
    return Paragraph(p, cell)

def blank(cell):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),  "0")
    sp.set(qn("w:after"),   "0")
    sp.set(qn("w:line"),    "120")
    sp.set(qn("w:lineRule"),"exact")
    pPr.append(sp); p.append(pPr); cell._tc.append(p)

# ═══════════════════════════════════════════════════════════
# Clickable checkbox (Word content control)
# ═══════════════════════════════════════════════════════════
def _checkbox():
    _CB[0] += 1
    cid = _CB[0]
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    a = OxmlElement("w:alias"); a.set(qn("w:val"),"Check Box"); sdtPr.append(a)
    t = OxmlElement("w:tag");   t.set(qn("w:val"),f"cb_{cid}"); sdtPr.append(t)
    i = OxmlElement("w:id");    i.set(qn("w:val"),str(cid));     sdtPr.append(i)
    cb  = etree.SubElement(sdtPr, f"{{{W14}}}checkbox")
    chk = etree.SubElement(cb, f"{{{W14}}}checked")
    chk.set(f"{{{W14}}}val","0")
    on = etree.SubElement(cb, f"{{{W14}}}checkedState")
    on.set(f"{{{W14}}}val","2612"); on.set(f"{{{W14}}}font","MS Gothic")
    off = etree.SubElement(cb, f"{{{W14}}}uncheckedState")
    off.set(f"{{{W14}}}val","2610"); off.set(f"{{{W14}}}font","MS Gothic")
    sdt.append(sdtPr)
    cnt = OxmlElement("w:sdtContent")
    r   = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf  = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"),"MS Gothic"); rf.set(qn("w:hAnsi"),"MS Gothic")
    rPr.append(rf)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"),str(FONT_SZ*2)); rPr.append(sz)
    r.append(rPr)
    tx = OxmlElement("w:t"); tx.text="☐"; r.append(tx)
    cnt.append(r); sdt.append(cnt)
    return sdt

def chk_line(cell, label, italic=False):
    para = cell_new_para(cell)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    no_space(para)
    para._p.append(_checkbox())
    run = para.add_run("  " + label)
    run.font.name=FONT; run.font.size=Pt(FONT_SZ); run.italic=italic
    _set_font(_rPr(run), FONT)

def note(cell, text):
    p = cell_new_para_auto(cell)
    srun(p, text, italic=True, color=C_TEXT_MID, size=FONT_SZ-1)

def field(cell, label="", w=32):
    p = cell_new_para_auto(cell)
    srun(p, label+"_"*w, italic=True, color=C_TEXT_MID, size=FONT_SZ-1)

# ═══════════════════════════════════════════════════════════
# Layout
# ═══════════════════════════════════════════════════════════
SN=500; ATT=4200; RSP=4588; TOTAL=SN+ATT+RSP  # 9288

def make_table(doc):
    t = doc.add_table(rows=1, cols=3)
    tbl_align_center(t)
    tbl_width(t, TOTAL)
    tbl_borders(t, C_BORDER)
    tbl_clear_style(t)
    # Column header row
    for cell, lbl, w, al in zip(
        t.rows[0].cells,
        ["S.N","Attributes","Response"],
        [SN,ATT,RSP],
        [WD_ALIGN_PARAGRAPH.CENTER,WD_ALIGN_PARAGRAPH.LEFT,WD_ALIGN_PARAGRAPH.LEFT]
    ):
        cell_shade(cell, C_MID_BLUE)
        cell_w(cell, w)
        cell_margins(cell, top=80, bottom=80, left=120, right=80)
        cell_valign(cell, "center")
        p = cell.paragraphs[0]; p.alignment=al; no_space(p)
        srun(p, lbl, bold=True, size=FONT_SZ, color=C_WHITE)
    return t

def q_row(tbl, sn, question, builder, tint=False):
    row = tbl.add_row()
    bg  = C_NEAR_WHITE if tint else C_WHITE
    bg2 = C_LIGHT_BLUE if tint else "EAF2FB"
    # S.N
    c0=row.cells[0]; cell_shade(c0,bg2); cell_w(c0,SN)
    cell_margins(c0,80,80,60,60); cell_valign(c0,"top")
    p=c0.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; no_space(p)
    srun(p,str(sn),bold=True,size=FONT_SZ,color=C_MID_BLUE)
    # Attribute
    c1=row.cells[1]; cell_shade(c1,bg); cell_w(c1,ATT)
    cell_margins(c1,80,80,120,80); cell_valign(c1,"top")
    p2=c1.paragraphs[0]; p2.alignment=WD_ALIGN_PARAGRAPH.LEFT; no_space(p2)
    srun(p2,question,size=FONT_SZ,color=C_TEXT_DARK)
    # Response
    c2=row.cells[2]; cell_shade(c2,bg); cell_w(c2,RSP)
    cell_margins(c2,80,80,120,80); cell_valign(c2,"top")
    for op in list(c2._tc.findall(qn("w:p"))): c2._tc.remove(op)
    builder(c2)
    row_h(row, 18)

# ═══════════════════════════════════════════════════════════
# Section header
# ═══════════════════════════════════════════════════════════
def sec_hdr(doc, title, icon=""):
    tbl = doc.add_table(rows=1, cols=1)
    tbl_align_center(tbl)
    tbl_width(tbl, TOTAL)
    tbl_borders(tbl, C_DARK_NAVY)
    tbl_clear_style(tbl)
    cell = tbl.rows[0].cells[0]
    cell_shade(cell, C_DARK_NAVY); cell_w(cell, TOTAL)
    cell_margins(cell,100,100,160,100); row_h(tbl.rows[0],22)
    cell_left_border(cell, C_GOLD)
    p = cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT; no_space(p)
    if icon: srun(p, icon+"  ", bold=True, size=FONT_SZ, color=C_WHITE)
    srun(p, title.upper(), bold=True, size=FONT_SZ, color=C_WHITE)
    g = doc.add_paragraph(); no_space(g); g.paragraph_format.space_after=Pt(2)

# ═══════════════════════════════════════════════════════════
# Response builders
# ═══════════════════════════════════════════════════════════
def r_yn(cell):
    chk_line(cell,"Yes"); chk_line(cell,"No")
    note(cell,"If Yes, please specify:"); field(cell,"",34)

def r_emp(cell):
    for o in ["Immediate (within 1–2 weeks)","Short-term (within 1 month)","Medium-term (1–3 months)","Long-term (>3 months)","Tentative date:","Not yet decided"]:
        chk_line(cell,o)
    
def r_emp1(cell):
    for o in ["< 500","500 – 1,000","1,000 – 5,000"]:
        chk_line(cell,o)
    chk_line(cell,"> 5,000"); field(cell,"  If > 5,000, specify: ",18)

def r_gov(cell):
    for o in ["Yes, centralised global office","Yes, regional offices",
              "No, decisions taken by IT / Legal / Other","No formal structure"]:
        chk_line(cell,o)
    note(cell,"Specify:"); field(cell,"",34)

def r_dec(cell):
    for o in ["Privacy Office","Legal & Compliance","IT Security","Business Unit Heads"]:
        chk_line(cell,o)
    chk_line(cell,"Other(please specify):"); field(cell,"  Specify: ",24)

def r_pol(short):
    def f(cell):
        for o in ["Existing framework in place (requires update)",
                  "Drafted but not implemented","Needs to be formulated from scratch"]:
            chk_line(cell,o)
        chk_line(cell,"Other(please specify):"); field(cell,"  Specify: ",24)
    return f

def r_opts(options, elaborate=False, other=True):
    def f(cell):
        for o in options:
            chk_line(cell,o)
        if other: chk_line(cell,"Other(please specify):"); field(cell,"  Specify: ",24)
        #if elaborate: note(cell,"Please elaborate:"); field(cell,"",34)
    return f

def r_disc(cell):
    chk_line(cell,"Yes"); chk_line(cell,"No")
    note(cell,"If Yes, please specify tool:"); field(cell,"",34)

def r_stor(cell):
    for o in ["On-premise","Cloud","Hybrid(On-premise + Cloud)"]: chk_line(cell,o)
    chk_line(cell,"Other(please specify):"); field(cell,"  Specify: ",24)

# ═══════════════════════════════════════════════════════════
# Header & Footer — pure XML paragraph, no table in header
# ═══════════════════════════════════════════════════════════
''' def add_hdr_ftr(doc, org_name):
    sec = doc.sections[0]

    # ── Header ──────────────────────────────────────────────
    hdr = sec.header; hdr.is_linked_to_previous=False
    he  = hdr._element
    for ch in list(he): he.remove(ch)

    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    # Center-justify the header text
    jc_el = OxmlElement("w:jc"); jc_el.set(qn("w:val"), "center"); pPr.append(jc_el)
    sp = OxmlElement("w:spacing"); sp.set(qn("w:before"),"80"); sp.set(qn("w:after"),"80")
    pPr.append(sp)
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom"); bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"6")
    bot.set(qn("w:space"),"1"); bot.set(qn("w:color"),C_GOLD.lstrip("#"))
    pBdr.append(bot); pPr.append(pBdr)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto")
    shd.set(qn("w:fill"),C_DARK_NAVY.lstrip("#")); pPr.append(shd)
    p.append(pPr)

    def hdr_run(text, color="FFFFFF", size=14):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), FONT); rf.set(qn("w:hAnsi"), FONT)
        rPr.append(rf)
        b = OxmlElement("w:b"); rPr.append(b)
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size)); rPr.append(sz)
        cl = OxmlElement("w:color"); cl.set(qn("w:val"), color); rPr.append(cl)
        r.append(rPr)
        t = OxmlElement("w:t"); t.text = text; t.set(XML_SPC, "preserve")
        r.append(t)
        return r

    # Single centered line: FIRM  |  Team  |  Pre-Scoping Questionnaire
    short_org = org_name if len(org_name) <= 22 else org_name[:20] + "…"
    p.append(hdr_run(
        f"PROTIVITI INDIA MEMBER FIRM | Data Privacy Team | Pre-Scoping Questionnaire | {short_org}"
    ))
    he.append(p)

    # ── Footer ───────────────────────────────────────────────
    ftr = sec.footer; ftr.is_linked_to_previous=False
    fe  = ftr._element
    for ch in list(fe): fe.remove(ch)

    fp = OxmlElement("w:p")
    fpPr = OxmlElement("w:pPr")
    jc=OxmlElement("w:jc"); jc.set(qn("w:val"),"center"); fpPr.append(jc)
    sp2=OxmlElement("w:spacing"); sp2.set(qn("w:before"),"60"); sp2.set(qn("w:after"),"60")
    fpPr.append(sp2)
    fpBdr=OxmlElement("w:pBdr")
    ftop=OxmlElement("w:top"); ftop.set(qn("w:val"),"single"); ftop.set(qn("w:sz"),"6")
    ftop.set(qn("w:space"),"1"); ftop.set(qn("w:color"),C_GOLD.lstrip("#"))
    fpBdr.append(ftop); fpPr.append(fpBdr)
    fp.append(fpPr)

    fr=OxmlElement("w:r"); frPr=OxmlElement("w:rPr")
    frf=OxmlElement("w:rFonts"); frf.set(qn("w:ascii"),FONT); frf.set(qn("w:hAnsi"),FONT)
    frPr.append(frf)
    fsz=OxmlElement("w:sz"); fsz.set(qn("w:val"),"16"); frPr.append(fsz)
    fcl=OxmlElement("w:color"); fcl.set(qn("w:val"),C_TEXT_MID.lstrip("#")); frPr.append(fcl)
    fr.append(frPr)
    ft=OxmlElement("w:t")
    ft.text=(f"CONFIDENTIAL · {org_name} · Protiviti India Member Firm · "
             f"Data Privacy Team · {datetime.now().strftime('%B %Y')}")
    ft.set(XML_SPC,"preserve"); fr.append(ft); fp.append(fr)
    fe.append(fp) '''

# ═══════════════════════════════════════════════════════════
# Page border
# ═══════════════════════════════════════════════════════════
def add_page_border(doc):
    sectPr = doc.sections[0]._sectPr
    pgBdr  = OxmlElement("w:pgBdr")
    for side in ["top","left","bottom","right"]:
        b=OxmlElement(f"w:{side}"); b.set(qn("w:val"),"single")
        b.set(qn("w:sz"),"12"); b.set(qn("w:space"),"24")
        b.set(qn("w:color"),C_MID_BLUE.lstrip("#")); pgBdr.append(b)
    for old in sectPr.findall(qn("w:pgBdr")): sectPr.remove(old)
    sectPr.append(pgBdr)

# ═══════════════════════════════════════════════════════════
# Cover block
# ═══════════════════════════════════════════════════════════
def _set_para_spacing(para, before, after, line, rule="auto"):
    pPr = _pPr(para)
    for old in pPr.findall(qn("w:spacing")): pPr.remove(old)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),  str(before))
    sp.set(qn("w:after"),   str(after))
    sp.set(qn("w:line"),    str(line))
    sp.set(qn("w:lineRule"),rule)
    pPr.append(sp)

def add_cover(doc, org_name, sector, logo_path=None):
    """
    Compact 2-col cover: left = Protiviti logo + tagline, right = title + org + date.
    Row height auto-fits content — no oversized box.
    """
    from docx.shared import Inches
    import io as _io

    LOGO_W  = 2000
    TITLE_W = TOTAL - LOGO_W   # 7288 DXA

    tbl = doc.add_table(rows=1, cols=2)
    tbl_align_center(tbl)
    tbl_width(tbl, TOTAL)
    tbl_borders(tbl, C_DARK_NAVY)
    tbl_clear_style(tbl)

    lc = tbl.rows[0].cells[0]
    rc = tbl.rows[0].cells[1]

    for cell, w in [(lc, LOGO_W), (rc, TITLE_W)]:
        cell_shade(cell, C_DARK_NAVY)
        cell_w(cell, w)
        cell_valign(cell, "center")

    cell_margins(lc, top=120, bottom=120, left=160, right=80)
    cell_margins(rc, top=120, bottom=120, left=80, right=160) 

    # Gold bottom stripe on both cells
    cell_bottom_border(lc, C_GOLD, sz="12")
    cell_bottom_border(rc, C_GOLD, sz="12")

    ''' ── Left: Logo image ──
    lp = lc.paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_para_spacing(lp, 0, 6, 240)
    run = lp.add_run()
    run.add_picture(_io.BytesIO(_LOGO_BYTES), width=Inches(1.15))'''

    # Logo only — no extra text below it

    # ── Right: Title only ──
    rp = rc.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_para_spacing(rp, 0, 0, 280)
    srun(rp, "Pre-Scoping Privacy Questionnaire", bold=True, size=20, color=C_WHITE)

    # Small gap after cover
    g = doc.add_paragraph(); no_space(g)
    g.paragraph_format.space_after = Pt(6)

# ═══════════════════════════════════════════════════════════
# Main export function
# ═══════════════════════════════════════════════════════════
def generate_questionnaire_docx(org_name: str, ai: dict) -> bytes:
    _CB[0]=1000
    short  = ai.get("short_name", org_name.split()[0])
    sector = ai.get("sector","")

    doc = Document()
    for sec in doc.sections:
        sec.page_width=Cm(21); sec.page_height=Cm(29.7)
        sec.top_margin=Cm(2.4); sec.bottom_margin=Cm(1.8)
        sec.left_margin=Cm(1.8); sec.right_margin=Cm(1.8)
        sec.header_distance=Cm(1.2); sec.footer_distance=Cm(1.0)

    sty=doc.styles["Normal"]
    sty.font.name=FONT; sty.font.size=Pt(FONT_SZ)
    sty.paragraph_format.space_before=Pt(0); sty.paragraph_format.space_after=Pt(3)

    # add_hdr_ftr(doc, org_name)
    add_page_border(doc)
    add_cover(doc, org_name, sector, logo_path=ai.get("logo_path"))

    # Section 1
    sec_hdr(doc,"Organisational Overview","🏢")
    t1=make_table(doc)
    q_row(t1,1,"Are there any subsidiaries, affiliates, or joint ventures to be included in this engagement?",r_yn)
    q_row(t1,2,"If your response above is “Yes”, please confirm whether the above mentioned entities have centralized Cybersecurity/IT, HR and Legal functions in place to support all business functions?",r_yn,tint=True)
    q_row(t1,3,"What is the approximate employee strength?",r_emp1)
    doc.add_paragraph()

    # Section 2
    sec_hdr(doc,"Governance & Accountability","⚖️")
    t2=make_table(doc)
    q_row(t2,1,"Has a Privacy Governance Committee or Privacy Office been set up?",r_gov)
    q_row(t2,2,"If your response to the above is “No”, please confirm who takes decisions on the use of personal or its related decision making?",r_dec,tint=True)
    q_row(t2,3,f"What is the current status of {short}'s privacy policy framework?",r_pol(short))
    doc.add_paragraph()

    # Section 3
    sec_hdr(doc,"Business Lines & Stakeholders","📊")
    t3=make_table(doc)
    q_row(t3,1,f"Which of the following are {short}'s core business lines?",r_opts(ai.get("business_lines",[]),elaborate=True))
    q_row(t3,2,"Which of these internal teams may potentially process personal data?",r_opts(ai.get("stakeholder_teams",[])),tint=True)
    doc.add_paragraph()

    # Section 4
    sec_hdr(doc,"Data Ecosystem","🖥️")
    t4=make_table(doc)
    q_row(t4,1,f"List all customer-facing interfaces used by {short}.",r_opts(ai.get("customer_interfaces",[]),elaborate=True))
    q_row(t4,2,"List all core systems / applications that process, store or manage personal data?",r_opts(ai.get("core_systems",[])),tint=True)
    q_row(t4,3,"Do you use any tools to identify, map or track personal data across systems?(E.g.,data discovery, data flow mapping, etc.) ",r_disc)
    q_row(t4,4,"Where is personal data stored and hosted?",r_stor,tint=True)
    doc.add_paragraph()
    
    # Section 5
    sec_hdr(doc,"Cross Border Data Transfer","🏢")
    t6=make_table(doc)
    q_row(t6,1,"Does any personal data processed by the organization get transferred or accessed from outside India? If yes, please specify the countries, entities involved, and purpose of transfer.",r_yn)
    doc.add_paragraph()

     # Section 6
    sec_hdr(doc,"ADDITIONAL DATA","🏢")
    t7=make_table(doc)
    q_row(t7,1,"When do you plan to initiate the engagement? Please provide a tentative start date.",r_emp)
    doc.add_paragraph()


    # Completion note
    nt=doc.add_table(1,1)
    tbl_align_center(nt); tbl_width(nt,TOTAL); tbl_borders(nt,C_GOLD); tbl_clear_style(nt)
    nc=nt.rows[0].cells[0]; cell_shade(nc,"FFF8E7"); cell_w(nc,TOTAL)
    cell_margins(nc,120,120,180,180)
    np_=nc.paragraphs[0]; np_.alignment=WD_ALIGN_PARAGRAPH.CENTER; no_space(np_)
    srun(np_,"Please complete all sections and return to the Data Privacy Team. All information will be treated as strictly confidential.",italic=True,size=FONT_SZ-1,color="7A5C00")

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()
