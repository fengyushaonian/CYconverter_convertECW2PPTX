#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ecw2pptx —— 一键将 .ecw 课件文件转换为标准 .pptx 演示文稿

用法:
    python3 ecw2pptx.py <input.ecw> [output.pptx]

功能:
    - 解析 ecw 包内 slide / layout / master / theme 的私有 XML
    - 将形状(文本/图片/组合/自定义几何/渐变/阴影)映射为 OOXML DrawingML
    - 重建 [Content_Types].xml / _rels 关系链 / presentation.xml
    - 内嵌 resource 图片，输出标准 OOXML .pptx(可由 PowerPoint / WPS / LibreOffice 打开)

仅依赖 Python 标准库(xml.etree.ElementTree / zipfile)。
"""

import sys
import os
import re
import zipfile
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------------------
# OOXML 命名空间
# ---------------------------------------------------------------------------
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------

def esc(text):
    """XML 文本/属性转义"""
    if text is None:
        return ""
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;"))


def find(el, tag):
    """返回 el 下第一个 tag 同名的直接子元素"""
    if el is None:
        return None
    for c in el:
        if c.tag == tag:
            return c
    return None


def findtext(el, tag, default=""):
    c = find(el, tag)
    if c is not None and c.text is not None:
        return c.text
    return default


def num_of_file(name, prefix):
    """从 'slide10.xml' 提取 10"""
    m = re.search(r"(\d+)", name)
    return int(m.group(1)) if m else 0


# ---------------------------------------------------------------------------
# ECW 私有 XML -> DrawingML 转换
# ---------------------------------------------------------------------------

def conv_color(el):
    """
    <color colorType="scheme|rgb|preset" colorValue=".." [lumMod=..] [shade=..] .../>
    -> <a:schemeClr|a:srgbClr|a:prstClr>
    """
    if el is None:
        return '<a:schemeClr val="tx1"/>'
    ctype = el.get("colorType", "scheme")
    val = el.get("colorValue", "000000")
    mods = ""
    for m in ("lumMod", "lumOff", "shade", "tint", "alpha", "satMod", "satOff", "hueMod", "hueOff"):
        v = el.get(m)
        if v is not None:
            mods += '<a:%s val="%s"/>' % (m, esc(v))
    if ctype == "rgb":
        return '<a:srgbClr val="%s">%s</a:srgbClr>' % (esc(val), mods)
    if ctype == "preset":
        return '<a:prstClr val="%s">%s</a:prstClr>' % (esc(val), mods)
    return '<a:schemeClr val="%s">%s</a:schemeClr>' % (esc(val), mods)


def conv_fill(el):
    """
    <background> 或任意填充容器 -> a:noFill / a:solidFill / a:gradFill
    直接子元素: noFill | solidFill(color) | gradFill(gsList, gs*, lin)
    """
    if el is None:
        return '<a:noFill/>'
    for c in el:
        if c.tag == "noFill":
            return "<a:noFill/>"
        if c.tag == "solidFill":
            return "<a:solidFill>%s</a:solidFill>" % conv_color(find(c, "color"))
        if c.tag == "gradFill":
            return conv_gradfill(c)
        if c.tag == "blipFill":
            return "<a:noFill/>"  # 本包未见，降级为无填充
    return "<a:noFill/>"


def conv_gradfill(el):
    stops = "".join(
        "<a:gs pos=\"%s\">%s</a:gs>" % (esc(s.get("pos", "0")), conv_color(find(s, "color")))
        for s in el if s.tag == "gs")
    lin = find(el, "lin")
    if lin is not None:
        lin_xml = '<a:lin ang="%s" scaled="%s"/>' % (
            esc(lin.get("angle", "0")), esc(lin.get("scaled", "0")))
        return "<a:gradFill><a:gsLst>%s</a:gsLst>%s</a:gradFill>" % (stops, lin_xml)
    path = find(el, "path")
    if path is not None:
        # 径向渐变：CT_PathGradientFill = gsLst + fillToRect（在 a:path 内）
        fr = find(path, "fillToRect")
        fr_xml = ""
        if fr is not None:
            attrs = "".join(
                ' %s="%s"' % (k, esc(fr.get(k))) for k in ("left", "top", "right", "bottom")
                if fr.get(k) is not None)
            fr_xml = "<a:fillToRect%s/>" % attrs
        path_xml = '<a:path path="%s"><a:gsLst>%s</a:gsLst>%s</a:path>' % (
            esc(path.get("pathType", path.get("path", "circle"))), stops, fr_xml)
        return "<a:gradFill>%s</a:gradFill>" % path_xml
    return "<a:gradFill><a:gsLst>%s</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>" % stops


def conv_outline(el):
    """<outline width=.. [cap] [cmpd]> noFill|solidFill </outline> -> <a:ln>"""
    if el is None:
        return ""
    attrs = ""
    for k, o in (("width", "w"), ("cap", "cap"), ("cmpd", "cmpd")):
        v = el.get(k)
        if v is not None:
            attrs += ' %s="%s"' % (o, esc(v))
    inner = ""
    for c in el:
        if c.tag == "noFill":
            inner = "<a:noFill/>"
        elif c.tag == "solidFill":
            inner = "<a:solidFill>%s</a:solidFill>" % conv_color(find(c, "color"))
        elif c.tag == "gradFill":
            inner = conv_gradfill(c)
    return "<a:ln%s>%s</a:ln>" % (attrs, inner)


def conv_prstgeom(el):
    """<prstGeom><prstGeomType>rect</prstGeomType><adjList/></prstGeom> -> a:prstGeom
    注意：ecw 的几何名是 prstGeomType 的【元素文本】，不是属性
    """
    pt = find(el, "prstGeomType")
    name = "rect"
    if pt is not None:
        name = (pt.text or "").strip() or pt.get("prstGeomType") or "rect"
    adj = find(el, "adjList")
    adj_xml = ""
    if adj is not None:
        adj_xml = "".join(
            '<a:adj name="%s" val="%s"/>' % (esc(a.get("name", "")), esc(a.get("val", "")))
            for a in adj if a.tag == "adj")
    return '<a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>' % (esc(name), adj_xml)


SVG_TOK = re.compile(r"[MmCcLlZz]|-?\d*\.?\d+(?:[eE][+-]?\d+)?")

# ecw 的 pathString 是相对坐标（设计单位），1 单位 = 6350 EMU
PATH_SCALE = 6350.0


def _scale_path_num(n):
    """pathString 数值 -> EMU 坐标（×6350）"""
    try:
        v = float(n) * PATH_SCALE
        return str(int(v)) if abs(v - round(v)) < 1e-6 else ("%g" % v)
    except (ValueError, TypeError):
        return str(n)


def svg_path_to_path_xml(path_string, w, h):
    """把 SVG pathString(M/L/C/Z) 转为 <a:path> 的子命令（坐标 ×6350 -> EMU）"""
    commands = []
    cur = None
    nums = []
    for tok in SVG_TOK.finditer(path_string or ""):
        t = tok.group(0)
        if t in "MmCcLlZz":
            if cur is not None:
                commands.append((cur, nums))
            cur = t
            nums = []
        else:
            nums.append(t)
    if cur is not None:
        commands.append((cur, nums))

    out = []
    for cmd, nums in commands:
        if cmd in ("M", "m"):
            for i in range(0, len(nums) - 1, 2):
                out.append('<a:moveTo><a:pt x="%s" y="%s"/></a:moveTo>'
                           % (_scale_path_num(nums[i]), _scale_path_num(nums[i + 1])))
            # M 之后的隐式坐标对按 L 处理
            if len(nums) > 2:
                for i in range(2, len(nums) - 1, 2):
                    out.append('<a:lnTo><a:pt x="%s" y="%s"/></a:lnTo>'
                               % (_scale_path_num(nums[i]), _scale_path_num(nums[i + 1])))
        elif cmd in ("L", "l"):
            for i in range(0, len(nums) - 1, 2):
                out.append('<a:lnTo><a:pt x="%s" y="%s"/></a:lnTo>'
                           % (_scale_path_num(nums[i]), _scale_path_num(nums[i + 1])))
        elif cmd in ("C", "c"):
            for i in range(0, len(nums) - 5, 6):
                x1, y1, x2, y2, x, y = nums[i:i + 6]
                out.append('<a:cubicBezTo><a:pt x="%s" y="%s"/><a:pt x="%s" y="%s"/>'
                           '<a:pt x="%s" y="%s"/></a:cubicBezTo>'
                           % tuple(_scale_path_num(v) for v in (x1, y1, x2, y2, x, y)))
        elif cmd in ("Z", "z"):
            out.append("<a:close/>")
    return '<a:path w="%s" h="%s">%s</a:path>' % (esc(w), esc(h), "".join(out))


def conv_customgeom(el):
    """<customGeom><gdLst><gd formula=../>..</gdLst><pathList><path pathString=../>..</pathList> -> a:custGeom"""
    gd = find(el, "gdLst")
    gd_xml = ""
    if gd is not None:
        gd_xml = "".join(
            '<a:gd name="%s" f="%s"/>' % (esc(g.get("name", "")), esc(g.get("formula", "")))
            for g in gd if g.tag == "gd")
    pl = find(el, "pathList")
    paths = ""
    r = b = 0
    if pl is not None:
        for p in pl:
            if p.tag == "path":
                w = p.get("width", "0")
                h = p.get("height", "0")
                try:
                    r = max(r, int(float(w)))
                    b = max(b, int(float(h)))
                except (ValueError, TypeError):
                    pass
                paths += svg_path_to_path_xml(p.get("pathString", ""), w, h)
    return ('<a:custGeom><a:avLst/><a:gdLst>%s</a:gdLst><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="%s" b="%s"/><a:pathLst>%s</a:pathLst></a:custGeom>'
            % (gd_xml, r, b, paths))


def conv_effectlist(el):
    """<effectList> reflection/outerShadow/innerShadow/softEdge/glow -> a:effectLst"""
    if el is None:
        return ""
    inner = ""
    for c in el:
        if c.tag == "reflection":
            attrs = ""
            for k, o in (("blurRad", "blurRad"), ("direction", "dir"),
                         ("endPos", "endPos"), ("startAlpha", "stA"),
                         ("endAlpha", "endA"), ("scaleX", "sx"),
                         ("scaleY", "sy"), ("alignment", "algn"),
                         ("rotWithShape", "rotWithShape"), ("fadeDirection", "fadeDir")):
                v = c.get(k)
                if v is not None:
                    attrs += ' %s="%s"' % (o, esc(v))
            inner += "<a:reflection%s/>" % attrs
        elif c.tag == "outerShadow":
            attrs = ""
            for k, o in (("blurRad", "blurRad"), ("direction", "dir"),
                         ("distance", "dist"), ("scaleX", "sx"),
                         ("scaleY", "sy"), ("alignment", "algn"),
                         ("rotWithShape", "rotWithShape")):
                v = c.get(k)
                if v is not None:
                    attrs += ' %s="%s"' % (o, esc(v))
            inner += "<a:outerShdw%s>%s</a:outerShdw>" % (
                attrs, conv_color(find(c, "color")))
        elif c.tag == "innerShadow":
            attrs = ""
            for k, o in (("blurRad", "blurRad"), ("direction", "dir"),
                         ("distance", "dist"), ("alignment", "algn"),
                         ("rotWithShape", "rotWithShape")):
                v = c.get(k)
                if v is not None:
                    attrs += ' %s="%s"' % (o, esc(v))
            inner += "<a:innerShdw%s>%s</a:innerShdw>" % (
                attrs, conv_color(find(c, "color")))
        elif c.tag == "softEdge":
            inner += '<a:softEdge rad="%s"/>' % esc(c.get("rad", "0"))
        elif c.tag == "glow":
            inner += "<a:glow>%s</a:glow>" % conv_color(find(c, "color"))
    return "<a:effectLst>%s</a:effectLst>" % inner


def get_em(el, name, default="0"):
    c = find(el, name)
    if c is not None and c.text is not None and c.text != "":
        return c.text
    return default


def conv_xfrm(el, group=False):
    """<transform> -> <a:xfrm>；group 模式附带 chOff/chExt"""
    if el is None:
        if group:
            return ('<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
                    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>')
        return '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
    rot = ""
    r = find(el, "rotation")
    if r is not None and r.text:
        rot = ' rot="%s"' % esc(r.text)
    off = '<a:off x="%s" y="%s"/>' % (esc(get_em(el, "left")), esc(get_em(el, "top")))
    ext = '<a:ext cx="%s" cy="%s"/>' % (esc(get_em(el, "width")), esc(get_em(el, "height")))
    if group:
        choff = '<a:chOff x="%s" y="%s"/>' % (
            esc(get_em(el, "childOffsetLeft")), esc(get_em(el, "childOffsetTop")))
        chext = '<a:chExt cx="%s" cy="%s"/>' % (
            esc(get_em(el, "childWidth")), esc(get_em(el, "childHeight")))
        return "<a:xfrm%s>%s%s%s%s</a:xfrm>" % (rot, off, ext, choff, chext)
    return "<a:xfrm%s>%s%s</a:xfrm>" % (rot, off, ext)


def conv_locks(el, kind):
    if el is None:
        return ""
    attrs = ""
    for k in ("noChangeAspect", "noGrp", "noMove", "noResize", "noRot",
              "noSelect", "noChangeArrowheads", "noEditPoints"):
        v = el.get(k)
        if v is not None:
            attrs += ' %s="%s"' % (k, esc(v))
    if kind == "pic":
        tag = "a:picLocks"
    elif kind == "grp":
        tag = "a:grpSpLocks"
    else:
        tag = "a:spLocks"
    return "<%s%s/>" % (tag, attrs)


def conv_spPr(props, kind="sp"):
    """<props> -> <p:spPr> 内部内容(xfrm/geom/fill/ln/effect/locks)"""
    if props is None:
        return ("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>"
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>')
    parts = []
    xf = find(props, "transform")
    parts.append(conv_xfrm(xf))

    pg = find(props, "prstGeom")
    cg = find(props, "customGeom")
    if pg is not None:
        parts.append(conv_prstgeom(pg))
    elif cg is not None:
        parts.append(conv_customgeom(cg))
    else:
        parts.append('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>')

    bg = find(props, "background")
    parts.append(conv_fill(bg) if bg is not None else "<a:noFill/>")

    ol = find(props, "outline")
    if ol is not None:
        nofill = find(ol, "noFill") is not None
        if nofill and (ol.get("width") in (None, "0")):
            parts.append('<a:ln w="0"><a:noFill/></a:ln>')
        else:
            parts.append(conv_outline(ol))

    ef = find(props, "effectList")
    if ef is not None:
        parts.append(conv_effectlist(ef))

    return "".join(parts)


# ---------------------------------------------------------------------------
# 文本
# ---------------------------------------------------------------------------

RUNPROPS_MAP = {
    "lang": "lang", "altLang": "altLang", "bold": "b", "italic": "i",
    "fontSize": "sz", "kern": "kern", "underline": "u", "strike": "strike",
    "cap": "cap", "spc": "spc", "noProof": "noProof", "smtClean": "smtClean",
    "baseline": "baseline", "dirty": "dirty", "spcPts": "spc",
}


def conv_runprops(el, tag="a:rPr"):
    """textRunProps/endParaRunProps -> a:rPr / a:endParaRPr
    子元素顺序遵循 CT_TextCharacterProperties：
    ln -> fill -> effectLst -> latin/ea/cs/sym
    """
    if el is None:
        return "<%s lang=\"zh-CN\"/>" % tag
    attrs = ""
    for k, o in RUNPROPS_MAP.items():
        v = el.get(k)
        if v is not None:
            attrs += ' %s="%s"' % (o, esc(v))
    inner = ""
    ol = find(el, "outline")
    if ol is not None and find(ol, "noFill") is None and ol.get("width") not in (None, "0"):
        inner += conv_outline(ol)
    bg = find(el, "background")
    if bg is not None:
        inner += conv_fill(bg)
    ef = find(el, "effectList")
    if ef is not None:
        inner += conv_effectlist(ef)
    f = find(el, "font")
    if f is not None:
        for ftag in ("latin", "ea", "cs", "sym"):
            fc = find(f, ftag)
            if fc is not None:
                tf = fc.get("typeface", "")
                # 过滤主题占位符(+mn-*/+mj-*)与空 typeface：此类值只允许出现在
                # theme 的 fontScheme 中，出现在 run 属性里属非法，省略即继承主题字体
                if tf and not tf.startswith("+"):
                    inner += '<a:%s typeface="%s"/>' % (ftag, esc(tf))
    return "<%s%s>%s</%s>" % (tag, attrs, inner, tag)


def _bullet_xml(bullet_el):
    """<bullet> -> a:buFont + a:buChar/a:buAutoNum 片段（按 OOXML 顺序）"""
    out = ""
    bf = find(bullet_el, "buFont")
    if bf is not None:
        tf = bf.get("typeface", "")
        if tf and not tf.startswith("+"):
            out += '<a:buFont typeface="%s"/>' % esc(tf)
    bc = find(bullet_el, "buChar")
    if bc is not None and bc.text:
        out += '<a:buChar char="%s"/>' % esc(bc.text)
    bau = find(bullet_el, "buAutoNum")
    if bau is not None:
        out += '<a:buAutoNum type="%s"/>' % esc(bau.get("type", "arabicPeriod"))
    return out


def conv_paragraph(el):
    """<paragraph> -> <a:p>"""
    pp = find(el, "paragraphProps")
    ppr = ""
    if pp is not None:
        attrs = ""
        for k, o in (("alignment", "algn"), ("marLeft", "marL"),
                     ("marRight", "marR"), ("indent", "indent"),
                     ("defTabSz", "defTabSz"), ("fontAlgn", "fontAlgn"),
                     ("eaLnBrk", "eaLnBrk"), ("latinLnBrk", "latinLnBrk"),
                     ("hangingPunct", "hangingPunct"), ("rtl", "rtl")):
            v = pp.get(k)
            if v is not None:
                attrs += ' %s="%s"' % (o, esc(v))
        inner = ""
        def_rpr = ""
        for c in pp:
            if c.tag in ("lnSpc", "spcBef", "spcAft"):
                if c.get("spcPct") is not None:
                    inner += "<a:%s><a:spcPct val=\"%s\"/></a:%s>" % (
                        c.tag, esc(c.get("spcPct")), c.tag)
                elif c.get("spcPts") is not None:
                    inner += "<a:%s><a:spcPts val=\"%s\"/></a:%s>" % (
                        c.tag, esc(c.get("spcPts")), c.tag)
            elif c.tag == "bullet":
                inner += _bullet_xml(c)
            elif c.tag == "bulletNone":
                inner += "<a:buNone/>"
            elif c.tag == "tabList":
                tabs = [t for t in c if t.tag == "tab"]
                if tabs:
                    inner += "<a:tabLst>%s</a:tabLst>" % "".join(
                        '<a:tab pos="%s" algn="%s"/>' % (
                            esc(t.get("pos", "0")), esc(t.get("algn", "l")))
                        for t in tabs)
            elif c.tag == "textRunProps":
                # paragraphProps 的默认 run 属性 -> a:defRPr（须位于 pPr 末尾）
                def_rpr = conv_runprops(c, "a:defRPr")
        inner += def_rpr
        ppr = "<a:pPr%s>%s</a:pPr>" % (attrs, inner)

    parts = [ppr] if ppr else []
    runs = []
    end_para = None
    for c in el:
        if c.tag == "endParaRunProps":
            # a:endParaRPr 必须位于段落内所有 run 之后
            end_para = conv_runprops(c, "a:endParaRPr")
        elif c.tag == "textRun":
            t = find(c, "text")
            rp = find(c, "textRunProps")
            rpr = conv_runprops(rp) if rp is not None else '<a:rPr lang="zh-CN"/>'
            runs.append("<a:r>%s<a:t>%s</a:t></a:r>" % (
                rpr, esc(t.text if t is not None and t.text else "")))
        elif c.tag == "textField":
            fid = c.get("id", "")
            ftype = c.get("type", "")
            fmap = {
                "datetimeFigureOut": "datetime1",
                "datetimeFigureIn": "datetime2",
                "datetimeFigureBoth": "datetime3",
                "slidenum": "slidenum",
                "slideTitle": "title",
                "datetimeFigureNum": "datetime1",
            }
            otype = fmap.get(ftype, ftype if ftype else "slidenum")
            t = find(c, "text")
            rp = find(c, "textRunProps")
            rpr = conv_runprops(rp) if rp is not None else '<a:rPr lang="zh-CN"/>'
            runs.append('<a:fld id="%s" type="%s">%s<a:t>%s</a:t></a:fld>' % (
                esc(fid), esc(otype), rpr,
                esc(t.text if t is not None and t.text else "")))
    parts.extend(runs)
    if end_para is not None:
        parts.append(end_para)
    return "<a:p>%s</a:p>" % "".join(parts)


def conv_textbody(el):
    """<textBody> -> <a:txBody>"""
    if el is None:
        return "<a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>"
    bp = find(el, "bodyProps")
    body = "<a:bodyPr/>"
    if bp is not None:
        attrs = ""
        for k, o in (("wrap", "wrap"), ("anchor", "anchor"),
                     ("anchorCtr", "anchorCtr"), ("topInset", "tIns"),
                     ("bottomInset", "bIns"), ("leftInset", "lIns"),
                     ("rightInset", "rIns"), ("rtlCol", "rtlCol"),
                     ("vert", "vert"), ("upright", "upright")):
            v = bp.get(k)
            if v is not None:
                attrs += ' %s="%s"' % (o, esc(v))
        af = find(bp, "autoFitType")
        afx = ""
        if af is not None:
            v = af.get("val", "")
            if v == "normAutofit":
                afx = "<a:normAutofit/>"
            elif v == "noAutofit":
                afx = "<a:noAutofit/>"
            elif v == "spAutoFit":
                afx = "<a:spAutoFit/>"
        body = "<a:bodyPr%s>%s</a:bodyPr>" % (attrs, afx)
    pl = find(el, "paragraphList")
    ps = ""
    if pl is not None:
        ps = "".join(conv_paragraph(p) for p in pl if p.tag == "paragraph")
    if not ps:
        ps = "<a:p/>"
    # textStyleList -> a:lstStyle（ecw 中默认段落样式/字体存放于此，必须转换）
    lst = ""
    tsl = find(el, "textStyleList")
    if tsl is not None:
        for ts in tsl:
            if ts.tag == "textStyle":
                lvl = ts.get("level", "0")
                try:
                    lvl_n = int(lvl) + 1
                except ValueError:
                    lvl_n = 1
                lst += conv_textstyle(ts, "a:lvl%dpPr" % max(1, min(lvl_n, 9)))
    return "<a:txBody>%s<a:lstStyle>%s</a:lstStyle>%s</a:txBody>" % (body, lst, ps)


# ---------------------------------------------------------------------------
# 形状
# ---------------------------------------------------------------------------

class Context:
    """转换上下文：收集每部分引用的图片，分配 rId"""

    def __init__(self, rid_start=1):
        self.images = []          # [(resource_path, media_name, rid)]
        self.rid_counter = rid_start

    def add_image(self, src):
        """src 如 image6.png；返回分配的 rId"""
        for path, media, rid in self.images:
            if media == src:
                return rid
        rid = "rId%d" % self.rid_counter
        self.rid_counter += 1
        self.images.append((src, src, rid))
        return rid


def conv_shape(el, ctx, media_set):
    """单个 <shape> -> p:sp / p:pic / p:grpSp"""
    sid = findtext(el, "id", "0")
    name = findtext(el, "name", "")
    descr = findtext(el, "descr", "")
    typ = findtext(el, "type", "shape")
    props = find(el, "props")

    # locks: 必须位于 cNvSpPr / cNvPicPr / cNvGrpSpPr（不得放进 spPr）
    locks_xml = ""
    ph_xml = ""
    if props is not None:
        lk = find(props, "locks")
        if lk is not None:
            locks_kind = "pic" if typ == "pic" else ("grp" if typ == "groupShape" else "sp")
            locks_xml = conv_locks(lk, locks_kind)
        ph = find(props, "placeholder")
        if ph is not None:
            attrs = ""
            for k in ("type", "sz", "idx", "orient"):
                v = ph.get(k)
                if v:
                    attrs += ' %s="%s"' % (k, esc(v))
            ph_xml = "<p:ph%s/>" % attrs
    nvpr = "<p:nvPr>%s</p:nvPr>" % ph_xml if ph_xml else "<p:nvPr/>"

    if typ == "pic":
        img = find(el, "image")
        src = None
        if img is not None:
            blip = find(img, "blip")
            if blip is not None:
                src = findtext(blip, "src")
        rid = None
        if src:
            if src in media_set:
                rid = ctx.add_image(src)
            else:
                rid = None
        nv = ('<p:nvPicPr><p:cNvPr id="%s" name="%s" descr="%s"/>'
              "<p:cNvPicPr>%s</p:cNvPicPr>%s</p:nvPicPr>") % (
            esc(sid), esc(name), esc(descr), locks_xml, nvpr)
        if rid:
            blip_fill = ('<p:blipFill><a:blip r:embed="%s"/>'
                         "<a:stretch><a:fillRect/></a:stretch></p:blipFill>") % rid
        else:
            blip_fill = "<p:blipFill><a:noFill/></p:blipFill>"
        sp = conv_spPr(props, "pic")
        return "<p:pic>%s%s<p:spPr>%s</p:spPr></p:pic>" % (nv, blip_fill, sp)

    if typ == "groupShape":
        sl = find(el, "shapeList")
        inner = ""
        if sl is not None:
            inner = "".join(conv_shape(s, ctx, media_set) for s in sl if s.tag == "shape")
        nv = ('<p:nvGrpSpPr><p:cNvPr id="%s" name="%s"/>'
              "<p:cNvGrpSpPr>%s</p:cNvGrpSpPr>%s</p:nvGrpSpPr>") % (
            esc(sid), esc(name), locks_xml, nvpr)
        xf = find(props, "transform") if props is not None else None
        grp = ("<p:grpSpPr>%s</p:grpSpPr>") % conv_xfrm(xf, group=True)
        return "<p:grpSp>%s%s%s</p:grpSp>" % (nv, grp, inner)

    # 普通形状 shape
    nv = ('<p:nvSpPr><p:cNvPr id="%s" name="%s" descr="%s"/>'
          "<p:cNvSpPr>%s</p:cNvSpPr>%s</p:nvSpPr>") % (
        esc(sid), esc(name), esc(descr), locks_xml, nvpr)
    sp = conv_spPr(props)
    tb = find(el, "textBody")
    tx = conv_textbody(tb) if tb is not None else ""
    return "<p:sp>%s<p:spPr>%s</p:spPr>%s</p:sp>" % (nv, sp, tx)


def conv_spTree(group_shape_el, ctx, media_set):
    """<groupShape>(顶层) -> <p:spTree>"""
    inner = ""
    if group_shape_el is not None:
        inner = "".join(
            conv_shape(s, ctx, media_set) for s in group_shape_el if s.tag == "shape")
    return ('<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + inner + "</p:spTree>")


# ---------------------------------------------------------------------------
# slide / layout / master / theme 组装
# ---------------------------------------------------------------------------

def conv_textstyle(ts, tag):
    """master/presentation 中的 <textStyle level=N> -> a:lvlNpPr / a:defPPr 内容"""
    attrs = ""
    for k, o in (("marLeft", "marL"), ("indent", "indent"),
                 ("alignment", "algn"), ("defTablSz", "defTabSz"),
                 ("eaLnBrk", "eaLnBrk"), ("latinLnBrk", "latinLnBrk"),
                 ("hangingPunct", "hangingPunct"), ("fontAlgn", "fontAlgn")):
        v = ts.get(k)
        if v is not None:
            attrs += ' %s="%s"' % (o, esc(v))
    inner = ""
    for c in ts:
        if c.tag in ("lnSpc", "spcBef", "spcAft"):
            if c.get("spcPct") is not None:
                inner += "<a:%s><a:spcPct val=\"%s\"/></a:%s>" % (
                    c.tag, esc(c.get("spcPct")), c.tag)
            elif c.get("spcPts") is not None:
                inner += "<a:%s><a:spcPts val=\"%s\"/></a:%s>" % (
                    c.tag, esc(c.get("spcPts")), c.tag)
        elif c.tag == "bullet":
            inner += _bullet_xml(c)
        elif c.tag == "buNone":
            inner += "<a:buNone/>"
        elif c.tag == "textRunProps":
            inner += conv_runprops(c, "a:defRPr")
    return "<%s%s>%s</%s>" % (tag, attrs, inner, tag)


def build_slide(el, ctx, media_set, layout_id_xml=""):
    sd = find(el, "slideData")
    bgx = ""
    if sd is not None:
        bg = find(sd, "background")
        if bg is not None:
            bgx = ("<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>" % conv_fill(bg))
    gs = find(sd, "groupShape") if sd is not None else None
    st = conv_spTree(gs, ctx, media_set)
    clr = find(el, "clrMapOvr")
    clr_xml = ""
    if clr is not None:
        if clr.get("masterClrMapping") == "1":
            clr_xml = "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
        else:
            clr_xml = "<p:clrMapOvr/>"
    return ('<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
            "<p:cSld>%s%s</p:cSld>%s%s</p:sld>") % (
        NS_A, NS_R, NS_P, bgx, st, clr_xml, layout_id_xml)


LAYOUT_TYPE_MAP = {
    "blank": "blank", "BLANK": "blank",
    "title": "title", "title slide": "title",
    "title and content": "twoObj", "title and content over content": "twoObj",
    "two content": "twoObj", "section header": "secHead",
    "comparison": "comparison", "title only": "titleOnly",
    "content with caption": "objTx", "picture with caption": "picTx",
    "title and vertical text": "vertTx", "vertical title and text": "vertTitleAndTx",
    "vertical title and text over content": "vertTxOverTx",
    "quote caption": "quoteCaption", "name card": "nameCard",
    "table": "tbl", "chart and text": "chartAndTx",
    "clip art and vertical text": "clipArtAndVertTx", "title and picture": "picTitle",
    "text and chart": "txAndChart", "text and clip art": "txAndClipArt",
    "content and caption": "txObj", "picture above text": "picAboveTx",
    "text over object": "txOverObj", "object over text": "objOverTx",
    "text over picture": "txOverPic", "object and caption": "objAndTx",
    "content over content": "obj", "object and two objects": "obj",
    "object": "obj", "obj": "obj", "content": "obj",
}


def guess_layout_type(name):
    key = (name or "").strip().lower()
    if key in LAYOUT_TYPE_MAP:
        return LAYOUT_TYPE_MAP[key]
    if "title" in key and "content" in key:
        return "twoObj"
    if "title" in key:
        return "title"
    return "blank"


def build_layout(el, ctx, media_set):
    name = findtext(el, "name", "")
    sd = find(el, "slideData")
    bgx = ""
    if sd is not None:
        bg = find(sd, "background")
        if bg is not None:
            bgx = ("<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>" % conv_fill(bg))
    gs = find(sd, "groupShape") if sd is not None else None
    st = conv_spTree(gs, ctx, media_set)
    clr = find(el, "clrMapOvr")
    clr_xml = ""
    if clr is not None:
        clr_xml = "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" \
            if clr.get("masterClrMapping") == "1" else "<p:clrMapOvr/>"
    ltype = guess_layout_type(name)
    return ('<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" preserve="1">'
            "<p:cSld>%s%s</p:cSld>%s<p:type>%s</p:type></p:sldLayout>"
            ) % (NS_A, NS_R, NS_P, bgx, st, clr_xml, ltype)


def build_master(el, ctx, media_set, layout_ids_xml):
    """layout_ids_xml: <p:sldLayoutId id=.. r:id=../> 片段，rId 由调用方分配"""
    sd = find(el, "slideData")
    bgx = ""
    if sd is not None:
        bgref = find(sd, "bgRef")
        if bgref is not None:
            bgx = ('<p:bg><p:bgRef idx="%s">%s</p:bgRef></p:bg>') % (
                esc(bgref.get("idx", "1001")), conv_color(find(bgref, "color")))
        else:
            bg = find(sd, "background")
            if bg is not None:
                bgx = ("<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>" % conv_fill(bg))
    gs = find(sd, "groupShape") if sd is not None else None
    st = conv_spTree(gs, ctx, media_set)

    clr = find(el, "clrMap")
    clr_xml = ('<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
               'accent1="accent1" accent2="accent2" accent3="accent3" '
               'accent4="accent4" accent5="accent5" accent6="accent6" '
               'hlink="hlink" folHlink="folHlink"/>')
    if clr is not None:
        m = {}
        for c in clr:
            if c.text:
                m[c.tag] = c.text
        if m:
            clr_xml = "<p:clrMap " + " ".join(
                '%s="%s"' % (k, esc(v)) for k, v in m.items()) + "/>"

    tx = ""
    has_title = has_body = has_other = False
    for src_tag, dst_tag in (("titleTextStyleList", "titleStyle"),
                             ("bodyTextStyleList", "bodyStyle"),
                             ("otherTextStyleList", "otherStyle")):
        lst = find(el, src_tag)
        if lst is not None:
            if src_tag == "titleTextStyleList":
                has_title = True
            elif src_tag == "bodyTextStyleList":
                has_body = True
            else:
                has_other = True
            lvls = []
            for ts in lst:
                if ts.tag != "textStyle":
                    continue
                try:
                    lvl = int(ts.get("level", "0")) + 1
                except ValueError:
                    lvl = 1
                if 1 <= lvl <= 9:
                    lvls.append(conv_textstyle(ts, "a:lvl%dpPr" % lvl))
            tx += "<p:%s>%s</p:%s>" % (dst_tag, "".join(lvls), dst_tag)
    # 缺失的样式组补空定义，保证 txStyles 三组齐全
    if not has_title:
        tx += '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>'
    if not has_body:
        tx += '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle>'
    if not has_other:
        tx += '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>'

    return ('<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
            "<p:cSld>%s%s</p:cSld>%s<p:sldLayoutIdLst>%s</p:sldLayoutIdLst>"
            "<p:txStyles>%s</p:txStyles></p:sldMaster>"
            ) % (NS_A, NS_R, NS_P, bgx, st, clr_xml, layout_ids_xml, tx)


def build_theme(el):
    """ecw <themeElements> -> a:theme"""
    name = findtext(el, "name", "默认设计模板")
    cs = find(el, "clrScheme")
    clr_xml = ""
    if cs is not None:
        order = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink"]
        entries = []
        for tag in order:
            c = find(cs, tag)
            if c is not None and c.text:
                entries.append('<a:%s><a:srgbClr val="%s"/></a:%s>' % (
                    tag, esc(c.text.strip()), tag))
        clr_xml = '<a:clrScheme name="%s">%s</a:clrScheme>' % (esc(name), "".join(entries))

    fs = find(el, "fontScheme")
    font_xml = ""
    if fs is not None:
        def font_part(tag):
            node = find(fs, tag)
            if node is None:
                return ""
            inner = ""
            for ftag in ("latin", "ea", "cs"):
                fc = find(node, ftag)
                if fc is not None:
                    tf = fc.get("typeface", "")
                    if tf:  # 过滤空 typeface
                        inner += '<a:%s typeface="%s"/>' % (ftag, esc(tf))
            return "<a:%s>%s</a:%s>" % (tag, inner, tag)
        font_xml = '<a:fontScheme name="%s">%s%s</a:fontScheme>' % (
            esc(name), font_part("majorFont"), font_part("minorFont"))

    fmt = find(el, "fmtScheme")
    fmt_xml = ""
    if fmt is not None:
        fl = find(fmt, "fillStyleList")
        fl_xml = ""
        if fl is not None:
            fl_xml = "<a:fillStyleLst>%s</a:fillStyleLst>" % "".join(
                conv_fill(c) if c.tag in ("solidFill", "gradFill", "noFill", "blipFill") else ""
                for c in fl)
        ol = find(fmt, "outlineStyleList")
        ol_xml = ""
        if ol is not None:
            parts = []
            for c in ol:
                if c.tag == "outline":
                    parts.append(conv_outline(c))
            ol_xml = "<a:lnStyleLst>%s</a:lnStyleLst>" % "".join(parts)
        ef = find(fmt, "effectStyleList")
        ef_xml = ""
        if ef is not None:
            parts = []
            for c in ef:
                if c.tag == "effectStyle":
                    eff = find(c, "effectList")
                    parts.append("<a:effectStyle><a:effectLst>%s</a:effectLst></a:effectStyle>"
                                 % (conv_effectlist(eff) if eff is not None else ""))
            ef_xml = "<a:effectStyleLst>%s</a:effectStyleLst>" % "".join(parts)
        bg = find(fmt, "bgFillStyleList")
        bg_xml = ""
        if bg is not None:
            bg_xml = "<a:bgFillStyleLst>%s</a:bgFillStyleLst>" % "".join(
                conv_fill(c) if c.tag in ("solidFill", "gradFill", "noFill", "blipFill") else ""
                for c in bg)
        fmt_xml = '<a:fmtScheme name="%s">%s%s%s%s</a:fmtScheme>' % (
            esc(name), fl_xml, ol_xml, ef_xml, bg_xml)
    else:
        fmt_xml = ('<a:fmtScheme name="%s"><a:fillStyleLst><a:solidFill>'
                   '<a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
                   '<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng"><a:solidFill>'
                   '<a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
                   '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
                   '</a:effectStyleLst><a:bgFillStyleLst><a:solidFill>'
                   '<a:schemeClr val="bg1"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
                   ) % esc(name)

    return ('<a:theme xmlns:a="%s" name="%s"><a:themeElements>%s%s%s</a:themeElements>'
            "<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
            ) % (NS_A, esc(name), clr_xml, font_xml, fmt_xml)


# ---------------------------------------------------------------------------
# 打包
# ---------------------------------------------------------------------------

CT_TYPES = {
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
    "xml": "application/xml",
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "emf": "image/x-emf",
    "wdp": "image/vnd.ms-photo",
}

CT_OVERRIDES = {
    "/ppt/presentation.xml":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "slide":
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    "slideLayout":
        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "slideMaster":
        "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "theme":
        "application/vnd.openxmlformats-officedocument.theme+xml",
}


def make_content_types(parts, media_exts):
    lines = ['<Types xmlns="%s">' % NS_CT]
    lines.append('<Default Extension="rels" ContentType="%s"/>' % CT_TYPES["rels"])
    lines.append('<Default Extension="xml" ContentType="application/xml"/>')
    for ext in sorted(media_exts):
        if ext in CT_TYPES:
            lines.append('<Default Extension="%s" ContentType="%s"/>' % (ext, CT_TYPES[ext]))
    for part in parts:
        content_type = None
        if part == "/ppt/presentation.xml":
            content_type = CT_OVERRIDES["/ppt/presentation.xml"]
        elif part == "/docProps/core.xml":
            content_type = ("application/vnd.openxmlformats-package.core-properties+xml")
        elif part == "/docProps/app.xml":
            content_type = ("application/vnd.openxmlformats-officedocument.extended-properties+xml")
        elif "/slides/slide" in part:
            content_type = CT_OVERRIDES["slide"]
        elif "/slideLayouts/" in part:
            content_type = CT_OVERRIDES["slideLayout"]
        elif "/slideMasters/" in part:
            content_type = CT_OVERRIDES["slideMaster"]
        elif "/theme/" in part:
            content_type = CT_OVERRIDES["theme"]
        if content_type:
            lines.append('<Override PartName="%s" ContentType="%s"/>' % (part, content_type))
    lines.append("</Types>")
    return XML_DECL + "\n".join(lines) + "\n"


def make_rels(rels):
    """rels: [(target, rtype, rid, target_mode?)]"""
    lines = ['<Relationships xmlns="%s">' % NS_REL]
    for target, rtype, rid, mode in rels:
        tm = ' TargetMode="%s"' % mode if mode else ""
        lines.append('<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (
            rid, rtype, target, tm))
    lines.append("</Relationships>")
    return XML_DECL + "\n".join(lines) + "\n"


RTYPE_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RTYPE_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_OFFICE_DOC = RTYPE_OFFICE + "/officeDocument"
REL_SLIDE = RTYPE_OFFICE + "/slide"
REL_LAYOUT = RTYPE_OFFICE + "/slideLayout"
REL_MASTER = RTYPE_OFFICE + "/slideMaster"
REL_THEME = RTYPE_OFFICE + "/theme"
REL_IMAGE = RTYPE_OFFICE + "/image"
REL_METADATA = RTYPE_OFFICE + "/slideMetadata"


def resolve_ecw_path(ref):
    """
    ecw 引用路径('/ppt/slideLayouts/slideLayout15.xml') -> 包内路径('layout/layout15.xml')
    """
    ref = (ref or "").replace("\\", "/").strip()
    if ref.startswith("/ppt/"):
        rel = ref[len("/ppt/"):]
        # 目录前缀与文件名都要替换：slideLayouts/xxx -> layout/xxx（slideLayout15.xml -> layout15.xml）
        rel = rel.replace("slideLayouts/", "layout/") \
                 .replace("slideLayout", "layout") \
                 .replace("slideMasters/", "master/") \
                 .replace("slideMaster", "master")
        return rel
    return ref


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    ecw_path = sys.argv[1]
    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        out_path = os.path.splitext(ecw_path)[0] + ".pptx"

    if not os.path.isfile(ecw_path):
        print("错误: 文件不存在: %s" % ecw_path)
        sys.exit(1)

    zin = zipfile.ZipFile(ecw_path)
    names = zin.namelist()

    # ---- 收集部件 ----
    slides = sorted([n for n in names if re.match(r"^slide/slide\d+\.xml$", n)],
                    key=lambda n: num_of_file(n, "slide"))
    layouts = sorted([n for n in names if re.match(r"^layout/layout\d+\.xml$", n)],
                     key=lambda n: num_of_file(n, "layout"))
    masters = sorted([n for n in names if re.match(r"^master/master\d+\.xml$", n)],
                     key=lambda n: num_of_file(n, "master"))
    themes = sorted([n for n in names if re.match(r"^theme/theme\d+\.xml$", n)],
                    key=lambda n: num_of_file(n, "theme"))
    media_names = set(os.path.basename(n) for n in names if n.startswith("resource/"))

    if not slides:
        print("错误: ecw 中未找到 slide/*.xml，可能不是有效的 ecw 课件包")
        sys.exit(1)

    # ---- 解析 XML ----
    def load(n):
        return ET.fromstring(zin.read(n))

    slide_els = {n: load(n) for n in slides}
    layout_els = {n: load(n) for n in layouts}
    master_els = {n: load(n) for n in masters}
    theme_els = {n: load(n) for n in themes}

    # ---- 关系解析 ----
    def ref_of(el, tag):
        return resolve_ecw_path(findtext(el, tag))

    # slide -> layout / layout -> master / master -> theme
    slide_layout_map = {}
    for n in slides:
        ref = ref_of(slide_els[n], "layoutID")
        # 转成标准输出路径 key: layout/layout15.xml
        slide_layout_map[n] = ref

    layout_master_map = {}
    for n in layouts:
        ref = ref_of(layout_els[n], "masterID")
        layout_master_map[n] = ref

    master_theme_map = {}
    for n in masters:
        ref = ref_of(master_els[n], "themeID")
        master_theme_map[n] = ref

    # master 下挂的 layouts（按数字排序）
    master_layouts = {}
    for n in layouts:
        m = layout_master_map.get(n)
        master_layouts.setdefault(m, []).append(n)
    for m in master_layouts:
        master_layouts[m].sort(key=lambda x: num_of_file(x, "layout"))

    # 输出部件名
    def out_slide(i):
        return "ppt/slides/slide%d.xml" % (i + 1)

    def out_layout(i):
        return "ppt/slideLayouts/slideLayout%d.xml" % (i + 1)

    def out_master(i):
        return "ppt/slideMasters/slideMaster%d.xml" % (i + 1)

    def out_theme(i):
        return "ppt/theme/theme%d.xml" % (i + 1)

    def rel_target(out_part, from_dir):
        """把包内绝对部件名转成相对于 from_dir 的关系目标
        out_part: 'ppt/slideLayouts/slideLayout5.xml'
        from_dir: 'ppt' / 'ppt/slides' / 'ppt/slideMasters' / 'ppt/slideLayouts'
        """
        rel = out_part[len("ppt/"):]
        depth = from_dir.count("/")
        return "../" * depth + rel

    # 建立 ecw 内文件名 -> 输出部件名 映射
    layout_out = {}
    for idx, n in enumerate(layouts):
        layout_out[n] = out_layout(idx)
    master_out = {}
    for idx, n in enumerate(masters):
        master_out[n] = out_master(idx)
    theme_out = {}
    for idx, n in enumerate(themes):
        theme_out[n] = out_theme(idx)

    def out_name_of_ref(ref):
        """'layout/layout15.xml' -> 'ppt/slideLayouts/slideLayoutN.xml'"""
        if ref in layout_out:
            return layout_out[ref]
        if ref in master_out:
            return master_out[ref]
        if ref in theme_out:
            return theme_out[ref]
        return None

    # ---- 构建 slide 内容 + 每 slide 图片 rels ----
    slide_bodies = []
    slide_rel_lists = []      # [(target, rtype, rid, mode)]
    all_media_exts = set()
    total_shapes = 0

    for idx, n in enumerate(slides):
        rels = []
        layout_id_xml = ""
        lref = slide_layout_map.get(n)
        lout = out_name_of_ref(lref)
        rid_start = 1
        if lout:
            layout_id_xml = '<p:sldLayoutId r:id="rId1"/>'
            rels.append((rel_target(lout, "ppt/slides"), REL_LAYOUT, "rId1", None))
            rid_start = 2
        ctx = Context(rid_start)
        body = build_slide(slide_els[n], ctx, media_names, layout_id_xml)
        slide_bodies.append(body)
        total_shapes += body.count("<p:sp>") + body.count("<p:pic>") + body.count("<p:grpSp>")
        for src, media, rid in ctx.images:
            rels.append(("../media/%s" % media, REL_IMAGE, rid, None))
            ext = media.rsplit(".", 1)[-1].lower() if "." in media else ""
            if ext:
                all_media_exts.add(ext)
        slide_rel_lists.append(rels)

    # ---- 构建 layout ----
    layout_bodies = {}
    layout_rel_lists = {}
    for n in layouts:
        rels = []
        mref = layout_master_map.get(n)
        mout = out_name_of_ref(mref)
        rid_start = 1
        if mout:
            rels.append((rel_target(mout, "ppt/slideLayouts"), REL_MASTER, "rId1", None))
            rid_start = 2
        ctx = Context(rid_start)
        body = build_layout(layout_els[n], ctx, media_names)
        layout_bodies[n] = body
        for src, media, rid in ctx.images:
            rels.append(("../media/%s" % media, REL_IMAGE, rid, None))
            ext = media.rsplit(".", 1)[-1].lower() if "." in media else ""
            if ext:
                all_media_exts.add(ext)
        layout_rel_lists[n] = rels

    # ---- 构建 master ----
    master_bodies = {}
    master_rel_lists = {}
    for n in masters:
        m_layouts = master_layouts.get(n, [])
        valid_layouts = [ln for ln in m_layouts if out_name_of_ref(ln)]
        layout_ids = []
        rels = []
        tref = master_theme_map.get(n)
        tout = out_name_of_ref(tref)
        if tout:
            rels.append((rel_target(tout, "ppt/slideMasters"), REL_THEME, "rId1", None))
        for i, ln in enumerate(valid_layouts, start=2):
            lout = out_name_of_ref(ln)
            layout_ids.append('<p:sldLayoutId id="%d" r:id="rId%d"/>' % (
                2147483648 + i * 1000, i))
            rels.append((rel_target(lout, "ppt/slideMasters"), REL_LAYOUT,
                         "rId%d" % i, None))
        ctx = Context(2 + len(valid_layouts))
        body = build_master(master_els[n], ctx, media_names, "".join(layout_ids))
        master_bodies[n] = body
        # master 内图片(如母版水印图)的 rId 与 XML 中 r:embed 完全一致
        for src, media, rid in ctx.images:
            rels.append(("../media/%s" % media, REL_IMAGE, rid, None))
            ext = media.rsplit(".", 1)[-1].lower() if "." in media else ""
            if ext:
                all_media_exts.add(ext)
        master_rel_lists[n] = rels

    # ---- theme ----
    theme_bodies = {n: build_theme(theme_els[n]) for n in themes}

    # ---- presentation.xml ----
    master_out_parts = [master_out[n] for n in masters]
    pres_rels = []
    rid_counter = 1
    master_id_xml = []
    for n in masters:
        pres_rels.append((rel_target(master_out[n], "ppt"), REL_MASTER,
                          "rId%d" % rid_counter, None))
        master_id_xml.append('<p:sldMasterId id="%d" r:id="rId%d"/>' % (
            2147483648 + rid_counter * 16, rid_counter))
        rid_counter += 1
    slide_id_xml = []
    for idx in range(len(slides)):
        pres_rels.append((rel_target(out_slide(idx), "ppt"), REL_SLIDE,
                          "rId%d" % rid_counter, None))
        slide_id_xml.append('<p:sldId id="%d" r:id="rId%d"/>' % (256 + idx, rid_counter))
        rid_counter += 1

    def default_text_style():
        sizes = {1: 3200, 2: 2800, 3: 2400, 4: 2000, 5: 1800, 6: 1800,
                 7: 1800, 8: 1800, 9: 1600}
        lvls = ""
        for i in range(1, 10):
            lvls += '<a:lvl%dpPr><a:defRPr sz="%d"/></a:lvl%dpPr>' % (i, sizes[i], i)
        return ("<p:defaultTextStyle><a:defPPr><a:defRPr lang=\"zh-CN\"/>%s"
                "</a:defPPr></p:defaultTextStyle>") % lvls

    presentation_xml = (
        '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" saveSubsetFonts="1">'
        "<p:sldMasterIdLst>%s</p:sldMasterIdLst>"
        "<p:sldIdLst>%s</p:sldIdLst>"
        '<p:sldSz cx="12192000" cy="6858000"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        "%s"
        "</p:presentation>"
    ) % (NS_A, NS_R, NS_P, "".join(master_id_xml), "".join(slide_id_xml),
         default_text_style())

    # ---- 写入 zip ----
    if os.path.exists(out_path):
        os.remove(out_path)
    zout = zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED)

    # 收集所有部件路径
    all_parts = ["/ppt/presentation.xml", "/docProps/core.xml", "/docProps/app.xml"]
    all_parts += ["/" + out_master(i) for i in range(len(masters))]
    all_parts += ["/" + out_layout(i) for i in range(len(layouts))]
    all_parts += ["/" + out_slide(i) for i in range(len(slides))]
    all_parts += ["/" + out_theme(i) for i in range(len(themes))]

    zout.writestr("[Content_Types].xml", make_content_types(all_parts, all_media_exts))
    zout.writestr("_rels/.rels", make_rels([
        ("ppt/presentation.xml", REL_OFFICE_DOC, "rId1", None),
        ("docProps/core.xml",
         "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
         "rId2", None),
        ("docProps/app.xml",
         "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
         "rId3", None)]))

    # 标准文档属性（提升 WPS / PowerPoint 兼容性）
    zout.writestr("docProps/core.xml", XML_DECL + (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>%s</dc:title><dc:creator>ecw2pptx</dc:creator>"
        "<cp:lastModifiedBy>ecw2pptx</cp:lastModifiedBy>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>'
        "</cp:coreProperties>\n") % esc(os.path.splitext(os.path.basename(ecw_path))[0]))
    zout.writestr("docProps/app.xml", XML_DECL + (
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>ecw2pptx</Application><Slides>%d</Slides>"
        "<PresentationFormat>宽屏</PresentationFormat></Properties>\n") % len(slides))

    zout.writestr("ppt/presentation.xml", XML_DECL + presentation_xml + "\n")
    zout.writestr("ppt/_rels/presentation.xml.rels", make_rels(pres_rels))

    # masters
    for idx, n in enumerate(masters):
        zout.writestr(out_master(idx), XML_DECL + master_bodies[n] + "\n")
        zout.writestr("ppt/slideMasters/_rels/slideMaster%d.xml.rels" % (idx + 1),
                      make_rels(master_rel_lists[n]))
    # layouts
    for idx, n in enumerate(layouts):
        zout.writestr(out_layout(idx), XML_DECL + layout_bodies[n] + "\n")
        zout.writestr("ppt/slideLayouts/_rels/slideLayout%d.xml.rels" % (idx + 1),
                      make_rels(layout_rel_lists[n]))
    # slides
    for idx in range(len(slides)):
        zout.writestr(out_slide(idx), XML_DECL + slide_bodies[idx] + "\n")
        zout.writestr("ppt/slides/_rels/slide%d.xml.rels" % (idx + 1),
                      make_rels(slide_rel_lists[idx]))
    # themes
    for idx, n in enumerate(themes):
        zout.writestr(out_theme(idx), XML_DECL + theme_bodies[n] + "\n")

    # media：收集 slide/layout/master 实际引用的图片
    copied = set()
    all_rel_lists = list(slide_rel_lists) + list(layout_rel_lists.values()) \
        + list(master_rel_lists.values())
    for rels in all_rel_lists:
        for target, rtype, rid, mode in rels:
            if rtype == REL_IMAGE:
                media = target.replace("../media/", "")
                if media in copied:
                    continue
                src_path = "resource/" + media
                if src_path in names:
                    zout.writestr("ppt/media/" + media, zin.read(src_path))
                    copied.add(media)
                else:
                    print("警告: 图片缺失: %s" % src_path)

    zout.close()
    zin.close()

    # ---- 报告 ----
    print("转换完成: %s" % out_path)
    print("  幻灯片: %d 页" % len(slides))
    print("  母版:   %d 个 (master: %s)" % (len(masters), ", ".join(masters)))
    print("  版式:   %d 个 (layout: %s)" % (len(layouts), ", ".join(layouts)))
    print("  主题:   %d 个" % len(themes))
    print("  内嵌图片: %d 个" % len(copied))
    print("  形状总数: %d" % total_shapes)


if __name__ == "__main__":
    main()
#（注：内容由AI生成）
