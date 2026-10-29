"""Schema order of WordprocessingML children (ECMA-376 Part 1). Word rejects a file as
"unreadable content" when children are out of order, so every insert goes through ensure()."""
from lxml import etree
from .pkg import W

ORDER = {
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
            "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
            "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
            "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
            "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "style": ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden",
              "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr",
              "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
             "tcPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
              "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
              "tblCaption", "tblDescription", "tblPrChange"],
    "tblBorders": ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
               "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
               "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
    "lvl": ["start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText", "lvlPicBulletId", "legacy",
            "lvlJc", "pPr", "rPr"],
    "abstractNum": ["nsid", "multiLevelType", "tmpl", "name", "styleLink", "numStyleLink", "lvl"],
    "numbering": ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"],
}
# trPr is an unbounded choice: any order is valid.


def q(local):
    return f"{{{W}}}{local}"


def local(el):
    return etree.QName(el).localname


def ensure(parent, name, attrs=None, first_in_group=False):
    """Return the child <w:name>, creating it in schema order if missing. attrs: {local: value} set on it."""
    el = parent.find(q(name))
    if el is None:
        el = etree.Element(q(name))
        insert_ordered(parent, el)
    if attrs:
        for k, v in attrs.items():
            el.set(q(k), str(v))
    return el


def insert_ordered(parent, el):
    order = ORDER.get(local(parent))
    if not order:
        parent.append(el)
        return el
    name = local(el)
    idx = order.index(name)
    later = set(order[idx + 1:])
    for i, ch in enumerate(parent):
        if not isinstance(ch.tag, str):
            continue
        if ch.tag.startswith(f"{{{W}}}") and local(ch) in later:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def remove(parent, *names):
    n = 0
    for name in names:
        for el in parent.findall(q(name)):
            parent.remove(el)
            n += 1
    return n


def check_order(root):
    """Return a list of order problems for the containers we know. Used by tests and before every save."""
    problems = []
    for parent_name, order in ORDER.items():
        for parent in root.iter(q(parent_name)):
            last = -1
            for ch in parent:
                if not isinstance(ch.tag, str) or not ch.tag.startswith(f"{{{W}}}"):
                    continue
                n = local(ch)
                if n not in order:
                    continue
                i = order.index(n)
                if parent_name == "sectPr" and n == "footerReference":
                    i = 0  # header and footer references may interleave
                if i < last:
                    problems.append(f"{parent_name}: {n} after {order[last]}")
                    break
                last = i
    return problems
