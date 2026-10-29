"""Word operations for the feasibility lab. Each op edits a Package in memory and returns a report with
'checks': read-back expressions that hidden Word evaluates on the saved file to prove the settings are real.

Every op declares ALLOWED: the fingerprint facets it may change. safewrite.apply refuses anything else.
"""
import random
import re
from copy import deepcopy

from lxml import etree
from ..pkg import W, NS
from ..ordered import ensure, insert_ordered, remove, q, local
from ..fingerprint import visible_text

WP = q("p"); WR = q("r"); WT = q("t")


def bgr(hexrgb):
    """Word COM colors are BGR integers."""
    r, g, b = int(hexrgb[0:2], 16), int(hexrgb[2:4], 16), int(hexrgb[4:6], 16)
    return b * 65536 + g * 256 + r


def norm(name):
    return " ".join(name.lower().replace(" - ", " ").split())


# ------------------------------------------------------------------ styles
def styles_root(pkg):
    part = pkg.word_part("/styles")
    pkg.touch(part)
    return pkg.tree(part)


def find_style(root, name):
    for s in root.findall("w:style", NS):
        n = s.find("w:name", NS)
        if n is not None and norm(n.get(q("val"))) == norm(name):
            return s
    return None


def new_style(root, name, stype="paragraph", based_on="Normal", next_style=None):
    sid = re.sub(r"[^A-Za-z0-9]", "", name) or "S"
    existing = {s.get(q("styleId")) for s in root.findall("w:style", NS)}
    base, i = sid, 1
    while sid in existing:
        i += 1
        sid = f"{base}{i}"
    s = etree.SubElement(root, q("style"))
    s.set(q("type"), stype)
    s.set(q("customStyle"), "1")
    s.set(q("styleId"), sid)
    ensure(s, "name", {"val": name})
    if based_on:
        b = find_style(root, based_on)
        ensure(s, "basedOn", {"val": b.get(q("styleId")) if b is not None else based_on})
    if next_style:
        ensure(s, "next", {"val": next_style})
    ensure(s, "qFormat")
    return s


def set_run_props(rpr, font=None, size=None, bold=None, italic=None, color=None):
    if font is not None:
        rf = ensure(rpr, "rFonts")
        for a in list(rf.attrib):
            if "Theme" in a or local_attr(a) in ("hint",):
                del rf.attrib[a]
        for a in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(q(a), font)
    if size is not None:
        ensure(rpr, "sz", {"val": int(round(size * 2))})
        ensure(rpr, "szCs", {"val": int(round(size * 2))})
    if bold is not None:
        for t in ("b", "bCs"):
            el = ensure(rpr, t)
            el.attrib.clear()
            if not bold:
                el.set(q("val"), "0")
    if italic is not None:
        for t in ("i", "iCs"):
            el = ensure(rpr, t)
            el.attrib.clear()
            if not italic:
                el.set(q("val"), "0")
    if color is not None:
        c = ensure(rpr, "color")
        c.attrib.clear()
        c.set(q("val"), color)


def local_attr(a):
    return a.split("}")[-1]


def set_para_props(ppr, before=None, after=None, keep_next=None, line=None, jc=None):
    if before is not None or after is not None or line is not None:
        sp = ensure(ppr, "spacing")
        for a in ("beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"):
            sp.attrib.pop(q(a), None)
        if before is not None:
            sp.set(q("before"), str(int(round(before * 20))))
        if after is not None:
            sp.set(q("after"), str(int(round(after * 20))))
        if line is not None:
            sp.set(q("line"), str(int(round(line * 240))))
            sp.set(q("lineRule"), "auto")
    if keep_next is not None:
        el = ensure(ppr, "keepNext")
        el.attrib.clear()
        if not keep_next:
            el.set(q("val"), "0")
    if jc is not None:
        ensure(ppr, "jc", {"val": jc})


def sync_linked(pkg, root, para_style):
    """Word keeps a linked character style ("Heading 1 Char") in step with its paragraph style's run
    properties. Do the same, so Word and the file agree."""
    lk = para_style.find("w:link", NS)
    if para_style.get(q("type")) != "paragraph" or lk is None:
        return None
    ch = next((x for x in root.findall("w:style", NS) if x.get(q("styleId")) == lk.get(q("val"))
               and x.get(q("type")) == "character"), None)
    if ch is None:
        return None
    src = para_style.find("w:rPr", NS)
    old = ch.find("w:rPr", NS)
    if src is None:
        return None
    new = deepcopy(src)
    remove(new, "rStyle")
    if old is not None and _sig(old) == _sig(new):
        return None
    pkg.mark(ch)
    if old is not None:
        ch.remove(old)
    insert_ordered(ch, new)
    return ch.find("w:name", NS).get(q("val"))


def op_styles(pkg, spec):
    """spec: {style name: {font,size,bold,italic,color,before,after,keep_next,line,jc,type,based_on}}"""
    root = styles_root(pkg)
    checks, created = [], []
    for name, props in spec.items():
        s = find_style(root, name)
        if s is None:
            s = new_style(root, name, props.get("type", "paragraph"), props.get("based_on", "Normal"))
            created.append(name)
        pkg.mark(s)
        rpr = ensure(s, "rPr")
        set_run_props(rpr, props.get("font"), props.get("size"), props.get("bold"), props.get("italic"), props.get("color"))
        if s.get(q("type")) == "paragraph":
            ppr = ensure(s, "pPr")
            set_para_props(ppr, props.get("before"), props.get("after"), props.get("keep_next"), props.get("line"), props.get("jc"))
        linked = sync_linked(pkg, root, s)
        e = f"doc.Styles({name!r})"
        if linked and "font" in props:
            checks.append({"expr": f"doc.Styles({linked!r}).Font.Name", "expect": props["font"]})
        if "font" in props: checks.append({"expr": e + ".Font.Name", "expect": props["font"]})
        if "size" in props: checks.append({"expr": e + ".Font.Size", "expect": float(props["size"])})
        if "bold" in props: checks.append({"expr": e + ".Font.Bold", "expect": -1 if props["bold"] else 0})
        if "italic" in props: checks.append({"expr": e + ".Font.Italic", "expect": -1 if props["italic"] else 0})
        if "color" in props: checks.append({"expr": e + ".Font.Color", "expect": bgr(props["color"])})
        if "before" in props: checks.append({"expr": e + ".ParagraphFormat.SpaceBefore", "expect": float(props["before"])})
        if "after" in props: checks.append({"expr": e + ".ParagraphFormat.SpaceAfter", "expect": float(props["after"])})
        if "keep_next" in props: checks.append({"expr": e + ".ParagraphFormat.KeepWithNext", "expect": -1 if props["keep_next"] else 0})
        if "based_on" in props: checks.append({"expr": e + ".BaseStyle.NameLocal", "expect": props["based_on"]})
    return {"created": created, "checks": checks}


OP_STYLES_ALLOWED = ()  # styles.xml is not a story part; no story facet may change


def styles_in_use(pkg):
    used = set()
    for story in pkg.stories():
        r = pkg.tree(story)
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for el in r.iter(q(tag)):
                used.add(el.get(q("val")))
    root = pkg.tree(pkg.word_part("/styles"))
    for s in root.findall("w:style", NS):
        for tag in ("basedOn", "next", "link"):
            el = s.find(f"w:{tag}", NS)
            if el is not None:
                used.add(el.get(q("val")))
    nb = pkg.word_part("/numbering")
    if nb:
        for el in pkg.tree(nb).iter(q("pStyle"), q("rStyle"), q("styleLink"), q("numStyleLink")):
            used.add(el.get(q("val")))
    st = pkg.word_part("/settings")
    if st:
        for el in pkg.tree(st).iter(q("defaultTableStyle"), q("attachedSchema")):
            used.add(el.get(q("val")))
    # styles named inside fields: STYLEREF "Name" and TOC \t "Name,1,Name2,2"
    names_in_fields = set()
    from ..fingerprint import fields as _fields
    for story in pkg.stories():
        for ins in _fields(pkg.tree(story)):
            for m in re.finditer(r'STYLEREF\s+"([^"]+)"', ins, re.I):
                names_in_fields.add(norm(m.group(1)))
            for m in re.finditer(r'\\t\s+"([^"]+)"', ins):
                for piece in m.group(1).split(","):
                    if not piece.strip().isdigit():
                        names_in_fields.add(norm(piece))
    for s_ in root.findall("w:style", NS):
        nm = s_.find("w:name", NS)
        if nm is not None and norm(nm.get(q("val"))) in names_in_fields:
            used.add(s_.get(q("styleId")))
    return used


def op_remove_unused_styles(pkg, names):
    root = styles_root(pkg)
    used = styles_in_use(pkg)
    removed, kept, checks = [], [], []
    for name in names:
        s = find_style(root, name)
        if s is None:
            continue
        if s.get(q("styleId")) in used or s.get(q("default")) in ("1", "true"):
            kept.append(name)
            continue
        pkg.mark(s)
        root.remove(s)
        removed.append(name)
    for name in removed:
        checks.append({"expr": f"style_in_use(doc, {name!r})", "expect": False})
    return {"removed": removed, "kept_in_use": kept, "checks": checks}


def op_hide_styles(pkg, names):
    root = styles_root(pkg)
    checks = []
    for name in names:
        s = find_style(root, name)
        if s is None:
            continue
        pkg.mark(s)
        remove(s, "qFormat")
        ensure(s, "semiHidden")
        ensure(s, "unhideWhenUsed")
        checks.append({"expr": f"doc.Styles({name!r}).QuickStyle", "expect": False})
    return {"hidden": names, "checks": checks}


# ------------------------------------------------------------------ heading numbering
HEADING_LEVELS = [  # (style, numFmt, lvlText, number at in, text at in, restart)
    ("Heading 1", "decimal", "%1", 0, 0.5, None),
    ("Heading 2", "decimal", "%1.%2", 0, 0.6, None),
    ("Heading 3", "decimal", "%1.%2.%3", 0, 0.75, None),
    ("Heading 4", "decimal", "%1.%2.%3.%4", 0, 0.9, None),
    ("Heading 5", "decimal", "%1.%2.%3.%4.%5", 0, 1.0, None),
    ("Heading 6", "decimal", "%1.%2.%3.%4.%5.%6", 0, 1.1, None),
    ("Heading 7", "upperLetter", "Appendix %7", 0, 1.25, 0),
    ("Heading 8", "decimal", "%7.%8", 0, 0.6, None),
    ("Heading 9", "decimal", "%7.%8.%9", 0, 0.75, None),
]


def _build_heading_abstract(sroot, levels, nsid, aid):
    a = etree.Element(q("abstractNum"))
    a.set(q("abstractNumId"), str(aid))
    ensure(a, "nsid", {"val": nsid})
    ensure(a, "multiLevelType", {"val": "multilevel"})
    for ilvl, (sname, fmt, text, num_at, text_at, restart) in enumerate(levels):
        st = find_style(sroot, sname)
        if st is None:
            raise ValueError(f"style {sname} missing")
        lvl = etree.SubElement(a, q("lvl"))
        lvl.set(q("ilvl"), str(ilvl))
        ensure(lvl, "start", {"val": 1})
        ensure(lvl, "numFmt", {"val": fmt})
        if restart is not None:
            ensure(lvl, "lvlRestart", {"val": restart})
        ensure(lvl, "pStyle", {"val": st.get(q("styleId"))})
        ensure(lvl, "suff", {"val": "tab"})
        ensure(lvl, "lvlText", {"val": text})
        ensure(lvl, "lvlJc", {"val": "left"})
        ppr = ensure(lvl, "pPr")
        tabs = ensure(ppr, "tabs")
        t = etree.SubElement(tabs, q("tab")); t.set(q("val"), "num"); t.set(q("pos"), str(int(text_at * 1440)))
        ensure(ppr, "ind", {"left": int(text_at * 1440), "hanging": int((text_at - num_at) * 1440)})
    return a


def _sig(el):
    """Prefix-independent signature of an element tree (namespace URI + local names, attributes, text)."""
    return (el.tag, tuple(sorted(el.attrib.items())), (el.text or "").strip(),
            tuple(_sig(c) for c in el if isinstance(c.tag, str)))


def _lvl_signature(abstract):
    """Levels of an abstractNum without its identity (nsid, tmpl), for comparing designs."""
    return tuple(_sig(l) for l in abstract.findall("w:lvl", NS))


def op_heading_numbering(pkg, levels=HEADING_LEVELS):
    nb_part = pkg.word_part("/numbering")
    pkg.touch(nb_part)
    nb = pkg.tree(nb_part)
    sroot = styles_root(pkg)
    built = _build_heading_abstract(sroot, levels, "00000000", "0")
    # already in place? then do nothing (idempotent)
    h1 = find_style(sroot, levels[0][0])
    cur = h1.find("w:pPr/w:numPr/w:numId", NS) if h1 is not None else None
    if cur is not None:
        num = next((n for n in nb.findall("w:num", NS) if n.get(q("numId")) == cur.get(q("val"))), None)
        absn = None
        if num is not None and num.find("w:lvlOverride", NS) is None:
            aidv = num.find("w:abstractNumId", NS).get(q("val"))
            absn = next((a for a in nb.findall("w:abstractNum", NS) if a.get(q("abstractNumId")) == aidv), None)
        linked = all((find_style(sroot, sn) is not None and
                      (find_style(sroot, sn).find("w:pPr/w:numPr/w:numId", NS) is not None) and
                      find_style(sroot, sn).find("w:pPr/w:numPr/w:numId", NS).get(q("val")) == cur.get(q("val")))
                     for sn, *_ in levels)
        if absn is not None and linked and _lvl_signature(absn) == _lvl_signature(built):
            return {"unchanged": "heading numbering already matches", "checks": [{"expr": "heading_list_strings(doc)", "expect": "computed"}]}
    abs_ids = [int(a.get(q("abstractNumId"))) for a in nb.findall("w:abstractNum", NS)]
    num_ids = [int(n.get(q("numId"))) for n in nb.findall("w:num", NS)]
    nsids = {a.find("w:nsid", NS).get(q("val")) for a in nb.findall("w:abstractNum", NS) if a.find("w:nsid", NS) is not None}
    aid, nid = max(abs_ids + [-1]) + 1, max(num_ids + [0]) + 1
    rnd = random.Random()
    nsid = "%08X" % rnd.getrandbits(32)
    while nsid in nsids:
        nsid = "%08X" % rnd.getrandbits(32)
    a = _build_heading_abstract(sroot, levels, nsid, aid)
    checks = []
    for ilvl, (sname, fmt, text, num_at, text_at, restart) in enumerate(levels):
        st = find_style(sroot, sname)
        pkg.mark(st)
        sppr = ensure(st, "pPr")
        numpr = ensure(sppr, "numPr")
        for ch in list(numpr):
            numpr.remove(ch)
        il = etree.SubElement(numpr, q("ilvl")); il.set(q("val"), str(ilvl))
        ni = etree.SubElement(numpr, q("numId")); ni.set(q("val"), str(nid))
        # a style's own indents would override the list level; remove them so the list rules
        remove(sppr, "ind")
        lv = ilvl + 1
        checks.append({"expr": f"doc.Styles({sname!r}).ListTemplate.ListLevels({lv}).NumberFormat", "expect": text})
        checks.append({"expr": f"doc.Styles({sname!r}).ListTemplate.ListLevels({lv}).TextPosition", "expect": float(text_at * 72)})
        checks.append({"expr": f"doc.Styles({sname!r}).ListTemplate.ListLevels({lv}).TrailingCharacter", "expect": 0})
    # styles based on a heading (for example TOC Heading) inherit its numbering: switch it off unless set
    heading_ids = {find_style(sroot, lv[0]).get(q("styleId")) for lv in levels}
    by_id = {s.get(q("styleId")): s for s in sroot.findall("w:style", NS)}
    unnumbered = []
    for s in sroot.findall("w:style", NS):
        if s.get(q("styleId")) in heading_ids or s.get(q("type")) != "paragraph":
            continue
        cur, seen = s, set()
        while cur is not None:
            b = cur.find("w:basedOn", NS)
            if b is None or b.get(q("val")) in seen:
                break
            seen.add(b.get(q("val")))
            if b.get(q("val")) in heading_ids:
                pkg.mark(s)
                sp = ensure(s, "pPr")
                if sp.find("w:numPr", NS) is None:
                    npr = ensure(sp, "numPr")
                    ni0 = etree.SubElement(npr, q("numId")); ni0.set(q("val"), "0")
                    unnumbered.append(s.find("w:name", NS).get(q("val")))
                break
            cur = by_id.get(b.get(q("val")))
    for name in unnumbered:
        checks.append({"expr": f"doc.Styles({name!r}).ListTemplate is None", "expect": True})
    # abstractNum elements must come before num elements
    first_num = nb.find("w:num", NS)
    if first_num is not None:
        first_num.addprevious(a)
    else:
        nb.append(a)
    n = etree.SubElement(nb, q("num"))
    n.set(q("numId"), str(nid))
    an = etree.SubElement(n, q("abstractNumId")); an.set(q("val"), str(aid))
    pkg.mark(a); pkg.mark(n)
    checks.append({"expr": "heading_list_strings(doc)", "expect": "computed"})
    return {"abstractNumId": aid, "numId": nid, "nsid": nsid, "numbering_switched_off": unnumbered, "checks": checks}


# ------------------------------------------------------------------ fields
def field_groups(container):
    """Top-level complex fields directly inside a paragraph: list of dicts with the runs of each part."""
    out, cur, depth = [], None, 0
    for r in container.iter(WR):
        if r.getparent() is not None and local(r.getparent()) not in ("p", "hyperlink", "smartTag", "ins", "fldSimple"):
            pass
        fc = r.find("w:fldChar", NS)
        if fc is not None:
            t = fc.get(q("fldCharType"))
            if t == "begin":
                depth += 1
                if depth == 1:
                    cur = {"begin": r, "instr": [], "result": [], "sep": None, "end": None}
                continue
            if t == "separate" and depth == 1:
                cur["sep"] = r
                continue
            if t == "end":
                depth -= 1
                if depth == 0 and cur is not None:
                    cur["end"] = r
                    out.append(cur)
                    cur = None
                continue
        if cur is not None and depth == 1:
            if r.find("w:instrText", NS) is not None and cur["sep"] is None:
                cur["instr"].append(r)
            elif cur["sep"] is not None:
                cur["result"].append(r)
    return out


def instr_of(g):
    return " ".join("".join(r.find("w:instrText", NS).text or "" for r in g["instr"]).split())


def set_instr(g, new):
    first = g["instr"][0]
    first.find("w:instrText", NS).text = f" {new} "
    first.find("w:instrText", NS).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for r in g["instr"][1:]:
        r.getparent().remove(r)
    g["instr"] = [first]


def op_hf_clean(pkg):
    """Headers and footers: STYLEREF loses \\* MERGEFORMAT; PAGE loses every \\* format switch."""
    changed = []
    for part in pkg.stories():
        root = pkg.tree(part)
        if local(root) not in ("hdr", "ftr"):
            continue
        touched = False
        for g in field_groups(root):
            ins = instr_of(g)
            word = ins.split(" ")[0].upper() if ins else ""
            new = ins
            if word == "STYLEREF":
                new = re.sub(r"\s*\\\*\s*MERGEFORMAT", "", ins, flags=re.I)
            elif word == "PAGE":
                new = "PAGE"
            if new != ins:
                pkg.mark(next(g["begin"].iterancestors(WP)))
                set_instr(g, new)
                changed.append((part, ins, new))
                touched = True
        for fs in root.iter(q("fldSimple")):
            ins = " ".join(fs.get(q("instr"), "").split())
            if ins.split(" ")[0].upper() == "PAGE" and ins.upper() != "PAGE":
                pkg.mark(next(fs.iterancestors(WP)))
                fs.set(q("instr"), " PAGE ")
                changed.append((part, ins, "PAGE"))
                touched = True
        if touched:
            pkg.touch(part)
    checks = [{"expr": "hf_field_codes(doc)", "expect": "no MERGEFORMAT on STYLEREF and no switch on PAGE"}]
    return {"rewritten": changed, "field_changes": changed, "checks": checks}


# ------------------------------------------------------------------ sections and page numbers
def body_sections(pkg):
    """[(sectPr element, [paragraphs of that section in order])]"""
    root = pkg.tree(pkg.main)
    body = root.find("w:body", NS)
    out, cur = [], []
    for el in body:
        if local(el) == "p":
            cur.append(el)
            sp = el.find("w:pPr/w:sectPr", NS)
            if sp is not None:
                out.append((sp, cur)); cur = []
        elif local(el) == "tbl":
            cur.append(el)
        elif local(el) == "sectPr":
            out.append((el, cur)); cur = []
    return out


def para_style(p):
    ps = p.find("w:pPr/w:pStyle", NS)
    return ps.get(q("val")) if ps is not None else None


def first_style(paras):
    for el in paras:
        if local(el) == "p" and (visible_text(el).strip("¶").strip() or para_style(el)):
            return para_style(el)
        if local(el) == "tbl":
            return "table"
    return None


def classify_sections(pkg):
    kinds = []
    seen_body = False
    in_appendix = False
    for i, (sp, paras) in enumerate(body_sections(pkg)):
        fs = first_style(paras)
        if i == 0:
            kinds.append("cover")
        elif fs == "TOCHeading":
            kinds.append("front")
        elif fs == "Heading1":
            kinds.append("chapter_start"); seen_body = True
        elif fs == "Heading7":
            kinds.append("appendix_start"); in_appendix = True
        else:
            kinds.append("appendix_cont" if in_appendix else ("chapter_cont" if seen_body else "front"))
    return kinds


PN = {  # kind -> pgNumType attributes ; chapStyle is the heading level
    "cover": None,
    "front": {"fmt": "lowerRoman", "start": "1"},
    "chapter_start": {"fmt": "decimal", "start": "1", "chapStyle": "1", "chapSep": "hyphen"},
    "chapter_cont": {"fmt": "decimal", "chapStyle": "1", "chapSep": "hyphen"},
    "appendix_start": {"fmt": "decimal", "start": "1", "chapStyle": "7", "chapSep": "hyphen"},
    "appendix_cont": {"fmt": "decimal", "chapStyle": "7", "chapSep": "hyphen"},
}


def op_page_numbers(pkg):
    pkg.touch(pkg.main)
    kinds = classify_sections(pkg)
    checks = []
    for i, ((sp, _), kind) in enumerate(zip(body_sections(pkg), kinds)):
        owner = sp.getparent().getparent() if local(sp.getparent()) == "pPr" else sp
        pkg.mark(owner)  # the paragraph that carries this section break (or the final sectPr)
        remove(sp, "pgNumType")
        spec = PN[kind]
        if spec:
            el = etree.Element(q("pgNumType"))
            for k, v in spec.items():
                el.set(q(k), v)
            insert_ordered(sp, el)
        n = i + 1
        e = f"doc.Sections({n}).Footers(1).PageNumbers"
        if spec:
            checks.append({"expr": e + ".NumberStyle", "expect": 2 if spec["fmt"] == "lowerRoman" else 0})
            checks.append({"expr": e + ".RestartNumberingAtSection", "expect": "start" in spec})
            if "start" in spec:
                checks.append({"expr": e + ".StartingNumber", "expect": 1})
            checks.append({"expr": e + ".IncludeChapterNumber", "expect": "chapStyle" in spec})
            if "chapStyle" in spec:
                checks.append({"expr": e + ".HeadingLevelForChapter", "expect": int(spec["chapStyle"]) - 1})
                checks.append({"expr": e + ".ChapterPageSeparator", "expect": 0})
    return {"kinds": kinds, "checks": checks}


# ------------------------------------------------------------------ tables
ROWTYPES = {  # role -> (tag style, fill or None, vAlign, repeat header)
    "caption": ("Caption", None, "center", True),
    "header": ("PE Tbl Header", "D9E2F3", "center", True),
    "subhead": ("PE Tbl Subhead", "D9D9D9", "center", False),
    "total": ("PE Tbl Total", None, "top", False),
    "note": ("PE Tbl Note", None, "top", False),
    "body": ("PE Tbl Body", None, "top", False),
}
TAG_STYLES = {
    "PE Tbl Header": {"font": "Arial", "size": 10, "bold": True, "before": 2, "after": 2, "keep_next": True},
    "PE Tbl Subhead": {"font": "Arial", "size": 10, "bold": True, "before": 1, "after": 1, "keep_next": True},
    "PE Tbl Total": {"font": "Arial", "size": 10, "bold": True, "before": 1, "after": 1},
    "PE Tbl Note": {"font": "Arial", "size": 9, "italic": True, "before": 1, "after": 1},
    "PE Tbl Body": {"font": "Arial", "size": 10, "before": 1, "after": 1},
}
DIRECT_RUN = ("rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color")


def cell_text(tc):
    return visible_text(tc).replace("¶", " ").strip()


def is_figure_table(tbl):
    grid = tbl.findall("w:tblGrid/w:gridCol", NS)
    return len(grid) == 1 and tbl.find(".//w:drawing", NS) is not None


def has_seq(el):
    return any("SEQ" in (t.text or "") for t in el.iter(q("instrText"))) or any("SEQ" in f.get(q("instr"), "") for f in el.iter(q("fldSimple")))


def row_role(tr, r, ncols):
    tcs = tr.findall("w:tc", NS)
    if r == 0 and (tr.find(".//w:pStyle[@w:val='Caption']", NS) is not None or has_seq(tr)):
        return "caption"
    if r == 1:
        return "header"
    fills = {(s.get(q("fill")) or "").upper() for s in tr.iter(q("shd"))}
    if fills & {"D9D9D9", "BFBFBF"}:
        return "subhead"
    first = cell_text(tcs[0]) if tcs else ""
    if re.match(r"^(Total|Subtotal)\b", first):
        return "total"
    if re.match(r"^(Note|Notes|Source)\b", first):
        return "note"
    if len(tcs) == 1 and ncols > 1:
        return "subhead"
    runs = [x for x in tr.iter(WR) if x.find("w:t", NS) is not None]
    if runs and all(x.find("w:rPr/w:b", NS) is not None and x.find("w:rPr/w:b", NS).get(q("val")) not in ("0", "false") for x in runs):
        return "subhead"
    return "body"


def fix_one_table(tbl, sid, grid_color="808080", grid_sz=4):
    """Apply row types to one w:tbl element in place. Returns the roles. Used by file mode and live mode."""
    ncols = len(tbl.findall("w:tblGrid/w:gridCol", NS))
    trs = tbl.findall("w:tr", NS)
    roles = [row_role(tr, r, ncols) for r, tr in enumerate(trs)]
    for tr, role in zip(trs, roles):
        tag, fill, valign, repeat = ROWTYPES[role]
        trpr = tr.find("w:trPr", NS)
        if repeat:
            if trpr is None:
                trpr = etree.Element(q("trPr")); tr.insert(0 if tr.find("w:tblPrEx", NS) is None else 1, trpr)
            if trpr.find("w:tblHeader", NS) is None:
                etree.SubElement(trpr, q("tblHeader"))
        elif trpr is not None:
            remove(trpr, "tblHeader")
        for tc in tr.findall("w:tc", NS):
            tcpr = tc.find("w:tcPr", NS)
            if tcpr is None:
                tcpr = etree.Element(q("tcPr")); tc.insert(0, tcpr)
            remove(tcpr, "shd")
            if fill:
                ensure(tcpr, "shd", {"val": "clear", "color": "auto", "fill": fill})
            ensure(tcpr, "vAlign", {"val": valign})
            for p in tc.findall("w:p", NS):
                ppr = p.find("w:pPr", NS)
                if ppr is None:
                    ppr = etree.Element(q("pPr")); p.insert(0, ppr)
                ensure(ppr, "pStyle", {"val": sid[tag]})
                remove(ppr, "spacing", "ind", "jc")
                mark = ppr.find("w:rPr", NS)
                if mark is not None:
                    remove(mark, *DIRECT_RUN)
                for r in p.iter(WR):
                    rpr = r.find("w:rPr", NS)
                    if rpr is not None:
                        remove(rpr, *DIRECT_RUN)
                        if len(rpr) == 0:
                            r.remove(rpr)
    tblpr = tbl.find("w:tblPr", NS)
    ensure(tblpr, "tblW", {"w": 5000, "type": "pct"})
    b = ensure(tblpr, "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ensure(b, side, {"val": "single", "sz": grid_sz, "space": 0, "color": grid_color})
    return roles


def tag_style_ids(styles_tree):
    return {name: find_style(styles_tree, name).get(q("styleId")) for name in list(TAG_STYLES) + ["Caption"]}


def op_fix_tables(pkg, grid_color="808080", grid_sz=4):
    op_styles(pkg, TAG_STYLES)
    sid = tag_style_ids(pkg.tree(pkg.word_part("/styles")))
    pkg.touch(pkg.main)
    root = pkg.tree(pkg.main)
    report = {"tables": 0, "figures_skipped": 0, "roles": {}, "checks": []}
    tindex = 0
    for tbl in root.iter(q("tbl")):
        if any(True for _ in tbl.iterancestors(q("tbl"))):
            continue
        tindex += 1
        if is_figure_table(tbl):
            report["figures_skipped"] += 1
            continue
        pkg.mark(tbl)
        roles = fix_one_table(tbl, sid, grid_color, grid_sz)
        for role in roles:
            report["roles"][role] = report["roles"].get(role, 0) + 1
        report["tables"] += 1
        report["checks"].append({"expr": f"table_check(doc, {tindex}, {roles!r})", "expect": "ok"})
    return report


def para_texts(root):
    return [(p, visible_text(p)) for p in root.iter(WP)]


def op_remove_empty_notes(pkg):
    pkg.touch(pkg.main)
    root = pkg.tree(pkg.main)
    before = para_texts(root)
    removed_paras, rows = set(), []
    for tbl in root.iter(q("tbl")):
        for tr in tbl.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) == 1 and re.fullmatch(r"\s*(Note|Notes|Source)\s*:?\s*", cell_text(tcs[0]) or ""):
                if tr.find(".//w:ins", NS) is not None or tr.find(".//w:del", NS) is not None:
                    continue
                rows.append(tr)
    for tr in rows:
        removed_paras.update(id(p) for p in tr.iter(WP))
        pkg.mark(tr)
        tr.getparent().remove(tr)
    expected = [t for (p, t) in before if id(p) not in removed_paras]
    after = [t for (p, t) in para_texts(root)]
    if after != expected:
        raise RuntimeError("remove_empty_notes: paragraph sequence check failed")
    return {"rows_removed": len(rows), "checks": [{"expr": "empty_note_rows(doc)", "expect": 0}]}


OP_TABLES_ALLOWED = ("pstyles", "para_format", "table_format", "direct_run_format")
OP_NOTES_ALLOWED = ("text", "text_len", "paragraphs", "tables", "pstyles", "para_format", "table_format", "direct_run_format")


# ------------------------------------------------------------------ captions
def make_field_runs(instr, result, rpr):
    def r_(child):
        r = etree.Element(WR)
        if rpr is not None:
            r.append(deepcopy(rpr))
        r.append(child)
        return r
    def fc(t):
        e = etree.Element(q("fldChar")); e.set(q("fldCharType"), t); return e
    it = etree.Element(q("instrText")); it.text = f" {instr} "; it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t = etree.Element(WT); t.text = result
    if result != result.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return [r_(fc("begin")), r_(it), r_(fc("separate")), r_(t), r_(fc("end"))]


def text_run(text, rpr):
    r = etree.Element(WR)
    if rpr is not None:
        r.append(deepcopy(rpr))
    t = etree.SubElement(r, WT); t.text = text
    if text != text.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def caption_paragraphs(root):
    out = []
    in_app = False
    body = root.find("w:body", NS)
    for p in body.iter(WP):
        st = para_style(p)
        if st == "Heading7":
            in_app = True
        if st == "Caption":
            txt = visible_text(p).lstrip("\u00b6")
            m = re.match(r"^(Table|Figure) ", txt)
            if m:
                out.append((p, m.group(1), in_app))
    return out


def op_captions(pkg):
    pkg.touch(pkg.main)
    root = pkg.tree(pkg.main)
    fixed, skipped = 0, []
    field_changes = []
    from ..fingerprint import fields as _fields
    for p, label, appendix in caption_paragraphs(root):
        if p.find(".//w:ins", NS) is not None or p.find(".//w:del", NS) is not None:
            skipped.append("tracked change in caption"); continue
        runs = [r for r in p if local(r) == "r"]
        # label run: first run whose text starts with the label
        lab = None
        for r in runs:
            t = r.find("w:t", NS)
            if t is not None and (t.text or "").startswith(label):
                lab = r; break
        if lab is None:
            skipped.append("label not in one run"); continue
        lt = lab.find("w:t", NS)
        rest = lt.text[len(label):]
        rpr = lab.find("w:rPr", NS)
        children = list(p)
        i = children.index(lab)
        zone = []
        if rest.strip():  # typed number in the label run: "Table 4-9"
            lt.text = label + " "
            lt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            number_text = rest.strip()
            tail = ""
            m = re.match(r"^([0-9A-Z]+(?:[-.][0-9]+)?)(.*)$", number_text)
            if not m:
                skipped.append("typed number not recognized"); lt.text = label + rest; continue
            number_text, tail = m.group(1), m.group(2)
            insert_at = i + 1
            if tail:
                p.insert(insert_at, text_run(tail, rpr))
        else:
            depth, j = 0, i + 1
            while j < len(children):
                c = children[j]
                if local(c) == "r":
                    fc = c.find("w:fldChar", NS)
                    t = c.find("w:t", NS)
                    if fc is not None:
                        depth += {"begin": 1, "end": -1}.get(fc.get(q("fldCharType")), 0)
                    elif depth == 0 and (c.find("w:tab", NS) is not None or (t is not None and (t.text or "").startswith((".", ":", " ")))):
                        break
                    zone.append(c)
                j += 1
            number_text = "".join(visible_text(z) for z in zone)
            holder = etree.Element(WP)
            for z in zone:
                holder.append(deepcopy(z))
            old_fields = _fields(holder)
            insert_at = children.index(zone[0]) if zone else i + 1
            for z in zone:
                p.remove(z)
        if "-" in number_text:
            chap, seqv = number_text.split("-", 1)
        else:
            skipped.append(f"number without chapter part: {number_text!r}")
            for k, z in enumerate(zone):
                p.insert(insert_at + k, z)
            continue
        if appendix:
            sref, seq = "STYLEREF 7 \\n \\t", f"SEQ {label}App \\* ARABIC \\s 7"
        else:
            sref, seq = "STYLEREF 1 \\s", f"SEQ {label} \\* ARABIC \\s 1"
        if rest.strip():
            old_fields = []
        new = make_field_runs(sref, chap, rpr) + [text_run("-", rpr)] + make_field_runs(seq, seqv, rpr)
        pkg.mark(p)
        olds, news = list(old_fields), [" ".join(sref.split()), " ".join(seq.split())]
        for o in olds:
            field_changes.append((pkg.main, o, None))
        for n_ in news:
            field_changes.append((pkg.main, None, n_))
        for k, el in enumerate(new):
            p.insert(insert_at + k, el)
        fixed += 1
    fc = [(st, o, n) for st, o, n in field_changes]
    return {"fixed": fixed, "skipped": skipped, "field_changes": _pair(fc),
            "checks": [{"expr": "caption_numbers(doc)", "expect": "computed"}]}


def _pair(changes):
    """Turn (story, old, None) and (story, None, new) items into the (story, old, new) form the guard reads;
    None stands for no field on that side."""
    out = []
    for st, o, n in changes:
        out.append((st, o if o is not None else "\0none", n if n is not None else "\0none"))
    return out


OP_CAPTIONS_ALLOWED = ("fields", "direct_run_format")


# ------------------------------------------------------------------ cross-reference audit (read-only)
def op_xref_audit(pkg):
    marks = set()
    for story in pkg.stories():
        for b in pkg.tree(story).iter(q("bookmarkStart")):
            marks.add(b.get(q("name")))
    rows = []
    for story in pkg.stories():
        root = pkg.tree(story)
        for g in field_groups(root):
            ins = instr_of(g)
            parts = ins.split()
            if not parts:
                continue
            kind = parts[0].upper()
            tgt = None
            if kind in ("REF", "PAGEREF", "NOTEREF") and len(parts) > 1:
                tgt = parts[1]
            elif kind == "HYPERLINK" and "\\l" in parts:
                tgt = parts[parts.index("\\l") + 1].strip('"')
            elif kind not in ("SEQ", "STYLEREF", "PAGE", "TOC", "NUMPAGES", "SECTIONPAGES") and len(parts) == 1 and parts[0] in marks:
                kind, tgt = "REF(short)", parts[0]
            if tgt:
                rows.append({"story": story, "kind": kind, "target": tgt, "status": "OK" if tgt in marks else "No target"})
    missing = sorted({r["target"] for r in rows if r["status"] != "OK"})
    return {"refs": len(rows), "missing": missing, "checks": [{"expr": "error_ref_targets(doc)", "expect": missing}]}
