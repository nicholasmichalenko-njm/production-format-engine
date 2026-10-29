"""Content fingerprint of a Word package, used to prove an operation changed only what it said.

Every key is "<story part>|<facet>". An operation declares which facets it may change;
any other difference is a violation and the write is refused.
"""
import fnmatch
import hashlib
from lxml import etree
from .pkg import W, R, NS

WT = f"{{{W}}}t"; WDT = f"{{{W}}}delText"; WTAB = f"{{{W}}}tab"; WBR = f"{{{W}}}br"; WCR = f"{{{W}}}cr"
WP = f"{{{W}}}p"; WFC = f"{{{W}}}fldChar"; WIT = f"{{{W}}}instrText"; WFS = f"{{{W}}}fldSimple"
WSYM = f"{{{W}}}sym"; WNBH = f"{{{W}}}noBreakHyphen"; WPTAB = f"{{{W}}}ptab"
REV_TAGS = ["ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "tblPrChange", "trPrChange",
            "tcPrChange", "sectPrChange", "numberingChange", "cellIns", "cellDel"]


def _h(s):
    return hashlib.sha256(s.encode("utf-8") if isinstance(s, str) else s).hexdigest()[:16]


def fields(root):
    """List of normalized field instructions in document order (complex and simple)."""
    out, stack = [], []
    for el in root.iter(WFC, WIT, WFS):
        if el.tag == WFS:
            out.append(" ".join(el.get(f"{{{W}}}instr", "").split()))
        elif el.tag == WFC:
            t = el.get(f"{{{W}}}fldCharType")
            if t == "begin":
                stack.append([])
            elif t == "end" and stack:
                out.append(" ".join("".join(stack.pop()).split()))
        elif el.tag == WIT and stack:
            stack[-1].append(el.text or "")
    return out


def visible_text(root):
    """Text the reader sees (field results included, field codes excluded, deletions excluded)."""
    parts = []
    depth_code = []  # per open field: True while in the code part
    for el in root.iter():
        tag = el.tag
        if tag == WFC:
            t = el.get(f"{{{W}}}fldCharType")
            if t == "begin":
                depth_code.append(True)
            elif t == "separate" and depth_code:
                depth_code[-1] = False
            elif t == "end" and depth_code:
                depth_code.pop()
            continue
        in_code = any(depth_code)
        if tag == WT and not in_code:
            parts.append(el.text or "")
        elif tag in (WTAB, WPTAB) and not in_code and el.getparent() is not None and el.getparent().tag == f"{{{W}}}r":
            parts.append("\t")
        elif tag in (WBR, WCR) and not in_code:
            parts.append("\n")
        elif tag == WNBH and not in_code:
            parts.append("‑")
        elif tag == WSYM and not in_code:
            parts.append("[sym:" + (el.get(f"{{{W}}}char") or "") + "]")
        elif tag == WP:
            parts.append("¶")
    return "".join(parts)


def deleted_text(root):
    return "".join(e.text or "" for e in root.iter(WDT))


def tables(root):
    sig = []
    for tbl in root.iter(f"{{{W}}}tbl"):
        rows = []
        for tr in tbl.findall("w:tr", NS):
            cells = []
            for tc in tr.findall("w:tc", NS):
                gs = tc.find("w:tcPr/w:gridSpan", NS)
                vm = tc.find("w:tcPr/w:vMerge", NS)
                cells.append(f"{gs.get(f'{{{W}}}val') if gs is not None else 1}"
                             f"{'v' + (vm.get(f'{{{W}}}val') or 'c') if vm is not None else ''}")
            rows.append(",".join(cells))
        depth = sum(1 for a in tbl.iterancestors(f"{{{W}}}tbl"))
        sig.append(f"d{depth}:" + "/".join(rows))
    return sig


def bookmarks(root):
    starts = {b.get(f"{{{W}}}id"): b.get(f"{{{W}}}name") for b in root.iter(f"{{{W}}}bookmarkStart")}
    ends = {b.get(f"{{{W}}}id") for b in root.iter(f"{{{W}}}bookmarkEnd")}
    names = sorted(n for n in starts.values() if n)
    unpaired = sorted(i for i in starts if i not in ends)
    return names, unpaired


def _ppr_no_sect(ppr):
    """Paragraph properties without the section properties (those are the 'sections' facet)."""
    sp = ppr.find(f"{{{W}}}sectPr")
    if sp is None:
        return etree.tostring(ppr)
    c = etree.fromstring(etree.tostring(ppr))
    c.remove(c.find(f"{{{W}}}sectPr"))
    return etree.tostring(c)


def bookmark_context(root):
    """name -> hash of the text of the paragraph where the bookmark starts (catches moved targets)."""
    out = {}
    for b in root.iter(f"{{{W}}}bookmarkStart"):
        p = next(b.iterancestors(WP), None)
        out[b.get(f"{{{W}}}name")] = _h(visible_text(p) if p is not None else "body")
    return tuple(sorted(out.items()))


def _c14n_all(root, *tags):
    return _h(b"".join(etree.tostring(e, method="c14n") for e in root.iter(*tags)))


MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"


def fingerprint(pkg):
    """Content facets. Formatting is not here on purpose: the unit guard (guard.py) checks every
    paragraph, property set, style and list element individually, which is stricter."""
    fp = {}
    for story in pkg.stories():
        root = pkg.tree(story)
        k = story + "|"
        vt = visible_text(root)
        fp[k + "text"] = _h(vt)
        fp[k + "text_len"] = len(vt)
        fp[k + "deleted_text"] = _h(deleted_text(root))
        fp[k + "paragraphs"] = sum(1 for _ in root.iter(WP))
        fp[k + "fields"] = tuple(fields(root))
        names, unpaired = bookmarks(root)
        fp[k + "bookmarks"] = tuple(names)
        fp[k + "bookmark_context"] = bookmark_context(root)
        fp[k + "bookmarks_unpaired"] = tuple(unpaired)
        fp[k + "tables"] = tuple(tables(root))
        fp[k + "drawings"] = sum(1 for _ in root.iter(f"{{{W}}}drawing")) + sum(1 for _ in root.iter(f"{{{W}}}pict"))
        fp[k + "drawing_xml"] = _c14n_all(root, f"{{{W}}}drawing", f"{{{W}}}pict", f"{{{W}}}object")
        fp[k + "objects"] = sum(1 for _ in root.iter(f"{{{W}}}object"))
        fp[k + "content_controls"] = sum(1 for _ in root.iter(f"{{{W}}}sdt"))
        fp[k + "content_control_props"] = _c14n_all(root, f"{{{W}}}sdtPr", f"{{{W}}}ffData")
        fp[k + "revisions"] = tuple((t, sum(1 for _ in root.iter(f"{{{W}}}{t}"))) for t in REV_TAGS)
        fp[k + "note_refs"] = sum(1 for _ in root.iter(f"{{{W}}}footnoteReference", f"{{{W}}}endnoteReference"))
        fp[k + "comment_refs"] = sum(1 for _ in root.iter(f"{{{W}}}commentReference"))
        fp[k + "alt_content"] = sum(1 for _ in root.iter(MC))
        fp[k + "hyperlinks"] = tuple((h.get(f"{{{R}}}id"), h.get(f"{{{W}}}anchor")) for h in root.iter(f"{{{W}}}hyperlink"))
        if story == pkg.main:
            fp[k + "section_count"] = sum(1 for _ in root.iter(f"{{{W}}}sectPr"))
    media, broken, external = [], [], []
    for part in list(pkg.raw):
        if part.endswith(".rels"):
            continue
        for rid, (typ, tgt, mode) in pkg.rels_of(part).items():
            if mode == "External":
                external.append(f"{part}:{rid}->{tgt}")
                continue
            if tgt not in pkg.raw:
                broken.append(f"{part}:{rid}->{tgt}")
            elif typ.endswith(("/image", "/oleObject", "/package", "/chart", "/diagramData")):
                media.append(f"{part}:{rid}:{tgt}:{_h(pkg.raw[tgt])}")
    fp["package|media"] = tuple(sorted(media))
    fp["package|external_links"] = tuple(sorted(external))
    fp["package|broken_rels"] = tuple(sorted(broken))
    fp["package|parts"] = tuple(sorted(pkg.raw))
    return fp


def diff(before, after, allowed=()):
    """Return the facets that changed and are not allowed. allowed is a list of glob patterns
    matched against the facet name (after the |) or the full key."""
    out = []
    for k in sorted(set(before) | set(after)):
        if before.get(k) == after.get(k):
            continue
        facet = k.split("|", 1)[1]
        if any(fnmatch.fnmatch(facet, a) or fnmatch.fnmatch(k, a) for a in allowed):
            continue
        out.append(k)
    return out
