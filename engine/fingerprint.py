"""Content fingerprint of a Word package, used to prove an operation changed only what it said.

Every key is "<story part>|<facet>". An operation declares which facets it may change;
any other difference is a violation and the write is refused.
"""
import copy
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


# ---------------------------------------------------------------------------------------------
# Logical view: the same facets, keyed by what the reader sees instead of by part name.
# Word re-saves a document under its own part layout: it gives every section its own copy of an
# inherited header or footer, renames header1.xml to header7.xml, adds an endnotes part that holds
# only the separator lines, and renumbers drawing and content-control ids. None of that changes the
# document. fingerprint() is keyed by part name, which is right for proving an engine operation
# (the engine never renames parts), but it reports every one of those as a change when comparing a
# file with Word's re-save of it. logical() resolves headers and footers per section, drops note
# separators and ignores ids, so a Word re-save of an unchanged document compares equal.

HF_TYPES = ("default", "first", "even")
NOTE_SKIP = ("separator", "continuationSeparator", "continuationNotice")


def _strip_ids(root, *tags):
    out = []
    for e in root.iter(*tags):
        c = etree.fromstring(etree.tostring(e))
        for i in list(c.iter(f"{{{W}}}id")):
            i.getparent().remove(i)
        for el in c.iter():
            for a in [a for a in el.attrib if a.endswith("}id") or a == "id"]:
                del el.attrib[a]
        out.append(etree.tostring(c, method="c14n"))
    return _h(b"".join(out))


def _facets(pkg, part, root):
    vt = visible_text(root)
    rels = pkg.rels_of(part) if part else {}
    names, _ = bookmarks(root)
    return {
        "text": _h(vt), "text_len": len(vt), "deleted_text": _h(deleted_text(root)),
        "paragraphs": sum(1 for _ in root.iter(WP)), "fields": tuple(fields(root)),
        "bookmarks": tuple(names), "bookmark_context": bookmark_context(root), "tables": tuple(tables(root)),
        "drawings": sum(1 for _ in root.iter(f"{{{W}}}drawing")) + sum(1 for _ in root.iter(f"{{{W}}}pict")),
        "objects": sum(1 for _ in root.iter(f"{{{W}}}object")),
        "content_controls": sum(1 for _ in root.iter(f"{{{W}}}sdt")),
        "content_control_props": _strip_ids(root, f"{{{W}}}sdtPr", f"{{{W}}}ffData"),
        "revisions": tuple((t, sum(1 for _ in root.iter(f"{{{W}}}{t}"))) for t in REV_TAGS),
        "note_refs": sum(1 for _ in root.iter(f"{{{W}}}footnoteReference", f"{{{W}}}endnoteReference")),
        "comment_refs": sum(1 for _ in root.iter(f"{{{W}}}commentReference")),
        "alt_content": sum(1 for _ in root.iter(MC)),
        "hyperlinks": tuple(((rels.get(h.get(f"{{{R}}}id")) or (None, None, None))[1], h.get(f"{{{W}}}anchor"))
                            for h in root.iter(f"{{{W}}}hyperlink")),
    }


def _empty_hf(f):
    return f is None or (f["text_len"] == 0 and not f["drawings"] and not f["fields"] and not f["tables"]
                         and not f["objects"] and not f["content_controls"])


def logical(pkg):
    fp = {}
    main = pkg.tree(pkg.main)
    for k, v in _facets(pkg, pkg.main, main).items():
        fp["body|" + k] = v
    fp["body|section_count"] = sum(1 for _ in main.iter(f"{{{W}}}sectPr"))
    rels = pkg.rels_of(pkg.main)
    settings = pkg.word_part("/settings")
    even = settings is not None and pkg.tree(settings).find("w:evenAndOddHeaders", NS) is not None
    sects = [s for s in main.iter(f"{{{W}}}sectPr") if s.getparent().tag != f"{{{W}}}pPrChange"
             and s.getparent().tag != f"{{{W}}}sectPrChange"]
    cache, eff = {}, {}
    for n, s in enumerate(sects, 1):
        title = s.find("w:titlePg", NS) is not None and s.find("w:titlePg", NS).get(f"{{{W}}}val") not in ("0", "false")
        for kind in ("header", "footer"):
            for ref in s.findall(f"w:{kind}Reference", NS):
                t = rels.get(ref.get(f"{{{R}}}id"))
                eff[(kind, ref.get(f"{{{W}}}type") or "default")] = t[1] if t else None
            for typ in HF_TYPES:
                if (typ == "first" and not title) or (typ == "even" and not even):
                    continue
                part = eff.get((kind, typ))
                if part not in cache:
                    cache[part] = _facets(pkg, part, pkg.tree(part)) if part in pkg.raw else None
                f = cache[part]
                if _empty_hf(f):
                    continue
                for k, v in f.items():
                    fp[f"sec{n}:{kind}:{typ}|{k}"] = v
    for kind in ("footnotes", "endnotes"):
        part = pkg.word_part("/" + kind)
        holder = etree.Element("holder")
        if part:
            for note in pkg.tree(part):
                if note.get(f"{{{W}}}type") not in NOTE_SKIP:
                    holder.append(copy.deepcopy(note))
        if len(holder):
            for k, v in _facets(pkg, part, holder).items():
                fp[f"{kind}|{k}"] = v
    cp = pkg.word_part("/comments")
    if cp:
        for k, v in _facets(pkg, cp, pkg.tree(cp)).items():
            fp["comments|" + k] = v
    media, external = set(), set()
    for part in list(pkg.raw):
        if part.endswith(".rels") or part.startswith("word/glossary"):
            continue
        for rid, (typ, tgt, mode) in pkg.rels_of(part).items():
            if mode == "External":
                if not typ.endswith("/hyperlink"):
                    external.add(tgt)
            elif tgt in pkg.raw and typ.endswith(("/image", "/oleObject", "/package", "/chart", "/diagramData")):
                media.add(_h(pkg.raw[tgt]))
    fp["package|media"] = tuple(sorted(media))
    fp["package|external_links"] = tuple(sorted(external))
    return fp
