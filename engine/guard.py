"""Unit guard: proves an operation changed only the elements it marked.

A unit is the smallest thing we hash:
  story parts (document, headers, footers, notes, comments, glossary):
      every w:p, every table/row/cell property set (tblPr, tblGrid, trPr, tcPr, tblPrEx), every sdtPr,
      every docPartPr, the attributes of every container (tbl, tr, tc, sdt, footnote, comment, docPart, body),
      and any other direct child of the root or the body (for example the final sectPr, w:background)
  styles part:     every w:style, docDefaults, latentStyles, the root attributes
  numbering part:  every abstractNum, num, numPicBullet, the root attributes
  any other part:  the whole part

Package.tree() records a snapshot of a part's units the first time the part is read, so the snapshot is
always taken before any change. After the operation, every unit that changed, appeared or disappeared must
be the marked element itself or sit under a marked element. Anything else is a violation.
"""
import hashlib
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STORY_ROOTS = {"document", "hdr", "ftr", "footnotes", "endnotes", "comments", "glossaryDocument"}
UNIT_TAGS = {f"{{{W}}}{t}" for t in ("p", "tblPr", "tblGrid", "trPr", "tcPr", "tblPrEx", "sdtPr", "docPartPr")}
CONTAINERS = {f"{{{W}}}{t}" for t in ("tbl", "tr", "tc", "sdt", "sdtContent", "footnote", "endnote", "comment",
                                       "docPart", "docPartBody", "docParts", "body")}
SKIP_CHILD = {f"{{{W}}}{t}" for t in ("p", "tbl", "sdt", "body", "docParts")}


def _hash(b):
    return hashlib.sha1(b).hexdigest()


def _c14n(el):
    return _hash(etree.tostring(el, method="c14n"))


def _attrs(el):
    return _hash((el.tag + "|" + "|".join(f"{k}={v}" for k, v in sorted(el.attrib.items()))).encode())


def units(root):
    """[(element, kind, hash)] for one part."""
    out = []
    loc = etree.QName(root).localname if isinstance(root.tag, str) else ""
    if root.tag.startswith(f"{{{W}}}") and loc in STORY_ROOTS:
        out.append((root, "attrs", _attrs(root)))
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            if el.tag in UNIT_TAGS:
                out.append((el, "c14n", _c14n(el)))
            elif el.tag in CONTAINERS:
                out.append((el, "attrs", _attrs(el)))
        for parent in [root] + root.findall(f"{{{W}}}body"):
            for ch in parent:
                if isinstance(ch.tag, str) and ch.tag not in SKIP_CHILD and ch.tag not in UNIT_TAGS and ch.tag not in CONTAINERS:
                    out.append((ch, "c14n", _c14n(ch)))
        return out
    if root.tag == f"{{{W}}}styles" or root.tag == f"{{{W}}}numbering":
        out.append((root, "attrs", _attrs(root)))
        for ch in root:
            if isinstance(ch.tag, str):
                out.append((ch, "c14n", _c14n(ch)))
        return out
    return [(root, "c14n", _c14n(root))]


def _hash_of(el, kind):
    return _c14n(el) if kind == "c14n" else _attrs(el)


class Snapshot:
    def __init__(self):
        self.parts = {}      # part -> {id(el): (el, kind, hash)}
        self.marks = set()   # ids of marked elements
        self.marked_parts = set()
        self._keep = []      # keep marked elements alive so ids stay unique

    def record(self, part, root):
        if part not in self.parts:
            self.parts[part] = {id(e): (e, k, h) for e, k, h in units(root)}

    def mark(self, el):
        self.marks.add(id(el))
        self._keep.append(el)

    def covered(self, el):
        if id(el) in self.marks:
            return True
        return any(id(a) in self.marks for a in el.iterancestors())

    def violations(self, pkg, limit=12):
        out = []
        for part in sorted(pkg.dirty):
            if part in self.marked_parts:
                continue
            if part not in self.parts:
                out.append(f"{part}: new part not declared")
                continue
            before = self.parts[part]
            root = pkg.trees.get(part)
            if root is None:
                continue
            seen = set()
            for el, kind, _ in units(root):
                seen.add(id(el))
                old = before.get(id(el))
                if old is None:
                    if not self.covered(el):
                        out.append(f"{part}: new {etree.QName(el).localname} not under a marked element")
                elif old[2] != _hash_of(el, kind) and not self.covered(el):
                    out.append(f"{part}: {etree.QName(el).localname} changed outside the declared elements")
                if len(out) >= limit:
                    return out
            for uid, (el, kind, _) in before.items():
                if uid not in seen and not self.covered(el):
                    out.append(f"{part}: {etree.QName(el).localname} removed outside the declared elements")
                    if len(out) >= limit:
                        return out
        for part, data in pkg.raw.items():
            if part in pkg.raw_at_load and part not in pkg.dirty and data is not pkg.raw_at_load[part] \
                    and part not in self.marked_parts:
                out.append(f"{part}: bytes replaced without declaration")
        for part in sorted(set(pkg.raw) - set(pkg.raw_at_load)):
            if part not in self.marked_parts and part not in self.parts and part not in pkg.dirty:
                out.append(f"{part}: new part not declared")
        for part in sorted(set(pkg.raw_at_load) - set(pkg.raw)):
            if part not in self.marked_parts:
                out.append(f"{part}: part removed without declaration")
        return out
