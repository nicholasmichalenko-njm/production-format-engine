"""Safe Open XML package layer.

Rules:
- Refuse anything we cannot fully understand: not a zip, OLE (encrypted or legacy binary),
  bad CRC, duplicate or unsafe entry names, oversized or bomb-like entries, missing
  content types or main part, any XML part that is not well formed, any DOCTYPE, Strict OOXML.
- Parts we do not edit are written back byte for byte, in the original order, with the
  original compression and timestamps.
- Only parts marked dirty are serialized again.
"""
import io
import posixpath
import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
NS = {"w": W, "r": R, "pr": PR, "ct": CT}

MAX_ENTRY = 1_500_000_000      # 1.5 GB uncompressed per entry
MAX_TOTAL = 1_500_000_000      # 1.5 GB uncompressed per package
MAX_RATIO = 200                # compressed bomb guard for large entries
OLE_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


class Refused(Exception):
    """The engine will not touch this file. The message says why, in plain words."""


def _parser():
    return etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False,
                           huge_tree=True, remove_blank_text=False, strip_cdata=False)


def parse_xml(data, name="part"):
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise Refused(f"{name} has a DOCTYPE or ENTITY declaration (not allowed in Office files).")
    try:
        return etree.fromstring(data, _parser())
    except etree.XMLSyntaxError as e:
        raise Refused(f"{name} is not well-formed XML: {e}")


def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class Package:
    def __init__(self, path):
        self.path = str(path)
        with open(self.path, "rb") as f:
            raw = f.read()
        self._load(raw)

    # ---------- loading and refusal rules ----------
    def _load(self, raw):
        if raw[:8] == OLE_MAGIC:
            raise Refused("This is an encrypted (password) file or a legacy binary .doc/.xls/.ppt. "
                          "Remove the password or save it as .docx/.xlsx/.pptx first.")
        if raw[:4] != b"PK\x03\x04":
            raise Refused("This is not an Office Open XML file (not a zip package).")
        try:
            zf = zipfile.ZipFile(io.BytesIO(raw))
        except zipfile.BadZipFile as e:
            raise Refused(f"The file is damaged (zip structure): {e}")
        names = [i.filename for i in zf.infolist()]
        if len(names) != len(set(n.lower() for n in names)):
            raise Refused("The package has duplicate part names.")
        total = 0
        for i in zf.infolist():
            n = i.filename
            if n.startswith("/") or "\\" in n or ".." in n.split("/") or ":" in n:
                raise Refused(f"Unsafe part name in package: {n!r}")
            if i.flag_bits & 0x1:
                raise Refused("The package uses zip encryption.")
            if i.file_size > MAX_ENTRY:
                raise Refused(f"Part {n} is too large ({i.file_size} bytes).")
            if i.file_size > 50_000_000 and i.compress_size and i.file_size / i.compress_size > MAX_RATIO:
                raise Refused(f"Part {n} looks like a zip bomb.")
            total += i.file_size
        if total > MAX_TOTAL:
            raise Refused("The package is too large.")
        try:
            bad = zf.testzip()
            if bad:
                raise Refused(f"The file is damaged: part {bad} fails its checksum.")
            self.infos = zf.infolist()
            self.raw = {i.filename: zf.read(i) for i in self.infos}
        except Refused:
            raise
        except Exception as e:  # zlib.error, EOFError, BadZipFile, NotImplementedError (unknown compression)
            raise Refused(f"The file is damaged and cannot be read safely ({type(e).__name__}: {e}).")
        if "[Content_Types].xml" not in self.raw:
            raise Refused("The package has no [Content_Types].xml.")
        self.trees = {}
        self.dirty = set()
        # Every XML part must be well formed, even if we will not edit it.
        for n, data in self.raw.items():
            ln = n.lower()
            if ln.endswith(".xml") or ln.endswith(".rels"):
                parse_xml(data, n)  # validation only; parsed again on demand
            # legacy .vml parts are often not well-formed XML; they are never edited and are kept byte for byte
        self.raw_at_load = dict(self.raw)
        from .guard import Snapshot
        self.snap = Snapshot()
        ct = self.tree("[Content_Types].xml")
        self.ct_over = {o.get("PartName"): o.get("ContentType") for o in ct.findall(f"{{{CT}}}Override")}
        self.ct_def = {d.get("Extension").lower(): d.get("ContentType") for d in ct.findall(f"{{{CT}}}Default")}
        rels = self.rels_of("")
        main = [t for (typ, t, mode) in rels.values() if typ.endswith("/officeDocument") and mode != "External"]
        if not main:
            raise Refused("The package has no main document part.")
        self.main = main[0].lstrip("/")
        if self.main not in self.raw:
            raise Refused(f"The main part {self.main} is missing.")
        root = self.tree(self.main)
        if root.tag.startswith(f"{{{STRICT_W}}}") or STRICT_W.encode() in self.raw[self.main][:3000]:
            raise Refused("Strict Open XML format is not supported. Save as a normal Word document.")
        for n in self.raw:
            if n.startswith("_xmlsignatures/"):
                self.signed = True
                break
        else:
            self.signed = False
        self.kind = self.content_type("/" + self.main) or ""
        self.is_word = "wordprocessingml" in self.kind or "ms-word" in self.kind
        self.has_macros = any(n.lower().endswith("vbaproject.bin") for n in self.raw)
        # features the engine cannot see into or that Word rewrites on open: refuse rather than guess
        if self.is_word:
            main_xml = self.raw[self.main]
            if b"altChunk" in main_xml:
                raise Refused("The document embeds another file as an altChunk. Word must convert it first (open and save in Word).")
            if b"dataBinding" in main_xml:
                raise Refused("The document has content controls bound to data. Word rewrites them on open; the engine leaves them alone.")
        if "docProps/custom.xml" in self.raw and b"_MarkAsFinal" in self.raw["docProps/custom.xml"]:
            raise Refused("The document is marked as final. Remove Mark as Final first.")

    # ---------- part access ----------
    def content_type(self, partname):
        if partname in self.ct_over:
            return self.ct_over[partname]
        ext = partname.rsplit(".", 1)[-1].lower()
        return self.ct_def.get(ext)

    def tree(self, name):
        if name not in self.trees:
            self.trees[name] = parse_xml(self.raw[name], name)
            if hasattr(self, "snap"):
                self.snap.record(name, self.trees[name])  # unit snapshot before anyone can change the tree
        return self.trees[name]

    # ---------- declaring what an operation changes ----------
    def mark(self, el):
        """Declare that el (and everything under it) may change, appear or disappear."""
        self.snap.mark(el)
        return el

    def mark_part(self, name):
        """Declare that a whole part may change, appear or disappear."""
        self.snap.marked_parts.add(name)

    def touch(self, name):
        if name not in self.raw:
            raise KeyError(name)
        self.tree(name)
        self.dirty.add(name)

    def has(self, name):
        return name in self.raw

    @staticmethod
    def rels_name(part):
        d, f = posixpath.split(part)
        return posixpath.join(d, "_rels", f + ".rels") if part else "_rels/.rels"

    def rels_of(self, part):
        """{rId: (type, resolved target part name or URL, mode)}"""
        rn = self.rels_name(part)
        if rn not in self.raw:
            return {}
        base = posixpath.dirname(part)
        out = {}
        for r in self.tree(rn).findall(f"{{{PR}}}Relationship"):
            mode = r.get("TargetMode")
            tgt = r.get("Target")
            if mode != "External":
                if tgt.startswith("/"):
                    tgt = tgt.lstrip("/")
                else:
                    tgt = posixpath.normpath(posixpath.join(base, tgt))
            out[r.get("Id")] = (r.get("Type"), tgt, mode)
        return out

    def related(self, part, suffix):
        return [t for (typ, t, mode) in self.rels_of(part).values() if typ.endswith(suffix) and mode != "External"]

    def word_part(self, suffix):
        p = self.related(self.main, suffix)
        return p[0] if p else None

    def stories(self):
        """Every part that holds document text: main, headers, footers, notes, comments, glossary."""
        out = [self.main]
        for suf in ("/header", "/footer", "/footnotes", "/endnotes", "/comments"):
            out += self.related(self.main, suf)
        g = self.word_part("/glossaryDocument")
        if g:
            out.append(g)
            for suf in ("/header", "/footer", "/footnotes", "/endnotes"):
                out += self.related(g, suf)
        return [p for p in out if p in self.raw]

    # ---------- adding parts ----------
    def add_part(self, name, data_or_tree, content_type):
        if name in self.raw:
            raise ValueError(f"part exists: {name}")
        self.raw[name] = b""
        info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        self.infos.append(info)
        if isinstance(data_or_tree, bytes):
            self.raw[name] = data_or_tree
        else:
            self.trees[name] = data_or_tree
            self.dirty.add(name)
        ct = self.tree("[Content_Types].xml")
        self.mark_part(name)
        self.mark_part("[Content_Types].xml")  # one unit: a new Override changes the whole part
        o = etree.SubElement(ct, f"{{{CT}}}Override")
        o.set("PartName", "/" + name)
        o.set("ContentType", content_type)
        self.ct_over["/" + name] = content_type
        self.dirty.add("[Content_Types].xml")

    def remove_part(self, name):
        self.mark_part(name)
        self.mark_part("[Content_Types].xml")
        self.raw.pop(name, None)
        self.trees.pop(name, None)
        self.dirty.discard(name)
        self.infos = [i for i in self.infos if i.filename != name]
        ct = self.tree("[Content_Types].xml")
        for o in ct.findall(f"{{{CT}}}Override"):
            if o.get("PartName") == "/" + name:
                ct.remove(o)
                self.dirty.add("[Content_Types].xml")
        self.ct_over.pop("/" + name, None)

    # ---------- saving ----------
    def to_bytes(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as out:
            for i in self.infos:
                n = i.filename
                data = serialize(self.trees[n]) if n in self.dirty else self.raw[n]
                zi = zipfile.ZipInfo(n, date_time=i.date_time)
                zi.compress_type = i.compress_type
                zi.external_attr = i.external_attr
                zi.create_system = i.create_system
                out.writestr(zi, data)
        return buf.getvalue()
