"""Rogue-operation suite: every case must be REFUSED with the original byte identical, or (for the
behavior cases) must give the exact expected result. No Word needed; runs on any machine.
Usage: python rogue.py <small.docx> <workdir> <results.json> [qp_master.dotx]
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from lxml import etree
from engine import safewrite
from engine.pkg import Package, Refused, W, NS
from engine.ops import word_ops as ops
from engine.ops import glossary as G
from engine.fingerprint import fields

WQ = lambda t: f"{{{W}}}{t}"


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def rezip(src, dst, edit, extra=()):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for i in zin.infolist():
            d = edit(i.filename, zin.read(i))
            if d is not None:
                z.writestr(i.filename, d)
        for n, d in extra:
            z.writestr(n, d)


def main():
    src, work, out = sys.argv[1:4]
    qp = sys.argv[4] if len(sys.argv) > 4 else None
    os.makedirs(work, exist_ok=True)
    results = []

    def case(name, expect, path, op, allowed=(), check=None, **kw):
        h0 = sha(path)
        res = {"case": name, "expect": expect}
        try:
            r = safewrite.apply(path, op, allowed=allowed, **kw)
            res["engine"] = "applied"; res["changed"] = r.changed
        except Refused as e:
            res["engine"] = "refused"; res["message"] = str(e)[:300]
        except Exception as e:
            res["engine"] = "error"; res["message"] = f"{type(e).__name__}: {e}"[:300]
        res["original_intact"] = sha(path) == h0
        if expect == "REFUSE":
            res["pass"] = res["engine"] == "refused" and res["original_intact"]
        else:
            try:
                res["check"] = check(path) if check else None
                res["pass"] = res["engine"] == "applied" and (check is None or res["check"] is True)
            except Exception:
                res["pass"] = False; res["check"] = traceback.format_exc()[-300:]
        results.append(res)
        print(json.dumps(res)[:400], flush=True)

    def fresh(name, base=src):
        p = os.path.join(work, name + os.path.splitext(base)[1])
        if os.path.exists(p):
            os.remove(p)
        shutil.copy2(base, p)
        return p

    # R01 styles op that also hides Normal text and deletes Heading 2
    def r01(pkg):
        rep = ops.op_styles(pkg, {"Caption": {"size": 9}})
        root = pkg.tree(pkg.word_part("/styles"))
        ops.ensure(ops.ensure(ops.find_style(root, "Normal"), "rPr"), "vanish")
        h2 = ops.find_style(root, "Heading 2"); h2.getparent().remove(h2)
        return rep
    case("R01_styles_collateral", "REFUSE", fresh("R01"), r01)

    # R02 heading numbering that also turns other lists into bullets and changes settings
    def r02(pkg):
        rep = ops.op_heading_numbering(pkg)
        nb = pkg.tree(pkg.word_part("/numbering"))
        for f_ in nb.iter(WQ("numFmt")):
            f_.set(WQ("val"), "bullet")
        st = pkg.word_part("/settings"); pkg.touch(st)
        etree.SubElement(pkg.tree(st), WQ("updateFields")).set(WQ("val"), "true")
        return rep
    case("R02_numbering_and_settings_collateral", "REFUSE", fresh("R02"), r02)

    # R03 fix tables that also strips body formatting and changes heading styles outside tables
    def r03(pkg):
        rep = ops.op_fix_tables(pkg)
        body = pkg.tree(pkg.main).find("w:body", NS)
        for p in body.findall("w:p", NS):
            ps = p.find("w:pPr/w:pStyle", NS)
            if ps is not None and ps.get(WQ("val")) == "Heading2":
                ps.set(WQ("val"), "Normal")
            for r in p.iter(WQ("r")):
                rpr = r.find("w:rPr", NS)
                if rpr is not None:
                    ops.remove(rpr, "b", "vertAlign", "rStyle")
        return rep
    case("R03_tables_body_collateral", "REFUSE", fresh("R03"), r03, ops.OP_TABLES_ALLOWED)

    # R04 change picture alt text and size while fixing tables
    def r04(pkg):
        rep = ops.op_fix_tables(pkg)
        for d in pkg.tree(pkg.main).iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"):
            d.set("descr", "changed")
        return rep
    case("R04_alt_text_change", "REFUSE", fresh("R04"), r04, ops.OP_TABLES_ALLOWED)

    # R05 move a cross-reference target bookmark into another paragraph
    def r05(pkg):
        rep = ops.op_captions(pkg)
        root = pkg.tree(pkg.main)
        b = next(x for x in root.iter(WQ("bookmarkStart")) if x.get(WQ("name"), "").startswith("_Ref1"))
        e = next(x for x in root.iter(WQ("bookmarkEnd")) if x.get(WQ("id")) == b.get(WQ("id")))
        b.getparent().remove(b); e.getparent().remove(e)
        target = root.find("w:body/w:p[5]", NS)
        target.insert(1, b); target.append(e)
        return rep
    case("R05_bookmark_moved", "REFUSE", fresh("R05"), r05, ops.OP_CAPTIONS_ALLOWED)

    # R06 a footer PAGEREF in a simple field must not become PAGE
    p6 = os.path.join(work, "R06.docx")
    def ed6(n, d):
        if n == "word/footer3.xml":
            return d.replace(b"</w:ftr>", b'<w:p><w:fldSimple w:instr=" PAGEREF _Ref100000101 \\h "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:ftr>')
        return d
    rezip(src, p6, ed6)
    def chk6(path):
        return any("PAGEREF _Ref100000101" in f for f in fields(Package(path).tree("word/footer3.xml")))
    case("R06_footer_pageref_kept", "APPLY", p6, ops.op_hf_clean, ("fields",), chk6)

    # R07 owner file in a folder whose name has brackets
    d7 = os.path.join(work, "Prop [Draft]")
    os.makedirs(d7, exist_ok=True)
    p7 = os.path.join(d7, "Volume_I.docx"); shutil.copy2(src, p7)
    open(os.path.join(d7, "~$lume_I.docx"), "wb").write(b"\x05user")
    case("R07_owner_file_bracket_folder", "REFUSE", p7, ops.op_hf_clean, ("fields",))

    # R08 hard kills at each write step (os._exit, no cleanup)
    for point in ("after_temp", "before_replace", "after_replace"):
        p8 = fresh("R08_" + point)
        h0 = sha(p8)
        code = (f"import sys; sys.path.insert(0, {os.path.dirname(HERE)!r}); from engine import safewrite; "
                f"from engine.ops import word_ops as ops; safewrite.apply({p8!r}, ops.op_hf_clean, allowed=('fields',))")
        env = dict(os.environ, PE_HARD_CRASH=point)
        rc = subprocess.run([sys.executable, "-c", code], env=env).returncode
        temps = [f for f in os.listdir(work) if f.startswith(safewrite.TMP_PREFIX) and f.endswith(os.path.basename(p8))]
        try:
            Package(p8); loads = True
        except Refused:
            loads = False
        intact = sha(p8) == h0
        res = {"case": f"R08_hard_kill_{point}", "exit_code": rc, "original_intact": intact, "file_loads": loads,
               "temp_left": len(temps)}
        # the next run must sweep an old temp and work normally
        for t in temps:
            old = time.time() - 3600
            os.utime(os.path.join(work, t), (old, old))
        try:
            r = safewrite.apply(p8, ops.op_hf_clean, allowed=("fields",))
            res["next_run"] = "applied" if r.changed else "nothing to change"; res["swept"] = r.swept
        except Refused as e:
            res["next_run"] = "refused: " + str(e)[:120]
        left = [f for f in os.listdir(work) if f.startswith(safewrite.TMP_PREFIX) and f.endswith(os.path.basename(p8))]
        if point == "after_replace":
            res["pass"] = rc == 77 and loads and not intact and not left  # replaced, valid, complete
        else:
            res["pass"] = rc == 77 and intact and loads and not left
        results.append(res); print(json.dumps(res), flush=True)

    # R09 point a picture at a different relationship
    def r09(pkg):
        rep = ops.op_fix_tables(pkg)
        b = next(pkg.tree(pkg.main).iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))
        b.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", "rId1")
        return rep
    case("R09_picture_relink", "REFUSE", fresh("R09"), r09, ops.OP_TABLES_ALLOWED)

    # R10 change a part nobody declared
    def r10(pkg):
        rep = ops.op_hf_clean(pkg)
        th = pkg.word_part("/theme"); pkg.touch(th)
        next(pkg.tree(th).iter("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")).set("typeface", "Comic Sans MS")
        return rep
    case("R10_undeclared_part_change", "REFUSE", fresh("R10"), r10, ("fields",))

    # R11 add a part nobody declared
    def r11(pkg):
        rep = ops.op_hf_clean(pkg)
        pkg.raw["customXml/item9.xml"] = b"<x/>"
        return rep
    case("R11_undeclared_new_part", "REFUSE", fresh("R11"), r11, ("fields",))

    # R12 caption fix that also rewrites a body cross-reference instruction
    def r12(pkg):
        rep = ops.op_captions(pkg)
        for it in pkg.tree(pkg.main).iter(WQ("instrText")):
            if (it.text or "").strip().startswith("REF _Ref"):
                p = next(it.iterancestors(WQ("p"))); pkg.mark(p)  # even if it marks the paragraph
                it.text = " REF _Ref100000102 \\h "
                break
        return rep
    case("R12_undeclared_field_change", "REFUSE", fresh("R12"), r12, ops.OP_CAPTIONS_ALLOWED)

    # R14 a style used only by name in a STYLEREF field must not be removed as unused
    p14 = os.path.join(work, "R14.docx")
    def ed14(n, d):
        if n == "word/styles.xml":
            return d.replace(b"</w:styles>", b'<w:style w:type="paragraph" w:customStyle="1" w:styleId="PEProjectName"><w:name w:val="PE Project Name"/><w:basedOn w:val="Normal"/></w:style></w:styles>')
        if n == "word/header3.xml":
            return d.replace(b"</w:hdr>", b'<w:p><w:fldSimple w:instr=" STYLEREF &quot;PE Project Name&quot; "><w:r><w:t>X</w:t></w:r></w:fldSimple></w:p></w:hdr>')
        return d
    rezip(src, p14, ed14)
    def chk14(path):
        return ops.find_style(Package(path).tree("word/styles.xml"), "PE Project Name") is not None
    case("R14_styleref_name_keeps_style", "APPLY", p14, ops.op_remove_unused_styles, (), chk14,
         op_args={"names": ["PE Project Name"]})

    # R15 a macro-enabled document keeps its macro part byte for byte
    p15 = os.path.join(work, "R15.docm")
    def ed15(n, d):
        if n == "[Content_Types].xml":
            return d.replace(b"wordprocessingml.document.main+xml", b"ms-word.document.macroEnabled.main+xml").replace(
                b"</Types>", b'<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
        return d
    vba = os.urandom(4096)
    rezip(src, p15, ed15, extra=[("word/vbaProject.bin", vba)])
    case("R15_macros_preserved", "APPLY", p15, ops.op_hf_clean, ("fields",),
         lambda path: zipfile.ZipFile(path).read("word/vbaProject.bin") == vba)

    # R16 to R20 refusal rules
    def mk(name, edit, extra=(), ext=".docx"):
        p = os.path.join(work, name + ext); rezip(src, p, edit, extra); return p
    case("R16_altchunk", "REFUSE", mk("R16", lambda n, d: d.replace(b"<w:body>", b'<w:body><w:altChunk r:id="rIdX"/>', 1) if n == "word/document.xml" else d), ops.op_hf_clean, ("fields",))
    case("R17_data_binding", "REFUSE", mk("R17", lambda n, d: d.replace(b"<w:body>", b'<w:body><w:sdt><w:sdtPr><w:dataBinding w:xpath="/a" w:storeItemID="{1}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>', 1) if n == "word/document.xml" else d), ops.op_hf_clean, ("fields",))
    case("R18_mark_as_final", "REFUSE", mk("R18", lambda n, d: d, extra=[("docProps/custom.xml", b'<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property name="_MarkAsFinal"/></Properties>')]), ops.op_hf_clean, ("fields",))
    case("R19_not_word", "REFUSE", mk("R19", lambda n, d: d.replace(b"wordprocessingml.document.main+xml", b"spreadsheetml.sheet.main+xml") if n == "[Content_Types].xml" else d), ops.op_hf_clean, ("fields",))
    case("R20_uppercase_xml_malformed", "REFUSE", mk("R20", lambda n, d: d, extra=[("customXml/ITEM5.XML", b"<a><b></a>")]), ops.op_hf_clean, ("fields",))

    # R13 glossary: rename one entry while changing another entry's body
    if qp and os.path.exists(qp):
        p13 = fresh("R13", qp)
        def r13(pkg):
            rep = G.op_rename(pkg, {"20 Main Footer": "21 Main Footer"})
            other = next(e for e in G.entries(pkg) if e["name"] == "80 Boilerplate Intro")
            next(other["el"].iter(WQ("t"))).text = "Changed boilerplate."
            return rep
        case("R13_glossary_collateral", "REFUSE", p13, r13, G.GLOSS_ALLOWED)

    summary = {"cases": len(results), "passed": sum(1 for r in results if r.get("pass")),
               "failed": [r["case"] for r in results if not r.get("pass")]}
    json.dump({"summary": summary, "results": results}, open(out, "w"), indent=1)
    print(json.dumps(summary))


if __name__ == "__main__":
    main()
