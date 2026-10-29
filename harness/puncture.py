"""Puncture tests. Every case states the required outcome:
  REFUSE  the engine must refuse with a plain message, and the original must stay byte identical
  SAFE    the engine may succeed; if it does, Word must open the result and the guarded facets must hold;
          if it refuses, the original must stay byte identical
Usage: python puncture.py <small.docx> <workdir> <results.json>
"""
import hashlib
import io
import json
import os
import random
import shutil
import stat
import sys
import time
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from lxml import etree
from engine import safewrite
from engine.pkg import Package, Refused, W
from engine.fingerprint import fingerprint, diff
from engine.ops import word_ops as ops
from wordsession import WordSession, retry

STRUCT = ["text", "text_len", "deleted_text", "fields", "bookmarks", "tables", "drawings", "objects",
          "content_controls", "revisions", "note_refs", "section_count", "alt_content"]


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def rezip(src, dst, edit):
    """Copy a package, letting edit(name, bytes) return new bytes, None to drop, or a list of extra (name, bytes)."""
    zin = zipfile.ZipFile(src)
    extra = []
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for i in zin.infolist():
            data = edit(i.filename, zin.read(i))
            if data is None:
                continue
            if isinstance(data, tuple):
                data, more = data
                extra += more
            z.writestr(i.filename, data)
        for n, d in extra:
            z.writestr(n, d)


def recipe_op(pkg):
    """A realistic combined change used by most cases: fix tables then captions then header/footer clean."""
    r1 = ops.op_fix_tables(pkg)
    r2 = ops.op_captions(pkg)
    r3 = ops.op_hf_clean(pkg)
    return {"tables": r1["tables"], "captions": r2["fixed"], "hf": len(r3["rewritten"]),
            "field_changes": r2.get("field_changes", []) + r3.get("field_changes", [])}


RECIPE_ALLOWED = ops.OP_TABLES_ALLOWED + ops.OP_CAPTIONS_ALLOWED + ("fields",)


def word_opens(w, path):
    try:
        doc = w.open(path, readonly=True)
        n = doc.Paragraphs.Count
        w.close(doc)
        return True, n
    except Exception as e:
        return False, str(e)[:200]


def word_resave_diff(w, path, work):
    doc = w.open(path, readonly=True)
    out = os.path.join(work, "~resave_" + os.path.basename(path))
    if not out.lower().endswith(".docx"):
        out += ".docx"
    retry(doc.SaveAs2, os.path.abspath(out), 16)
    w.close(doc)
    a, b = fingerprint(Package(path)), fingerprint(Package(out))
    return [k for k in diff(a, b) if k.split("|", 1)[1] in STRUCT]


def run_engine(path, op=recipe_op, allowed=RECIPE_ALLOWED, **kw):
    h0 = sha(path) if os.path.exists(path) else None
    try:
        r = safewrite.apply(path, op, allowed=allowed, **kw)
        return {"engine": "applied", "changed": r.changed, "report": r.report}, h0
    except Refused as e:
        return {"engine": "refused", "message": str(e)}, h0
    except Exception as e:
        return {"engine": "error", "message": f"{type(e).__name__}: {e}"}, h0


holder_ref = {}


def evaluate(case, expect, path, w, work, res, h0):
    d = os.path.dirname(path)
    res["temp_left"] = [f for f in os.listdir(d) if f.startswith("~pe_tmp")]
    if res["engine"] in ("refused", "error"):
        if case == "P17_changed_on_disk" and "after_mutation" in holder_ref:
            h0 = holder_ref["after_mutation"]  # the other program's change is expected; ours is not
        res["original_intact"] = (sha(path) == h0) if h0 else True
        res["pass"] = res["original_intact"] and not res["temp_left"] and (expect != "REFUSE" or res["engine"] == "refused" or "crash" in case)
        if expect == "SAFE" and res["engine"] == "error":
            res["pass"] = False  # an unhandled error is never acceptable
    else:
        ok, info = word_opens(w, path)
        res["word_opens_output"] = ok
        if ok:
            res["word_resave_structural_diff"] = word_resave_diff(w, path, work)
        res["pass"] = expect == "SAFE" and ok and not res.get("word_resave_structural_diff") and not res["temp_left"]
    return res


def main():
    src, work, out = sys.argv[1:4]
    os.makedirs(work, exist_ok=True)
    results = []
    w = WordSession()

    only = os.environ.get("PE_ONLY")

    def case(name, expect, build, **kw):
        if only and not any(name.startswith(o) for o in only.split(",")):
            return
        path = os.path.join(work, f"{name}.docx")
        try:
            if os.path.exists(path):
                os.chmod(path, stat.S_IWRITE); os.remove(path)
            built = build(path)
            path = built if isinstance(built, str) else path
            res = {"case": name, "expect": expect}
            r, h0 = run_engine(path, **kw)
            res.update(r)
            evaluate(name, expect, path, w, work, res, h0)
        except Exception:
            res = {"case": name, "expect": expect, "pass": False, "crash": traceback.format_exc()[-800:]}
        results.append(res)
        print(json.dumps(res, default=str)[:600], flush=True)

    # ---------- damaged and hostile packages (must REFUSE)
    raw = open(src, "rb").read()
    case("P01_truncated", "REFUSE", lambda p: open(p, "wb").write(raw[: len(raw) // 2]))

    def crc(p):
        b = bytearray(raw)
        zi = zipfile.ZipFile(io.BytesIO(raw)).getinfo("word/document.xml")
        off = zi.header_offset + 30 + len(zi.filename) + len(zi.extra) + 200
        b[off] ^= 0xFF
        open(p, "wb").write(bytes(b))
    case("P02_bad_checksum", "REFUSE", crc)
    case("P03_malformed_xml", "REFUSE", lambda p: rezip(src, p, lambda n, d: d.replace(b"</w:body>", b"</w:bod>") if n == "word/document.xml" else d))
    case("P04_doctype_entity", "REFUSE", lambda p: rezip(src, p, lambda n, d: d.replace(b"<w:document", b'<!DOCTYPE x [<!ENTITY e SYSTEM "file:///c:/windows/win.ini">]><w:document', 1) if n == "word/document.xml" else d))
    case("P05_zip_bomb", "REFUSE", lambda p: rezip(src, p, lambda n, d: (d, [("word/media/bomb.bin", b"\0" * 200_000_000)]) if n == "word/document.xml" else d))
    case("P06_path_traversal", "REFUSE", lambda p: rezip(src, p, lambda n, d: (d, [("../evil.txt", b"x")]) if n == "word/document.xml" else d))
    case("P07_signed", "REFUSE", lambda p: rezip(src, p, lambda n, d: (d, [("_xmlsignatures/sig1.xml", b"<Signature/>")]) if n == "word/document.xml" else d))
    case("P08_not_office", "REFUSE", lambda p: open(p, "wb").write(b"This is a text file renamed to docx." * 100))

    # ---------- files saved by Word in formats or states the engine must refuse
    def word_saveas(fmt, password=None, protect=None, ext=".docx"):
        def b(p):
            doc = w.open(src, readonly=True)
            if protect is not None:
                doc.Protect(protect, False, "pw")
            target = p if ext == ".docx" else p[:-5] + ext
            if password:
                retry(doc.SaveAs2, os.path.abspath(target), fmt, False, password)
            else:
                retry(doc.SaveAs2, os.path.abspath(target), fmt)
            w.close(doc)
            return target
        return b
    case("P09_password", "REFUSE", word_saveas(16, password="secret"))
    case("P10_legacy_doc", "REFUSE", word_saveas(0, ext=".doc"))
    case("P11_strict_ooxml", "REFUSE", word_saveas(24))
    case("P12_restrict_editing", "REFUSE", word_saveas(16, protect=3))

    def readonly(p):
        shutil.copy2(src, p); os.chmod(p, stat.S_IREAD)
    case("P13_readonly_file", "REFUSE", readonly)

    # ---------- open in Word
    holder = holder_ref
    def open_in_word(p):
        shutil.copy2(src, p)
        holder["doc"] = w.open(p, readonly=False)
    case("P14_open_in_word", "REFUSE", open_in_word)
    try:
        w.close(holder["doc"])
    except Exception:
        pass

    # ---------- crash and race safety
    def copy(p):
        shutil.copy2(src, p)
    case("P15_crash_after_temp", "REFUSE", copy, _fail_at="after_temp")
    case("P16_crash_before_replace", "REFUSE", copy, _fail_at="before_replace")

    def mutate_during(pkg):
        with open(pkg.path, "ab") as f:
            f.write(b"\0")  # another program changes the file while the engine works
        holder["after_mutation"] = sha(pkg.path)
        return recipe_op(pkg)
    case("P17_changed_on_disk", "REFUSE", copy, op=mutate_during)

    def rogue_text(pkg):
        recipe_op(pkg)
        t = next(pkg.tree(pkg.main).iter(f"{{{W}}}t"))
        t.text = (t.text or "") + " sneaky edit"
        return {}
    case("P18_guard_text_change", "REFUSE", copy, op=rogue_text)

    def rogue_drop_bookmark(pkg):
        recipe_op(pkg)
        b = next(pkg.tree(pkg.main).iter(f"{{{W}}}bookmarkStart"))
        b.getparent().remove(b)
        return {}
    case("P19_guard_bookmark_loss", "REFUSE", copy, op=rogue_drop_bookmark)

    def rogue_order(pkg):
        recipe_op(pkg)
        for rpr in pkg.tree(pkg.main).iter(f"{{{W}}}rPr"):
            if rpr.getparent().tag == f"{{{W}}}r":
                etree.SubElement(rpr, f"{{{W}}}sz").set(f"{{{W}}}val", "20")
                etree.SubElement(rpr, f"{{{W}}}b")  # b after sz: schema order broken
                break
        return {}
    case("P20_guard_schema_order", "REFUSE", copy, op=rogue_order)

    # ---------- content the engine must carry through untouched (SAFE)
    def plain_paras(doc, n):
        """n plain body paragraphs: not in a table, Normal style, with text."""
        out, i = [], 10
        total = doc.Paragraphs.Count
        while len(out) < n and i < total:
            pr = doc.Paragraphs(i)
            if not pr.Range.Information(12) and pr.Style.NameLocal == "Normal" and len(pr.Range.Text) > 60:
                out.append(pr.Range)
            i += 3
        return out

    def with_word_content(kind):
        def b(p):
            shutil.copy2(src, p)
            doc = w.open(p, readonly=False, window=True)
            r1, r2, r3 = plain_paras(doc, 3)
            if kind == "content_controls":
                doc.ContentControls.Add(0, doc.Range(r1.Start, r1.End - 1))   # rich text
                doc.ContentControls.Add(1, doc.Range(r2.Start, r2.End - 1))   # plain text
                doc.ContentControls.Add(0, doc.Tables(2).Range)               # around a whole table
            elif kind == "textbox":
                tb = doc.Shapes.AddTextbox(1, 72, 72, 200, 60, r1)
                tb.TextFrame.TextRange.InsertAfter("Text box content survives.")
                doc.Shapes.AddShape(Type=1, Left=300, Top=72, Width=80, Height=40, Anchor=r2)
            elif kind == "ole":
                import openpyxl
                xl = os.path.join(work, "embed.xlsx")
                wb = openpyxl.Workbook(); wb.active["A1"] = "embedded"; wb.save(xl)
                r1.Collapse(0)
                doc.InlineShapes.AddOLEObject(FileName=os.path.abspath(xl), LinkToFile=False, DisplayAsIcon=False, Range=r1)
            elif kind == "comments":
                doc.Comments.Add(r1, "Reviewer comment survives.")
            doc.Save()
            w.close(doc)
        return b
    for k in ("content_controls", "textbox", "ole", "comments"):
        if k == "ole":
            continue  # embedding an Excel object hangs hidden Word on an invisible dialog; see report
        case(f"P2{['content_controls','textbox','ole','comments'].index(k)+1}_{k}", "SAFE", with_word_content(k))

    def unicode_text(p):
        uni = "Arabic \u0627\u0644\u0646\u0638\u0627\u0645, Hebrew \u05de\u05e2\u05e8\u05db\u05ea, CJK \u822a\u5929\u5668, emoji \U0001F680, combining e\u0301, NBSP\u00a0here, \u00b5m and \u2264 5 \u00b0C"
        def ed(n, d):
            if n != "word/document.xml":
                return d
            r = etree.fromstring(d)
            for i, t in enumerate(r.iter(f"{{{W}}}t")):
                if i % 50 == 7:
                    t.text = (t.text or "") + " " + uni
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            return etree.tostring(r, xml_declaration=True, encoding="UTF-8", standalone=True)
        rezip(src, p, ed)
    case("P25_unicode", "SAFE", unicode_text)

    # ---------- idempotency: the same recipe twice gives the same document
    def idem(p):
        shutil.copy2(src, p)
        safewrite.apply(p, recipe_op, allowed=RECIPE_ALLOWED)
        holder["first"] = {n: Package(p).raw[n] for n in ("word/document.xml", "word/styles.xml")}
        holder["first_sha"] = sha(p)
    case("P26_idempotent", "SAFE", idem)
    if only:
        w.quit()
        json.dump({"results": results}, open(out, "w"), indent=1, default=str)
        print(json.dumps({"passed": sum(1 for r in results if r.get("pass")), "of": len(results)}))
        return
    second = os.path.join(work, "P26_idempotent.docx")
    results[-1]["second_run_identical"] = holder.get("first_sha") == sha(second)
    results[-1]["second_run_changed_flag"] = results[-1].get("changed")
    results[-1]["pass"] = results[-1]["pass"] and results[-1]["second_run_identical"]

    # ---------- fuzz: damaged-but-openable inputs must never get worse
    rnd = random.Random(7)
    base_xml = zipfile.ZipFile(src).read("word/document.xml")
    fuzz = {"variants": 0, "word_rejects_input": 0, "engine_refused": 0, "engine_applied": 0,
            "made_worse": 0, "errors": 0, "examples": []}
    for k in range(30):
        r = etree.fromstring(base_xml)
        body = r.find(f"{{{W}}}body")
        els = [e for e in body.iter() if e is not body and e.getparent() is not None]
        for _ in range(3):
            e = rnd.choice(els)
            act = rnd.choice(["drop", "dup", "swap", "attr"])
            par = e.getparent()
            if par is None:
                continue
            if act == "drop":
                par.remove(e)
            elif act == "dup":
                e.addnext(etree.fromstring(etree.tostring(e)))
            elif act == "swap" and e.getnext() is not None:
                nx = e.getnext(); par.remove(nx); e.addprevious(nx)
            elif act == "attr" and e.attrib:
                a = rnd.choice(list(e.attrib)); e.set(a, "999999")
        data = etree.tostring(r, xml_declaration=True, encoding="UTF-8", standalone=True)
        p = os.path.join(work, f"F{k:02d}.docx")
        rezip(src, p, lambda n, d: data if n == "word/document.xml" else d)
        fuzz["variants"] += 1
        ok_in, _ = word_opens(w, p)
        if not ok_in:
            fuzz["word_rejects_input"] += 1
        res, h0 = run_engine(p)
        if res["engine"] == "refused":
            fuzz["engine_refused"] += 1
            if sha(p) != h0:
                fuzz["made_worse"] += 1
            continue
        if res["engine"] == "error":
            fuzz["errors"] += 1
            fuzz["examples"].append(res["message"][:160])
            if sha(p) != h0:
                fuzz["made_worse"] += 1
            continue
        fuzz["engine_applied"] += 1
        ok_out, info = word_opens(w, p)
        if ok_in and not ok_out:
            fuzz["made_worse"] += 1
            fuzz["examples"].append(f"F{k:02d}: Word opened input but not output: {info}")
    fuzz["pass"] = fuzz["made_worse"] == 0
    results.append({"case": "P27_fuzz_30", "expect": "never worse", **fuzz})
    print(json.dumps(results[-1])[:800], flush=True)

    w.quit()
    summary = {"cases": len(results), "passed": sum(1 for r in results if r.get("pass")),
               "failed": [r["case"] for r in results if not r.get("pass")]}
    json.dump({"summary": summary, "results": results}, open(out, "w"), indent=1, default=str)
    print(json.dumps(summary))


if __name__ == "__main__":
    main()
