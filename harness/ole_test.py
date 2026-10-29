"""Gate M7: embedded and linked objects (Excel in a proposal volume) must survive every operation, and Word
must never hang or reach outside the PC while FormatCheck checks a document that holds them.

Objects injected into a copy of a fixture:
  E1  embedded Excel worksheet in a body paragraph (w:object, package part in word/embeddings)
  E2  embedded Excel worksheet inside the first table cell
  L1  LINK field, automatic update, to an Excel file that exists on this PC
  L2  LINK field, automatic update, to a local path that does not exist
  L3  LINK field, automatic update, to a network path that does not answer (\\\\10.255.255.1)
  L4  linked OLE object (o:OLEObject Type="Link", UpdateMode="Always") to a path that does not exist

Checks:
  1. Word opens the raw file and saves it (watchdog, own process). This is the Word-written test volume.
  2. Engine: every op in memory and the full recipe through safewrite. Every embedded part byte identical,
     every w:object element identical, every LINK instruction identical, every external relationship identical.
  3. FormatCheck's own Word check on the result, with a time limit: opens, no re-save drift, no hang.
  4. Word lists the objects in the result (type, ProgID, link source) without activating them: same as input.
  5. No Excel or Word process is left behind, and the Excel file behind L1 is not changed.
Usage: python ole_test.py <fixture.docx> <logo.png> <workdir> <results.json> [--no-word]
"""
import copy
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(HERE)
sys.path.insert(0, LAB); sys.path.insert(0, HERE); sys.path.insert(0, os.path.join(LAB, "fieldcheck"))
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
T_PKG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
T_IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
T_OLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NSMAP = {"w": W, "r": R, "v": V, "o": O}


def q(ns, tag):
    return f"{{{ns}}}{tag}"


def make_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in (["Labor category", "Hours", "Rate"], ["Engineer III", 1880, 142.5], ["Analyst II", 960, 98.0]):
        ws.append(row)
    wb.save(path)


def object_xml(shape_n, img_rid, obj_rid, typ, extra=""):
    return (f'<w:r xmlns:w="{W}" xmlns:r="{R}" xmlns:v="{V}" xmlns:o="{O}"><w:object w:dxaOrig="4320" w:dyaOrig="1440">'
            f'<v:shape id="_x0000_i{shape_n}" type="#_x0000_t75" style="width:216pt;height:72pt" o:ole="">'
            f'<v:imagedata r:id="{img_rid}" o:title=""/></v:shape>'
            f'<o:OLEObject Type="{typ}" ProgID="Excel.Sheet.12" ShapeID="_x0000_i{shape_n}" DrawAspect="Content" '
            f'ObjectID="_17000000{shape_n}" r:id="{obj_rid}"{extra}/></w:object></w:r>')


def link_field_runs(instr, result):
    instr = instr.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")
    return [
        f'<w:r xmlns:w="{W}"><w:fldChar w:fldCharType="begin"/></w:r>',
        f'<w:r xmlns:w="{W}"><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>',
        f'<w:r xmlns:w="{W}"><w:fldChar w:fldCharType="separate"/></w:r>',
        f'<w:r xmlns:w="{W}"><w:t>{result}</w:t></w:r>',
        f'<w:r xmlns:w="{W}"><w:fldChar w:fldCharType="end"/></w:r>',
    ]


def inject(src, dst, xlsx, logo, unc):
    zi = zipfile.ZipFile(src)
    parts = {i.filename: zi.read(i.filename) for i in zi.infolist()}
    infos = {i.filename: i for i in zi.infolist()}
    doc = etree.fromstring(parts["word/document.xml"])
    rels = etree.fromstring(parts["word/_rels/document.xml.rels"])
    ct = etree.fromstring(parts["[Content_Types].xml"])
    used = {r.get("Id") for r in rels}

    def rel(typ, target, external=False):
        n = 9000 + len(used)
        rid = f"rId{n}"
        used.add(rid)
        e = etree.SubElement(rels, q(PR, "Relationship"), Id=rid, Type=typ, Target=target)
        if external:
            e.set("TargetMode", "External")
        return rid

    xbytes = open(xlsx, "rb").read()
    parts["word/embeddings/Microsoft_Excel_Worksheet_E1.xlsx"] = xbytes
    parts["word/embeddings/Microsoft_Excel_Worksheet_E2.xlsx"] = xbytes
    parts["word/media/ole_preview.png"] = open(logo, "rb").read()
    exts = {d.get("Extension").lower() for d in ct.findall(q(CT, "Default"))}
    if "xlsx" not in exts:
        etree.SubElement(ct, q(CT, "Default"), Extension="xlsx", ContentType=XLSX_CT)
    if "png" not in exts:
        etree.SubElement(ct, q(CT, "Default"), Extension="png", ContentType="image/png")
    img = rel(T_IMG, "media/ole_preview.png")
    e1 = rel(T_PKG, "embeddings/Microsoft_Excel_Worksheet_E1.xlsx")
    e2 = rel(T_PKG, "embeddings/Microsoft_Excel_Worksheet_E2.xlsx")
    l4 = rel(T_OLE, "file:///C:\\PE_missing\\Linked_Missing.xlsx", external=True)

    body = doc.find(q(W, "body"))
    # anchor on a body paragraph outside every field (not inside a TOC result), with text and no field
    depth, free = 0, []
    for p in body.findall(q(W, "p")):
        start = depth
        has_fld = False
        for e in p.iter(q(W, "fldChar")):
            has_fld = True
            k = e.get(q(W, "fldCharType"))
            depth += (k == "begin") - (k == "end")
        if start == 0 and depth == 0 and not has_fld and p.find(f".//{q(W, 't')}") is not None:
            free.append(p)
    anchor = free[len(free) // 3]

    def new_par(runs, label):
        p = etree.Element(q(W, "p"))
        r = etree.SubElement(p, q(W, "r")); t = etree.SubElement(r, q(W, "t")); t.text = label + " "
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for x in runs:
            p.append(etree.fromstring(x))
        return p

    xl = xlsx.replace("\\", "\\\\")
    new = [
        new_par([object_xml(1025, img, e1, "Embed")], "E1"),
        new_par(link_field_runs(f'LINK Excel.Sheet.12 "{xl}" "Sheet1!R1C1:R3C3" \\a \\t', "L1 linked table"), "L1"),
        new_par(link_field_runs('LINK Excel.Sheet.12 "C:\\\\PE_missing\\\\Missing.xlsx" "Sheet1!R1C1:R3C3" \\a \\t', "L2 missing"), "L2"),
        new_par(link_field_runs(f'LINK Excel.Sheet.12 "{unc}" "Sheet1!R1C1:R3C3" \\a \\t', "L3 network"), "L3"),
        new_par([object_xml(1027, img, l4, "Link", ' UpdateMode="Always"')], "L4"),
        new_par(link_field_runs('INCLUDETEXT "C:\\\\PE_missing\\\\Boilerplate.docx"', "L5 included text"), "L5"),
        new_par(link_field_runs('INCLUDEPICTURE "C:\\\\PE_missing\\\\Figure.png" \\d', "L6 linked picture"), "L6"),
    ]
    for p in reversed(new):
        anchor.addnext(p)
    tc = body.find(f".//{q(W, 'tc')}")
    if tc is not None:
        tc.find(q(W, "p")).append(etree.fromstring(object_xml(1026, img, e2, "Embed")))
    parts["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for name, b in parts.items():
            zo.writestr(infos.get(name) or zipfile.ZipInfo(name, (2026, 10, 29, 0, 0, 0)), b, zipfile.ZIP_DEFLATED)


def fresh_pair(a, b, work):
    import uuid
    t = uuid.uuid4().hex[:8]
    fa, fb = os.path.join(work, f"fc_{t}_b.docx"), os.path.join(work, f"fc_{t}_a.docx")
    shutil.copyfile(a, fa); shutil.copyfile(b, fb)
    return fa, fb


def drop_paragraphs(src, dst, labels):
    """Copy a docx without the body paragraphs whose text starts with one of the labels (E1, L1 ...)."""
    zi = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zi.infolist():
            b = zi.read(i.filename)
            if i.filename == "word/document.xml" and labels:
                doc = etree.fromstring(b)
                for p in list(doc.find(q(W, "body")).findall(q(W, "p"))):
                    t = "".join(x.text or "" for x in p.iter(q(W, "t")))
                    if any(t.startswith(l + " ") for l in labels):
                        p.getparent().remove(p)
                b = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
            zo.writestr(i, b)


# ---------------------------------------------------------------- package-level evidence
def evidence(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    doc = etree.fromstring(z.read("word/document.xml"))
    objs = []
    for o in doc.iter(q(W, "object")):
        c = copy.deepcopy(o)
        objs.append(hashlib.sha256(etree.tostring(c, method="c14n")).hexdigest()[:16])
    instr, cur, depth = [], [], 0
    for e in doc.iter(q(W, "fldChar"), q(W, "instrText")):
        if e.tag == q(W, "fldChar"):
            k = e.get(q(W, "fldCharType"))
            if k == "begin":
                depth += 1
                if depth == 1:
                    cur = []
            elif k in ("separate", "end") and depth == 1 and cur is not None:
                s = "".join(cur).strip()
                if s.startswith("LINK"):
                    instr.append(s)
                cur = None
            if k == "end":
                depth -= 1
        elif depth == 1 and cur is not None:
            cur.append(e.text or "")
    for fs in doc.iter(q(W, "fldSimple")):
        if fs.get(q(W, "instr"), "").strip().startswith("LINK"):
            instr.append(fs.get(q(W, "instr")).strip())
    emb = {n: hashlib.sha256(z.read(n)).hexdigest()[:16] for n in names if n.startswith("word/embeddings/")}
    ext = []
    for n in names:
        if n.endswith(".rels"):
            for r in etree.fromstring(z.read(n)):
                if r.get("TargetMode") == "External":
                    ext.append(f"{n}:{r.get('Id')}:{r.get('Target')}")
    oleobj = [(e.get("Type"), e.get("ProgID"), e.get("UpdateMode")) for e in doc.iter(q(O, "OLEObject"))]
    return {"objects": objs, "ole": oleobj, "link_instr": [" ".join(i.split()) for i in instr],
            "embeddings": emb, "external": sorted(ext)}


# ---------------------------------------------------------------- Word steps, each in its own process
def pids(image):
    out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"], capture_output=True, text=True,
                         creationflags=0x08000000).stdout
    return {int(l.split('","')[1]) for l in out.splitlines() if l.startswith('"')}


def run_isolated(args, timeout):
    """Run a worker; end only the Word it started (exact PID from the worker), at once if a dialog shows."""
    import fieldcheck as FC
    out = args[-1]
    pid_file = out + ".pid"
    pre_x = pids("EXCEL.EXE")
    t = time.perf_counter()
    p = subprocess.Popen([sys.executable, os.path.abspath(__file__), "--worker"] + args, creationflags=0x08000000)
    r = {}
    while p.poll() is None:
        if time.perf_counter() - t > timeout:
            r["timed_out"] = True
            break
        try:
            d = FC.visible_dialogs(int(open(pid_file).read().strip()))
        except Exception:
            d = []
        if d:
            r["blocked_by_dialog"] = d[0]
            break
        time.sleep(1)
    if r:
        p.kill()
        r["word_ended"] = FC.end_word(pid_file)
    r["seconds"] = round(time.perf_counter() - t, 1)
    if os.path.exists(pid_file):
        os.remove(pid_file)
    time.sleep(2)
    r["excel_started"] = len(pids("EXCEL.EXE") - pre_x)
    return r


def worker(mode, a, b, out):
    import pythoncom, win32com.client
    import fieldcheck as FC
    pythoncom.CoInitialize()
    pre = FC.winword_pids()
    app = win32com.client.DispatchEx("Word.Application")
    mine = [x for x in FC.winword_pids() - pre if FC.is_automation_word(x)]
    if len(mine) == 1:
        open(out + ".pid", "w").write(str(mine[0]))
    res = {}
    try:
        app.Visible = False; app.DisplayAlerts = 0; app.ScreenUpdating = False; app.AutomationSecurity = 3
        orig = bool(app.Options.UpdateLinksAtOpen)
        res["update_links_at_open_option"] = orig
        app.Options.UpdateLinksAtOpen = False
        t = time.perf_counter()
        doc = app.Documents.Open(os.path.abspath(a), False, mode == "list", False, "", "", True, "", "", 0, 0, False, False, 0, True)
        res["open_s"] = round(time.perf_counter() - t, 1)
        ext_types = {56, 68, 67, 36, 45, 46, 55}
        if mode == "normalize":
            doc.SaveAs2(os.path.abspath(b), 16)
        elif mode in ("update_locked", "update_unlocked"):
            res["external_before"] = [(f.Type, f.Result.Text.strip()[:30]) for f in doc.Fields if f.Type in ext_types]
            if mode == "update_locked":
                for f in doc.Fields:
                    if f.Type in FC.EXTERNAL_FIELD_TYPES:
                        f.Locked = True
            t = time.perf_counter()
            doc.Fields.Update()
            res["update_s"] = round(time.perf_counter() - t, 1)
            res["external_after"] = [(f.Type, f.Result.Text.strip()[:30]) for f in doc.Fields if f.Type in ext_types]
            res["results_kept"] = res["external_before"] == res["external_after"]
        else:
            rows = []
            for s in doc.InlineShapes:
                row = {"type": s.Type}
                try:
                    row["progid"] = s.OLEFormat.ProgID
                except Exception:
                    row["progid"] = None
                try:
                    row["source"] = os.path.basename(s.LinkFormat.SourceFullName) if s.Type == 2 else None
                except Exception:
                    row["source"] = "n/a"
                if s.Type in (1, 2):
                    rows.append(row)
            res["ole_shapes"] = rows
            res["link_fields"] = sum(1 for f in doc.Fields if f.Type == 56)  # wdFieldLink
            res["link_results"] = [f.Result.Text.strip()[:30] for f in doc.Fields if f.Type == 56]
        doc.Close(0)
    except Exception as e:
        res["error"] = type(e).__name__ + " " + str(e)[:200]
    finally:
        try:
            app.Options.UpdateLinksAtOpen = res.get("update_links_at_open_option", True)
            app.NormalTemplate.Saved = True
            app.Quit(0)
        except Exception:
            pass
    json.dump(res, open(out, "w"))


def word_step(mode, a, b, work, timeout):
    import uuid
    fresh = os.path.join(work, f"j_{uuid.uuid4().hex[:8]}.docx")  # fresh name: Word remembers names after a forced stop
    shutil.copyfile(a, fresh)
    a = fresh
    out = os.path.join(work, f"w_{mode}_{int(time.time() * 1000)}.json")
    r = run_isolated([mode, a, b or "", out], timeout)
    if os.path.exists(out):
        r.update(json.load(open(out))); os.remove(out)
    return r


# ---------------------------------------------------------------- main
def main():
    fixture, logo, work, out = sys.argv[1:5]
    use_word = "--no-word" not in sys.argv
    os.makedirs(work, exist_ok=True)
    res = {"steps": {}}
    if use_word:
        import fieldcheck as FC0
        res["automation_word_at_start"] = sum(1 for x in pids("WINWORD.EXE") if FC0.is_automation_word(x))
    try:
        xlsx = os.path.abspath(os.path.join(work, "Linked_Source.xlsx"))
        make_xlsx(xlsx)
        x_hash = hashlib.sha256(open(xlsx, "rb").read()).hexdigest()
        raw = os.path.join(work, "ole_raw.docx")
        inject(fixture, raw, xlsx, logo, r"\\10.255.255.1\share\Network.xlsx")
        docs = {"raw": raw}
        if use_word:
            wd = os.path.join(work, "ole_word.docx")
            res["steps"]["word_normalize"] = word_step("normalize", raw, wd, work, 300)
            if os.path.exists(wd):
                docs["word"] = wd
        import fieldcheck as FC
        from engine.pkg import Package, Refused
        for label, src in docs.items():
            r = {}
            try:
                Package(src)
                r["load"] = "accepted"
            except Refused as e:
                r["load"] = "refused: " + str(e)[:200]
                res["steps"][label] = r
                continue
            ev0 = evidence(src)
            r["input"] = {"objects": len(ev0["objects"]), "ole": ev0["ole"], "links": len(ev0["link_instr"]),
                          "embeddings": len(ev0["embeddings"]), "external": len(ev0["external"])}
            r["dry"] = {k: v.get("result") + ("" if v.get("result") != "refused" else ": " + v.get("why", "")) for k, v in FC.dry_ops(src).items()}
            after = os.path.join(work, f"ole_{label}_after.docx")
            shutil.copyfile(src, after)
            steps = FC.recipe(after)
            r["recipe"] = {s["op"]: ("ok" if s["ok"] else s["why"]) for s in steps}
            ev1 = evidence(after)
            r["preserved"] = {k: ev0[k] == ev1[k] for k in ("objects", "ole", "link_instr", "embeddings", "external")}
            if use_word:
                r["formatcheck_word"] = FC.run_word(*fresh_pair(src, after, work), 600)
                r["list_before"] = word_step("list", src, None, work, 300)
                r["list_after"] = word_step("list", after, None, work, 300)
                sb = r["list_before"].get("ole_shapes"); sa = r["list_after"].get("ole_shapes")
                r["word_objects_same"] = sb is not None and sb == sa and len(sb) >= 3
                r["update_unlocked"] = word_step("update_unlocked", after, None, work, 300)
                r["update_locked"] = word_step("update_locked", after, None, work, 300)
            res["steps"][label] = r
        if use_word:
            out_ = subprocess.run([sys.executable, os.path.join(LAB, "fieldcheck", "fieldcheck.py"), "--option-check",
                                   "1" if res["steps"].get("word_normalize", {}).get("update_links_at_open_option", True) else "0"],
                                  capture_output=True, text=True, timeout=120).stdout
            res["word_link_option_check"] = json.loads(out_.strip().splitlines()[-1])
        res["linked_source_unchanged"] = hashlib.sha256(open(xlsx, "rb").read()).hexdigest() == x_hash
        if use_word:
            import fieldcheck as FC2
            res["automation_word_left"] = sum(1 for x in pids("WINWORD.EXE") if FC2.is_automation_word(x))
    except Exception:
        res["crash"] = traceback.format_exc()
    json.dump(res, open(out, "w"), indent=1, default=str)
    print(json.dumps(res, default=str)[:6000])


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        worker(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    else:
        main()
