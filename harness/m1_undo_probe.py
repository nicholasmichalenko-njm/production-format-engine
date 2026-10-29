"""How does Undo behave for a live table fix done from outside Word?"""
import json, os, shutil, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from lxml import etree
from engine.pkg import Package, NS, W
from engine.ops import word_ops as ops
from m1_live import flat_document, table_signature, body_tables
from wordsession import WordSession, retry

work = os.path.join(os.path.dirname(HERE), "results", "v2", "m1")
prep = os.path.join(work, "live_base.docx")
baseline = os.path.join(work, "live_baseline.docx")
sid = ops.tag_style_ids(Package(prep).tree(Package(prep).word_part("/styles")))
base_sig = table_signature(body_tables(baseline)[3])
out = {}
w = WordSession()
try:
    for variant in ("readonly_doc", "editable_doc", "editable_window"):
        cp = os.path.join(work, f"undo_probe_{variant}.docx")
        shutil.copy2(prep, cp)
        doc = w.open(cp, readonly=(variant == "readonly_doc"), window=(variant == "editable_window"))
        rng = doc.Tables(4).Range
        root, docxml = flat_document(rng.WordOpenXML)
        ops.fix_one_table(next(docxml.iter(f"{{{W}}}tbl")), sid)
        xml = etree.tostring(root, encoding="unicode")
        doc.UndoClear()
        ur = w.app.UndoRecord
        ur.StartCustomRecord("PE fix table")
        rec = ur.IsRecordingCustomRecord
        start = rng.Start
        doc.Tables(4).Delete()
        doc.Range(start, start).InsertXML(xml)
        ur.EndCustomRecord()
        steps = []
        for k in range(1, 6):
            okk = doc.Undo(1)
            p = os.path.join(work, f"undo_probe_{variant}_{k}.docx")
            retry(doc.SaveAs2, os.path.abspath(p), 16)
            sig = table_signature(body_tables(p)[3])
            steps.append({"undo": k, "returned": bool(okk), "restored": sig == base_sig})
            if sig == base_sig or not okk:
                break
        out[variant] = {"recording": bool(rec), "steps": steps}
        doc.Saved = True
        w.close(doc)
        print(variant, json.dumps(out[variant]), flush=True)
finally:
    w.quit()
json.dump(out, open(os.path.join(work, "undo_probe.json"), "w"), indent=1)
