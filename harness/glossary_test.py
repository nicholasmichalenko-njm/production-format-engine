"""Gate M4: Quick Parts edited in the XML must be accepted by Word exactly as intended.

1. Word builds a Quick Parts master template with 8 entries (tables, fields, pictures, alignment tabs).
2. Word builds an empty effort template (no Quick Parts).
3. Engine, through safewrite: rename, recategorize, describe, duplicate, delete on the master;
   provision 5 entries (2 with pictures) into the empty effort template; provision again (conflict, no copy).
4. Word opens each result, lists BuildingBlockEntries, checks name, category, description, and inserts every
   entry into a new document: text and picture counts must match the XML.
Usage: python glossary_test.py <workdir> <logo.png> <results.json>
"""
import json
import os
import shutil
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from engine import safewrite
from engine.pkg import Package, Refused
from engine.ops import glossary as G
from wordsession import WordSession, retry

TYPE_QP = 1  # wdTypeQuickParts


def build_master(w, path, logo):
    d0 = w.app.Documents.Add()
    retry(d0.SaveAs2, os.path.abspath(path), 14)  # wdFormatXMLTemplate
    w.close(d0)
    doc = w.app.Documents.Add(os.path.abspath(path))  # a document attached to the new template
    tpl = doc.AttachedTemplate
    if os.path.normcase(tpl.FullName) != os.path.normcase(os.path.abspath(path)):
        raise RuntimeError(f"attached template is {tpl.FullName}; refusing to write building blocks there")
    specs = []

    def add(name, cat, build, desc):
        r = doc.Content
        r.Collapse(0)
        start = r.Start
        build(r)
        block = doc.Range(start, doc.Content.End - 1)
        tpl.BuildingBlockEntries.Add(name, TYPE_QP, cat, block, desc, 1)
        specs.append(name)
        doc.Content.InsertParagraphAfter()

    def hdr(r):
        t = doc.Tables.Add(r, 1, 3)
        t.Cell(1, 1).Range.InlineShapes.AddPicture(os.path.abspath(logo), False, True)
        f = t.Cell(1, 2).Range
        f.Collapse(1); doc.Fields.Add(f, -1, 'STYLEREF "Title" \\* CHARFORMAT', False)
        t.Cell(1, 3).Range.Text = "EX-26-R-0001"
    def ftr(r):
        r.Text = "Example Agency\tProprietary\t"
        e = doc.Range(r.End, r.End); doc.Fields.Add(e, 33, "", False)  # PAGE
    def ftr_pl(r):
        r.Text = "Example Agency (page limited)\t"
        e = doc.Range(r.End, r.End); doc.Fields.Add(e, 33, "", False)
    def cover(r):
        r.InlineShapes.AddPicture(os.path.abspath(logo), False, True)
        r2 = doc.Content; r2.Collapse(0); r2.InsertAfter("Example Program Technical Volume")
    def kp(r):
        t = doc.Tables.Add(r, 4, 3)
        t.Cell(1, 1).Merge(t.Cell(1, 3))
        c = t.Cell(1, 1).Range; c.Text = "Table "
        e = doc.Range(c.End - 1, c.End - 1); doc.Fields.Add(e, -1, "SEQ Table \\* ARABIC", False)
        for i, h in enumerate(["Name", "Role", "Clearance"]):
            t.Cell(2, i + 1).Range.Text = h
    def fig(r):
        t = doc.Tables.Add(r, 2, 1)
        t.Cell(1, 1).Range.InlineShapes.AddPicture(os.path.abspath(logo), False, True)
        t.Cell(2, 1).Range.Text = "Figure 1. Placeholder"
    def intro(r):
        r.Text = "Sierra Space will deliver the program with a proven, traceable approach."
    def close(r):
        r.Text = "We look forward to supporting the mission."
    add("10 Main Header", "2 Headers", hdr, "Logo, title, solicitation [QP:0001]")
    add("20 Main Footer", "3 Footers", ftr, "Client and page [QP:0002]")
    add("20 Main Footer [PL]", "3 Footers", ftr_pl, "Page-limited variant [QP:0003]")
    add("30 Blue Cover", "Covers", cover, "Cover with image [QP:0004]")
    add("60 Key Personnel Table", "Tables", kp, "Table with caption row [QP:0005]")
    add("70 Figure Placeholder", "Figures", fig, "Figure table [QP:0006]")
    add("80 Boilerplate Intro", "Text", intro, "Intro paragraph [QP:0007]")
    add("81 Boilerplate Close", "Text", close, "Closing paragraph [QP:0008]")
    tpl.Save()
    doc.Saved = True
    w.close(doc)
    return specs


def word_entries(w, path, insert=True):
    """BuildingBlockEntries of a template as Word sees them, and the result of inserting each."""
    doc = w.open(path, readonly=True)
    tpl = [t for t in w.app.Templates if os.path.normcase(t.FullName) == os.path.normcase(os.path.abspath(path))]
    out = []
    if tpl:
        t = tpl[0]
        w.app.Templates.LoadBuildingBlocks()
        n = t.BuildingBlockEntries.Count
        target = w.app.Documents.Add() if insert else None
        for i in range(1, n + 1):
            e = t.BuildingBlockEntries(i)
            row = {"name": e.Name, "category": e.Category.Name, "type": e.Type.Index, "description": e.Description}
            if insert:
                target.Content.Delete()
                rng = target.Content; rng.Collapse(1)
                ins = e.Insert(rng, True)
                row["inserted_chars"] = len(target.Content.Text.strip())
                row["inserted_pictures"] = target.InlineShapes.Count
                row["inserted_tables"] = target.Tables.Count
            out.append(row)
        if insert:
            target.Saved = True
            w.close(target)
    w.close(doc)
    return sorted(out, key=lambda r: (r["category"], r["name"]))


def xml_entries(path):
    return sorted([{k: e[k] for k in ("name", "category", "description", "pictures")} for e in G.entries(Package(path))],
                  key=lambda r: (r["category"], r["name"]))


def main():
    work, logo, out = sys.argv[1:4]
    os.makedirs(work, exist_ok=True)
    res = {"steps": []}
    master = os.path.join(work, "QP_Master.dotx")
    effort = os.path.join(work, "Effort_Template.dotx")
    for p in (master, effort):
        if os.path.exists(p):
            os.remove(p)
    w = WordSession()
    try:
        res["built"] = build_master(w, master, logo)
        d = w.app.Documents.Add(); retry(d.SaveAs2, os.path.abspath(effort), 14); w.close(d)
        res["word_master_before"] = word_entries(w, master, insert=False)
        w.quit(); time.sleep(2)  # templates are released only when Word closes
        steps = [
            ("rename", master, G.op_rename, {"mapping": {"10 Main Header": "11 Main Header", "81 Boilerplate Close": "82 Boilerplate Close"}}),
            ("recategorize", master, G.op_recategorize, {"mapping": {"30 Blue Cover": "1 Covers"}}),
            ("describe", master, G.op_describe, {"mapping": {"70 Figure Placeholder": "Figure table, picture in row 1 [QP:0006]"}}),
            ("duplicate", master, G.op_duplicate, {"name": "80 Boilerplate Intro", "new_name": "80 Boilerplate Intro [PL]"}),
            ("delete", master, G.op_delete, {"names": ["20 Main Footer [PL]"]}),
            ("provision", effort, G.op_provision, {"src_path": master, "names": ["11 Main Header", "20 Main Footer", "30 Blue Cover", "60 Key Personnel Table", "70 Figure Placeholder"]}),
            ("provision_again_conflict", effort, G.op_provision, {"src_path": master, "names": ["20 Main Footer"]}),
        ]
        for name, path, fn, args in steps:
            try:
                r = safewrite.apply(path, fn, allowed=G.GLOSS_ALLOWED, op_args=args)
                res["steps"].append({"step": name, "ok": True, "changed": r.changed, "report": r.report, "seconds": round(r.seconds, 2)})
            except Refused as e:
                res["steps"].append({"step": name, "ok": False, "refused": str(e)})
        w = WordSession()
        for label, path in (("master", master), ("effort", effort)):
            got = word_entries(w, path)
            exp = xml_entries(path)
            res[f"word_{label}"] = got
            res[f"xml_{label}"] = exp
            res[f"{label}_match"] = [(g["name"], g["category"], g["description"]) for g in got] == [(e["name"], e["category"], e["description"]) for e in exp]
            res[f"{label}_pictures_match"] = all(g["inserted_pictures"] >= e["pictures"] for g, e in zip(got, exp))
            res[f"{label}_all_insert"] = all(g["inserted_chars"] > 0 or g["inserted_pictures"] > 0 for g in got)
            res[f"{label}_entries"] = len(got)
        # non-vacuous: the master ends with 8 entries (8 built, 1 duplicated, 1 deleted), the effort with the 5 provisioned
        res["entry_counts_ok"] = res.get("master_entries") == 8 and res.get("effort_entries") == 5
    except Exception:
        res["crash"] = traceback.format_exc()
    finally:
        try:
            w.quit()
        except Exception:
            pass
    json.dump(res, open(out, "w"), indent=1, default=str)
    print(json.dumps({k: v for k, v in res.items() if not k.startswith(("word_", "xml_"))}, default=str)[:3000])


if __name__ == "__main__":
    main()
