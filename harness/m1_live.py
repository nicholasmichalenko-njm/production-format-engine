"""Gate M1: the live bridge. Word is driven from outside (as the app would drive the user's open Word).

For each of N tables:
  A. round trip   read Table.Range.WordOpenXML, InsertXML it back unchanged, inside one custom undo record
  B. transform    read the same XML, fix the table with the engine's fix_one_table, InsertXML it back
  C. undo         after B, one Undo must give back the original document
Each result is saved by Word and compared with a baseline saved by Word with no change.
Also: B's table must match what the file engine produces for the same table.
Usage: python m1_live.py <small.docx> <workdir> <results.json> [N]
"""
import json
import os
import shutil
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from lxml import etree
from engine import safewrite
from engine.pkg import Package, NS, W, parse_xml
from engine.fingerprint import fingerprint, diff, tables as table_sigs, visible_text
from engine.ops import word_ops as ops
from wordsession import WordSession, retry

PKGNS = "http://schemas.microsoft.com/office/2006/xmlPackage"
FACETS = ["text", "text_len", "deleted_text", "fields", "bookmarks", "bookmark_context", "tables", "drawings",
          "drawing_xml", "revisions", "paragraphs", "section_count", "note_refs", "hyperlinks", "content_controls"]


def table_signature(tbl):
    """What matters about a fixed table: grid shape, text, paragraph styles, fills, header rows, borders, width."""
    rows = []
    for tr in tbl.findall("w:tr", NS):
        hdr = tr.find("w:trPr/w:tblHeader", NS) is not None
        cells = []
        for tc in tr.findall("w:tc", NS):
            shd = tc.find("w:tcPr/w:shd", NS)
            styles = tuple(ops.para_style(p) for p in tc.findall("w:p", NS))
            cells.append((styles, (shd.get(f"{{{W}}}fill") if shd is not None else None), visible_text(tc)))
        rows.append((hdr, tuple(cells)))
    tw = tbl.find("w:tblPr/w:tblW", NS)
    b = tbl.find("w:tblPr/w:tblBorders/w:insideH", NS)
    return (tw.get(f"{{{W}}}type") if tw is not None else None, tw.get(f"{{{W}}}w") if tw is not None else None,
            b.get(f"{{{W}}}color") if b is not None else None, tuple(rows))


WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"


def drawings_without_session_ids(path):
    """Picture XML with Word's per-session tracking IDs (wp14:anchorId, wp14:editId) removed.
    Word writes new random values for these every time a document with a window is saved."""
    p = Package(path)
    out = []
    for d in p.tree(p.main).iter(f"{{{W}}}drawing"):
        c = etree.fromstring(etree.tostring(d))
        for el in c.iter():
            for a in [a for a in el.attrib if a.startswith(f"{{{WP14}}}")]:
                del el.attrib[a]
        out.append(etree.tostring(c, method="c14n"))
    return out


def body_tables(path):
    p = Package(path)
    return [t for t in p.tree(p.main).find("w:body", NS).findall("w:tbl", NS)]


def flat_document(xml):
    root = parse_xml(xml.encode("utf-8") if isinstance(xml, str) else xml, "flat OPC")
    for part in root.findall(f"{{{PKGNS}}}part"):
        if part.get(f"{{{PKGNS}}}name") == "/word/document.xml":
            return root, part.find(f"{{{PKGNS}}}xmlData")[0]
    raise RuntimeError("no document part in WordOpenXML")


def main():
    src, work, out = sys.argv[1:4]
    n = int(sys.argv[4]) if len(sys.argv) > 4 else 8
    os.makedirs(work, exist_ok=True)
    res = {"cases": []}
    # prepare: tag styles must exist in the document (file mode, through safewrite)
    prep = os.path.join(work, "live_base.docx")
    shutil.copy2(src, prep)
    safewrite.apply(prep, lambda pkg: ops.op_styles(pkg, ops.TAG_STYLES))
    # file-mode reference for every table
    ref = os.path.join(work, "live_filemode_ref.docx")
    shutil.copy2(prep, ref)
    safewrite.apply(ref, ops.op_fix_tables, allowed=ops.OP_TABLES_ALLOWED)
    w = WordSession()
    try:
        # the baseline is saved under the same conditions as the tests (a document window): with a window,
        # Word stamps wp14 anchor and edit IDs on every picture when it saves
        doc = w.open(prep, readonly=False, window=True)
        baseline = os.path.join(work, "live_baseline.docx")
        retry(doc.SaveAs2, os.path.abspath(baseline), 16)
        w.close(doc)
        base_fp = fingerprint(Package(baseline))
        base_drawings = drawings_without_session_ids(baseline)
        ref_tables = body_tables(ref)
        sid = ops.tag_style_ids(Package(prep).tree(Package(prep).word_part("/styles")))
        # pick N data tables (not figure tables), spread through the document
        cand = [i for i, t in enumerate(body_tables(baseline), 1) if not ops.is_figure_table(t)]
        picks = cand[:: max(1, len(cand) // n)][:n]
        for ti in picks:
            row = {"table": ti}
            for mode in ("roundtrip", "transform", "transform_then_undo"):
                # a document window is required: without one Word records the change but Undo does nothing
                doc = w.open(prep, readonly=False, window=True)
                try:
                    rng = doc.Tables(ti).Range
                    t = time.perf_counter()
                    xml = rng.WordOpenXML
                    row[f"{mode}_read_s"] = round(time.perf_counter() - t, 3)
                    if mode != "roundtrip":
                        root, docxml = flat_document(xml)
                        tbls = docxml.iter(f"{{{W}}}tbl")
                        tbl = next(tbls)
                        ops.fix_one_table(tbl, sid)
                        xml = etree.tostring(root, encoding="unicode")
                    ur = w.app.UndoRecord
                    ur.StartCustomRecord("PE fix table")
                    t = time.perf_counter()
                    # Replace, not insert into: InsertXML on Table.Range nests the new table in cell 1
                    start = rng.Start
                    doc.Tables(ti).Delete()
                    doc.Range(start, start).InsertXML(xml)
                    row[f"{mode}_insert_s"] = round(time.perf_counter() - t, 3)
                    ur.EndCustomRecord()
                    if mode == "transform_then_undo":
                        doc.Undo(1)
                    outp = os.path.join(work, f"live_t{ti}_{mode}.docx")
                    retry(doc.SaveAs2, os.path.abspath(outp), 16)
                finally:
                    doc.Saved = True
                    w.close(doc)
                fp = fingerprint(Package(outp))
                d = [k for k in diff(base_fp, fp) if k.split("|", 1)[1] in FACETS and not k.endswith("|drawing_xml")]
                if drawings_without_session_ids(outp) != base_drawings:
                    d.append("pictures differ (ignoring Word session IDs)")
                got_t = body_tables(outp)
                base_t = body_tables(baseline)
                if mode != "transform":
                    # every table, formatting included, must equal the baseline
                    same = len(got_t) == len(base_t) and all(table_signature(a) == table_signature(b) for a, b in zip(base_t, got_t))
                    if not same:
                        d.append("table formatting differs from baseline")
                row[f"{mode}_diff"] = d
                if mode == "transform":
                    got = body_tables(outp)
                    row["transform_matches_file_mode"] = (len(got) == len(ref_tables) and
                                                         table_signature(got[ti - 1]) == table_signature(ref_tables[ti - 1]))
                    # all other tables untouched
                    base_t = body_tables(baseline)
                    row["other_tables_untouched"] = all(table_signature(a) == table_signature(b)
                                                        for k, (a, b) in enumerate(zip(base_t, got), 1) if k != ti)
            res["cases"].append(row)
            print(json.dumps(row), flush=True)
        # COM call cost from outside the process
        doc = w.open(prep, readonly=True)
        t = time.perf_counter()
        cnt = min(2000, doc.Paragraphs.Count)
        ps = doc.Paragraphs
        for i in range(1, cnt + 1, 1):
            _ = ps(i).Range.Start
        res["com_paragraph_reads"] = {"count": cnt, "seconds": round(time.perf_counter() - t, 2), "note": "Paragraphs(i) indexing"}
        t = time.perf_counter(); k = 0
        for para in doc.Paragraphs:
            _ = para.Range.Start; k += 1
            if k >= cnt:
                break
        res["com_paragraph_iter"] = {"count": k, "seconds": round(time.perf_counter() - t, 2), "note": "for-each iteration"}
        w.close(doc)
    except Exception:
        res["crash"] = traceback.format_exc()
    finally:
        w.quit()
    rt = [c for c in res["cases"]]
    res["summary"] = {
        "tables": len(rt),
        "roundtrip_clean": sum(1 for c in rt if not c.get("roundtrip_diff")),
        "transform_clean_structure": sum(1 for c in rt if not c.get("transform_diff")),
        "transform_matches_file_mode": sum(1 for c in rt if c.get("transform_matches_file_mode")),
        "other_tables_untouched": sum(1 for c in rt if c.get("other_tables_untouched")),
        "undo_restores_original": sum(1 for c in rt if not c.get("transform_then_undo_diff")),
    }
    json.dump(res, open(out, "w"), indent=1)
    print(json.dumps(res["summary"]), json.dumps(res.get("com_paragraph_reads")), json.dumps(res.get("com_paragraph_iter")), res.get("crash", "")[:1500])


if __name__ == "__main__":
    main()
