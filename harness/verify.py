"""Run engine operations through safewrite on copies of a fixture, then prove the result with real Word.

For every case:
  A. engine: safewrite.apply on a fresh copy (all refusal rules, fingerprint guard, snapshot, atomic replace)
  B. word open: hidden Word opens the output with OpenAndRepair off; any error = FAIL
  C. word resave (no field update): Word saves a copy; our fingerprint of Word's file must equal ours for
     the structural facets (text, fields, bookmarks, tables, sections, drawings, revisions). A silent repair
     by Word would show here.
  D. settings read-back: each check expression is evaluated against the live Word object model
  E. computed checks after Word updates all fields: heading numbers, caption numbers, broken references,
     page number labels in the PDF
Usage: python verify.py <fixture.docx> <case name|all> <results.json> [--sample N] [--pdf]
"""
import json
import os
import re
import shutil
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)
from engine import safewrite
from engine.pkg import Package, Refused, NS, W
from engine.fingerprint import fingerprint, diff, visible_text
from engine.ops import word_ops as ops
from wordsession import WordSession, retry

STRUCT = ["text", "text_len", "deleted_text", "fields", "bookmarks", "bookmark_context", "tables", "drawings", "drawing_xml",
          "objects", "content_controls", "content_control_props", "revisions", "note_refs", "comment_refs",
          "section_count", "paragraphs", "alt_content", "hyperlinks", "media", "external_links"]


def oracle_roles(tbl):
    """Expected row roles from the fixture's own content (independent of the engine's detection)."""
    out = []
    for r, tr in enumerate(tbl.findall("w:tr", NS)):
        first = ops.cell_text(tr.findall("w:tc", NS)[0])
        whole = " ".join(ops.cell_text(tc) for tc in tr.findall("w:tc", NS))
        if r == 0 and re.match(r"^(Table|Figure) ", whole):
            out.append("caption")
        elif first == "Item":
            out.append("header")
        elif first.startswith("Phase"):
            out.append("subhead")
        elif first.startswith("Total"):
            out.append("total")
        elif first.startswith("Note"):
            out.append("note")
        else:
            out.append("body")
    return out

CASES = {
    "styles": (ops.op_styles, {"spec": {
        "Normal": {"font": "Arial", "size": 11, "after": 6},
        "Heading 1": {"font": "Arial", "size": 16, "bold": True, "color": "1F4E79", "before": 24, "after": 6, "keep_next": True},
        "Heading 2": {"font": "Arial", "size": 13, "bold": True, "color": "1F4E79", "before": 12, "after": 4},
        "Caption": {"font": "Arial", "size": 9, "bold": True, "color": "000000", "before": 0, "after": 0},
        "PE Win Theme": {"font": "Arial", "size": 11, "color": "1F4E79", "before": 6, "after": 6, "based_on": "Normal"}}},
        ops.OP_STYLES_ALLOWED),
    "remove_unused_styles": (ops.op_remove_unused_styles, {"names": ["Light Shading Accent 1", "Medium Grid 3 Accent 6", "Book Title", "Quote", "Intense Quote"]}, ()),
    "hide_styles": (ops.op_hide_styles, {"names": ["Title", "Subtitle", "Intense Quote", "Book Title"]}, ()),
    "heading_numbering": (ops.op_heading_numbering, {}, ()),
    "hf_clean": (ops.op_hf_clean, {}, ("fields",)),
    "page_numbers": (ops.op_page_numbers, {}, ("sections",)),
    "fix_tables": (ops.op_fix_tables, {}, ops.OP_TABLES_ALLOWED),
    "remove_empty_notes": (ops.op_remove_empty_notes, {}, ops.OP_NOTES_ALLOWED),
    "captions": (ops.op_captions, {}, ops.OP_CAPTIONS_ALLOWED),
    "xref_audit": (ops.op_xref_audit, {}, ()),
}
RECIPE = ["hf_clean", "styles", "heading_numbering", "page_numbers", "fix_tables", "remove_empty_notes", "captions", "xref_audit"]


# ---------------------------------------------------------------- read-back helpers (Word side)
def style_in_use(doc, name):
    try:
        return bool(doc.Styles(name).InUse)
    except Exception:
        return False


def expected_heading_strings(levels_seq):
    counts = [0] * 9
    out = []
    for lv in levels_seq:
        i = lv - 1
        counts[i] += 1
        for d in range(i + 1, 9):
            if d != 6:  # Heading 7 (appendix) never restarts
                counts[d] = 0
        if lv <= 6:
            out.append(".".join(str(counts[k]) for k in range(lv)))
        elif lv == 7:
            out.append("Appendix " + chr(64 + counts[6]))
        elif lv == 8:
            out.append(f"{chr(64 + counts[6])}.{counts[7]}")
        else:
            out.append(f"{chr(64 + counts[6])}.{counts[7]}.{counts[8]}")
    return out


def heading_list_strings(doc):
    rows = []
    for p in doc.ListParagraphs:
        name = p.Style.NameLocal
        m = re.fullmatch(r"Heading ([1-9])", name)
        if m:
            rng = p.Range
            rows.append((rng.Start, int(m.group(1)), rng.ListFormat.ListString))
    rows.sort()  # Word returns ListParagraphs in reverse document order
    levels = [r[1] for r in rows]
    got = [r[2] for r in rows]
    exp = expected_heading_strings(levels)
    bad = [(i, g, e) for i, (g, e) in enumerate(zip(got, exp)) if g != e]
    # every heading paragraph in the file must be numbered by Word (a heading that lost its number would
    # otherwise drop out of ListParagraphs unseen)
    pk = Package(doc.FullName)
    in_file = sum(1 for p in pk.tree(pk.main).iter(f"{{{W}}}p")
                  if (ops.para_style(p) or "").startswith("Heading") and (ops.para_style(p) or "")[7:].isdigit())
    return {"headings": len(got), "headings_in_file": in_file, "mismatch": len(bad) + abs(in_file - len(got)),
            "first_bad": bad[:5]}


def table_check(doc, tindex, roles, rows_to_check=None):
    t = doc.Tables(tindex)
    errs = []
    # Word reports PreferredWidth as 9999999 (undefined) on tables with merged cells; the XML check covers it
    if t.PreferredWidthType != 2 or (abs(t.PreferredWidth - 100) > 0.01 and t.PreferredWidth != 9999999):
        errs.append(f"width {t.PreferredWidthType}/{t.PreferredWidth}")
    for side in (-1, -2, -3, -4, -5, -6):  # top, left, bottom, right, inside horizontal, inside vertical
        b = t.Borders(side)
        if b.Color != ops.bgr("808080") or b.LineWidth != 4:
            errs.append(f"border {side}: {b.Color}/{b.LineWidth}")
    for c in t.Range.Cells:
        ri = c.RowIndex - 1
        role = roles[ri] if ri < len(roles) else "?"
        tag, fill, valign, repeat = ops.ROWTYPES[role]
        sn = c.Range.Paragraphs(1).Style.NameLocal
        if sn != tag:
            errs.append(f"r{ri+1} style {sn} != {tag}")
        col = c.Shading.BackgroundPatternColor
        want = ops.bgr(fill) if fill else -16777216
        if col != want:
            errs.append(f"r{ri+1} fill {col} != {want}")
        if role != "caption":
            exp_style = ops.TAG_STYLES[tag]
            f = c.Range.Font
            if f.Name != "Arial" or abs(f.Size - exp_style["size"]) > 0.01:
                errs.append(f"r{ri+1} font {f.Name} {f.Size}")
        if len(errs) > 6:
            break
    for r in (1, 2):
        try:
            if not t.Rows(r).HeadingFormat:
                errs.append(f"row {r} not repeat header")
        except Exception:
            errs_xml = "xml"  # rows not addressable with merged cells: checked in Word's saved XML instead
    return "ok" if not errs else errs


# ---------------------------------------------------------------- read-back helpers (Word-written XML)
def fields_with_results(root):
    out, stack = [], []
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        if tag == "fldChar":
            t = el.get(f"{{{W}}}fldCharType")
            if t == "begin":
                stack.append({"instr": "", "res": "", "phase": "instr"})
            elif t == "separate" and stack:
                stack[-1]["phase"] = "res"
            elif t == "end" and stack:
                f = stack.pop()
                out.append((" ".join(f["instr"].split()), f["res"]))
                if stack and stack[-1]["phase"] == "res":
                    stack[-1]["res"] += f["res"]
        elif tag == "instrText" and stack:
            stack[-1]["instr"] += el.text or ""
        elif tag == "t" and stack and stack[-1]["phase"] == "res":
            stack[-1]["res"] += el.text or ""
    return out


def word_chapter_numbers(doc):
    """Heading 1 and Heading 7 numbers exactly as Word displays them, in document order."""
    rows = []
    for p in doc.ListParagraphs:
        n = p.Style.NameLocal
        if n in ("Heading 1", "Heading 7"):
            rng = p.Range
            rows.append((rng.Start, n, rng.ListFormat.ListString))
    rows.sort()
    return [(n, s) for _, n, s in rows]


def xml_caption_numbers(path, chapters=None):
    """Expected caption number = the chapter number Word shows on the heading + the sequence in that chapter."""
    pkg = Package(path)
    root = pkg.tree(pkg.main)
    h1 = h7 = 0
    shown1 = [s for n, s in (chapters or []) if n == "Heading 1"]
    shown7 = [s.split()[-1] for n, s in (chapters or []) if n == "Heading 7"]
    seq = {}
    total = bad = 0
    examples = []
    for p in root.iter(f"{{{W}}}p"):
        st = ops.para_style(p)
        if st == "Heading1":
            h1 += 1
        elif st == "Heading7":
            h7 += 1
        if st != "Caption":
            continue
        txt = visible_text(p).replace("¶", "")
        m = re.match(r"^(Table|Figure) ([^.\t]+)", txt)
        if not m:
            continue
        label, num = m.groups()
        part = (label, "app", h7) if h7 else (label, "body", h1)
        seq[part] = seq.get(part, 0) + 1
        c1 = shown1[h1 - 1] if shown1 and 0 < h1 <= len(shown1) else str(h1)
        c7 = shown7[h7 - 1] if shown7 and 0 < h7 <= len(shown7) else chr(64 + h7)
        exp = f"{c7}-{seq[part]}" if h7 else f"{c1}-{seq[part]}"
        total += 1
        if num.strip() != exp:
            bad += 1
            if len(examples) < 6:
                examples.append((label, num, exp))
    return {"captions": total, "wrong": bad, "examples": examples}


def xml_error_refs(path):
    pkg = Package(path)
    out = set()
    for story in pkg.stories():
        for instr, res in fields_with_results(pkg.tree(story)):
            if "Error!" in res:
                parts = instr.split()
                out.add(parts[1] if len(parts) > 1 else instr)
    return sorted(out)


def xml_hf_codes(path):
    pkg = Package(path)
    bad = []
    for story in pkg.stories():
        r = pkg.tree(story)
        if r.tag.split("}")[-1] not in ("hdr", "ftr"):
            continue
        for instr, res in fields_with_results(r):
            u = instr.upper()
            if u.startswith("STYLEREF") and "MERGEFORMAT" in u:
                bad.append(instr)
            if u.startswith("PAGE") and "\\*" in u:
                bad.append(instr)
    return bad


def xml_empty_notes(path):
    pkg = Package(path)
    n = 0
    for tr in pkg.tree(pkg.main).iter(f"{{{W}}}tr"):
        tcs = tr.findall("w:tc", NS)
        if len(tcs) == 1 and re.fullmatch(r"\s*(Note|Notes|Source)\s*:?\s*", ops.cell_text(tcs[0])):
            n += 1
    return n


def pdf_page_labels(pdf):
    import pymupdf as fitz
    d = fitz.open(pdf)
    labels = []
    for pg in d:
        h = pg.rect.height
        words = [w for w in pg.get_text("words") if w[1] > h - 72 * 0.9]
        words.sort(key=lambda w: (w[1], w[0]))
        labels.append(words[-1][4] if words else "")
    return labels


def check_page_labels(labels, n_chapters, n_appendix):
    """Rules: cover has no number; front i, ii...; chapters c-1.. consecutive, c increments by 1;
    appendices X-1.. consecutive, X increments A, B... Landscape pages continue."""
    problems = []
    def to_roman(n):
        vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
                (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
        out = ""
        for v_, s_ in vals:
            while n >= v_:
                out += s_; n -= v_
        return out
    roman = [to_roman(i) for i in range(1, 400)]
    k = 0
    if labels and labels[0] not in ("",) and not re.fullmatch(r"[ivx]+|\d+-\d+", labels[0]):
        pass
    # cover pages: label not numeric
    while k < len(labels) and not re.fullmatch(r"[ivxlcdm]+|[0-9A-Z]+-\d+|Appendix", labels[k]):
        k += 1
    covers = k
    r = 0
    while k < len(labels) and re.fullmatch(r"[ivxlcdm]+", labels[k]):
        if labels[k] != roman[r]:
            problems.append((k + 1, labels[k], roman[r]))
        r += 1; k += 1
    prev = None
    chapters, apps = set(), set()
    for i in range(k, len(labels)):
        m = re.fullmatch(r"(?:Appendix\s*)?([0-9]+|[A-Z])-(\d+)", labels[i])
        if not m:
            problems.append((i + 1, labels[i], "c-n")); prev = None; continue
        c, n = m.group(1), int(m.group(2))
        (apps if c.isalpha() else chapters).add(c)
        if prev:
            pc, pn = prev
            same = c == pc and n == pn + 1
            nxt = n == 1 and ((c.isdigit() and pc.isdigit() and int(c) == int(pc) + 1) or
                              (c.isalpha() and pc.isalpha() and ord(c) == ord(pc) + 1) or
                              (c == "A" and pc.isdigit()))
            if not (same or nxt):
                problems.append((i + 1, labels[i], f"after {pc}-{pn}"))
        prev = (c, n)
    if len(chapters) != n_chapters:
        problems.append(("chapters", len(chapters), n_chapters))
    if len(apps) != n_appendix:
        problems.append(("appendices", len(apps), n_appendix))
    if covers != 1:
        problems.append(("cover pages", covers, 1))
    if r < 1:
        problems.append(("front matter pages", r, ">=1"))
    return {"pages": len(labels), "cover_pages": covers, "front_pages": r, "chapters_seen": len(chapters),
            "appendices_seen": len(apps), "expected_chapters": n_chapters, "expected_appendices": n_appendix,
            "problems": len(problems), "first_problems": problems[:8]}


# ---------------------------------------------------------------- runner
def run_case(fixture, names, outdir, w, sample=None, pdf=False, manifest=None):
    os.makedirs(outdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(fixture))[0]
    work = os.path.join(outdir, f"{base}__{'+'.join(names) if len(names) < 3 else 'recipe'}.docx")
    shutil.copy2(fixture, work)
    res = {"case": names, "file": work, "steps": []}
    checks = []
    for name in names:
        fn, args, allowed = CASES[name]
        t = time.perf_counter()
        try:
            r = safewrite.apply(work, fn, allowed=allowed, op_args=args)
            rep = r.report
            checks += rep.pop("checks", [])
            res["steps"].append({"op": name, "ok": True, "changed": r.changed, "seconds": round(time.perf_counter() - t, 2),
                                 "dirty": getattr(r, "dirty", []), "report": rep})
        except Refused as e:
            res["steps"].append({"op": name, "ok": False, "refused": str(e)})
            return res
    # B + C: Word opens and resaves with no update
    t = time.perf_counter()
    try:
        doc = w.open(work, readonly=False)
    except Exception as e:
        res["word_open"] = f"FAIL {e}"
        return res
    res["word_open"] = "ok"
    resaved = work.replace(".docx", "__word_resaved.docx")
    retry(doc.SaveAs2, os.path.abspath(resaved), 16)
    ours, words = fingerprint(Package(work)), fingerprint(Package(resaved))
    d = [k for k in diff(ours, words) if k.split("|", 1)[1] in STRUCT]
    res["word_resave_structural_diff"] = d[:20]
    # and against the original input: only what the ops declared may differ
    declared = set()
    for name in names:
        declared.update(CASES[name][2])
    fin = fingerprint(Package(fixture))
    d2 = [k for k in diff(fin, words) if k.split("|", 1)[1] in STRUCT and k.split("|", 1)[1] not in declared]
    res["word_resave_vs_input_diff"] = d2[:20]
    # independent table expectations from the fixture's content
    fx_tables = Package(fixture).tree(Package(fixture).main).find("w:body", NS).findall("w:tbl", NS)
    oracle = {i: oracle_roles(t) for i, t in enumerate(fx_tables, 1) if not ops.is_figure_table(t)}
    rs = Package(resaved)
    hdr_bad = 0
    for i, tbl_ in enumerate(rs.tree(rs.main).find("w:body", NS).findall("w:tbl", NS), 1):
        if i in oracle and any(c["expr"].startswith("table_check") for c in checks):
            trs = tbl_.findall("w:tr", NS)
            want = [r in ("caption", "header") for r in [x for x in oracle[i]][:len(trs)]]
            have = [tr.find("w:trPr/w:tblHeader", NS) is not None for tr in trs]
            hdr_bad += want != have
    res["word_saved_repeat_header_mismatch"] = hdr_bad
    if any(c["expr"].startswith("table_check") for c in checks):
        rp = Package(resaved)
        widths = {"pct5000": 0, "other": 0}
        for tbl in rp.tree(rp.main).find("w:body", NS).findall("w:tbl", NS):
            if ops.is_figure_table(tbl):
                continue
            tw = tbl.find("w:tblPr/w:tblW", NS)
            ok = tw is not None and tw.get(f"{{{W}}}type") == "pct" and tw.get(f"{{{W}}}w") in ("5000", "100%")
            widths["pct5000" if ok else "other"] += 1
        res["word_saved_table_widths"] = widths
    # D: settings read-back
    env = {"doc": doc, "style_in_use": style_in_use, "heading_list_strings": heading_list_strings, "table_check": table_check}
    out = []
    computed = []
    tables_seen = 0
    for c in checks:
        e = c["expr"]
        if c["expect"] == "computed" or e.startswith(("hf_field_codes", "caption_numbers", "error_ref_targets", "empty_note_rows")):
            computed.append(c); continue
        if e.startswith("table_check"):
            tables_seen += 1
            if sample and (tables_seen - 1) % sample:
                continue
            ti = int(re.match(r"table_check\(doc, (\d+)", e).group(1))
            eng_roles = eval(e[e.index("["):e.rindex("]") + 1])
            if oracle.get(ti) != eng_roles:
                out.append({"expr": f"roles of table {ti}", "expect": oracle.get(ti), "got": eng_roles, "ok": False})
            e = f"table_check(doc, {ti}, {oracle.get(ti)!r})"
        try:
            got = eval(e, env)
            ok = (got == c["expect"]) if not isinstance(c["expect"], float) else abs(float(got) - c["expect"]) < 0.01
        except Exception as ex:
            got, ok = f"ERROR {ex}", False
        out.append({"expr": e, "expect": c["expect"], "got": got if not hasattr(got, "_oleobj_") else str(got), "ok": ok})
    res["readback"] = {"checked": len(out), "failed": [x for x in out if not x["ok"]][:20]}
    # E: computed checks after a full Word update
    res["word_update_s"] = round(w.update_all(doc), 2)
    if any(c["expr"].startswith("heading_list_strings") for c in computed):
        res["heading_numbers"] = heading_list_strings(doc)
    chapters = word_chapter_numbers(doc) if any(c["expr"].startswith("caption_numbers") for c in computed) else None
    updated = work.replace(".docx", "__word_updated.docx")
    retry(doc.SaveAs2, os.path.abspath(updated), 16)
    if pdf:
        pdfp = updated.replace(".docx", ".pdf")
        t = time.perf_counter()
        retry(doc.ExportAsFixedFormat, os.path.abspath(pdfp), 17)
        res["pdf_s"] = round(time.perf_counter() - t, 2)
    w.close(doc)
    for c in computed:
        e = c["expr"]
        if e.startswith("caption_numbers"):
            res["caption_numbers"] = xml_caption_numbers(updated, chapters)
            up = Package(updated)
            toc_entries = sum(1 for p in up.tree(up.main).iter(f"{{{W}}}p") if ops.para_style(p) == "TableofFigures")
            body_tables = sum(1 for p, label, app in ops.caption_paragraphs(up.tree(up.main)) if label == "Table" and not app)
            res["list_of_tables"] = {"entries": toc_entries, "body_table_captions": body_tables, "match": toc_entries == body_tables}
            res["caption_numbers"]["chapter_numbers_seen"] = [s for n, s in chapters][:6]
        elif e.startswith("error_ref_targets"):
            got = xml_error_refs(updated)
            res["xref"] = {"engine_missing": c["expect"], "word_error_targets": got, "match": got == sorted(c["expect"])}
        elif e.startswith("hf_field_codes"):
            res["hf_bad_codes"] = xml_hf_codes(updated)
        elif e.startswith("empty_note_rows"):
            res["empty_note_rows_left"] = xml_empty_notes(updated)
    if pdf and manifest:
        m = json.load(open(manifest))
        n_ch = len({c["chapter"] for c in m["captions"] if not c["appendix"]})
        res["page_labels"] = check_page_labels(pdf_page_labels(pdfp), n_ch, 2)
    res["word_seconds"] = round(time.perf_counter() - t, 2)
    return res


def main():
    fixture, which, out = sys.argv[1:4]
    sample = int(sys.argv[sys.argv.index("--sample") + 1]) if "--sample" in sys.argv else None
    pdf = "--pdf" in sys.argv
    manifest = sys.argv[sys.argv.index("--manifest") + 1] if "--manifest" in sys.argv else None
    outdir = os.path.join(os.path.dirname(os.path.abspath(out)), "cases")
    results = []
    with WordSession() as w:
        if which == "all":
            for name in CASES:
                try:
                    results.append(run_case(fixture, [name], outdir, w, sample))
                except Exception:
                    results.append({"case": [name], "crash": traceback.format_exc()})
                print(json.dumps(results[-1], default=str)[:1500], flush=True)
        else:
            names = RECIPE if which == "recipe" else which.split("+")
            try:
                results.append(run_case(fixture, names, outdir, w, sample, pdf, manifest))
            except Exception:
                results.append({"case": names, "crash": traceback.format_exc()})
            print(json.dumps(results[-1], default=str)[:4000], flush=True)
    json.dump(results, open(out, "w"), indent=1, default=str)


if __name__ == "__main__":
    main()
