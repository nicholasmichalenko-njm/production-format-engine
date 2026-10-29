"""Gate M5: the engine's style merge must give the same Word settings as Word's CopyStylesFromTemplate.

1. Build a template that disagrees with the document on purpose: new Heading 1 to 3, Caption and Normal
   values, the team heading numbering (Appendix at Heading 7), and a new custom style. Word saves it.
2. Copy A: engine merge through safewrite.  Copy B: Word's CopyStylesFromTemplate, saved by Word.
3. Word opens both and reports, for every style in the template, the font, size, bold, italic, color,
   spacing, indents, alignment, base and next style, and its list level; and the number of every heading.
4. Pass when A and B agree on every property and every heading number.
Usage: python m5_merge.py <document.docx> <workdir> <results.json>
"""
import json, os, re, shutil, sys, traceback
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from engine import safewrite
from engine.pkg import Package, Refused, NS, W
from engine.ops import word_ops as ops
from engine.ops.merge import op_merge_styles
from wordsession import WordSession, retry

PROPS = ["Font.Name", "Font.Size", "Font.Bold", "Font.Italic", "Font.Color", "ParagraphFormat.SpaceBefore",
         "ParagraphFormat.SpaceAfter", "ParagraphFormat.LeftIndent", "ParagraphFormat.FirstLineIndent",
         "ParagraphFormat.Alignment", "ParagraphFormat.KeepWithNext", "ParagraphFormat.LineSpacing"]


def style_report(doc, names):
    out = {}
    for n in names:
        try:
            s = doc.Styles(n)
        except Exception:
            out[n] = "missing"; continue
        r = {}
        for p in PROPS:
            try:
                v = s
                for part in p.split("."):
                    v = getattr(v, part)
                r[p] = round(v, 2) if isinstance(v, float) else v
            except Exception:
                r[p] = None
        for p, f in (("BaseStyle", lambda: s.BaseStyle.NameLocal), ("NextParagraphStyle", lambda: s.NextParagraphStyle.NameLocal),
                     ("ListLevel", lambda: s.ListLevelNumber), ("ListFormat", lambda: s.ListTemplate.ListLevels(s.ListLevelNumber).NumberFormat if s.ListTemplate else None),
                     ("ListTextPos", lambda: s.ListTemplate.ListLevels(s.ListLevelNumber).TextPosition if s.ListTemplate else None)):
            try:
                r[p] = f()
            except Exception:
                r[p] = None
        out[n] = r
    return out


def heading_numbers(doc):
    rows = []
    for p in doc.ListParagraphs:
        n = p.Style.NameLocal
        if re.fullmatch(r"Heading [1-9]", n):
            rows.append((p.Range.Start, n, p.Range.ListFormat.ListString))
    rows.sort()
    return [(n, s) for _, n, s in rows]


def main():
    src, work, out = sys.argv[1:4]
    os.makedirs(work, exist_ok=True)
    res = {}
    # 1. template that disagrees with the document
    tpl_raw = os.path.join(work, "m5_template_raw.docx")
    shutil.copy2(src, tpl_raw)
    safewrite.apply(tpl_raw, ops.op_styles, op_args={"spec": {
        "Normal": {"font": "Georgia", "size": 10.5, "after": 4},
        "Heading 1": {"font": "Georgia", "size": 18, "bold": True, "color": "1F3864", "before": 18, "after": 6, "keep_next": True},
        "Heading 2": {"font": "Georgia", "size": 14, "bold": True, "color": "2E74B5", "before": 12, "after": 4},
        "Heading 3": {"font": "Georgia", "size": 12, "italic": True, "color": "2E74B5"},
        "Caption": {"font": "Arial Narrow", "size": 9, "bold": True, "color": "000000", "before": 2, "after": 6},
        "PE Body Text": {"font": "Georgia", "size": 11, "before": 0, "after": 8, "based_on": "Normal"}}})
    safewrite.apply(tpl_raw, ops.op_heading_numbering)
    template = os.path.join(work, "m5_template.dotx")
    names = []
    w = WordSession()
    try:
        d = w.open(tpl_raw)
        retry(d.SaveAs2, os.path.abspath(template), 14)
        w.close(d)
        tp = Package(template)
        names = [s.find("w:name", NS).get(f"{{{W}}}val") for s in tp.tree(tp.word_part("/styles")).findall("w:style", NS)
                 if s.get(f"{{{W}}}type") in ("paragraph", "character")]
        res["template_styles"] = len(names)
        # 2A engine
        a = os.path.join(work, "m5_A_engine.docx"); shutil.copy2(src, a)
        try:
            r = safewrite.apply(a, op_merge_styles, op_args={"template_path": template})
            res["engine"] = {k: v for k, v in r.report.items() if k != "checks"}
            res["engine_seconds"] = round(r.seconds, 2)
        except Refused as e:
            res["engine"] = {"refused": str(e)[:400]}
        # 2B Word
        b = os.path.join(work, "m5_B_word.docx"); shutil.copy2(src, b)
        d = w.open(b, window=True)
        d.CopyStylesFromTemplate(os.path.abspath(template))
        d.Save()
        w.close(d)
        # 3 compare in Word
        reps = {}
        for label, p in (("A", a), ("B", b)):
            d = w.open(p, readonly=True, window=True)
            reps[label] = {"styles": style_report(d, names), "headings": heading_numbers(d)}
            w.close(d)
        diffs = []
        for n in names:
            sa, sb = reps["A"]["styles"].get(n), reps["B"]["styles"].get(n)
            if sa != sb:
                if isinstance(sa, dict) and isinstance(sb, dict):
                    diffs.append({"style": n, "props": {k: (sa[k], sb[k]) for k in sa if sa[k] != sb.get(k)}})
                else:
                    diffs.append({"style": n, "A": sa if not isinstance(sa, dict) else "present", "B": sb if not isinstance(sb, dict) else "present"})
        res["styles_compared"] = len(names)
        res["style_differences"] = diffs
        ha, hb = reps["A"]["headings"], reps["B"]["headings"]
        res["headings"] = {"A": len(ha), "B": len(hb), "same": ha == hb, "first_A": ha[:4], "first_B": hb[:4],
                           "first_mismatch": next(((i, x, y) for i, (x, y) in enumerate(zip(ha, hb)) if x != y), None)}
        # the merge must also have reached the expected template values
        res["engine_heading1"] = reps["A"]["styles"].get("Heading 1")
    except Exception:
        res["crash"] = traceback.format_exc()
    finally:
        w.quit()
    res["pass"] = (not res.get("crash") and "refused" not in res.get("engine", {}) and not res.get("style_differences")
                   and res.get("headings", {}).get("same"))
    json.dump(res, open(out, "w"), indent=1, default=str)
    print(json.dumps({k: v for k, v in res.items() if k not in ("engine_heading1",)}, default=str)[:3000])


if __name__ == "__main__":
    main()
