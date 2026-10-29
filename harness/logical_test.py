"""The logical view must call a Word re-save of an unchanged document equal, and still catch real changes.

For each document: Word re-saves it (no field update). Pass when
1. logical(original) == logical(Word's re-save), and
2. each puncture made on Word's re-save is reported: header text changed in one part, footer text changed,
   body text changed, a picture removed from the body, a footnote's text changed (when the document has one).
Usage: python logical_test.py <results.json> <doc.docx> [<doc.docx> ...]
"""
import io, json, os, re, shutil, sys, tempfile, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from engine.pkg import Package
from engine.fingerprint import logical, diff
from wordsession import WordSession, retry


def rewrite(src, dst, part_pred, fn):
    """Copy a zip, applying fn to the first part matching part_pred whose bytes fn changes."""
    done = None
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zi.infolist():
            b = zi.read(i.filename)
            if done is None and part_pred(i.filename):
                nb = fn(b)
                if nb != b:
                    b, done = nb, i.filename
            zo.writestr(i, b)
    return done


def first_text(b, new):
    return re.sub(rb"(<w:t(?: [^>]*)?>)([^<]+)(</w:t>)", lambda m: m.group(1) + new + m.group(3), b, count=1)


def main():
    out, docs = sys.argv[1], sys.argv[2:]
    work = tempfile.mkdtemp(prefix="logical_")
    res = {"docs": []}
    w = WordSession()
    try:
        for i, src in enumerate(docs, 1):
            a = os.path.join(work, f"d{i}.docx"); shutil.copyfile(src, a)
            rs = os.path.join(work, f"d{i}_word.docx")
            d = w.open(a, readonly=True); retry(d.SaveAs2, os.path.abspath(rs), 16); w.close(d)
            base = logical(Package(a))
            r = {"doc": f"Doc-{i}", "keys": len(base), "resave_diff": diff(base, logical(Package(rs)))}
            pun = {
                "header_text": (lambda n: re.fullmatch(r"word/header\d+\.xml", n), lambda b: first_text(b, b"PUNCTURE")),
                "footer_text": (lambda n: re.fullmatch(r"word/footer\d+\.xml", n), lambda b: first_text(b, b"PUNCTURE")),
                "body_text": (lambda n: n == "word/document.xml", lambda b: first_text(b, b"PUNCTURE")),
                "picture_removed": (lambda n: n == "word/document.xml",
                                    lambda b: re.sub(rb"<w:drawing>.*?</w:drawing>", b"", b, count=1, flags=re.S)),
                "footnote_text": (lambda n: n == "word/footnotes.xml",
                                  lambda b: re.sub(rb'(<w:footnote (?:(?!w:type=)[^>])*w:id="[1-9]\d*"[^>]*>.*?<w:t(?: [^>]*)?>)([^<]+)',
                                                   lambda m: m.group(1) + b"PUNCTURE", b, count=1, flags=re.S)),
            }
            r["punctures"] = {}
            for name, (pred, fn) in pun.items():
                p = os.path.join(work, f"d{i}_{name}.docx")
                part = rewrite(rs, p, pred, fn)
                if part is None:
                    r["punctures"][name] = "not applicable"
                    continue
                dd = diff(base, logical(Package(p)))
                r["punctures"][name] = {"caught": bool(dd), "keys": dd[:4]}
            r["pass"] = not r["resave_diff"] and all(v == "not applicable" or v["caught"] for v in r["punctures"].values())
            res["docs"].append(r)
    finally:
        w.quit()
        shutil.rmtree(work, ignore_errors=True)
    res["pass"] = all(d["pass"] for d in res["docs"]) and bool(res["docs"])
    json.dump(res, open(out, "w"), indent=1, default=str)
    print(json.dumps(res, default=str)[:4000])


if __name__ == "__main__":
    main()
