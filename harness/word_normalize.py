"""Open a raw fixture in hidden Word, update all fields and TOCs, save as a Word-written .docx, export PDF.
Also gate M2: timings on the large fixture.
Usage: python word_normalize.py <raw.docx> <out.docx> <out.pdf> <result.json>
"""
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(__file__))
from wordsession import WordSession, retry

raw, out, pdf, res = sys.argv[1:5]
r = {"input": raw}
t0 = time.perf_counter()
with WordSession() as w:
    r["word_build"] = w.build
    t = time.perf_counter(); doc = w.open(raw); r["open_s"] = round(time.perf_counter() - t, 2)
    r["update_fields_and_tocs_s"] = round(w.update_all(doc), 2)
    t = time.perf_counter(); doc.Repaginate(); r["repaginate_s"] = round(time.perf_counter() - t, 2)
    t = time.perf_counter(); r["pages"] = doc.ComputeStatistics(2); r["pages_s"] = round(time.perf_counter() - t, 2)
    r["sections"] = doc.Sections.Count
    r["tables"] = doc.Tables.Count
    r["fields"] = doc.Fields.Count
    # field results that show an error, in the main story
    errs = 0
    for f in doc.Fields:
        try:
            if "Error!" in (f.Result.Text or ""):
                errs += 1
        except Exception:
            pass
    r["fields_with_error_result"] = errs
    t = time.perf_counter(); retry(doc.SaveAs2, os.path.abspath(out), 16); r["save_s"] = round(time.perf_counter() - t, 2)
    t = time.perf_counter(); retry(doc.ExportAsFixedFormat, os.path.abspath(pdf), 17); r["pdf_s"] = round(time.perf_counter() - t, 2)
    w.close(doc)
r["total_s"] = round(time.perf_counter() - t0, 2)
json.dump(r, open(res, "w"), indent=1)
print(json.dumps(r))
