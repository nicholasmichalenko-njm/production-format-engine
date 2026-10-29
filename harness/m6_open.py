"""Gate M6 (local part): the engine must never write a file that something else has open.
Cases:
  A  another program holds the file open with no sharing
  B  another program reads it and shares read and write (a viewer, an indexer)
  C  Word has it open read-only (Word writes no owner file then)
  D  Word has it open normally (owner file present)
  E  a stale owner file from a crash, no program has it open
  F  closed, no owner file (control: engine must write)
Co-authoring through OneDrive or SharePoint AutoSave cannot be made on this PC; it stays open for the work PC.
Usage: python m6_open.py <small.docx> <workdir> <results.json>
"""
import json
import os
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE)); sys.path.insert(0, HERE)
from engine import safewrite
from engine.pkg import Refused
from engine.ops import word_ops as ops
from wordsession import WordSession

HOLD = r'''
import sys, time, win32file
share = int(sys.argv[2])
h = win32file.CreateFile(sys.argv[1], win32file.GENERIC_READ, share, None, win32file.OPEN_EXISTING, 0, None)
print("held", flush=True)
time.sleep(float(sys.argv[3]))
h.Close()
'''


def try_write(path):
    try:
        r = safewrite.apply(path, ops.op_hf_clean, allowed=("fields",))
        return {"engine": "applied", "changed": r.changed}
    except Refused as e:
        return {"engine": "refused", "message": str(e)}


def main():
    src, work, out = sys.argv[1:4]
    os.makedirs(work, exist_ok=True)
    holder = os.path.join(work, "hold.py")
    open(holder, "w").write(HOLD)
    res = []

    def fresh(name):
        p = os.path.join(work, name + ".docx")
        if os.path.exists(p):
            os.remove(p)
        shutil.copy2(src, p)
        return p

    for name, share, expect in (("A_exclusive_holder", 0, "refused"), ("B_shared_reader", 3, "refused")):
        p = fresh(name)
        proc = subprocess.Popen([sys.executable, holder, p, str(share), "8"], stdout=subprocess.PIPE, text=True)
        proc.stdout.readline()
        r = try_write(p)
        proc.wait()
        res.append({"case": name, "expect": expect, **r, "pass": r["engine"] == expect})

    w = WordSession()
    try:
        p = fresh("C_word_readonly")
        d = w.open(p, readonly=True)
        r = try_write(p)
        r["owner_file_present"] = bool(safewrite.owner_files(p))
        w.close(d)
        res.append({"case": "C_word_readonly", "expect": "refused", **r, "pass": r["engine"] == "refused"})

        p = fresh("D_word_open")
        d = w.open(p, readonly=False)
        r = try_write(p)
        r["owner_file_present"] = bool(safewrite.owner_files(p))
        w.close(d)
        res.append({"case": "D_word_open", "expect": "refused", **r, "pass": r["engine"] == "refused"})
    finally:
        w.quit()

    p = fresh("E_stale_owner_file")
    stale = os.path.join(work, "~$" + os.path.basename(p)[2:])
    open(stale, "wb").write(b"\x0bcrashed user" + b"\0" * 40)
    r = try_write(p)
    res.append({"case": "E_stale_owner_file", "expect": "refused (asks the user)", **r, "pass": r["engine"] == "refused"})
    os.remove(stale)

    p = fresh("F_closed_control")
    r = try_write(p)
    res.append({"case": "F_closed_control", "expect": "applied", **r, "pass": r["engine"] == "applied"})

    json.dump(res, open(out, "w"), indent=1)
    for x in res:
        print(json.dumps(x))


if __name__ == "__main__":
    main()
