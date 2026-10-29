"""The only path by which the engine changes a file on disk.

apply(path, op) does, in order:
 1. Sweep stale engine temp files; refuse if the file is open anywhere we can see (fail closed).
 2. Load the package with every refusal rule of pkg.Package; refuse signed, protected or non-Word files.
 3. Fingerprint content, run the operation, fingerprint again.
 4. Unit guard: every paragraph, property set, style, list element or part that changed must be under an
    element the operation marked (guard.py).
 5. Content facets: only the facets the operation declares may change. If it declares 'fields', the
    changed field instructions must equal its declared field_changes exactly.
 6. Schema order check on every changed part.
 7. Write a temp file in the same folder, load it again with all refusal rules, fingerprint it again,
    and prove every untouched part is byte identical.
 8. Refuse if the original changed on disk meanwhile; re-check open state.
 9. Copy the original to _PE_Snapshots, verify the copy, atomically replace.
An exception or a process kill at any step leaves the original untouched.
"""
import collections
import datetime
import glob
import hashlib
import os
import shutil
import sys
import time
import urllib.parse

from .pkg import Package, Refused, W, NS, serialize, parse_xml
from .fingerprint import fingerprint, diff
from .ordered import check_order

SNAP_DIR = "_PE_Snapshots"
TMP_PREFIX = "~pe_tmp_"


def _crash_point(name):
    """Test hook: PE_HARD_CRASH=<name> ends the process right here, with no cleanup (a real kill)."""
    if os.environ.get("PE_HARD_CRASH") == name:
        os._exit(77)


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def owner_files(path):
    d, f = os.path.split(os.path.abspath(path))
    cands = {"~$" + f, "~$" + f[1:], "~$" + f[2:]}
    return [p for p in glob.glob(os.path.join(glob.escape(d), "~$*")) if os.path.basename(p) in cands]


def exclusive_probe(path):
    """True if no other process has the file open at all (Windows share mode 0)."""
    if sys.platform != "win32":
        return True
    import pywintypes
    import win32file
    for attempt in range(2):
        try:
            h = win32file.CreateFile(path, win32file.GENERIC_READ | win32file.GENERIC_WRITE, 0, None,
                                     win32file.OPEN_EXISTING, 0, None)
            h.Close()
            return True
        except pywintypes.error as e:
            if e.winerror in (32, 33):
                if attempt == 0:
                    time.sleep(1.0)
                    continue
                return False
            raise
    return False


def open_in_running_word(path):
    """Fail closed. True if any registered document is this file by local path, or is a web address
    (OneDrive, SharePoint) with the same file name. An error while looking counts as open."""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        pythoncom.CoInitialize()  # this thread may not have COM set up (or a caller released it)
    except Exception:
        return True
    try:
        target = os.path.normcase(os.path.abspath(path))
        base = os.path.basename(path).lower()
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        for mk in rot:
            try:
                name = mk.GetDisplayName(ctx, None)
            except Exception:
                continue
            if os.path.normcase(name) == target:
                return True
            if name.lower().startswith(("http://", "https://")):
                tail = urllib.parse.unquote(name.split("?")[0].rstrip("/").rsplit("/", 1)[-1]).lower()
                if tail == base:
                    return True
    except Exception:
        return True
    # no CoUninitialize here: releasing COM would disconnect any Word session the caller holds
    return False


def sweep_stale_temps(folder):
    """Remove engine temp files left by a killed process. Never touches anything else."""
    removed = 0
    for p in glob.glob(os.path.join(glob.escape(folder), TMP_PREFIX + "*")):
        try:
            if time.time() - os.path.getmtime(p) < 600:
                continue  # may belong to an engine run in progress
            os.remove(p)
            removed += 1
        except OSError:
            pass  # still in use by a live engine process
    return removed


def open_checks(path):
    if not os.path.isfile(path):
        raise Refused("The file does not exist.")
    if not os.access(path, os.W_OK):
        raise Refused("The file is read-only.")
    if owner_files(path):
        raise Refused("Word has this file open, or left an owner file (~$) after a crash. Close it, or delete the ~$ file if Word is closed.")
    if open_in_running_word(path):
        raise Refused("Word has this file open (here or as a OneDrive/SharePoint document with the same name). Close it, or use the live mode.")
    if not exclusive_probe(path):
        raise Refused("Another program has this file open. Close it and try again.")


def check_package_flags(pkg, allow_protected=False):
    if not pkg.is_word:
        raise Refused("This is not a Word document or template.")
    if pkg.signed:
        raise Refused("The file has a digital signature. A change would break it.")
    st = pkg.word_part("/settings")
    if st and not allow_protected:
        s = pkg.tree(st)
        dp = s.find("w:documentProtection", NS)
        if dp is not None and dp.get(f"{{{W}}}enforcement") in ("1", "true", "on"):
            raise Refused("The file has Restrict Editing turned on. Turn it off first.")
        if s.find("w:writeProtection", NS) is not None:
            raise Refused("The file is write-protected (password to modify).")


def _canon(tree):
    from lxml import etree
    return etree.tostring(tree, method="c14n")


def check_field_changes(before, after, declared):
    """The field instructions that changed must be exactly the declared (story, old, new) list."""
    problems = []
    stories = {k.split("|")[0] for k in list(before) + list(after) if k.endswith("|fields")}
    dec = collections.defaultdict(lambda: (collections.Counter(), collections.Counter()))
    for story, old, new in declared:
        if old not in (None, "\0none"):
            dec[story][0][old] += 1
        if new not in (None, "\0none"):
            dec[story][1][new] += 1
    for s in stories:
        b = collections.Counter(before.get(s + "|fields", ()))
        a = collections.Counter(after.get(s + "|fields", ()))
        gone, came = b - a, a - b
        want_gone, want_came = dec[s][0] - dec[s][1], dec[s][1] - dec[s][0]
        if gone != want_gone or came != want_came:
            problems.append(f"{s}: fields changed {dict(gone)} -> {dict(came)}, declared {dict(want_gone)} -> {dict(want_came)}")
    return problems


class Result:
    def __init__(self, **kw):
        self.__dict__.update(kw)

    def as_dict(self):
        return dict(self.__dict__)


def apply(path, op, *, allowed=(), op_args=None, dry_run=False, allow_protected=False, _fail_at=None):
    """op(pkg, **op_args) -> report dict. allowed: content facet globs the op may change."""
    path = os.path.abspath(path)
    t0 = time.perf_counter()
    d, f = os.path.split(path)
    swept = sweep_stale_temps(d)
    open_checks(path)
    st0 = os.stat(path)
    h0 = sha(path)
    pkg = Package(path)
    check_package_flags(pkg, allow_protected)
    before = fingerprint(pkg)
    try:
        report = op(pkg, **(op_args or {})) or {}
    except Refused:
        raise
    except Exception as e:
        raise Refused(f"The operation stopped on this file ({type(e).__name__}: {e}). Nothing was written.")
    # a part that serializes to what it was read from is not a change
    for n in list(pkg.dirty):
        if n in pkg.raw_at_load and n in pkg.trees and _canon(pkg.trees[n]) == _canon(parse_xml(pkg.raw_at_load[n], n)):
            pkg.dirty.discard(n)
    violations = pkg.snap.violations(pkg)
    if violations:
        raise Refused("The change would touch elements it did not declare: " + "; ".join(violations[:6]))
    after = fingerprint(pkg)
    bad = diff(before, after, allowed)
    if bad:
        raise Refused("The change would alter content it should not touch: " + ", ".join(bad[:12]))
    if any(a == "fields" or a.endswith("|fields") for a in allowed):
        fp = check_field_changes(before, after, report.get("field_changes", []))
        if fp:
            raise Refused("Field changes do not match what the operation declared: " + "; ".join(fp[:3]))
    for part in pkg.dirty:
        if part.lower().endswith(".xml") and part in pkg.trees:
            old = set(check_order(parse_xml(pkg.raw_at_load[part], part))) if part in pkg.raw_at_load else set()
            new_problems = set(check_order(pkg.tree(part))) - old
            if new_problems:
                raise Refused(f"Element order problem in {part}: {sorted(new_problems)[:3]}")
    if not pkg.dirty and set(pkg.raw) == set(pkg.raw_at_load):
        return Result(path=path, changed=False, report=report, swept=swept, seconds=time.perf_counter() - t0)
    data = pkg.to_bytes()
    tmp = os.path.join(d, f"{TMP_PREFIX}{os.getpid()}_{f}")
    try:
        with open(tmp, "wb") as out:
            out.write(data)
            out.flush()
            os.fsync(out.fileno())
        _crash_point("after_temp")
        if _fail_at == "after_temp":
            raise RuntimeError("simulated crash after temp write")
        check = Package(tmp)
        disk = fingerprint(check)
        bad2 = diff(after, disk, ())
        if bad2:
            raise Refused("The written file differs from the checked result: " + ", ".join(bad2[:12]))
        for n, raw in pkg.raw_at_load.items():
            if n not in pkg.dirty and n in check.raw and raw != check.raw[n]:
                raise Refused(f"Untouched part {n} changed on write.")
        if dry_run:
            os.remove(tmp)
            return Result(path=path, changed=True, dry_run=True, dirty=sorted(pkg.dirty), report=report,
                          seconds=time.perf_counter() - t0)
        st1 = os.stat(path)
        if (st1.st_mtime_ns, st1.st_size) != (st0.st_mtime_ns, st0.st_size) or sha(path) != h0:
            raise Refused("The file changed on disk while the engine worked. Nothing was written.")
        open_checks(path)
        sd = os.path.join(d, SNAP_DIR)
        os.makedirs(sd, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        stem, ext = os.path.splitext(f)
        snap = os.path.join(sd, f"{stem}__{stamp}{ext}")
        shutil.copy2(path, snap)
        if sha(snap) != h0:
            raise Refused("The snapshot copy failed its check. Nothing was written.")
        _crash_point("before_replace")
        if _fail_at == "before_replace":
            raise RuntimeError("simulated crash before replace")
        os.replace(tmp, path)
        tmp = None
        _crash_point("after_replace")
        return Result(path=path, changed=True, snapshot=snap, dirty=sorted(pkg.dirty), report=report,
                      macros_kept=pkg.has_macros, swept=swept, sha_before=h0, sha_after=sha(path),
                      seconds=time.perf_counter() - t0)
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)


def restore(path, snapshot):
    """Put a snapshot back. The current file becomes a snapshot first. The snapshot must load cleanly."""
    path = os.path.abspath(path)
    open_checks(path)
    d, f = os.path.split(path)
    stem, ext = os.path.splitext(f)
    sd = os.path.join(d, SNAP_DIR)
    if os.path.dirname(os.path.abspath(snapshot)) != sd or not os.path.basename(snapshot).startswith(stem + "__"):
        raise Refused("That snapshot does not belong to this file.")
    Package(snapshot)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    shutil.copy2(path, os.path.join(sd, f"{stem}__{stamp}_before_restore{ext}"))
    tmp = os.path.join(d, f"{TMP_PREFIX}restore_{os.getpid()}_{f}")
    shutil.copy2(snapshot, tmp)
    if sha(tmp) != sha(snapshot):
        os.remove(tmp)
        raise Refused("The restore copy failed its check.")
    os.replace(tmp, path)
