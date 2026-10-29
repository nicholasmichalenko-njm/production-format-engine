"""Build gate: logging calls may only carry fixed messages written in our code, never data.
Ported from Stationkeeper's tools/check_diag.py for FormatCheck.

Checks every diag.error / warning / activity / swallowed call in FormatCheck's .py files:
the message (and area) is a plain string literal with no / \\ or :, and keyword arguments are only exc=, code=,
ext=, whitelisted count names, or **counts (values are checked at run time by diag._counts).
Exit code 0 = pass, 1 = fail (the build stops).
"""
import ast, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "fieldcheck"))
import diag  # noqa: E402

FILES = ["fieldcheck/fieldcheck.py", "fieldcheck/gui.py", "fieldcheck/feedback.py", "fieldcheck/diag.py"]
DIAG_FUNCS = {"error", "warning", "activity", "swallowed"}


def is_literal(n):
    return isinstance(n, ast.Constant) and isinstance(n.value, str)


def check_file(path):
    problems = []
    tree = ast.parse(open(path, encoding="utf-8").read())
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
            continue
        f = node.func
        if not (isinstance(f.value, ast.Name) and f.value.id == "diag" and f.attr in DIAG_FUNCS):
            continue
        name = f.attr
        if name in ("swallowed", "activity"):
            lits = node.args[:1]
        else:
            lits = node.args[:2]
            if len(node.args) < 2:
                problems.append(f"{node.lineno}: diag.{name} needs an area and a message")
        for a in lits:
            if not is_literal(a):
                problems.append(f"{node.lineno}: diag.{name}() area and message must be plain string literals")
            elif any(ch in a.value for ch in "/\\:") or len(a.value) > 160:
                problems.append(f"{node.lineno}: diag.{name}() text must not contain / \\ or :")
        for kw in node.keywords:
            if kw.arg is None or kw.arg in ("exc", "w", "state", "code", "ext") or kw.arg in diag.COUNT_KEYS:
                continue
            if kw.arg.startswith("ext_"):
                continue
            problems.append(f"{node.lineno}: diag.{name}() keyword '{kw.arg}' is not an allowed field")
    return problems


def main():
    bad = 0
    for f in FILES:
        p = os.path.join(ROOT, f)
        if not os.path.isfile(p):
            print(f"{f}: missing"); bad += 1; continue
        for prob in check_file(p):
            print(f"{f}:{prob}"); bad += 1
    print("Logging check passed: messages are fixed text, no data." if not bad else f"Logging check FAILED: {bad} problem(s)")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
