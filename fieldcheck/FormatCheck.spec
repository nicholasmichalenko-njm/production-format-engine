# PyInstaller spec for FormatCheck (read-only field check). Build: pyinstaller --noconfirm fieldcheck\FormatCheck.spec
# Always build from this spec; CLI flags would regenerate it and drop the excludes.
import os
ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
a = Analysis(
    [os.path.join(SPECPATH, "fieldcheck.py")],
    pathex=[ROOT],
    hiddenimports=["engine", "engine.pkg", "engine.guard", "engine.fingerprint", "engine.ordered", "engine.safewrite",
                   "engine.ops", "engine.ops.word_ops", "engine.ops.glossary", "engine.ops.merge",
                   "win32timezone", "win32com.client", "pythoncom", "pywintypes", "win32file"],
    excludes=["pymupdf", "fitz", "openpyxl", "pptx", "PIL", "numpy", "setuptools", "pip", "unittest", "pydoc_data",
              "ssl", "_ssl", "http", "xmlrpc", "ftplib", "smtplib", "imaplib", "poplib", "nntplib", "urllib.request",
              "email.mime", "asyncio", "sqlite3"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="FormatCheck", console=True, upx=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="FormatCheck")
