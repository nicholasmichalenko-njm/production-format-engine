"""Hidden Word session for the lab. Rules:
- DispatchEx: always a new Word process, never the user's Word.
- Macros force-disabled for every file we open (AutomationSecurity = 3).
- No change to Word Options (they persist to the user's profile).
- Normal.dotm marked saved before Quit so it is never written.
- Retries COM calls that Word rejects while busy.
"""
import os
import time

import pythoncom
import pywintypes
import win32com.client

BUSY = (-2147418111, -2147417846)  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


def retry(fn, *a, tries=40, **kw):
    for i in range(tries):
        try:
            return fn(*a, **kw)
        except pywintypes.com_error as e:
            if e.hresult in BUSY and i < tries - 1:
                time.sleep(0.25)
                continue
            raise


class WordSession:
    def __init__(self):
        pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = 0
        self.app.ScreenUpdating = False
        self.app.AutomationSecurity = 3
        self.build = self.app.Build

    def open(self, path, readonly=False, window=False):
        """window=True gives the document a window (Word itself stays hidden); some editing calls need one."""
        self.app.AutomationSecurity = 3
        return retry(self.app.Documents.Open, os.path.abspath(path), False, readonly, False,
                     "", "", True, "", "", 0, 0, window, False, 0, True)
        # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
        # Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair,
        # DocumentDirection, NoEncodingDialog

    def update_all(self, doc):
        """Update every field in every story, then every TOC fully."""
        t = time.perf_counter()
        for story in doc.StoryRanges:
            s = story
            while s is not None:
                retry(s.Fields.Update)
                try:
                    s = s.NextStoryRange
                except pywintypes.com_error:
                    s = None
        for i in range(1, doc.TablesOfContents.Count + 1):
            retry(doc.TablesOfContents(i).Update)
        for i in range(1, doc.TablesOfFigures.Count + 1):
            retry(doc.TablesOfFigures(i).Update)
        return time.perf_counter() - t

    def close(self, doc, save=False):
        try:
            doc.Close(-1 if save else 0)
        except pywintypes.com_error:
            pass

    def quit(self):
        try:
            for i in range(self.app.Documents.Count, 0, -1):
                self.app.Documents(i).Close(0)
        except Exception:
            pass
        try:
            self.app.NormalTemplate.Saved = True
        except Exception:
            pass
        try:
            self.app.Quit(0)
        except Exception:
            pass
        self.app = None
        pythoncom.CoUninitialize()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.quit()
