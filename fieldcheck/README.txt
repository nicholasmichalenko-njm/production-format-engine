FormatCheck 0.2.0 - read-only field check for the production format engine

WHAT IT DOES
FormatCheck tests the format engine on real proposal volumes without changing them.
- It reads each Word file in the folder you pick one time, to make a local copy.
- Everything happens on the copies, in a work folder next to FormatCheck.exe:
  FormatCheck_runs\run_<date-time>
- It runs every engine operation on each copy and records whether the safety guard held.
- If Word is installed, Word opens each copy and each result in a hidden session with macros off,
  and reports pages, broken cross-references, and any change Word made when it saved.
- At the end it checks that every original is byte for byte the same as before.

HOW TO RUN
1. Unzip the FormatCheck folder to a local folder on the PC (for example Documents).
   Do not put it in a OneDrive or SharePoint folder: the copies of your volumes go next to it.
2. Double-click FormatCheck.exe and pick the folder with the volumes.
   Or from a command prompt:  FormatCheck.exe "C:\path\to\volumes" [--recursive] [--no-word] [--keep]
3. The report opens in your browser when the run ends. A 500-page volume takes about 5 minutes with Word.

WHAT TO SHARE
FormatCheck_report.html and FormatCheck_report.json hold counts only: no document text and no file
names. Documents show as Doc-1, Doc-2 and so on.
KEY_stays_on_this_PC.txt maps Doc-N to the file. Keep it on this PC.
The local copies are deleted at the end of the run (unless you use --keep).

WHAT IT NEVER DOES
- It never writes to the folder you pick.
- It never uses the network.
- It never opens your own Word window and never changes Normal.dotm.
- Each Word job has a time limit. If Word stops answering or shows a dialog, FormatCheck ends only
  the Word it started (identified by its process number), never a Word you opened.

LINKED OBJECTS AND LINK FIELDS
Volumes often hold linked Excel tables, INCLUDETEXT or INCLUDEPICTURE fields. FormatCheck never
updates them: they would reach other files or network shares, and Word replaces a linked table with
"Error! Not a valid link." when the source is not reachable. The report counts them per document.
A hidden Word stops on a security notice for such files, so FormatCheck's own Word turns off
"Update automatic links at open" for its session and sets it back before it closes. At the end of
the run a fresh Word confirms the setting is what it was, and the report says so.
