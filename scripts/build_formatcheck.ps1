# Build FormatCheck, test the built kit on fixtures, and zip it.
# Usage: powershell -File scripts\build_formatcheck.ps1
$ErrorActionPreference = "Continue"  # PyInstaller logs to stderr; failures are checked explicitly below
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$ver = (Select-String -Path fieldcheck\fieldcheck.py -Pattern '^VERSION = "(.+)"').Matches[0].Groups[1].Value
New-Item -ItemType Directory -Force results | Out-Null
Remove-Item -Recurse -Force dist\FormatCheck -ErrorAction SilentlyContinue
& .\venv\Scripts\pyinstaller.exe --noconfirm --distpath dist --workpath build fieldcheck\FormatCheck.spec *> results\build_formatcheck.log
if (-not (Test-Path dist\FormatCheck\FormatCheck.exe)) { throw "build failed, see results\build_formatcheck.log" }
Copy-Item fieldcheck\README.txt dist\FormatCheck\README.txt -Force

# kit test: Word-written volumes, a file written by another tool, a protected file and a damaged file
$T = "results\kit_test_$ver"
Remove-Item -Recurse -Force $T -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $T | Out-Null
Copy-Item fixtures\small.docx "$T\Volume I.docx"
Copy-Item fixtures\large.docx "$T\Volume II.docx"
Copy-Item fixtures\raw_small.docx "$T\Volume III not from Word.docx"
# embedded and linked Excel objects, link fields to missing and network files (built by harness\ole_test.py)
if (Test-Path results\ole_small\ole_word.docx) { Copy-Item results\ole_small\ole_word.docx "$T\Volume IV linked objects.docx" }
if (Test-Path results\v2\puncture\P09_password.docx) { Copy-Item results\v2\puncture\P09_password.docx "$T\Protected.docx" }
[IO.File]::WriteAllBytes("$Root\$T\Damaged.docx", [IO.File]::ReadAllBytes("$Root\fixtures\small.docx")[0..29999])
$before = Get-ChildItem $T | ForEach-Object { "$($_.Name)|$((Get-FileHash $_.FullName).Hash)" }
$wordBefore = @(Get-Process WINWORD -ErrorAction SilentlyContinue).Id
& dist\FormatCheck\FormatCheck.exe $T --timeout 900 *> results\kit_test_$ver.log
$after = Get-ChildItem $T | ForEach-Object { "$($_.Name)|$((Get-FileHash $_.FullName).Hash)" }
$left = @(@(Get-Process WINWORD -ErrorAction SilentlyContinue).Id | Where-Object { $wordBefore -notcontains $_ }).Count
$run = Get-ChildItem dist\FormatCheck\FormatCheck_runs -Directory | Sort-Object Name | Select-Object -Last 1
$j = Get-Content (Join-Path $run.FullName FormatCheck_report.json) -Raw | ConvertFrom-Json
$names = Select-String -Path (Join-Path $run.FullName "FormatCheck_report.*") -Pattern "Volume|Protected|Damaged" -Quiet
$summary = [ordered]@{
  version = $ver
  originals_identical = -not (Compare-Object $before $after)
  word_left_running = $left
  file_names_in_report = [bool]$names
  word_link_option_check = $j.word_link_option_check
  documents = @($j.documents | ForEach-Object { [ordered]@{ id = $_.id; load = $_.load
      recipe_ok = if ($_.recipe) { -not ($_.recipe | Where-Object { -not $_.ok }) } else { $null }
      drift_before = @($_.word.before.resave_drift | Select-Object -Unique) -join ","
      drift_after = @($_.word.after.resave_drift | Select-Object -Unique) -join ","
      pages = "$($_.word.before.pages)/$($_.word.after.pages)"
      word_stopped = "$($_.word.blocked_by_dialog)$($_.word.timed_out)"
      linked_fields = $_.word.before.linked_fields_locked
      option_restored = $_.word.links_option_restored } })
}
$summary | ConvertTo-Json -Depth 4 | Set-Content results\kit_summary_$ver.json
Remove-Item -Recurse -Force dist\FormatCheck\FormatCheck_runs
$zip = "FormatCheck_$ver.zip"
Remove-Item $zip -ErrorAction SilentlyContinue
Compress-Archive -Path dist\FormatCheck -DestinationPath $zip
Get-Content results\kit_summary_$ver.json
"zip $zip $((Get-FileHash $zip).Hash)"
