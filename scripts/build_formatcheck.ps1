# Build FormatCheck, test the built kit on fixtures, and zip it.
# Usage: powershell -File scripts\build_formatcheck.ps1
$ErrorActionPreference = "Continue"  # PyInstaller logs to stderr; failures are checked explicitly below
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$ver = (Select-String -Path fieldcheck\fieldcheck.py -Pattern '^VERSION = "(.+)"').Matches[0].Groups[1].Value
New-Item -ItemType Directory -Force results | Out-Null
# gates: logging calls carry fixed text only; diagnostics, feedback guard and export tests pass
& .\venv\Scripts\python.exe tools\check_diag.py
if ($LASTEXITCODE -ne 0) { throw "logging check failed" }
& .\venv\Scripts\python.exe tests\test_formatcheck_diag.py *> results\test_formatcheck_diag.log
if ($LASTEXITCODE -ne 0) { throw "diagnostics tests failed, see results\test_formatcheck_diag.log" }
(Get-Content results\test_formatcheck_diag.log -Tail 1)
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
# windowed exe: wait for it explicitly
Start-Process -FilePath dist\FormatCheck\FormatCheck.exe -ArgumentList "`"$Root\$T`"", "--timeout", "900", "--no-open" -Wait
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
# what the frozen exe logged during the kit test (content-free by construction)
$ev = @(); if (Test-Path dist\FormatCheck\logs\events.jsonl) { $ev = @(Get-Content dist\FormatCheck\logs\events.jsonl | ForEach-Object { $_ | ConvertFrom-Json }) }
$act = @(); if (Test-Path dist\FormatCheck\logs\activity.jsonl) { $act = @(Get-Content dist\FormatCheck\logs\activity.jsonl | ForEach-Object { $_ | ConvertFrom-Json }) }
$summary["logged_events"] = @($ev | ForEach-Object { "$($_.level) $($_.area): $($_.msg) $($_.type) $(@($_.where)[-1])" })
$summary["activity_doc_done"] = @($act | Where-Object { $_.action -eq "doc_done" }).Count
# window smoke: the exe opens its window, logs app_started, closes on request
$p = Start-Process -FilePath dist\FormatCheck\FormatCheck.exe -PassThru
Start-Sleep 6
$p.Refresh()
$summary["window_title"] = $p.MainWindowTitle
$summary["window_closed"] = $p.CloseMainWindow()
$p.WaitForExit(15000) | Out-Null
$summary["window_exited"] = $p.HasExited
$summary["window_logged_start_close"] = [bool](Select-String dist\FormatCheck\logs\activity.jsonl -Pattern '"app_started"' -Quiet) -and [bool](Select-String dist\FormatCheck\logs\activity.jsonl -Pattern '"app_closed"' -Quiet)
$summary | ConvertTo-Json -Depth 4 | Set-Content results\kit_summary_$ver.json
Remove-Item -Recurse -Force dist\FormatCheck\FormatCheck_runs, dist\FormatCheck\logs, dist\FormatCheck\feedback -ErrorAction SilentlyContinue
$zip = "FormatCheck_$ver.zip"
Remove-Item $zip -ErrorAction SilentlyContinue
Compress-Archive -Path dist\FormatCheck -DestinationPath $zip
Get-Content results\kit_summary_$ver.json
"zip $zip $((Get-FileHash $zip).Hash)"
