$ErrorActionPreference = "Continue"
Set-Location (Join-Path $PSScriptRoot "..")
$env:PYTHONWARNINGS = "ignore"
$R = "results\run_$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item -ItemType Directory -Force -Path $R | Out-Null
$py = ".\venv\Scripts\python.exe"
function Step($name, $argsList) {
  $t = Get-Date
  "START $name $t" | Out-File -Append "$R\progress.txt"
  & $py @argsList *> "$R\$name.log"
  "END $name $((Get-Date) - $t)" | Out-File -Append "$R\progress.txt"
}
Step "small_all"    @("harness\verify.py", "fixtures\small.docx", "all", "$R\small_all.json")
Step "small_recipe" @("harness\verify.py", "fixtures\small.docx", "recipe", "$R\small_recipe.json", "--pdf", "--manifest", "fixtures\manifest_small.json")
Step "glossary"     @("harness\glossary_test.py", "$R\qp", "fixtures\Logo.png", "$R\glossary.json")
Step "puncture"     @("harness\puncture.py", "fixtures\small.docx", "$R\puncture", "$R\puncture.json")
Step "rogue"        @("harness\rogue.py", "fixtures\small.docx", "$R\rogue", "$R\rogue.json")
Step "m1_live"      @("harness\m1_live.py", "fixtures\small.docx", "$R\m1", "$R\m1.json", "8")
Step "m5_merge"     @("harness\m5_merge.py", "fixtures\small.docx", "$R\m5", "$R\m5.json")
Step "m6_open"      @("harness\m6_open.py", "fixtures\small.docx", "$R\m6", "$R\m6.json")
Step "xl_recipe"    @("harness\verify.py", "fixtures\xl.docx", "recipe", "$R\xl_recipe.json", "--sample", "10", "--pdf", "--manifest", "fixtures\manifest_xl.json")
"ALL DONE $(Get-Date)" | Out-File -Append "$R\progress.txt"
