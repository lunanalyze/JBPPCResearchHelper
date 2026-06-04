$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$MakeNsis = "C:\Program Files (x86)\NSIS\makensis.exe"

if (-not (Test-Path -LiteralPath $Python)) {
  throw "Python runtime not found: $Python"
}
if (-not (Test-Path -LiteralPath $MakeNsis)) {
  throw "NSIS makensis not found: $MakeNsis"
}

Set-Location $Root

& $Python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --noconsole `
  --name PPCResearchHelper `
  --add-data "report_prompt.md;." `
  --add-data "ppc_report_template.docx;." `
  app.py
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

& $MakeNsis installer.nsi
if ($LASTEXITCODE -ne 0) {
  throw "NSIS failed with exit code $LASTEXITCODE"
}

Write-Host "Installer: $Root\dist\PPCResearchHelperSetup.exe"
