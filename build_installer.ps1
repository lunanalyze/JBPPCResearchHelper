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

# 버전의 단일 원본은 updater.py 의 APP_VERSION 이다. 여기서 읽어 NSIS 로 넘긴다 —
# 앱이 스스로 보고하는 버전과 제어판 '프로그램 추가/제거' 의 버전이 어긋나면,
# 어느 쪽이 맞는지 확인할 방법이 없어진다.
$VersionLine = Select-String -LiteralPath (Join-Path $Root "updater.py") -Pattern '^APP_VERSION\s*=\s*"([^"]+)"'
if (-not $VersionLine) { throw "updater.py 에서 APP_VERSION 을 찾지 못했습니다" }
$AppVersion = $VersionLine.Matches[0].Groups[1].Value
Write-Host "APP_VERSION: $AppVersion"

# updater\apply.ps1 은 자동 업데이트가 exe 를 갈아끼울 때 쓰는 교체 스크립트다.
# exe 안의 Python 은 교체 대상 자신이라 쓸 수 없어, Windows 기본 PowerShell 로 돈다.
& $Python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --noconsole `
  --name PPCResearchHelper `
  --icon "PPC.ico" `
  --add-data "report_prompt.md;." `
  --add-data "ppc_report_template.docx;." `
  --add-data "PPC.ico;." `
  --add-data "updater\apply.ps1;updater" `
  app.py
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

& $MakeNsis "/DAPP_VERSION=$AppVersion" installer.nsi
if ($LASTEXITCODE -ne 0) {
  throw "NSIS failed with exit code $LASTEXITCODE"
}

Write-Host "Installer: $Root\dist\PPCResearchHelperSetup.exe"
Write-Host "다음: .\build_release.ps1 로 업데이트 팩과 latest.json 을 만드세요."
