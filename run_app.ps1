$ErrorActionPreference = "Stop"

$python = "python"
$bundled = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (Test-Path -LiteralPath $bundled) {
    $python = $bundled
}

& $python -B (Join-Path $PSScriptRoot "app.py")
