$ErrorActionPreference = "Stop"
try {
  $utf8 = [System.Text.UTF8Encoding]::new($false)
  [Console]::InputEncoding = $utf8
  [Console]::OutputEncoding = $utf8
  $OutputEncoding = $utf8
} catch {}
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$skillDir = Resolve-Path (Join-Path $scriptDir "..")
$python = $env:TACTILE_WINDOWS_PYTHON
if (-not $python) {
  $python = "python"
}

& $python (Join-Path $skillDir "scripts\windows_interface.py") @args
exit $LASTEXITCODE
