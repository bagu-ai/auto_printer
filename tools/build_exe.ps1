param([string]$Python = '')
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build_release.ps1') -Python $Python
exit $LASTEXITCODE
