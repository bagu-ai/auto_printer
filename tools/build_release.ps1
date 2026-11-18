param([string]$Python = '')
$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
Set-Location -LiteralPath $projectRoot
if (-not $Python) { $Python = Join-Path $projectRoot '.venv\Scripts\python.exe' }
if (-not (Test-Path -LiteralPath $Python)) {
    throw "Missing build environment: $Python. Create .venv and install requirements-build.txt, or pass -Python with an existing environment."
}
$distRoot = Join-Path $projectRoot 'dist\v1.3.3'
$buildRun = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$workRoot = Join-Path $projectRoot ('build\' + $buildRun)
$env:TEMP = Join-Path $projectRoot '.tmp'
$env:TMP = $env:TEMP
$env:PYINSTALLER_CONFIG_DIR = Join-Path $projectRoot '.cache\pyinstaller'
$env:PYTHONDONTWRITEBYTECODE = '1'
New-Item -ItemType Directory -Force -Path $distRoot,$workRoot,$env:TEMP | Out-Null
$exeName = '标签自动打印工具_防卡住版'
# Folder distribution avoids C-drive extraction and repeated spawn extraction.
$buildArgs = @('-m','PyInstaller','--noconfirm','--clean','--onedir','--windowed',
    '--name',$exeName,'--distpath',$distRoot,'--workpath',$workRoot,'--specpath',$workRoot,
    '--hidden-import','label_printer.core.bartender_backend',
    '--hidden-import','label_printer.core.diagnostics',
    '--exclude-module','tkinter','--exclude-module','matplotlib','--exclude-module','scipy',
    '--exclude-module','PySide6.QtWebEngineCore','--exclude-module','PySide6.QtWebEngineWidgets',
    'run_app.py')
& $Python @buildArgs
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed: $LASTEXITCODE" }
$packageRoot = Join-Path $distRoot $exeName
$templateRoot = Join-Path $packageRoot '模板'
New-Item -ItemType Directory -Force -Path $templateRoot | Out-Null
Get-ChildItem -LiteralPath (Join-Path $projectRoot '模板') -Filter '*.btw' -File | Copy-Item -Destination $templateRoot
Copy-Item -LiteralPath (Join-Path $projectRoot '修复版使用说明.txt') -Destination $packageRoot
$launcher = [string]::Join([Environment]::NewLine, @('@echo off','cd /d "%~dp0"','set LABEL_PRINTER_WAIT_FOR_COMPLETION=0','set LABEL_PRINTER_FAST_MODE=0','start "" "%~dp0标签自动打印工具_防卡住版.exe"',''))
$encoding = [System.Text.Encoding]::GetEncoding(936)
[System.IO.File]::WriteAllText((Join-Path $packageRoot '启动标签打印.cmd'), $launcher, $encoding)
Write-Output "Built: $packageRoot"
# No running application is closed, no print queue is modified, no user folder is deleted.
