# 打包为单个 exe 并组装发布 zip（在开发机上运行；目标机无需 Python）。
# 由 build_exe.bat 调用，也可直接：powershell -NoProfile -ExecutionPolicy Bypass -File tools\build_exe.ps1
# 注意：本脚本含中文文件名，必须保持 UTF-8 编码（PowerShell 原生支持）。
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

$version = 'v1.2.1'
$exeName = '标签自动打印工具'
$py = 'build_venv\Scripts\python.exe'

# [0/2] 干净打包环境（PySide6 锁 6.8.0.2 —— Qt 6.11 依赖 ICU，会和 Anaconda 的 icuuc.dll 冲突）
if (-not (Test-Path 'build_venv\Scripts\pyinstaller.exe')) {
    Write-Host '[0/2] 首次运行，创建干净打包环境 build_venv ...'
    python -m venv build_venv
    & $py -m pip install -q 'PySide6==6.8.0.2' pandas xlrd openpyxl pywin32 pyinstaller
}

# [1/2] PyInstaller 打包
Write-Host '[1/2] PyInstaller 打包...'
& 'build_venv\Scripts\pyinstaller.exe' --noconfirm --clean --onefile --windowed `
    --name $exeName run_app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 失败 ($LASTEXITCODE)" }

# [2/2] 组装安装包（只打包正式模板，排除 *old* 等历史副本）
Write-Host '[2/2] 组装安装包...'
if (Test-Path 'package') { Remove-Item -Recurse -Force package }
New-Item -ItemType Directory -Force 'package\模板' | Out-Null
Copy-Item "dist\$exeName.exe" 'package\'
Copy-Item '模板\中文版无云南巴菰公司信息标签模板（60-30mm）.btw' 'package\模板\'
Copy-Item '模板\英文版带云南巴菰公司信息标签模板（60-30mm）(1).btw' 'package\模板\'
Copy-Item '使用说明.txt' 'package\' -ErrorAction SilentlyContinue
Copy-Item "docs\${exeName}_使用说明书_$version.docx" 'package\'
Copy-Item "docs\${exeName}_使用说明书_$version.pdf" 'package\'

Get-Process | Where-Object { $_.ProcessName -match '标签自动打印' } | Stop-Process -Force -ErrorAction SilentlyContinue
$zip = "${exeName}_$version.zip"
Compress-Archive -Path 'package\*' -DestinationPath $zip -Force

Write-Host ''
Write-Host '============================================'
Write-Host " 完成：$zip"
Write-Host ' 解压后双击 exe 即可使用（需 BarTender）'
Write-Host '============================================'
