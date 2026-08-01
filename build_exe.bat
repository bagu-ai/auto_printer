@echo off
rem 打包为单个 exe（在开发机上运行；目标机无需 Python）
rem 使用干净的 build_venv 打包，避免 Anaconda 全家桶拖垮 PyInstaller
setlocal
cd /d %~dp0

if not exist build_venv\Scripts\pyinstaller.exe (
  echo [0/2] 首次运行，创建干净打包环境 build_venv ...
  python -m venv build_venv || goto :err
  rem 注意：PySide6 锁 6.8.0.2 —— Qt 6.11 依赖 ICU，会和 Anaconda 的 icuuc.dll 冲突导致 DLL 加载失败
  build_venv\Scripts\python.exe -m pip install -q "PySide6==6.8.0.2" pandas xlrd openpyxl pywin32 pyinstaller || goto :err
)

echo [1/2] PyInstaller 打包...
build_venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed ^
  --name "标签自动打印工具" ^
  run_app.py || goto :err

echo [2/2] 组装安装包...
if exist package rmdir /s /q package
mkdir package\模板
copy /y "dist\标签自动打印工具.exe" package\ >nul
copy /y "模板\*.btw" package\模板\ >nul
copy /y "使用说明.txt" package\ 2>nul
powershell -NoProfile -Command "Get-Process | Where-Object { $_.ProcessName -match '标签自动打印' } | Stop-Process -Force -ErrorAction SilentlyContinue; Compress-Archive -Path 'package\*' -DestinationPath '标签自动打印工具_v1.2.zip' -Force" || goto :err

echo.
echo ============================================
echo  完成：标签自动打印工具_v1.2.zip
echo  解压后双击 exe 即可使用（需 BarTender）
echo ============================================
goto :eof

:err
echo 打包失败，请检查上方错误信息
exit /b 1
