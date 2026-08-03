@echo off
rem Build shim: all logic (and Chinese filenames) lives in tools\build_exe.ps1.
rem This .bat must stay pure ASCII -- cmd.exe cannot parse non-ASCII reliably.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\build_exe.ps1"
exit /b %errorlevel%
