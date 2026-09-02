@echo off
cd /d "%~dp0"
set PYINSTALLER_CONFIG_DIR=%~dp0.pyinstaller
set TEMP=%~dp0.tmp
set TMP=%~dp0.tmp
if not exist "%TEMP%" mkdir "%TEMP%"
if not exist "%PYINSTALLER_CONFIG_DIR%" mkdir "%PYINSTALLER_CONFIG_DIR%"
echo starting > build_status.txt
python run_build.py
echo exitcode=%ERRORLEVEL% >> build_status.txt
