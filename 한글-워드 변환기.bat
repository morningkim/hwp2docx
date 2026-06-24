@echo off
cd /d "%~dp0"
if exist "C:\Python313\pythonw.exe" (
    start "" "C:\Python313\pythonw.exe" "%~dp0hwp2docx_gui.py"
) else (
    start "" pythonw "%~dp0hwp2docx_gui.py"
)
