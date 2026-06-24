@echo off
rem 한글(HWP) <-> 워드(DOCX/HWPX) 변환기 GUI 실행 (콘솔 없이 창으로 뜸)
cd /d "%~dp0"
if exist "C:\Python313\pythonw.exe" (
    start "" "C:\Python313\pythonw.exe" "%~dp0hwp2docx_gui.py"
) else (
    start "" pythonw "%~dp0hwp2docx_gui.py"
)
