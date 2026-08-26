@echo off
REM Build a standalone docx2pdf.exe with PyInstaller
setlocal
cd /d "%~dp0.."

pip install pyinstaller pywin32 || goto :err

REM shadow the pyinstaller-hooks-contrib hook for the *other* PyPI "docx2pdf"
REM package which calls copy_metadata('docx2pdf') and breaks on our repo
if not exist build\hooks mkdir build\hooks
> build\hooks\hook-docx2pdf.py echo hiddenimports = ["docx2pdf.core", "docx2pdf.cli", "docx2pdf"]

python -m PyInstaller --onefile --console --name docx2pdf ^
  --paths "%~dp0.." ^
  --additional-hooks-dir "%~dp0..\build\hooks" ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  launcher.py || goto :err

echo.
echo Built: dist\docx2pdf.exe
echo Try:  dist\docx2pdf.exe convert --list
exit /b 0

:err
echo Build failed.
exit /b 1