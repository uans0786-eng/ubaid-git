@echo off
chcp 65001 >nul
REM Drag a PDF file onto this .bat file
if "%~1"=="" (
  echo PDF file ko is .bat file par drag karke chhorein.
  pause
  exit /b
)
set /p MARGIN=Word ka neeche ka margin (misal: 0.5in ya 1.27cm): 
set /p PAGES=Kin safhat par kaam ho (misal: 5-120, khali chhorein = tamam): 
set /p SKIP=Kaun se safhat na chheren (misal: 23,47 - khali bhi chhor sakte hain): 
set ARGS=--bottom-margin %MARGIN%
if not "%PAGES%"=="" set ARGS=%ARGS% --pages %PAGES%
if not "%SKIP%"=="" set ARGS=%ARGS% --skip %SKIP%
python "%~dp0bottom_align.py" "%~1" "%~dpn1_barabar.pdf" %ARGS%
pause
