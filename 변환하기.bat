@echo off
chcp 65001 >nul
title HTML to HWPX Converter

rem === input HTML: drag&drop (%1) or file dialog ===
set "HTML=%~1"
if "%HTML%"=="" (
  for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Title='Select HTML file'; $d.Filter='HTML|*.html;*.htm|All|*.*'; if($d.ShowDialog() -eq 'OK'){[Console]::Out.Write($d.FileName)}"`) do set "HTML=%%F"
)
if "%HTML%"=="" (
  echo No file selected.
  pause
  exit /b
)

echo.
echo ============================================
echo  Converting: %HTML%
echo ============================================
echo  (Hangul window may flash briefly - do NOT close it.)
echo.

python "%~dp0engine\build.py" "%HTML%"
if errorlevel 1 echo.
if errorlevel 1 echo [ERROR] Conversion failed. Check messages above. Need: Python, Hangul, Gyeonggi fonts.

echo.
pause
