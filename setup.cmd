@echo off
rem One-time setup for the excel365 connector (called by launch.cmd; output goes to setup.log).
rem 1. Find Python 3.10+; install Python 3.12 for this user with winget if there is none.
rem 2. Install openpyxl into %LOCALAPPDATA%\excel365\lib (private to this connector).
setlocal EnableDelayedExpansion
set "DATA=%LOCALAPPDATA%\excel365"
set "LIB=%DATA%\lib"
set "CHECK=import sys; print(sys.executable if sys.version_info >= (3, 10) else '')"
echo === excel365 setup %DATE% %TIME% ===

call :findpy
if not defined PY (
  echo Python 3.10+ not found - installing Python 3.12 for this user with winget
  where winget >nul 2>&1 || (echo winget not available - install Python 3.12 from python.org, then restart Claude & exit /b 1)
  winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
  call :findpy
  if not defined PY (echo Python install did not complete & exit /b 1)
)
echo python: !PY!

"!PY!" -m pip --version >nul 2>&1 || "!PY!" -m ensurepip --user
if not exist "%LIB%" mkdir "%LIB%"
"!PY!" -m pip install --disable-pip-version-check --no-warn-script-location --upgrade --target "%LIB%" "openpyxl>=3.1,<4" || (echo pip install openpyxl failed & exit /b 1)
"!PY!" -c "import sys; sys.path.insert(0, r'%LIB%'); import openpyxl; print('openpyxl', openpyxl.__version__)" || (echo openpyxl import check failed & exit /b 1)

> "%DATA%\python.txt" (echo|set /p="!PY!")
echo %DATE% %TIME%> "%DATA%\ready-1"
echo setup complete
exit /b 0

:findpy
set "PY="
rem py launcher first, then python on PATH (skipping the Microsoft Store alias), then default install folders
for /f "delims=" %%E in ('py -3 -c "%CHECK%" 2^>nul') do if not defined PY set "PY=%%E"
if defined PY exit /b 0
for /f "delims=" %%P in ('where python 2^>nul ^| findstr /v /i "WindowsApps"') do (
  if not defined PY for /f "delims=" %%E in ('"%%P" -c "%CHECK%" 2^>nul') do set "PY=%%E"
)
if defined PY exit /b 0
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "%ProgramFiles%\Python3*") do (
  if not defined PY if exist "%%D\python.exe" for /f "delims=" %%E in ('"%%D\python.exe" -c "%CHECK%" 2^>nul') do set "PY=%%E"
)
exit /b 0
