@echo off
rem Starts the excel365 MCP server. First run: installs Python (if missing) and openpyxl via setup.cmd.
rem stdout belongs to the MCP protocol, so setup output goes to %LOCALAPPDATA%\excel365\setup.log.
setlocal
set "ROOT=%~dp0"
set "DATA=%LOCALAPPDATA%\excel365"
if not exist "%DATA%" mkdir "%DATA%"
if not exist "%DATA%\ready-1" (
  call "%ROOT%setup.cmd" >> "%DATA%\setup.log" 2>&1
  if errorlevel 1 (
    echo excel365 setup failed - see %DATA%\setup.log 1>&2
    exit /b 1
  )
)
set /p PY=<"%DATA%\python.txt"
set "PYTHONIOENCODING=utf-8"
"%PY%" "%ROOT%server\xl_mcp.py"
