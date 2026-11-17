@echo off
chcp 65001 >nul
REM Alaamri MS Project MCP - one-click setup (Developed by Dr Ibrahim M.A. Alaamri, MIT)
REM Run this file from inside the unzipped alaamri-msproject-mcp folder.
cd /d "%~dp0"
where python >nul 2>nul || (echo [!] Python is not installed. Install Python 3.10+ from python.org and tick "Add python.exe to PATH". & pause & exit /b 1)
echo [1/3] Creating virtual environment...
python -m venv .venv || (pause & exit /b 1)
echo [2/3] Installing the server...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -e . || (pause & exit /b 1)
echo [3/3] Done. Add this to %%APPDATA%%\Claude\claude_desktop_config.json :
set "P=%~dp0"
set "P=%P:\=\\%"
echo.
echo {
echo   "mcpServers": {
echo     "alaamri-msproject": {
echo       "command": "%P%.venv\\Scripts\\python.exe",
echo       "args": ["%P%server.py"]
echo     }
echo   }
echo }
echo.
echo Then fully quit and reopen Claude Desktop.
pause
