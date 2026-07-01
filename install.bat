@echo off
REM AlphaCAM MCP Bridge — Windows Install Script
REM Run as Administrator for VBA add-in installation

echo === AlphaCAM MCP Bridge Installer ===
echo.

REM 1. Install Python dependencies
echo [1/4] Installing Python packages...
pip install -r requirements.txt
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: pip install failed. Make sure Python and pip are installed.
    pause
    exit /b 1
)

REM 2. Test imports
echo [2/4] Testing imports...
python -c "import mcp; import win32com; print('OK')" 
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Import test failed.
    pause
    exit /b 1
)

REM 3. Install VBA add-ins (CCC features)
echo [3/4] Installing VBA add-ins...
python install_vba.py
IF %ERRORLEVEL% NEQ 0 (
    echo WARNING: VBA installation skipped or failed.
    echo You can manually run: python install_vba.py
)

REM 4. Test AlphaCAM connection
echo [4/4] Testing AlphaCAM connection...
python -c "import sys; sys.path.insert(0, '.'); from alphacam_com import AlphaCAM; a=AlphaCAM(); print(a.get_info())"
IF %ERRORLEVEL% NEQ 0 (
    echo WARNING: Could not connect to AlphaCAM.
    echo Make sure AlphaCAM 2016 R1 is installed and has been run at least once.
)

echo.
echo === Installation Complete ===
echo.
echo Available tools: %cd%\.env.example
echo Copy to .env and edit to customize: copy .env.example .env
echo.
echo To start the MCP server with stdio (for Reasonix config):
echo   python server.py
echo.
echo To start with SSE (HTTP, port 8080):
echo   python server.py --port 8080
echo.
pause
