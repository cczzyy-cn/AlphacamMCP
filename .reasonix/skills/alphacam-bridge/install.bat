@echo off
REM AlphaCAM MCP Bridge — Windows Install Script
REM Run as Administrator

echo Installing AlphaCAM MCP Bridge...

REM 1. Install Python dependencies
echo [1/3] Installing Python packages...
pip install mcp pywin32
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: pip install failed. Make sure Python and pip are installed.
    pause
    exit /b 1
)

REM 2. Test import
echo [2/3] Testing imports...
python -c "import mcp; import win32com; print('OK')" 
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Import test failed.
    pause
    exit /b 1
)

REM 3. Test AlphaCAM connection
echo [3/3] Testing AlphaCAM connection...
python -c "import sys; sys.path.insert(0, '.'); from alphacam_com import AlphaCAM; a=AlphaCAM(); print(a.get_info())"
IF %ERRORLEVEL% NEQ 0 (
    echo WARNING: Could not connect to AlphaCAM.
    echo Make sure AlphaCAM 2016 R1 is installed and has been run at least once.
)

echo.
echo === Installation Complete ===
echo.
echo To start the MCP server with stdio (for Reasonix config):
echo   python server.py
echo.
echo To start with SSE (HTTP, port 8080):
echo   python server.py --port 8080
echo.
pause
