@echo off
title Enhanced Power BI Report Tools v2.0
color 0A

echo.
echo ========================================================
echo   ENHANCED POWER BI REPORT TOOLS v2.0 - PLUGIN ARCHITECTURE
echo   Built by Reid Havens of Analytic Endeavors
echo   Website: https://www.analyticendeavors.com
echo ========================================================
echo.
echo Enhanced with comprehensive security audit logging
echo Professional-grade Power BI report management suite
echo NEW: Plugin-based architecture with automatic tool discovery
echo.
echo FEATURES:
echo    Advanced Page Copy - Duplicate pages with bookmarks
echo    Report Merger - Combine multiple reports
echo    Theme Management - Intelligent conflict resolution
echo    Tool Manager - Automatic plugin discovery
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ and try again.
    pause
    exit /b 1
)

if exist "main.py" (
    echo Using consolidated main application main.py...
    python main.py
) else if exist "enhanced_main.py" (
    echo Using enhanced plugin architecture enhanced_main.py...
    python enhanced_main.py
) else (
    echo ERROR: No main application file found
    echo Please ensure main.py exists.
    pause
    exit /b 1
)

if errorlevel 1 (
    echo.
    echo ERROR: Application encountered an error
    echo Check the console output above for details.
    pause
    exit /b 1
)

echo.
echo Enhanced Power BI Report Tools session completed!
echo Visit us at: https://www.analyticendeavors.com
pause
