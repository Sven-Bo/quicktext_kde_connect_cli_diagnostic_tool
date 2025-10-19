@echo off
REM =====================================================
REM QuickText - KDE Connect Diagnostic Tool
REM Version 1.0
REM By Sven Bosau
REM https://pythonandvba.com/quicktext
REM 
REM This tool diagnoses KDE Connect CLI configuration
REM for use with QuickText Excel tool.
REM 
REM What it does:
REM - Tests command execution in CMD and PowerShell
REM - Checks if kdeconnect-cli.exe files exist
REM - Verifies PATH environment variable
REM - Shows system information (Windows version, etc.)
REM - Creates diagnostic log file
REM - Opens log folder for easy sharing
REM 
REM Safe to run - only reads, does not modify anything.
REM 
REM For setup help: https://pythonandvba.com/quicktext
REM =====================================================

REM CRITICAL: Keep window open no matter what happens
REM This wrapper ensures the window never closes automatically
if not "%1"=="WRAPPED" (
    cmd /k "%~f0" WRAPPED
    exit
)

setlocal enabledelayedexpansion
title KDE Connect Diagnostic Tool for QuickText

REM Setup colors (safe operation, just visual)
color 0B 2>nul

REM Setup temp files for command testing and logging
set "TESTFILE=%TEMP%\quicktext_kde_test.txt"
set "LOGFILE=%TEMP%\quicktext_kde_diagnostic.log"

REM Define all paths at the beginning to avoid variable expansion issues
set "WINDOWSAPPS_PATH=%LOCALAPPDATA%\Microsoft\WindowsApps"
set "WINDOWSAPPS_CLI=%WINDOWSAPPS_PATH%\kdeconnect-cli.exe"
set "DESKTOP_PATH=C:\Program Files\KDE Connect\bin"
set "DESKTOP_CLI=%DESKTOP_PATH%\kdeconnect-cli.exe"

REM Initialize log file and show system info
echo QuickText - KDE Connect Diagnostic > "!LOGFILE!"
echo Version 1.0 >> "!LOGFILE!"
echo By Sven Bosau >> "!LOGFILE!"
echo ======================================== >> "!LOGFILE!"
echo Date: %date% %time% >> "!LOGFILE!"
echo User: %USERNAME% >> "!LOGFILE!"
echo Computer: %COMPUTERNAME% >> "!LOGFILE!"
ver >> "!LOGFILE!"
echo Architecture: %PROCESSOR_ARCHITECTURE% >> "!LOGFILE!"
echo ======================================== >> "!LOGFILE!"
echo. >> "!LOGFILE!"

cls
echo.
echo ========================================================================
echo                   QuickText - KDE Connect Diagnostic
echo                              Version 1.0
echo                            By Sven Bosau
echo ========================================================================
echo.
echo  Website: https://pythonandvba.com/quicktext
echo.
echo  This tool checks if KDE Connect CLI is properly configured
echo  for use with QuickText Excel tool.
echo.
echo ========================================================================
echo.
echo SYSTEM INFORMATION:
echo   Date/Time    : %date% %time%
echo   User         : %USERNAME%
echo   Computer     : %COMPUTERNAME%
echo   Architecture : %PROCESSOR_ARCHITECTURE%
echo.
echo ========================================================================

REM =====================================================
REM TEST 1: Test the EXACT VBA command
REM This is the actual command that Excel VBA executes
REM =====================================================
echo.
echo.
echo ========================================================================
echo TEST 1: Command Execution Test (CMD)
echo ========================================================================
echo [TEST 1] Command Execution Test >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo Testing the QuickText command:
echo %comspec% /c set LANG=en_US.UTF-8 ^&^& kdeconnect-cli --list-devices ^> %%TEMP%%\test.txt
echo.
echo This tests: UTF-8 encoding, temp file writing, and CLI execution
echo.

REM Clean up any previous test file
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

REM Execute the EXACT command that VBA uses
REM This tests: UTF-8 encoding, temp file writing, and CLI execution
%comspec% /c set LANG=en_US.UTF-8 && kdeconnect-cli --list-devices > "!TESTFILE!" 2>&1

REM Check if command executed and created output file
if not exist "!TESTFILE!" (
    echo [FAIL] Command failed to execute or write to temp folder
    echo.
    echo ERROR: Cannot write to TEMP folder or command failed completely.
    echo.
    goto :checkCliExists
)

REM Read the output from the test file
echo.

echo Output:
echo Output: >> "!LOGFILE!"
echo --------------------
echo -------------------- >> "!LOGFILE!"
type "!TESTFILE!"
type "!TESTFILE!" >> "!LOGFILE!"
echo --------------------
echo -------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Check what the output contains
findstr /C:"is not recognized" "!TESTFILE!" >nul 2>&1
if !errorlevel! equ 0 (
    echo [FAIL] kdeconnect-cli is not recognized as a command
    echo [FAIL] kdeconnect-cli is not recognized as a command >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
    echo This means the CLI is not in your PATH.
    echo This means the CLI is not in your PATH. >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
    goto :testPowerShell
)

REM Check for device output (success case)
findstr /C:"device" "!TESTFILE!" >nul 2>&1
if !errorlevel! equ 0 (
    echo [PASS] Command executed successfully!
    echo.
    
    REM Count devices
    findstr /C:"0 devices found" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [INFO] 0 devices found - no phone is currently connected
        echo.
    ) else (
        findstr /C:"1 device" "!TESTFILE!" >nul 2>&1
        if !errorlevel! equ 0 (
            echo [INFO] 1 device found and connected
            echo.
        ) else (
            echo [INFO] Multiple devices found
            echo.
        )
    )
    
    REM Check for Qt warning (non-critical)
    findstr /C:"QEventDispatcherWin32" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [WARNING] Qt warning detected - this is safe to ignore
        echo.
    )
    
    echo [PASS] CMD test completed successfully
    echo.
    goto :testPowerShell
)

REM Check for other errors
findstr /C:"error" /C:"Error" /C:"ERROR" "!TESTFILE!" >nul 2>&1
if !errorlevel! equ 0 (
    echo [ERROR] Command produced an error
    echo.
    echo See output above for details.
    echo.
    goto :testPowerShell
)

REM Unknown output
echo [WARNING] Unexpected output from command
echo.
goto :testPowerShell

REM =====================================================
REM TEST 1B: Test command in PowerShell
REM =====================================================
:testPowerShell
echo.
echo.
echo ========================================================================
echo TEST 1B: Command Execution Test (PowerShell)
echo ========================================================================
echo [TEST 1B] Testing command in PowerShell >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Clean up previous test file
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

REM Test in PowerShell
powershell -Command "kdeconnect-cli --list-devices" > "!TESTFILE!" 2>&1

if exist "!TESTFILE!" (
    echo PowerShell Output:
    echo PowerShell Output: >> "!LOGFILE!"
    echo --------------------
    echo -------------------- >> "!LOGFILE!"
    type "!TESTFILE!"
    type "!TESTFILE!" >> "!LOGFILE!"
    echo --------------------
    echo -------------------- >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
    
    REM Check if it worked in PowerShell
    findstr /C:"is not recognized" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FAIL] Command also fails in PowerShell
        echo [FAIL] Command also fails in PowerShell >> "!LOGFILE!"
        echo.
        echo. >> "!LOGFILE!"
    ) else (
        findstr /C:"device" "!TESTFILE!" >nul 2>&1
        if !errorlevel! equ 0 (
            echo [PASS] Command works in PowerShell!
            echo [PASS] Command works in PowerShell! >> "!LOGFILE!"
            echo.
            echo. >> "!LOGFILE!"
        )
    )
)

goto :checkCliExists

REM =====================================================
REM TEST 2: Check if CLI executable file exists
REM =====================================================
:checkCliExists
echo.
echo.
echo ========================================================================
echo TEST 2: File Existence Check
echo ========================================================================
echo [TEST 2] Checking if kdeconnect-cli.exe exists >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

set "CLI_FOUND=0"
set "CLI_PATH=NONE"

REM Check WindowsApps location (Microsoft Store version)
echo Checking: %WINDOWSAPPS_CLI%
if exist "%WINDOWSAPPS_CLI%" (
    set "CLI_FOUND=1"
    set "CLI_PATH=%WINDOWSAPPS_CLI%"
    echo [FOUND] Microsoft Store version
    echo [FOUND] Microsoft Store version >> "!LOGFILE!"
    echo Location: %WINDOWSAPPS_CLI%
    echo Location: %WINDOWSAPPS_CLI% >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
) else (
    echo [NOT FOUND] Microsoft Store version not installed
    echo.
)

REM Check Program Files location (Desktop installer version)
echo Checking: %DESKTOP_CLI%
if exist "%DESKTOP_CLI%" (
    set "CLI_FOUND=1"
    if "!CLI_PATH!"=="NONE" set "CLI_PATH=%DESKTOP_CLI%"
    echo [FOUND] Desktop installer version
    echo [FOUND] Desktop installer version >> "!LOGFILE!"
    echo Location: %DESKTOP_CLI%
    echo Location: %DESKTOP_CLI% >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
) else (
    echo [NOT FOUND] Desktop installer version not installed
    echo.
)

REM Check if CLI was found and decide next step
if "!CLI_FOUND!"=="1" goto :checkPath
if "!CLI_FOUND!"=="0" goto :notInstalled

REM Should never reach here
echo ERROR: Unexpected CLI_FOUND value: !CLI_FOUND!
goto :finalReport

:notInstalled
echo [FAIL] kdeconnect-cli.exe not found on your system
echo.
echo ========================================
echo  RESULT: KDE CONNECT NOT INSTALLED
echo ========================================
echo.
echo You need to install KDE Connect first.
echo Choose one of the following options:
echo.
echo ========================================
echo Option 1 - Microsoft Store (Recommended)
echo ========================================
echo 1. Open Microsoft Store
echo 2. Search for "KDE Connect"
echo 3. Click Install
echo 4. Restart your computer
echo.
echo The Microsoft Store version should work
echo automatically after installation.
echo.
echo ========================================
echo Option 2 - Desktop Installer
echo ========================================
echo 1. Visit https://kdeconnect.kde.org/download.html
echo 2. Download Windows installer
echo 3. Run installer
echo 4. Add to PATH (see instructions below)
echo 5. Restart your computer
echo.
echo PATH Setup Instructions:
echo    a. Press Win + R, type sysdm.cpl, press Enter
echo    b. Click the Advanced tab
echo    c. Click Environment Variables button
echo    d. Under User variables, find and select Path
echo    e. Click Edit
echo    f. Click New
echo    g. Add: C:\Program Files\KDE Connect\bin
echo    h. Click OK on all dialogs
echo.
goto :finalReport

:checkPath
echo CLI found, proceeding to PATH check...
echo.

REM =====================================================
REM TEST 3: Check if CLI is in PATH
REM =====================================================
echo.
echo.
echo ========================================================================
echo TEST 3: PATH Environment Variable Check
echo ========================================================================
echo [TEST 3] Checking if CLI location is in PATH >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Check if command is accessible
where kdeconnect-cli >nul 2>&1
if errorlevel 1 (
    echo [FAIL] kdeconnect-cli is NOT in PATH
    echo [FAIL] kdeconnect-cli is NOT in PATH >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
) else (
    echo [PASS] kdeconnect-cli is in PATH
    echo [PASS] kdeconnect-cli is in PATH >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
)

REM Show current user PATH environment variable
echo ------------------------------------------------------------------------
echo Current User PATH Environment Variable:
echo ------------------------------------------------------------------------
echo Current User PATH Environment Variable >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do (
    echo %%b
    echo %%b >> "!LOGFILE!"
)
echo.

REM Check if required paths are in PATH
echo ------------------------------------------------------------------------
echo PATH Analysis:
echo ------------------------------------------------------------------------
echo PATH Analysis >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Check for WindowsApps path
echo Checking for: %WINDOWSAPPS_PATH%
echo Checking for: %WINDOWSAPPS_PATH% >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do (
    echo %%b | findstr /I /C:"%WINDOWSAPPS_PATH%" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FOUND] WindowsApps path is in user PATH
        echo [FOUND] WindowsApps path is in user PATH >> "!LOGFILE!"
    ) else (
        echo [NOT FOUND] WindowsApps path is NOT in user PATH
        echo [NOT FOUND] WindowsApps path is NOT in user PATH >> "!LOGFILE!"
    )
)
echo.

REM Check for Desktop installer path
echo Checking for: %DESKTOP_PATH%
echo Checking for: %DESKTOP_PATH% >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do (
    echo %%b | findstr /I /C:"%DESKTOP_PATH%" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FOUND] Desktop installer path is in user PATH
        echo [FOUND] Desktop installer path is in user PATH >> "!LOGFILE!"
    ) else (
        echo [NOT FOUND] Desktop installer path is NOT in user PATH
        echo [NOT FOUND] Desktop installer path is NOT in user PATH >> "!LOGFILE!"
    )
)
echo.

REM =====================================================
REM FINAL REPORT
REM =====================================================
:finalReport
echo.
echo.
echo ========================================================================
echo                         DIAGNOSTIC COMPLETE
echo ========================================================================
echo.
echo  DIAGNOSTIC COMPLETE >> "!LOGFILE!"
echo  ======================================== >> "!LOGFILE!"
echo. >> "!LOGFILE!"
echo  All diagnostic tests have been completed.
echo.
echo  Log file saved to:
echo  !LOGFILE!
echo.
echo ------------------------------------------------------------------------
echo  Share this log file with QuickText support for assistance:
echo ------------------------------------------------------------------------
echo  1. Press any key to open the log folder
echo  2. Find: quicktext_kde_diagnostic.log
echo  3. Send it via email or support ticket
echo.
echo ========================================================================

:cleanup
REM Clean up test file (safe operation)
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

echo Press any key to open log folder...
pause >nul
start "" "%TEMP%"
goto :eof