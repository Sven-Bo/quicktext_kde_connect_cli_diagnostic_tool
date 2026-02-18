@echo off
REM =====================================================
REM QuickText - KDE Connect Diagnostic Tool
REM Version 1.2
REM By Sven Bosau
REM https://pythonandvba.com/quicktext
REM 
REM This tool diagnoses KDE Connect CLI configuration
REM for use with QuickText Excel tool.
REM 
REM What it does:
REM - Tests direct CLI execution (how QuickText v1.6+ works)
REM - Tests command execution in CMD and PowerShell (legacy)
REM - Checks if kdeconnect-cli.exe files exist
REM - Verifies PATH environment variable (User + System)
REM - Shows system information (Windows version, Excel, etc.)
REM - Creates diagnostic log file
REM - Opens log folder for easy sharing
REM 
REM Safe to run - only reads, does not modify anything.
REM 
REM For setup help: https://pythonandvba.com/docs/quicktext/troubleshooting/quicktext-run-kde-connect-diagnostic-tool/
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
echo Version 1.1 >> "!LOGFILE!"
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
echo                              Version 1.1
echo                            By Sven Bosau
echo ========================================================================
echo.
echo  Docs: https://pythonandvba.com/docs/quicktext/troubleshooting/quicktext-run-kde-connect-diagnostic-tool/
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

REM Detect installed Excel versions and bitness
set "EXCEL_INFO=Not detected"
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul') do (
    set "EXCEL_BITNESS=%%b"
)
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport 2^>nul') do (
    set "EXCEL_VERSION=%%b"
)
if defined EXCEL_VERSION (
    set "EXCEL_INFO=!EXCEL_VERSION! (!EXCEL_BITNESS!)"
) else (
    REM Try MSI-based Office detection
    for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot" /v Path 2^>nul') do (
        set "EXCEL_INFO=Office 16.0 (MSI install)"
    )
)
echo   Excel        : !EXCEL_INFO!
echo Excel: !EXCEL_INFO! >> "!LOGFILE!"
echo.
echo ========================================================================

REM =====================================================
REM TEST 1: Direct CLI execution (how QuickText v1.6+ works)
REM QuickText uses CreateProcessW to call kdeconnect-cli
REM directly - NOT through cmd.exe or PowerShell.
REM This test mirrors that exact behavior.
REM =====================================================
echo.
echo.
echo ========================================================================
echo TEST 1: Direct CLI Execution (QuickText v1.6+ method)
echo ========================================================================
echo [TEST 1] Direct CLI Execution (CreateProcess method) >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo QuickText v1.6+ calls kdeconnect-cli.exe directly via CreateProcess
echo (without cmd.exe or PowerShell). This test mirrors that behavior.
echo.

REM Clean up any previous test file
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

REM Test 1a: Direct --version call (mirrors QuickText IsInstalled check)
echo --- Test 1a: kdeconnect-cli --version ---
echo --- Test 1a: kdeconnect-cli --version --- >> "!LOGFILE!"
kdeconnect-cli --version > "!TESTFILE!" 2>&1
set "DIRECT_VERSION_EXIT=!errorlevel!"

if exist "!TESTFILE!" (
    echo Output:
    echo Output: >> "!LOGFILE!"
    echo --------------------
    echo -------------------- >> "!LOGFILE!"
    type "!TESTFILE!"
    type "!TESTFILE!" >> "!LOGFILE!"
    echo --------------------
    echo -------------------- >> "!LOGFILE!"
    echo Exit code: !DIRECT_VERSION_EXIT!
    echo Exit code: !DIRECT_VERSION_EXIT! >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"

    REM Check for errors FIRST (error messages also contain 'kdeconnect-cli')
    findstr /C:"is not recognized" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FAIL] kdeconnect-cli is not recognized when called directly
        echo [FAIL] Direct call not recognized >> "!LOGFILE!"
        echo.
        echo This is the EXACT error QuickText would encounter.
        echo This is the EXACT error QuickText would encounter. >> "!LOGFILE!"
        echo The CLI executable is not in the system PATH.
        echo The CLI executable is not in the system PATH. >> "!LOGFILE!"
    ) else (
        REM No error found - check for valid version output
        findstr /C:"kdeconnect-cli" "!TESTFILE!" >nul 2>&1
        if !errorlevel! equ 0 (
            echo [PASS] Direct --version call works! This is what QuickText uses.
            echo [PASS] Direct --version call works >> "!LOGFILE!"
        ) else (
            echo [FAIL] Direct --version call returned unexpected output
            echo [FAIL] Unexpected output from direct call >> "!LOGFILE!"
        )
    )
) else (
    echo [FAIL] Direct --version call failed completely (no output)
    echo [FAIL] Direct --version call failed completely >> "!LOGFILE!"
    echo.
    echo This means kdeconnect-cli.exe cannot be found by the system.
    echo This means kdeconnect-cli.exe cannot be found by the system. >> "!LOGFILE!"
    echo QuickText would show: "CreateProcess Failed, Code: 2"
    echo QuickText would show: CreateProcess Failed, Code: 2 >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"

REM Test 1b: Direct --list-devices call
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

echo --- Test 1b: kdeconnect-cli --list-devices ---
echo --- Test 1b: kdeconnect-cli --list-devices --- >> "!LOGFILE!"
kdeconnect-cli --list-devices > "!TESTFILE!" 2>&1
set "DIRECT_LIST_EXIT=!errorlevel!"

if exist "!TESTFILE!" (
    echo Output:
    echo Output: >> "!LOGFILE!"
    echo --------------------
    echo -------------------- >> "!LOGFILE!"
    type "!TESTFILE!"
    type "!TESTFILE!" >> "!LOGFILE!"
    echo --------------------
    echo -------------------- >> "!LOGFILE!"
    echo Exit code: !DIRECT_LIST_EXIT!
    echo Exit code: !DIRECT_LIST_EXIT! >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"

    REM Check for errors FIRST before checking for success
    findstr /C:"is not recognized" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FAIL] kdeconnect-cli is not recognized when called directly
        echo [FAIL] Direct --list-devices not recognized >> "!LOGFILE!"
    ) else (
        findstr /C:"device" "!TESTFILE!" >nul 2>&1
        if !errorlevel! equ 0 (
            echo [PASS] Direct --list-devices call works!
            echo [PASS] Direct --list-devices call works >> "!LOGFILE!"
            echo.

            REM Count devices
            findstr /C:"0 devices found" "!TESTFILE!" >nul 2>&1
            if !errorlevel! equ 0 (
                echo [INFO] 0 devices found - no phone is currently connected
                echo [INFO] 0 devices found >> "!LOGFILE!"
            ) else (
                findstr /C:"1 device" "!TESTFILE!" >nul 2>&1
                if !errorlevel! equ 0 (
                    echo [INFO] 1 device found and connected
                    echo [INFO] 1 device found >> "!LOGFILE!"
                ) else (
                    echo [INFO] Multiple devices found
                    echo [INFO] Multiple devices found >> "!LOGFILE!"
                )
            )

            REM Check for Qt warning (non-critical)
            findstr /C:"QEventDispatcherWin32" "!TESTFILE!" >nul 2>&1
            if !errorlevel! equ 0 (
                echo [WARNING] Qt warning detected - this is safe to ignore
                echo [WARNING] Qt warning detected - safe to ignore >> "!LOGFILE!"
            )
        ) else (
            echo [FAIL] Direct --list-devices call did not return device info
            echo [FAIL] Direct --list-devices failed >> "!LOGFILE!"
        )
    )
) else (
    echo [FAIL] Direct --list-devices call failed completely (no output)
    echo [FAIL] Direct --list-devices failed completely >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"

REM =====================================================
REM TEST 2: Legacy command execution tests (CMD + PowerShell)
REM These were used by older QuickText versions.
REM Useful to compare: if these pass but Test 1 fails,
REM it indicates a PATH issue specific to direct execution.
REM =====================================================
echo.
echo.
echo ========================================================================
echo TEST 2: Legacy Command Execution (CMD + PowerShell)
echo ========================================================================
echo [TEST 2] Legacy Command Execution Tests >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo These tests use cmd.exe and PowerShell (older QuickText method).
echo If these pass but Test 1 fails, it means the CLI works through
echo a shell but not via direct execution (CreateProcess).
echo.

REM Test 2a: CMD
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

echo --- Test 2a: CMD ---
echo --- Test 2a: CMD --- >> "!LOGFILE!"
%comspec% /c kdeconnect-cli --list-devices > "!TESTFILE!" 2>&1

if exist "!TESTFILE!" (
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

    findstr /C:"is not recognized" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FAIL] CMD: kdeconnect-cli is not recognized
        echo [FAIL] CMD: not recognized >> "!LOGFILE!"
    ) else (
        findstr /C:"device" "!TESTFILE!" >nul 2>&1
        if !errorlevel! equ 0 (
            echo [PASS] CMD: Command works!
            echo [PASS] CMD: works >> "!LOGFILE!"
        ) else (
            echo [WARNING] CMD: Unexpected output
            echo [WARNING] CMD: Unexpected output >> "!LOGFILE!"
        )
    )
) else (
    echo [FAIL] CMD: No output produced
    echo [FAIL] CMD: No output >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"

REM Test 2b: PowerShell
if exist "!TESTFILE!" del "!TESTFILE!" 2>nul

echo --- Test 2b: PowerShell ---
echo --- Test 2b: PowerShell --- >> "!LOGFILE!"
powershell -Command "kdeconnect-cli --list-devices" > "!TESTFILE!" 2>&1

if exist "!TESTFILE!" (
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

    findstr /C:"is not recognized" "!TESTFILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FAIL] PowerShell: kdeconnect-cli is not recognized
        echo [FAIL] PowerShell: not recognized >> "!LOGFILE!"
    ) else (
        findstr /C:"device" "!TESTFILE!" >nul 2>&1
        if !errorlevel! equ 0 (
            echo [PASS] PowerShell: Command works!
            echo [PASS] PowerShell: works >> "!LOGFILE!"
        ) else (
            echo [WARNING] PowerShell: Unexpected output
            echo [WARNING] PowerShell: Unexpected output >> "!LOGFILE!"
        )
    )
) else (
    echo [FAIL] PowerShell: No output produced
    echo [FAIL] PowerShell: No output >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"

REM =====================================================
REM TEST 3: Check if CLI executable file exists
REM =====================================================
echo.
echo.
echo ========================================================================
echo TEST 3: File Existence Check
echo ========================================================================
echo [TEST 3] Checking if kdeconnect-cli.exe exists >> "!LOGFILE!"
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
    echo [NOT FOUND] Microsoft Store version >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
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
    echo [NOT FOUND] Desktop installer version >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
)

REM Check if CLI was found and decide next step
if "!CLI_FOUND!"=="1" goto :checkPath
if "!CLI_FOUND!"=="0" goto :notInstalled

REM Should never reach here
echo ERROR: Unexpected CLI_FOUND value: !CLI_FOUND!
goto :finalReport

:notInstalled
echo [FAIL] kdeconnect-cli.exe not found on your system
echo [FAIL] kdeconnect-cli.exe not found >> "!LOGFILE!"
echo.
echo ========================================
echo  RESULT: KDE CONNECT CLI NOT INSTALLED
echo ========================================
echo  RESULT: KDE CONNECT CLI NOT INSTALLED >> "!LOGFILE!"
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
REM TEST 4: Check if CLI is in PATH
REM =====================================================
echo.
echo.
echo ========================================================================
echo TEST 4: PATH Environment Variable Check
echo ========================================================================
echo [TEST 4] Checking if CLI location is in PATH >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Check if command is accessible via 'where'
echo Resolved CLI location(s) via 'where':
echo Resolved CLI location(s) via 'where': >> "!LOGFILE!"
where kdeconnect-cli > "!TESTFILE!" 2>&1
if errorlevel 1 (
    echo [FAIL] kdeconnect-cli is NOT in PATH
    echo [FAIL] kdeconnect-cli is NOT in PATH >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
) else (
    type "!TESTFILE!"
    type "!TESTFILE!" >> "!LOGFILE!"
    echo [PASS] kdeconnect-cli is in PATH
    echo [PASS] kdeconnect-cli is in PATH >> "!LOGFILE!"
    echo.
    echo. >> "!LOGFILE!"
)
echo.

REM Show current User PATH environment variable
echo ------------------------------------------------------------------------
echo User PATH (HKCU\Environment):
echo ------------------------------------------------------------------------
echo User PATH (HKCU\Environment): >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do (
    echo %%b
    echo %%b >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"

REM Show current System PATH environment variable
echo ------------------------------------------------------------------------
echo System PATH (HKLM):
echo ------------------------------------------------------------------------
echo System PATH (HKLM): >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do (
    echo %%b
    echo %%b >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"

REM Check if required paths are in PATH
echo ------------------------------------------------------------------------
echo PATH Analysis:
echo ------------------------------------------------------------------------
echo PATH Analysis >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Check for WindowsApps path in User PATH
echo Checking User PATH for: %WINDOWSAPPS_PATH%
echo Checking User PATH for: %WINDOWSAPPS_PATH% >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do (
    echo %%b | findstr /I /C:"%WINDOWSAPPS_PATH%" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FOUND] WindowsApps path is in User PATH
        echo [FOUND] WindowsApps path is in User PATH >> "!LOGFILE!"
    ) else (
        echo [NOT FOUND] WindowsApps path is NOT in User PATH
        echo [NOT FOUND] WindowsApps path is NOT in User PATH >> "!LOGFILE!"
    )
)
echo.

REM Check for Desktop installer path in User PATH
echo Checking User PATH for: %DESKTOP_PATH%
echo Checking User PATH for: %DESKTOP_PATH% >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do (
    echo %%b | findstr /I /C:"%DESKTOP_PATH%" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FOUND] Desktop installer path is in User PATH
        echo [FOUND] Desktop installer path is in User PATH >> "!LOGFILE!"
    ) else (
        echo [NOT FOUND] Desktop installer path is NOT in User PATH
        echo [NOT FOUND] Desktop installer path is NOT in User PATH >> "!LOGFILE!"
    )
)
echo.

REM Check for Desktop installer path in System PATH
echo Checking System PATH for: %DESKTOP_PATH%
echo Checking System PATH for: %DESKTOP_PATH% >> "!LOGFILE!"
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do (
    echo %%b | findstr /I /C:"%DESKTOP_PATH%" >nul 2>&1
    if !errorlevel! equ 0 (
        echo [FOUND] Desktop installer path is in System PATH
        echo [FOUND] Desktop installer path is in System PATH >> "!LOGFILE!"
    ) else (
        echo [NOT FOUND] Desktop installer path is NOT in System PATH
        echo [NOT FOUND] Desktop installer path is NOT in System PATH >> "!LOGFILE!"
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

REM =====================================================
REM DIAGNOSIS: Interpret results and provide actionable hints
REM =====================================================

REM Collect results from log
set "DIAG_DIRECT_VERSION=UNKNOWN"
set "DIAG_DIRECT_LIST=UNKNOWN"
set "DIAG_CMD=UNKNOWN"
set "DIAG_PS=UNKNOWN"
set "DIAG_FILE_STORE=UNKNOWN"
set "DIAG_FILE_DESKTOP=UNKNOWN"
set "DIAG_WHERE=UNKNOWN"

findstr /C:"[PASS] Direct --version call works" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_DIRECT_VERSION=PASS") else (set "DIAG_DIRECT_VERSION=FAIL")

findstr /C:"[PASS] Direct --list-devices call works" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_DIRECT_LIST=PASS") else (set "DIAG_DIRECT_LIST=FAIL")

findstr /C:"[PASS] CMD: works" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_CMD=PASS") else (set "DIAG_CMD=FAIL")

findstr /C:"[PASS] PowerShell: works" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_PS=PASS") else (set "DIAG_PS=FAIL")

findstr /C:"[FOUND] Microsoft Store version" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_FILE_STORE=FOUND") else (set "DIAG_FILE_STORE=NOT_FOUND")

findstr /C:"[FOUND] Desktop installer version" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_FILE_DESKTOP=FOUND") else (set "DIAG_FILE_DESKTOP=NOT_FOUND")

findstr /C:"[PASS] kdeconnect-cli is in PATH" "!LOGFILE!" >nul 2>&1
if !errorlevel! equ 0 (set "DIAG_WHERE=PASS") else (set "DIAG_WHERE=FAIL")

REM Write summary + diagnosis to both screen and log
echo ========================================================================
echo                            DIAGNOSIS
echo ========================================================================
echo. >> "!LOGFILE!"
echo ======================================== >> "!LOGFILE!"
echo DIAGNOSIS >> "!LOGFILE!"
echo ======================================== >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo  Test Results:
echo  Test Results: >> "!LOGFILE!"
echo    Direct CLI (v1.6+) : !DIAG_DIRECT_VERSION!
echo    Direct CLI (v1.6+) : !DIAG_DIRECT_VERSION! >> "!LOGFILE!"
echo    CMD (legacy)       : !DIAG_CMD!
echo    CMD (legacy)       : !DIAG_CMD! >> "!LOGFILE!"
echo    PowerShell (legacy): !DIAG_PS!
echo    PowerShell (legacy): !DIAG_PS! >> "!LOGFILE!"
echo    File exists (Store): !DIAG_FILE_STORE!
echo    File exists (Store): !DIAG_FILE_STORE! >> "!LOGFILE!"
echo    File exists (Dsktp): !DIAG_FILE_DESKTOP!
echo    File exists (Dsktp): !DIAG_FILE_DESKTOP! >> "!LOGFILE!"
echo    In PATH (where)    : !DIAG_WHERE!
echo    In PATH (where)    : !DIAG_WHERE! >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

REM Now provide the actual diagnosis using goto-based branching
REM (batch cmd.exe does not reliably support else-if chains)
echo ------------------------------------------------------------------------
echo  Diagnosis:
echo ------------------------------------------------------------------------
echo  Diagnosis: >> "!LOGFILE!"
echo ---------------------------------------- >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"

if "!DIAG_DIRECT_VERSION!"=="PASS" if "!DIAG_DIRECT_LIST!"=="PASS" goto :diag_ok
if "!DIAG_DIRECT_VERSION!"=="PASS" goto :diag_partial
if "!DIAG_FILE_STORE!"=="NOT_FOUND" if "!DIAG_FILE_DESKTOP!"=="NOT_FOUND" goto :diag_not_installed
if "!DIAG_WHERE!"=="FAIL" goto :diag_path_issue
goto :diag_unknown

:diag_ok
echo  [OK] Everything works! CLI is installed and accessible.
echo  [OK] Everything works! CLI is installed and accessible. >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo  If QuickText still has issues, the problem is likely:
echo  If QuickText still has issues, the problem is likely: >> "!LOGFILE!"
echo    - Phone not connected / KDE Connect app not running on phone
echo    - Phone not connected / KDE Connect app not running on phone >> "!LOGFILE!"
echo    - Phone and PC not on the same Wi-Fi network
echo    - Phone and PC not on the same Wi-Fi network >> "!LOGFILE!"
echo    - Device not paired in KDE Connect
echo    - Device not paired in KDE Connect >> "!LOGFILE!"
goto :diag_done

:diag_partial
echo  [PARTIAL] --version works but --list-devices fails.
echo  [PARTIAL] --version works but --list-devices fails. >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo  Possible cause: KDE Connect service is not running.
echo  Possible cause: KDE Connect service is not running. >> "!LOGFILE!"
echo  Fix: Start KDE Connect on the PC, then retry.
echo  Fix: Start KDE Connect on the PC, then retry. >> "!LOGFILE!"
goto :diag_done

:diag_not_installed
echo  [NOT INSTALLED] KDE Connect CLI executable not found anywhere.
echo  [NOT INSTALLED] KDE Connect CLI executable not found anywhere. >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
echo  Fix: Install KDE Connect from the Microsoft Store (recommended)
echo  Fix: Install KDE Connect from the Microsoft Store >> "!LOGFILE!"
echo  or download from https://kdeconnect.kde.org/download.html
echo  or download from https://kdeconnect.kde.org/download.html >> "!LOGFILE!"
echo  Then restart the computer.
echo  Then restart the computer. >> "!LOGFILE!"
goto :diag_done

:diag_path_issue
echo  [PATH ISSUE] CLI executable exists but is NOT in PATH.
echo  [PATH ISSUE] CLI executable exists but is NOT in PATH. >> "!LOGFILE!"
echo  This is why QuickText shows "CreateProcess Failed, Code: 2"
echo  This is why QuickText shows CreateProcess Failed, Code: 2 >> "!LOGFILE!"
echo.
echo. >> "!LOGFILE!"
if "!DIAG_FILE_STORE!"=="FOUND" (
    echo  Found: Microsoft Store version
    echo  Found: Microsoft Store version >> "!LOGFILE!"
    echo  Fix: Add WindowsApps to User PATH:
    echo  Fix: Add WindowsApps to User PATH: >> "!LOGFILE!"
    echo    1. Press Win+R, type sysdm.cpl, press Enter
    echo    1. Press Win+R, type sysdm.cpl, press Enter >> "!LOGFILE!"
    echo    2. Advanced tab ^> Environment Variables
    echo    2. Advanced tab, Environment Variables >> "!LOGFILE!"
    echo    3. Under User variables, edit Path
    echo    3. Under User variables, edit Path >> "!LOGFILE!"
    echo    4. Add: %%LOCALAPPDATA%%\Microsoft\WindowsApps
    echo    4. Add: %%LOCALAPPDATA%%\Microsoft\WindowsApps >> "!LOGFILE!"
    echo    5. Click OK, then restart computer
    echo    5. Click OK, then restart computer >> "!LOGFILE!"
)
if "!DIAG_FILE_DESKTOP!"=="FOUND" (
    echo  Found: Desktop installer version
    echo  Found: Desktop installer version >> "!LOGFILE!"
    echo  Fix: Add KDE Connect bin to PATH:
    echo  Fix: Add KDE Connect bin to PATH: >> "!LOGFILE!"
    echo    1. Press Win+R, type sysdm.cpl, press Enter
    echo    1. Press Win+R, type sysdm.cpl, press Enter >> "!LOGFILE!"
    echo    2. Advanced tab ^> Environment Variables
    echo    2. Advanced tab, Environment Variables >> "!LOGFILE!"
    echo    3. Under User variables, edit Path
    echo    3. Under User variables, edit Path >> "!LOGFILE!"
    echo    4. Add: C:\Program Files\KDE Connect\bin
    echo    4. Add: C:\Program Files\KDE Connect\bin >> "!LOGFILE!"
    echo    5. Click OK, then restart computer
    echo    5. Click OK, then restart computer >> "!LOGFILE!"
)
echo.
echo. >> "!LOGFILE!"
echo  Alternative: In QuickText Settings, enter the full path to
echo  Alternative: In QuickText Settings, enter the full path to >> "!LOGFILE!"
echo  kdeconnect-cli.exe in the KDE Connect Path field. No restart needed.
echo  kdeconnect-cli.exe in the KDE Connect Path field. No restart needed. >> "!LOGFILE!"
goto :diag_done

:diag_unknown
echo  [UNKNOWN] Unexpected combination of results.
echo  [UNKNOWN] Unexpected combination of results. >> "!LOGFILE!"
echo  Please review the full log details above.
echo  Please review the full log details above. >> "!LOGFILE!"
goto :diag_done

:diag_done

echo.
echo. >> "!LOGFILE!"
echo ======================================== >> "!LOGFILE!"
echo DIAGNOSTIC COMPLETE >> "!LOGFILE!"
echo ======================================== >> "!LOGFILE!"
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
