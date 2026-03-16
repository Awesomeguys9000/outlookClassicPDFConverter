@echo off
:: ============================================================
:: Attachment PDF Converter - Installer
:: Must be run as Administrator
:: ============================================================

echo.
echo ==========================================
echo  Attachment PDF Converter - Installer
echo ==========================================
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click the script and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

:: Set install directory
set "INSTALL_DIR=%ProgramFiles%\AttachmentPdfConverter"

:: Create install directory
echo [1/3] Creating install directory...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Copy files
echo [2/3] Copying files...
copy /Y "%~dp0AttachmentPdfConverter.dll" "%INSTALL_DIR%\" >nul
copy /Y "%~dp0AttachmentPdfConverter.dll.config" "%INSTALL_DIR%\" >nul
copy /Y "%~dp0Microsoft.Office.Interop.Outlook.dll" "%INSTALL_DIR%\" >nul
copy /Y "%~dp0Microsoft.Office.Interop.Word.dll" "%INSTALL_DIR%\" >nul
copy /Y "%~dp0Microsoft.Office.Interop.Excel.dll" "%INSTALL_DIR%\" >nul

:: Register the COM add-in using RegAsm
echo [3/3] Registering COM add-in...
set "REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if not exist "%REGASM%" set "REGASM=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"

"%REGASM%" "%INSTALL_DIR%\AttachmentPdfConverter.dll" /codebase /tlb
if %errorlevel% neq 0 (
    echo.
    echo ERROR: COM registration failed. Make sure .NET Framework 4.8 is installed.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo  Installation complete!
echo ==========================================
echo.
echo The add-in has been installed. Please:
echo   1. Close Outlook if it is open
echo   2. Re-open Outlook
echo   3. Compose a new email to see the
echo      "Convert to PDF" button on the ribbon
echo.
pause
