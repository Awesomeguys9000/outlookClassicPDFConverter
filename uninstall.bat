@echo off
:: ============================================================
:: Attachment PDF Converter - Uninstaller
:: Must be run as Administrator
:: ============================================================

echo.
echo ==========================================
echo  Attachment PDF Converter - Uninstaller
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

set "INSTALL_DIR=%ProgramFiles%\AttachmentPdfConverter"

:: Unregister the COM add-in
echo [1/2] Unregistering COM add-in...
set "REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if not exist "%REGASM%" set "REGASM=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"

if exist "%INSTALL_DIR%\AttachmentPdfConverter.dll" (
    "%REGASM%" "%INSTALL_DIR%\AttachmentPdfConverter.dll" /unregister
)

:: Remove the Outlook add-in registry key (in case it wasn't cleaned up)
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\AttachmentPdfConverter.Connect" /f >nul 2>&1

:: Remove install directory
echo [2/2] Removing files...
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"

echo.
echo ==========================================
echo  Uninstall complete!
echo ==========================================
echo.
echo Please restart Outlook if it is running.
echo.
pause
