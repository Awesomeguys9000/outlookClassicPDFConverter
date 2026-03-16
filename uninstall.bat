@echo off
:: ============================================================
:: Attachment PDF Converter - Per-User Uninstaller (No Admin)
:: ============================================================

echo.
echo ==========================================
echo  Attachment PDF Converter - Uninstaller
echo ==========================================
echo.

set "INSTALL_DIR=%LocalAppData%\AttachmentPdfConverter"
set "CLSID={F1A2B3C4-D5E6-4F78-9A0B-C1D2E3F4A5B6}"

:: Remove Outlook add-in registration
echo [1/3] Removing Outlook add-in registration...
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\AttachmentPdfConverter.Connect" /f >nul 2>&1

:: Remove COM registration
echo [2/3] Removing COM registration...
reg delete "HKCU\Software\Classes\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\AttachmentPdfConverter.Connect" /f >nul 2>&1

:: Remove installed files
echo [3/3] Removing files...
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"

echo.
echo ==========================================
echo  Uninstall complete!
echo ==========================================
echo.
echo Please restart Outlook if it is running.
echo.
pause
