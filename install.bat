@echo off
:: ============================================================
:: Attachment PDF Converter - Per-User Installer (No Admin)
:: ============================================================

echo.
echo ==========================================
echo  Attachment PDF Converter - Installer
echo ==========================================
echo.

:: Set install directory under user's AppData
set "INSTALL_DIR=%LocalAppData%\AttachmentPdfConverter"

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

:: Register COM class under HKCU (no admin needed)
echo [3/3] Registering COM add-in (current user)...

set "CLSID={F1A2B3C4-D5E6-4F78-9A0B-C1D2E3F4A5B6}"
set "DLL_PATH=%INSTALL_DIR%\AttachmentPdfConverter.dll"

:: Register the COM class under HKCU\Software\Classes\CLSID
reg add "HKCU\Software\Classes\CLSID\%CLSID%" /ve /d "AttachmentPdfConverter.Connect" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /ve /d "mscoree.dll" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "ThreadingModel" /d "Both" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "Class" /d "AttachmentPdfConverter.Connect" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "Assembly" /d "AttachmentPdfConverter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "CodeBase" /d "%DLL_PATH%" /f >nul

:: Register the ProgId
reg add "HKCU\Software\Classes\AttachmentPdfConverter.Connect" /ve /d "AttachmentPdfConverter.Connect" /f >nul
reg add "HKCU\Software\Classes\AttachmentPdfConverter.Connect\CLSID" /ve /d "%CLSID%" /f >nul

:: Register as an Outlook add-in
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AttachmentPdfConverter.Connect" /v "FriendlyName" /d "Attachment PDF Converter" /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AttachmentPdfConverter.Connect" /v "Description" /d "Converts email attachments to PDF using Microsoft Print to PDF" /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AttachmentPdfConverter.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul

echo.
echo ==========================================
echo  Installation complete!
echo ==========================================
echo.
echo Installed to: %INSTALL_DIR%
echo.
echo Please:
echo   1. Close Outlook if it is open
echo   2. Re-open Outlook
echo   3. Compose a new email to see the
echo      "Convert to PDF" button on the ribbon
echo.
pause
