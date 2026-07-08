@echo off
setlocal enableextensions
set "STAGE=%TEMP%\AbasOutlookAddinSetup"

:: --- Adminrechte sicherstellen (sonst neu starten mit Elevation) ---
net session >nul 2>&1
if %errorlevel%==0 goto :install

echo Diese Installation benoetigt Administratorrechte.
echo Es erscheint gleich eine Sicherheitsabfrage (UAC) - bitte mit "Ja" bestaetigen.
if not exist "%STAGE%" mkdir "%STAGE%"
copy /Y "%~dp0AbasOutlookAddin.dll" "%STAGE%\AbasOutlookAddin.dll" >nul
copy /Y "%~f0" "%STAGE%\install_colleague.bat" >nul
powershell -NoProfile -Command "Start-Process -FilePath '%STAGE%\install_colleague.bat' -Verb RunAs"
exit /b 0

:install
set "INSTALL_DIR=%ProgramFiles%\ABAS Outlook Addin"
set "SRC=%~dp0AbasOutlookAddin.dll"
set "REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

echo ============================================
echo   ABAS Outlook Drag ^& Drop Add-in - Setup
echo ============================================
echo.

echo [1/5] Outlook schliessen (falls offen)...
taskkill /IM OUTLOOK.EXE >nul 2>&1
timeout /t 3 /nobreak >nul
taskkill /IM OUTLOOK.EXE /F >nul 2>&1

echo [2/5] Dateien kopieren nach "%INSTALL_DIR%"...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /Y "%SRC%" "%INSTALL_DIR%\AbasOutlookAddin.dll" >nul
if not exist "%INSTALL_DIR%\AbasOutlookAddin.dll" (
    echo FEHLER: DLL konnte nicht kopiert werden.
    pause & exit /b 1
)

echo [3/5] COM registrieren...
"%REGASM%" "%INSTALL_DIR%\AbasOutlookAddin.dll" /codebase /tlb:"%INSTALL_DIR%\AbasOutlookAddin.tlb" /nologo
if %errorlevel% neq 0 (
    echo FEHLER bei der COM-Registrierung.
    pause & exit /b 1
)

echo [4/5] Outlook-Add-in registrieren (alle Benutzer)...
set "KEY=HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect"
reg add "%KEY%" /v Description  /t REG_SZ    /d "ABAS Outlook Drag & Drop Add-in" /f >nul
reg add "%KEY%" /v FriendlyName /t REG_SZ    /d "ABAS Drag & Drop" /f >nul
reg add "%KEY%" /v LoadBehavior /t REG_DWORD /d 3 /f >nul

echo [5/5] Resilienz-Eintrag setzen (Add-in bleibt aktiv)...
reg add "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "AbasOutlookAddin.Connect" /t REG_DWORD /d 1 /f >nul

echo.
echo ============================================
echo   Installation erfolgreich abgeschlossen!
echo   Bitte Outlook (klassisch) starten.
echo ============================================
echo.

:: Staging aufraeumen (nur vorhanden bei Elevation-Neustart)
rd /s /q "%STAGE%" >nul 2>&1
pause
