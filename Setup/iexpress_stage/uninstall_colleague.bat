@echo off
setlocal enableextensions

:: --- Adminrechte sicherstellen (sonst neu starten mit Elevation) ---
net session >nul 2>&1
if %errorlevel%==0 goto :uninstall

echo Diese Deinstallation benoetigt Administratorrechte.
echo Es erscheint gleich eine Sicherheitsabfrage (UAC) - bitte mit "Ja" bestaetigen.
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b 0

:uninstall
set "INSTALL_DIR=%ProgramFiles%\ABAS Outlook Addin"
set "REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
set "CLSID={A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
set "PROGID=AbasOutlookAddin.Connect"

echo ============================================
echo   ABAS Outlook Drag ^& Drop Add-in - Deinstallation
echo ============================================
echo.

echo [1/5] Outlook schliessen (falls offen)...
taskkill /IM OUTLOOK.EXE >nul 2>&1
timeout /t 3 /nobreak >nul
taskkill /IM OUTLOOK.EXE /F >nul 2>&1

echo [2/5] COM-Registrierung aufheben...
if exist "%INSTALL_DIR%\AbasOutlookAddin.dll" (
    "%REGASM%" "%INSTALL_DIR%\AbasOutlookAddin.dll" /unregister /nologo
)
:: CLSID/ProgID zusaetzlich direkt entfernen (maschinen- UND benutzerweit)
reg delete "HKCR\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKCR\%PROGID%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\%PROGID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Classes\CLSID\%CLSID%" /f >nul 2>&1

echo [3/5] Outlook-Add-in-Eintraege entfernen (HKCU + HKLM + 32-bit)...
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\%PROGID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\%PROGID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\%PROGID%" /f >nul 2>&1

echo [4/5] Resilienz-/Deaktivierungslisten bereinigen...
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "%PROGID%" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "%PROGID%" /f >nul 2>&1

echo [5/5] Programmdateien und Benutzerdaten loeschen...
if exist "%INSTALL_DIR%" rd /s /q "%INSTALL_DIR%"
if exist "%LOCALAPPDATA%\AbasOutlookAddin" rd /s /q "%LOCALAPPDATA%\AbasOutlookAddin"

echo.
echo ============================================
echo   Deinstallation abgeschlossen.
echo   WICHTIG: Outlook war geschlossen - bitte
echo   erst jetzt wieder starten.
echo ============================================
echo.
pause
