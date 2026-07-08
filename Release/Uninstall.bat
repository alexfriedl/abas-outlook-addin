@echo off
setlocal enableextensions
set "STAGE=%TEMP%\AbasOutlookAddinUninstall"

:: ============================================================
:: ABAS Outlook Drag ^& Drop Add-in - Deinstallation
:: Entfernt ALLE bekannten Installationen restlos:
::   - beide Programmpfade (mit/ohne Leerzeichen)
::   - HKCU- UND HKLM-Registrierung (alte und neue Installer)
:: Fordert Administratorrechte bei Bedarf selbst an (UAC).
:: ============================================================

:: --- Adminrechte sicherstellen (sonst neu starten mit Elevation) ---
net session >nul 2>&1
if %errorlevel%==0 goto :uninstall

echo Diese Deinstallation benoetigt Administratorrechte.
echo Es erscheint gleich eine Sicherheitsabfrage (UAC) - bitte mit "Ja" bestaetigen.
if not exist "%STAGE%" mkdir "%STAGE%"
copy /Y "%~f0" "%STAGE%\Uninstall.bat" >nul
powershell -NoProfile -Command "Start-Process -FilePath '%STAGE%\Uninstall.bat' -Verb RunAs"
exit /b 0

:uninstall
set "NEWDIR=%ProgramFiles%\ABAS Outlook Addin"
set "OLDDIR=%ProgramFiles%\AbasOutlookAddin"
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

echo [2/5] COM-Registrierung aufheben (beide Pfade)...
if exist "%NEWDIR%\AbasOutlookAddin.dll" "%REGASM%" "%NEWDIR%\AbasOutlookAddin.dll" /unregister /nologo
if exist "%OLDDIR%\AbasOutlookAddin.dll" "%REGASM%" "%OLDDIR%\AbasOutlookAddin.dll" /unregister /nologo
:: CLSID/ProgID zusaetzlich direkt entfernen (maschinen- UND benutzerweit)
reg delete "HKCR\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKCR\%PROGID%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\%PROGID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Classes\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Classes\WOW6432Node\CLSID\%CLSID%" /f >nul 2>&1

echo [3/5] Outlook-Add-in-Eintraege entfernen (HKCU + HKLM + 32-bit)...
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\%PROGID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\%PROGID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\%PROGID%" /f >nul 2>&1

echo [4/5] Resilienz-/Deaktivierungslisten bereinigen...
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "%PROGID%" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "%PROGID%" /f >nul 2>&1

echo [5/5] Programmdateien und Benutzerdaten loeschen...
if exist "%NEWDIR%" rd /s /q "%NEWDIR%"
if exist "%OLDDIR%" rd /s /q "%OLDDIR%"
if exist "%LOCALAPPDATA%\AbasOutlookAddin" rd /s /q "%LOCALAPPDATA%\AbasOutlookAddin"
if exist "%STAGE%" rd /s /q "%STAGE%" 2>nul

echo.
echo ============================================
echo   Deinstallation abgeschlossen.
echo   Zum Aktualisieren jetzt Install.bat ausfuehren.
echo   Outlook war geschlossen - bitte erst danach starten.
echo ============================================
echo.
pause
