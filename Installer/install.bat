@echo off
:: ============================================================
:: ABAS Outlook Add-in – Build & Installations-Script
:: Als Administrator ausführen!
:: ============================================================

setlocal

set INSTALL_DIR=C:\Program Files\AbasOutlookAddin
set DLL_PATH=%INSTALL_DIR%\AbasOutlookAddin.dll
set BUILD_OUTPUT=..\AbasOutlookAddin\bin\Release\AbasOutlookAddin.dll

echo === ABAS Outlook Add-in Installation ===
echo.

:: 1. Build prüfen
if not exist "%BUILD_OUTPUT%" (
    echo FEHLER: DLL nicht gefunden: %BUILD_OUTPUT%
    echo Bitte zuerst in Visual Studio im Release-Modus bauen.
    pause
    exit /b 1
)

:: 2. Installationsverzeichnis erstellen
echo [1/4] Verzeichnis erstellen...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: 3. DLL kopieren
echo [2/4] Dateien kopieren...
copy /Y "%BUILD_OUTPUT%" "%DLL_PATH%"
copy /Y "..\AbasOutlookAddin\bin\Release\AbasOutlookAddin.pdb" "%INSTALL_DIR%\" 2>nul

:: 4. COM registrieren
echo [3/4] COM registrieren...
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "%DLL_PATH%" /codebase /tlb:"%INSTALL_DIR%\AbasOutlookAddin.tlb"
if %ERRORLEVEL% neq 0 (
    echo FEHLER bei COM-Registrierung!
    pause
    exit /b 1
)

:: 5. Outlook Add-in Eintrag setzen
echo [4/4] Outlook Add-in registrieren...
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect" /v "Description" /t REG_SZ /d "ABAS Outlook Drag & Drop Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect" /v "FriendlyName" /t REG_SZ /d "ABAS Drag & Drop" /f
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f

echo.
echo === Installation abgeschlossen! ===
echo Bitte Outlook neu starten.
echo Log-Dateien: %%LOCALAPPDATA%%\AbasOutlookAddin\Logs\
echo.
pause
