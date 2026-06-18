@echo off
:: ============================================================
:: ABAS Outlook Add-in – WiX Build Script
:: Erstellt MSI + EXE-Bundle (mit .NET-Prerequisite)
::
:: Voraussetzung: WiX Toolset 3.11 installiert
::   https://github.com/wixtoolset/wix3/releases
::
:: Aufruf: build_msi.bat [Release|Debug]
:: ============================================================

setlocal
set CONFIG=%1
if "%CONFIG%"=="" set CONFIG=Release

:: WiX Toolset Pfad (Standard-Installation)
set WIX_BIN=C:\Program Files (x86)\WiX Toolset v3.11\bin
if not exist "%WIX_BIN%\candle.exe" (
    echo FEHLER: WiX Toolset nicht gefunden unter: %WIX_BIN%
    echo Download: https://github.com/wixtoolset/wix3/releases
    pause
    exit /b 1
)

:: Sicherstellen dass DLL gebaut wurde
set DLL=..\AbasOutlookAddin\bin\%CONFIG%\AbasOutlookAddin.dll
if not exist "%DLL%" (
    echo FEHLER: DLL nicht gefunden: %DLL%
    echo Bitte erst in Visual Studio im %CONFIG%-Modus bauen.
    pause
    exit /b 1
)

echo === ABAS Outlook Add-in – MSI Build ===
echo Konfiguration: %CONFIG%
echo.

:: Ausgabeverzeichnis
if not exist "out" mkdir out

:: -------------------------------------------------------
:: Schritt 1: MSI kompilieren
:: -------------------------------------------------------
echo [1/4] Kompiliere Product.wxs...
"%WIX_BIN%\candle.exe" ^
    Product.wxs ^
    -ext WixUtilExtension ^
    -out out\Product.wixobj ^
    -dConfiguration=%CONFIG%

if %ERRORLEVEL% neq 0 (
    echo FEHLER beim Kompilieren von Product.wxs
    pause
    exit /b 1
)

:: -------------------------------------------------------
:: Schritt 2: MSI linken
:: -------------------------------------------------------
echo [2/4] Erstelle MSI...
"%WIX_BIN%\light.exe" ^
    out\Product.wixobj ^
    -ext WixUIExtension ^
    -ext WixUtilExtension ^
    -cultures:de-DE ^
    -loc de-DE.wxl ^
    -out out\AbasOutlookAddin.msi ^
    -sw1076

if %ERRORLEVEL% neq 0 (
    echo FEHLER beim Erstellen der MSI
    pause
    exit /b 1
)

echo     ✓ MSI erstellt: out\AbasOutlookAddin.msi

:: -------------------------------------------------------
:: Schritt 3: EXE-Bundle kompilieren (mit .NET Check)
:: -------------------------------------------------------
echo [3/4] Kompiliere Bundle.wxs...
"%WIX_BIN%\candle.exe" ^
    Bundle.wxs ^
    -ext WixBalExtension ^
    -ext WixNetFxExtension ^
    -out out\Bundle.wixobj

if %ERRORLEVEL% neq 0 (
    echo FEHLER beim Kompilieren von Bundle.wxs
    echo (Bundle wird uebersprungen, MSI ist trotzdem verwendbar)
    goto :sign
)

:: -------------------------------------------------------
:: Schritt 4: EXE-Bundle linken
:: -------------------------------------------------------
echo [4/4] Erstelle Setup-EXE...
"%WIX_BIN%\light.exe" ^
    out\Bundle.wixobj ^
    -ext WixBalExtension ^
    -ext WixNetFxExtension ^
    -out out\AbasOutlookAddinSetup.exe

if %ERRORLEVEL% neq 0 (
    echo FEHLER beim Erstellen der Setup-EXE
    echo (MSI ist trotzdem verwendbar)
    goto :sign
)

echo     ✓ Setup-EXE erstellt: out\AbasOutlookAddinSetup.exe

:: -------------------------------------------------------
:sign
:: Optional: Signieren (auskommentiert – Zertifikat anpassen)
:: -------------------------------------------------------
echo.
echo --- Optionale Code-Signierung ---
echo Zum Signieren (verhindert SmartScreen-Warnung):
echo.
echo   signtool.exe sign /f IhrZertifikat.pfx /p IhrPasswort
echo               /fd sha256 /tr http://timestamp.digicert.com /td sha256
echo               out\AbasOutlookAddin.msi
echo               out\AbasOutlookAddinSetup.exe
echo.

:: Ergebnis
echo ============================================
echo BUILD ERFOLGREICH
echo ============================================
echo.
echo Ausgabe-Dateien:
if exist "out\AbasOutlookAddinSetup.exe" (
    echo   Setup-EXE (empfohlen): out\AbasOutlookAddinSetup.exe
)
echo   MSI (GPO/SCCM):         out\AbasOutlookAddin.msi
echo.
echo GPO Silent-Rollout:
echo   msiexec.exe /i AbasOutlookAddin.msi /qn /log install.log
echo.
pause
