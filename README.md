# ABAS Outlook Drag & Drop Add-in

Ermöglicht das Drag & Drop von Outlook-Elementen (E-Mails, Anhänge, Kontakte, Termine)
direkt in den **ABAS Windows-Client** – ohne API-Hooking, ohne globale Systemeingriffe.

---

## Funktionsweise

```
Outlook (Selektion)
      │
      ▼
  Mausbewegung erkannt (NativeWindow Subclassing)
      │
      ▼
  Element als Datei extrahiert (temp, benutzerspezifisch)
      │
      ▼
  Standard Windows CF_HDROP DataObject erstellt
      │
      ▼
  OLE DoDragDrop() → ABAS empfängt echte Dateipfade
      │
      ▼
  Temp-Dateien nach 30s sicher gelöscht
```

**Kein globales API-Hooking** – im Gegensatz zu OutlookFileDrag wird kein DLL-Injection
verwendet. Stattdessen: NativeWindow-Subclassing nur auf dem Outlook-Prozess.

---

## Unterstützte Elemente

| Element       | Wird gespeichert als |
|--------------|---------------------|
| E-Mail        | .msg                |
| Anhang        | Originalformat      |
| Kontakt       | .vcf                |
| Termin        | .ics                |
| Aufgabe       | .msg                |

---

## Voraussetzungen

- **Windows** 10 oder 11 (64-bit)
- **Outlook** 2016, 2019 oder Microsoft 365 (klassisch, nicht neue Outlook-App)
- **.NET Framework** 4.7.2 (auf Windows 10/11 vorinstalliert)
- **Visual Studio** 2019/2022 mit "Office/SharePoint development" Workload (zum Bauen)

---

## Build

### 1. Strong Name Key erstellen (einmalig)
```cmd
sn.exe -k AbasOutlookAddin\AbasOutlookAddin.snk
```

### 2. In Visual Studio öffnen
`AbasOutlookAddin\AbasOutlookAddin.csproj` öffnen und im **Release**-Modus bauen.

### 3. Signieren (empfohlen für Unternehmenseinsatz)
```cmd
signtool.exe sign /f IhrZertifikat.pfx /p Passwort /t http://timestamp.digicert.com ^
    bin\Release\AbasOutlookAddin.dll
```
Ein Code-Signing-Zertifikat verhindert Warnmeldungen beim Installieren.

---

## Installation

### Einzelplatz
```cmd
cd Installer
install.bat          (Als Administrator ausführen)
```
Danach **Outlook neu starten**.

### Massenrollout via GPO
1. DLL auf Netzlaufwerk oder per SCCM/Intune verteilen
2. Registry-Key setzen (HKLM statt HKCU):
```
HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect
LoadBehavior = 3
```
3. COM via RegAsm auf jedem Client registrieren (Startup-Script)

### Silent-Installation via MSI (empfohlen)
Das Projekt kann als WiX-Installer verpackt werden:
```cmd
candle.exe Setup.wxs
light.exe Setup.wixobj -o AbasOutlookAddin.msi
msiexec.exe /i AbasOutlookAddin.msi /qn
```

---

## Verwendung

1. Outlook öffnen
2. E-Mail oder Element in der Liste anklicken und **gedrückt halten**
3. Bei E-Mails mit Anhängen erscheint ein Dialog:
   - **Ja** = E-Mail als .msg ablegen
   - **Nein** = Anhänge ablegen
4. In das ABAS-Fenster ziehen und loslassen ✓

---

## Sicherheit

| Aspekt                    | Diese Lösung         | OutlookFileDrag      |
|--------------------------|---------------------|---------------------|
| API-Hooking              | ❌ Nein              | ✅ Ja (EasyHook)     |
| Signierung               | ✅ Möglich (SNK/PFX) | ⚠️ Selbstsigniert    |
| Temp-Pfad                | Benutzerspezifisch   | System-Temp          |
| Quellcode prüfbar        | ✅ Vollständig        | ✅ Open Source        |
| GPO-rolloutfähig         | ✅ Ja                 | ✅ Ja                 |
| AV-Fehlalarme            | Unwahrscheinlich     | Möglich (Hooking)    |
| Aktiver Support          | Ihr Code             | Seit 2018 eingestellt|

**Temp-Verzeichnis:** `%LOCALAPPDATA%\AbasOutlookAddin\Temp\` (nur aktueller Benutzer)  
**Log-Verzeichnis:** `%LOCALAPPDATA%\AbasOutlookAddin\Logs\`

---

## Deinstallation

```cmd
:: COM-Registrierung entfernen
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" ^
    "C:\Program Files\AbasOutlookAddin\AbasOutlookAddin.dll" /unregister

:: Outlook-Eintrag entfernen
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect" /f

:: Dateien löschen
rmdir /s /q "C:\Program Files\AbasOutlookAddin"
```

---

## Lizenz

MIT – frei verwendbar, anpassbar und intern weitergabefähig.
