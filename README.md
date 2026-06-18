# ABAS Outlook Drag & Drop Add-in

Ermöglicht das Drag & Drop von Outlook-Elementen (E-Mails, Anhänge, Kontakte, Termine)
direkt in den **ABAS Windows-Client** – ohne API-Hooking, ohne globale Systemeingriffe.

---

## Funktionsweise

```
Outlook (Selektion)
      │
      ▼
  Maus-Drag erkannt (thread-lokaler WH_MOUSE-Hook auf dem Outlook-UI-Thread)
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

**Kein globales API-Hooking, keine DLL-Injection** – im Gegensatz zu OutlookFileDrag
(EasyHook) wird nichts in fremde Prozesse injiziert. Die Drag-Erkennung läuft über einen
**thread-lokalen `WH_MOUSE`-Hook** (`SetWindowsHookEx` mit der Thread-ID des Outlook-UI-Threads).
Dieser Hook gilt ausschließlich für den Outlook-eigenen Thread, sieht dadurch zuverlässig auch
die Maus-Events der Kind-Fenster (E-Mail-Liste) und kommt ohne fest verdrahtete Fensterklassen aus.

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

### 2. Bauen
**Visual Studio:** `AbasOutlookAddin\AbasOutlookAddin.csproj` öffnen und im **Release**-Modus bauen.

**Kommandozeile (MSBuild):**
```cmd
msbuild AbasOutlookAddin\AbasOutlookAddin.csproj /t:Rebuild /p:Configuration=Release /p:RegisterForCOMInterop=false
```
> Das Projekt nutzt `LangVersion 8.0`. Die Verweise auf die Outlook-PIA und `Extensibility`
> werden portabel aus dem **GAC** aufgelöst – kein maschinenspezifischer Pfad nötig, aber die
> Office-PIAs müssen installiert sein (VS-Workload „Office/SharePoint-Entwicklung").
> `RegisterForCOMInterop=false` vermeidet die (Admin-pflichtige) COM-Registrierung beim Bauen –
> diese erledigt erst die Installation.

### 3. Signieren (empfohlen für Unternehmenseinsatz)
```cmd
signtool.exe sign /f IhrZertifikat.pfx /p Passwort /t http://timestamp.digicert.com ^
    bin\Release\AbasOutlookAddin.dll
```
Ein Code-Signing-Zertifikat verhindert Warnmeldungen beim Installieren.

---

## Installation

### Empfohlen: Setup.exe (zum Weitergeben an Endanwender)
Eine selbstentpackende Installationsdatei liegt nach dem Build unter
`Setup\out\AbasOutlookAddinSetup.exe`. Sie kann z. B. per E-Mail verteilt werden.

1. `AbasOutlookAddinSetup.exe` per **Rechtsklick → Als Administrator ausführen**
   (sie fordert die Adminrechte sonst selbst via UAC an).
2. Das Setup schließt Outlook, kopiert die DLL nach `C:\Program Files\ABAS Outlook Addin\`,
   registriert COM (`RegAsm /codebase`) und trägt das Add-in unter `HKLM` ein (für alle Benutzer).
3. **Outlook (klassisch) starten** – das Add-in lädt automatisch.

> Die EXE ist **nicht code-signiert**. Beim Start erscheint daher ggf. eine SmartScreen-Warnung
> („Weitere Informationen" → „Trotzdem ausführen"). Für den produktiven Rollout ein
> Code-Signing-Zertifikat verwenden (siehe unten). Neu bauen lässt sich die EXE über
> `Setup\iexpress_stage\setup.sed` mit dem Windows-Bordmittel `iexpress`.

### Einzelplatz (manuell)
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

> **Hinweis:** Das Add-in hat **keine sichtbare Oberfläche** (kein Menüband-Button, kein Symbol).
> Es arbeitet unsichtbar im Hintergrund und reagiert nur auf das Ziehen mit der Maus.

---

## Add-in prüfen (ist es aktiv?)

Da es keine sichtbare UI gibt, lässt sich der Status so kontrollieren:

1. **COM-Add-Ins-Liste:** Outlook → *Datei → Optionen → Add-Ins*. Unten bei *Verwalten:*
   **COM-Add-Ins** auswählen → *Gehe zu…*. Der Eintrag **„ABAS Drag & Drop"** muss
   **angehakt** sein. Steht er unter *Deaktivierte Anwendungs-Add-Ins*, wieder aktivieren.
2. **Log-Datei** (sicherster Nachweis): `%LOCALAPPDATA%\AbasOutlookAddin\Logs\addin_JJJJMMTT.log`.
   Beim Start steht dort `ABAS Outlook Add-in erfolgreich geladen.` und
   `Maus-Ueberwachung installiert`. Bei einem Drag erscheint `Drag gestartet mit N Element(e)`.
3. **Funktionstest ohne ABAS:** `Test\AbasDropTest.exe` starten (akzeptiert nur echte
   Dateipfade/CF_HDROP) und eine E-Mail hineinziehen – erscheint der Dateipfad, funktioniert alles.

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

## Sicherheits-Audit (2026-06-18)

Code-Review des gesamten Add-ins. **Keine kritischen oder hohen Befunde.** Bereits umgesetzte
Schutzmaßnahmen:

| Schutz | Umsetzung |
|--------|-----------|
| Path-Traversal | `SanitizeFileName` entfernt `..`, `/`, `\`, ungültige Zeichen + reservierte Namen (CON, PRN …) |
| TOCTOU / Symlinks | `ValidateTempFilePath` prüft kanonischen Pfad + blockiert Reparse-Points; erneute Prüfung vor dem Löschen |
| Datenisolation | Temp-Verzeichnis mit restriktiver ACL (nur aktueller Benutzer), Crash-Recovery-Cleanup |
| Log-Injection | `SanitizeLogMessage` entfernt Steuerzeichen/Newlines, begrenzt Länge |
| Ressourcen | Temp-Limit (200) + verzögertes Cleanup (30 s) nach jedem Drag |
| Hooking | Thread-**lokaler** `WH_MOUSE`-Hook (kein globaler Hook, keine Injection), sauberes `UnhookWindowsHookEx` beim Entladen |

**Empfehlungen (niedrige Priorität):**
- **Code-Signing:** DLL **und** `Setup.exe` mit einem Zertifikat signieren → keine SmartScreen-/AV-Warnungen, Manipulationsschutz.
- **Strong-Name-Pinning:** `VerifyAssemblyIntegrity` prüft nur, *dass* ein Strong Name vorhanden ist, nicht *welcher*. Optional den erwarteten Public-Key-Token fest hinterlegen (bindet die DLL an einen festen Signaturschlüssel).
- **Hinweis Datenschutz:** E-Mail-Inhalte liegen für max. 30 s als Datei im (ACL-geschützten) Temp-Verzeichnis.

---

## Deinstallation

### Empfohlen: Skript (macht die Setup.exe-Installation rückgängig)
`Setup\iexpress_stage\uninstall_colleague.bat` per **Rechtsklick → Als Administrator
ausführen** (fordert UAC sonst selbst an). Es schließt Outlook, hebt die COM-Registrierung
auf, entfernt die `HKLM`-Einträge und löscht den Programmordner sowie die
benutzerspezifischen Temp-/Log-Daten.

### Manuell (als Administrator)
```cmd
:: COM-Registrierung entfernen
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" ^
    "%ProgramFiles%\ABAS Outlook Addin\AbasOutlookAddin.dll" /unregister

:: Outlook-Einträge entfernen (HKLM, da Setup für alle Benutzer installiert)
reg delete "HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "AbasOutlookAddin.Connect" /f

:: Dateien löschen
rmdir /s /q "%ProgramFiles%\ABAS Outlook Addin"
rmdir /s /q "%LOCALAPPDATA%\AbasOutlookAddin"
```

---

## Lizenz

MIT – frei verwendbar, anpassbar und intern weitergabefähig.
