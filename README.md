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
3. Standardmäßig wird die E-Mail als **`.msg`** abgelegt (die Anhänge sind darin enthalten).
4. Zum **zusätzlichen** Ablegen aller Anhänge als separate Dateien: beim Losziehen die
   **Strg-Taste** gedrückt halten (passend zur Windows-Konvention „Strg+Ziehen = Kopieren").
   Es wird dann `.msg` **+** alle Anhänge abgelegt. Signatur-Logos und andere im Text
   eingebettete Bilder werden dabei **nicht** mit abgelegt (ab v1.4.0).
5. In das ABAS-Fenster ziehen und loslassen ✓

### Anhänge direkt ins DMS ziehen (ab v1.4.0)

Ein **einzelner Anhang** lässt sich jetzt direkt aus der E-Mail ins ABAS ziehen – ohne Umweg
über die `.msg`:

1. Anhang im **Lesebereich** oder in der **geöffneten E-Mail** anklicken (Mehrfachauswahl mit
   Strg/Shift möglich)
2. Von dort aus ins ABAS-Fenster ziehen ✓

Hintergrund: Outlooks eigener Anhang-Drag liefert *virtuelle* Dateien
(`FileGroupDescriptor`/`FileContents`). Der ABAS-Client nimmt aber nur echte Dateipfade
(`CF_HDROP`) an – deshalb legt das Add-in die markierten Anhänge selbst als Temp-Dateien ab
und reicht deren Pfade weiter.

- Der Anhang-Drag startet **nur**, wenn Outlook markierte Anhänge meldet
  (`Explorer.AttachmentSelection` bzw. `Inspector.AttachmentSelection`) **und** der Mauszeiger
  beim Losziehen auch wirklich auf einem dieser Anhänge steht (ab v1.5.0, siehe unten). Sonst
  verhält sich das Add-in wie bisher und Outlook macht seinen eigenen Drag.
- Anhänge werden **nie** aus der Quell-Mail entfernt – das Verschieben-Verhalten aus v1.3.0
  greift hier bewusst nicht.
- **OLE-Objekte** (z. B. eingebettete Excel-Bereiche) lassen sich technisch nicht als Datei
  speichern und werden übersprungen (steht im Log). Eingebettete E-Mails landen als `.msg`.

#### Kein Kapern der Textmarkierung (ab v1.5.0)

Gemeldetes Verhalten: Man hängt einen Anhang an E-Mail 1, will danach in E-Mail 2 Text
markieren – der Text wird nicht markiert, stattdessen hängt plötzlich der Anhang aus E-Mail 1
an E-Mail 2. Ursache waren zwei Dinge:

1. `GetAttachmentSelection` fiel auf die **jeweils andere Quelle** zurück. Fand sich im
   angeklickten Fenster keine Anhang-Auswahl, wurde die des Lesebereichs bzw. eines anderen
   geöffneten Fensters genommen – ein Anhang aus einem fremden Fenster also.
2. Es wurde gar nicht geprüft, **worauf** der Zeiger zeigt. Jedes Ziehen außerhalb der
   Nachrichtenliste wurde zum Anhang-Drag, sobald Outlook irgendwo eine Anhang-Auswahl
   meldete – auch mitten im Fließtext.

Beides ist behoben. Die Auswahl wird nur noch aus dem Fenster geholt, in dem der Klick begann,
und der Drag startet nur, wenn unter dem Mauszeiger tatsächlich einer der markierten Anhänge
liegt. Geprüft wird das über die Barrierefreiheits-Schnittstelle (`AccessibleObjectFromPoint`)
samt übergeordneter Elemente – nötig, weil Anhangbereich und Nachrichtentext im **selben**
Fenster liegen (beide Fensterklasse `_WwG`), sich über die Fensterklasse also nicht
unterscheiden lassen. Lehnt das Add-in ab, macht Outlook wie gewohnt seinen eigenen Drag; im
Log steht eine Zeile `Anhang-Drag NICHT gestartet: …`.

Zusätzlich startet ein Drag jetzt nur noch, wenn die **Maustaste wirklich gedrückt** ist. Der
thread-lokale Hook verpasst gelegentlich ein `WM_LBUTTONUP`; danach löste die nächste
Mausbewegung einen „Phantom-Drag" mit der gesamten aktuellen Auswahl aus.

### Verschieben innerhalb Outlook (v1.3.0, abgesichert ab v1.5.0)

Wird eine E-Mail **innerhalb von Outlook** auf einen anderen Ordner gezogen, wird sie
**verschoben** statt kopiert. Technisch importiert Outlook die abgelegte `.msg` als **neues**
Element; das Original muss deshalb vom Add-in entfernt werden.

Der Drop muss dafür erkennbar ein interner Ordner-Move sein
(`ExplorerWrapper.TryCompleteInternalMove`):

- kein Strg gehalten (Strg = weiterhin kopieren),
- das Ziel-Fenster gehört zum **Outlook-Prozess selbst** (der ABAS-Client ist ein anderer
  Prozess und kann so **nie** ein Löschen auslösen),
- der Drop landete im **Outlook-Hauptfenster** (gleiches Wurzelfenster wie der Explorer) –
  ein Verfassen-/Inspector-Fenster ist ein eigenes Top-Level-Fenster und fällt heraus, sodass
  eine als **Anhang** in eine neue Mail gezogene E-Mail **nicht** gelöscht wird,
- das Ziel ist nicht die Nachrichtenliste selbst.

#### Ankunftsnachweis statt Vertrauen (ab v1.5.0)

Die vier Bedingungen oben beschreiben nur die **Geometrie** des Drops. Sie sagen nichts
darüber, ob Outlook die `.msg` auch tatsächlich irgendwo importiert hat. Bis v1.4.2 wurde
allein daraufhin gelöscht – ein Drop, der nichts importiert, entfernte die Mail ersatzlos.

Ab v1.5.0 gilt: **erst suchen, dann löschen.** `DragDropHandler.ScheduleVerifiedMove` entfernt
ein Quell-Element nur, wenn im Postfach eine Kopie existiert, die

1. **außerhalb des Quellordners** liegt,
2. **dieselbe Identität** hat – `PR_INTERNET_MESSAGE_ID`, und wenn die fehlt (bei IMAP-Konten
   regelmäßig der Fall) Betreff + Empfangszeit + Absender,
3. **neu** ist, also nach dem Beginn dieses Drags angelegt wurde.

Punkt 3 ist der entscheidende: Ohne ihn gilt jede alte Kopie derselben Mail – etwa in
„Gesendete Elemente" oder im Archiv – als Nachweis, und die Quell-Mail wäre weg, obwohl der
Drop nichts bewirkt hat. Gesucht wird in **allen** eingebundenen Postfächern (Ordner-Budget
120), damit auch ein Drop in ein anderes Konto erkannt wird. Findet sich kein Nachweis, bleibt
das Original erhalten – dann liegt die Mail eben doppelt, aber nichts ist verloren.

Die Prüfung läuft verzögert, weil Outlook den Import erst abschließen muss. Sie kostet im
Normalfall wenige Millisekunden (gemessen: 25 ms für ein Element), weil beim ersten Treffer
abgebrochen wird.

#### Schadensbegrenzung: Mengengrenze

Ein einzelner Drop darf höchstens **25** Quell-Elemente entfernen (`MaxAutoDelete`). Wird die
Grenze überschritten – etwa weil versehentlich mit Strg+A alles markiert war – bleibt
**nichts** gelöscht; im Log steht eine `SICHERHEIT:`-Zeile. Die Kopien im Zielordner liegen
dann doppelt vor, aber der Posteingang ist unversehrt.

#### Papierkorb

**Ab v1.5.0 bleibt die Quell-Mail in „Gelöschte Elemente" liegen** (wiederherstellbar). In
v1.4.2 wurde sie endgültig entfernt; damit war jedes Fehlverhalten unwiderruflich. Wer das
alte Verhalten will, setzt `PurgeFromTrash=1` – dann wird die Mail in den Papierkorb
**verschoben** und dort gelöscht (kein Suchen über die Message-ID mehr nötig, das lief bei
IMAP-Konten ohnehin ins Leere). Hat das Postfach keinen Papierkorb – typisch für IMAP –
bleibt es beim normalen Löschen.

#### Einstellungen

```cmd
:: pro Benutzer (HKCU sticht HKLM)
reg add "HKCU\Software\ABAS Outlook Addin" /v InternalMove   /t REG_DWORD /d 0  /f
reg add "HKCU\Software\ABAS Outlook Addin" /v MaxAutoDelete  /t REG_DWORD /d 25 /f
reg add "HKCU\Software\ABAS Outlook Addin" /v PurgeFromTrash /t REG_DWORD /d 0  /f

:: oder unternehmensweit
reg add "HKLM\SOFTWARE\ABAS Outlook Addin" /v InternalMove /t REG_DWORD /d 0 /f
```

| Wert | Standard | Bedeutung |
|------|----------|-----------|
| `InternalMove`   | 1  | 1 = interner Drop verschiebt, 0 = kopiert (Quell-Mail bleibt immer) |
| `MaxAutoDelete`  | 25 | Obergrenze automatisch entfernter Quell-Elemente pro Drop; 0 = nie löschen |
| `PurgeFromTrash` | 0  | 1 = Quell-Mail endgültig aus dem Papierkorb entfernen |

Alle Werte werden beim Start bzw. bei der ersten Verwendung gelesen und im Log protokolliert.

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
   `Maus-Ueberwachung installiert`. Bei einem Drag erscheint `Drag gestartet mit N Element(e)`,
   beim Ziehen eines Anhangs `Anhang-Drag gestartet mit N Anhang/Anhaengen`.
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

### Empfohlen: `Uninstall.bat` aus dem Release-Paket
Im Verteil-Paket (`Release\AbasOutlookAddin_Install.zip`) liegt neben `Install.bat`
auch **`Uninstall.bat`**. Einfach **doppelklicken** (fordert UAC selbst an). Das Skript
schließt Outlook und entfernt **alle** bekannten Installationen restlos – **beide**
Programmpfade (mit/ohne Leerzeichen) und **beide** Registry-Hives (`HKCU` **und**
`HKLM`). Damit werden auch Altlasten aus älteren Installer-Versionen (parallele 1.0.0/1.1.0-
Installationen) sauber beseitigt. Anschließend Temp-/Log-Daten gelöscht.

> **Update-Ablauf:** erst `Uninstall.bat`, dann `Install.bat` ausführen – so ist
> garantiert nur die neue Version registriert und kein alter Eintrag bleibt aktiv.

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
