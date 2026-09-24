# ABAS Outlook Drag & Drop Add-in

Ermöglicht das Drag & Drop von Outlook-Elementen (E-Mails, Anhänge, Kontakte, Termine)
direkt in den **ABAS Windows-Client** – ohne API-Hooking, ohne globale Systemeingriffe.

---

## Was man damit machen kann

Das Add-in hat genau einen Zweck: **Mails und Anhänge aus Outlook ins ABAS bekommen.** Ohne
das Add-in geht das nicht, weil ABAS nur echte Dateien annimmt und Outlook beim Ziehen keine
echte Datei herausgibt (technischer Hintergrund unter [Funktionsweise](#funktionsweise)).

> **Ab v1.6.0 wird mit der RECHTEN Maustaste gezogen.** Die linke Maustaste gehört wieder
> vollständig Outlook – Verschieben, Text markieren, Outlooks eigener Anhang-Drag laufen
> unverändert, das Add-in bekommt davon gar nichts mit.

**1. Eine Mail ins ABAS ziehen**  
Mail in der Liste anklicken, dann mit **gedrückter rechter Maustaste** ins ABAS ziehen.
Dort landet die komplette Mail als Datei, also mit Absender, Betreff, Text und allem, was
drin hängt.

**2. Mehrere Mails auf einmal ziehen**  
Mit Strg oder Shift mehrere markieren und zusammen rüberziehen (ebenfalls rechte Taste).

**3. Mail plus Anhänge einzeln ablegen**  
Beim Ziehen zusätzlich die **Strg-Taste** gedrückt halten. Dann kommt nicht nur die Mail an,
sondern zusätzlich jeder Anhang als eigene Datei. Praktisch, wenn im ABAS die Rechnung als
PDF gebraucht wird und nicht nur die Mail drumherum.

**4. Nur einen einzelnen Anhang ablegen**  
Den Anhang direkt anklicken – im Lesebereich oder in der geöffneten Mail – und von dort mit
der rechten Maustaste ins ABAS ziehen. Dann kommt genau diese eine Datei an, ohne die Mail
drumherum.

Kontakte, Termine und Aufgaben lassen sich genauso ziehen, siehe
[Unterstützte Elemente](#unterstützte-elemente).

Innerhalb von Outlook ändert sich nichts: Eine Mail mit der **linken** Maustaste auf einen
anderen Ordner gezogen wird von Outlook selbst verschoben (siehe
[Verschieben innerhalb Outlook](#verschieben-innerhalb-outlook-ab-v160-wieder-outlook-selbst)).
Ein **Rechtsklick ohne Ziehen** öffnet wie gewohnt das Kontextmenü.

Das Add-in hat **keine Oberfläche** – kein Button, kein Symbol im Menüband. Es arbeitet
unsichtbar mit und reagiert nur auf das Ziehen mit der rechten Maustaste.

---

## Funktionsweise

```
Outlook (Selektion)
      │
      ▼
  Drag mit der RECHTEN Maustaste erkannt (thread-lokaler WH_MOUSE-Hook,
  WM_RBUTTONDOWN/-UP; die linke Taste wird nicht angefasst)
      │
      ▼
  Element als Datei extrahiert (temp, benutzerspezifisch)
      │
      ▼
  Standard Windows CF_HDROP DataObject erstellt
      │
      ▼
  OLE DoDragDrop() mit eigenem QueryContinueDrag an der rechten Taste
  → ABAS empfängt echte Dateipfade
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
2. E-Mail oder Element in der Liste anklicken (linke Taste, wie immer)
3. Mit der **rechten Maustaste** darauf drücken, gedrückt halten und losziehen
4. Standardmäßig wird die E-Mail als **`.msg`** abgelegt (die Anhänge sind darin enthalten).
5. Zum **zusätzlichen** Ablegen aller Anhänge als separate Dateien: beim Losziehen die
   **Strg-Taste** gedrückt halten (passend zur Windows-Konvention „Strg+Ziehen = Kopieren").
   Es wird dann `.msg` **+** alle Anhänge abgelegt. Signatur-Logos und andere im Text
   eingebettete Bilder werden dabei **nicht** mit abgelegt (ab v1.4.0).
6. In das ABAS-Fenster ziehen und loslassen ✓

> Wird die rechte Maustaste nur **geklickt**, ohne den Zeiger zu bewegen, passiert nichts
> Besonderes – Outlook öffnet sein Kontextmenü. Das Add-in wird erst aktiv, wenn bei
> gedrückter rechter Taste die Ziehen-Schwelle von Windows überschritten wird.

### Anhänge direkt ins DMS ziehen (ab v1.4.0)

Ein **einzelner Anhang** lässt sich jetzt direkt aus der E-Mail ins ABAS ziehen – ohne Umweg
über die `.msg`:

1. Anhang im **Lesebereich** oder in der **geöffneten E-Mail** anklicken (Mehrfachauswahl mit
   Strg/Shift möglich)
2. Von dort aus mit der **rechten Maustaste** ins ABAS-Fenster ziehen ✓

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
thread-lokale Hook verpasst gelegentlich ein `WM_RBUTTONUP`; danach löste die nächste
Mausbewegung einen „Phantom-Drag" mit der gesamten aktuellen Auswahl aus.

### Verschieben innerhalb Outlook (ab v1.6.0 wieder Outlook selbst)

**Das Add-in fasst das Verschieben nicht mehr an.** Wer eine Mail mit der **linken**
Maustaste auf einen anderen Ordner zieht, löst Outlooks eigenes Verschieben aus – genau
wie ohne Add-in, ohne Umweg über eine `.msg` und ohne Verzögerung.

Damit ist auch die gesamte Löschmechanik entfallen, die es von v1.3.0 bis v1.5.0 gab:
Das Add-in legte die Mail als `.msg` ab, Outlook importierte sie als **neues** Element, und
das Original musste anschließend entfernt werden. Das brachte zwei Fehlerbilder mit sich:

- **Kopie statt Verschieben.** Der Maus-Hook ist thread-lokal und griff in *allen*
  Outlook-Fenstern, die Abschlussprüfung verglich den Drop aber gegen das *eine* beim Laden
  gebundene Explorer-Fenster. Mit einem zweiten geöffneten Fenster wurde das Ziehen also
  übernommen, der Abschluss aber verworfen – die Mail lag danach doppelt vor.
- **Unbegrenzte Verzögerung.** Das Entfernen des Originals lief verzögert über Outlooks
  UI-Thread (`BeginInvoke`). War der beschäftigt, wartete die Arbeit: gemessen 33 Sekunden
  bis mehrere Minuten statt der geplanten 2,5 Sekunden.

**Das Add-in entfernt ab v1.6.0 unter keinen Umständen mehr ein Element.** Die Einstellungen
`InternalMove`, `MaxAutoDelete` und `PurgeFromTrash` sind damit ersatzlos entfallen – ebenso
der Registry-Schlüssel `Software\ABAS Outlook Addin`, den das Add-in nicht mehr liest.

> **Hinweis:** Das Add-in hat **keine sichtbare Oberfläche** (kein Menüband-Button, kein Symbol).
> Es arbeitet unsichtbar im Hintergrund und reagiert nur auf das Ziehen mit der **rechten** Maustaste.

---

## Add-in prüfen (ist es aktiv?)

Da es keine sichtbare UI gibt, lässt sich der Status so kontrollieren:

1. **COM-Add-Ins-Liste:** Outlook → *Datei → Optionen → Add-Ins*. Unten bei *Verwalten:*
   **COM-Add-Ins** auswählen → *Gehe zu…*. Der Eintrag **„ABAS Drag & Drop"** muss
   **angehakt** sein. Steht er unter *Deaktivierte Anwendungs-Add-Ins*, wieder aktivieren.
2. **Log-Datei** (sicherster Nachweis): `%LOCALAPPDATA%\AbasOutlookAddin\Logs\addin_JJJJMMTT.log`.
   Beim Start steht dort `ABAS Outlook Add-in erfolgreich geladen (Ziehen mit der RECHTEN
   Maustaste; die linke bleibt bei Outlook).` und
   `Maus-Ueberwachung installiert`. Bei einem Drag erscheint `Drag gestartet mit N Element(e)`,
   beim Ziehen eines Anhangs `Anhang-Drag gestartet mit N Anhang/Anhaengen`.
3. **Funktionstest ohne ABAS:** `Test\AbasDropTest.exe` starten (akzeptiert nur echte
   Dateipfade/CF_HDROP) und eine E-Mail mit der rechten Maustaste hineinziehen – erscheint
   der Dateipfad, funktioniert alles.

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
