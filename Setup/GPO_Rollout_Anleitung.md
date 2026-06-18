# GPO-Rollout Anleitung – ABAS Outlook Add-in

Schritt-für-Schritt Anleitung für den unternehmensweiten Rollout via
**Active Directory Gruppenrichtlinie (GPO)**.

---

## Voraussetzungen

- Active Directory Domäne
- WiX-Build erfolgreich durchgeführt (`out\AbasOutlookAddin.msi`)
- MSI ggf. mit Code-Signing-Zertifikat signiert
- Netzwerkfreigabe für die MSI-Datei

---

## Schritt 1: MSI auf Netzwerkfreigabe bereitstellen

```
\\IhrServer\Software\AbasOutlookAddin\AbasOutlookAddin.msi
```

Berechtigungen auf die Freigabe:
- **Domänen-Computer**: Lesen ✓
- **Domänen-Benutzer**: Lesen ✓
- **Administratoren**: Vollzugriff ✓

---

## Schritt 2: GPO erstellen (Gruppenrichtlinienverwaltung)

1. **Server-Manager** → **Tools** → **Gruppenrichtlinienverwaltung** öffnen
2. Auf die gewünschte **OU** (Organisationseinheit) mit den Ziel-Computern rechtsklicken
3. **„Gruppenrichtlinienobjekt hier erstellen und verknüpfen"**
4. Name: `ABAS Outlook Add-in Rollout`

---

## Schritt 3: Software-Installation konfigurieren

1. GPO rechtsklicken → **Bearbeiten**
2. Navigieren zu:
   ```
   Computerkonfiguration
     → Richtlinien
       → Softwareeinstellungen
         → Softwareinstallation
   ```
3. Rechtsklick → **Neu** → **Paket...**
4. MSI-Pfad eingeben (UNC-Pfad!):
   ```
   \\IhrServer\Software\AbasOutlookAddin\AbasOutlookAddin.msi
   ```
5. Bereitstellungsmethode: **Zugewiesen** (Assigned) ✓

> **Wichtig:** Immer den UNC-Netzwerkpfad `\\Server\Share\...` verwenden,
> niemals einen lokalen Pfad `C:\...`

---

## Schritt 4: Zielcomputer festlegen (Sicherheitsfilterung)

In der GPO unter **Bereich** → **Sicherheitsfilterung**:

- Standardmäßig: `Authentifizierte Benutzer` (alle Domänencomputer)
- Für selektiven Rollout: Bestimmte **Sicherheitsgruppe** hinzufügen,
  z.B. `GRP_ABAS_Benutzer`

---

## Schritt 5: COM-Registrierung via Startup-Skript

Da GPO-MSI die RegAsm-CustomAction nicht immer korrekt ausführt,
empfiehlt sich ein **Computerstartskript**:

1. GPO-Editor → `Computerkonfiguration → Windows-Einstellungen → Skripts → Start`
2. PowerShell-Skript hinzufügen: `Register-AbasAddin.ps1`

```powershell
# Register-AbasAddin.ps1
# Wird als SYSTEM beim Computerstart ausgeführt

$dllPath = "C:\Program Files\Ihr Unternehmen\ABAS Outlook Addin\AbasOutlookAddin.dll"
$regAsm  = "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$logPath = "C:\Windows\Temp\AbasAddinInstall.log"

if (Test-Path $dllPath) {
    $result = & $regAsm $dllPath /codebase /nologo 2>&1
    Add-Content $logPath "$(Get-Date) RegAsm: $result"
} else {
    Add-Content $logPath "$(Get-Date) DLL nicht gefunden: $dllPath"
}
```

---

## Schritt 6: Rollout testen

Auf einem Test-Client:
```cmd
:: Gruppenrichtlinien sofort anwenden (ohne Neustart)
gpupdate /force

:: GPO-Anwendung prüfen
gpresult /r

:: Nach Neustart prüfen ob MSI installiert wurde
reg query "HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\AbasOutlookAddin.Connect"
```

Dann **Outlook starten** → Add-in sollte in der Ribbon-Leiste erscheinen
(Datei → Optionen → Add-ins).

---

## Deinstallation via GPO

1. GPO-Editor → Softwareinstallation → Paket rechtsklicken → **Entfernen**
2. Option: **„Software sofort von Benutzern und Computern deinstallieren"** ✓

Oder per SCCM/Intune:
```cmd
msiexec.exe /x {11223344-5566-7788-AABB-CCDDEEFF0011} /qn /log uninstall.log
```

---

## Troubleshooting

| Problem | Lösung |
|---------|--------|
| Add-in erscheint nicht in Outlook | `gpupdate /force`, Outlook neu starten |
| Add-in von Outlook deaktiviert | Registry-Key `DoNotDisableAddinList` prüfen (in MSI enthalten) |
| RegAsm schlägt fehl | .NET 4.7.2 installiert? `reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release` |
| MSI lässt sich nicht installieren | UNC-Pfad korrekt? Computerkonto hat Lesezugriff? |
| Outlook fragt nach COM-Zulassung | Code-Signing-Zertifikat verwenden |

---

## Log-Dateien

| Datei | Inhalt |
|-------|--------|
| `%LOCALAPPDATA%\AbasOutlookAddin\Logs\addin_YYYYMMDD.log` | Add-in Laufzeit-Log |
| `C:\Windows\Temp\AbasAddinInstall.log` | RegAsm Startup-Skript |
| `%TEMP%\AbasOutlookAddin_install.log` | MSI-Installationslog |

MSI-Log manuell erzeugen:
```cmd
msiexec.exe /i AbasOutlookAddin.msi /qn /log "%TEMP%\AbasOutlookAddin_install.log"
```
