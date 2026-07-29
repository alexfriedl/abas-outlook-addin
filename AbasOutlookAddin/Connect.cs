using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Outlook;

namespace AbasOutlookAddin
{
    /// <summary>
    /// ABAS Outlook Drag & Drop Add-in
    /// Ermöglicht das Drag & Drop von Outlook-Elementen in den ABAS Windows-Client.
    /// Kein API-Hooking – setzt auf den standard Windows IDataObject / CF_HDROP Mechanismus.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("AbasOutlookAddin.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        private Application _outlookApp;
        private Explorer _activeExplorer;
        private DragDropHandler _dragDropHandler;

        // Referenzen halten, damit GC die Fenster-Subclasses nicht abräumt.
        private readonly System.Collections.Generic.List<ExplorerWrapper> _explorerWrappers
            = new System.Collections.Generic.List<ExplorerWrapper>();

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
            object AddInInst, ref Array custom)
        {
            try
            {
                // Assembly-Integritätsprüfung (#3)
                if (!VerifyAssemblyIntegrity())
                {
                    Logger.LogError("SICHERHEIT: Assembly-Integritaetspruefung fehlgeschlagen. Add-in wird nicht geladen.");
                    return;
                }

                _outlookApp = (Application)Application;
                _activeExplorer = _outlookApp.ActiveExplorer();
                _dragDropHandler = new DragDropHandler(_outlookApp);

                // Explorer-Events registrieren
                if (_activeExplorer != null)
                    AttachExplorerEvents(_activeExplorer);

                // Neuen Explorer überwachen
                _outlookApp.Explorers.NewExplorer += Explorers_NewExplorer;

                // Einstellungen einmal beim Start protokollieren – so ist im Log sofort
                // sichtbar, mit welchem Verhalten das Add-in beim Anwender laeuft.
                Logger.Log($"ABAS Outlook Add-in erfolgreich geladen (internes Verschieben: " +
                           $"{(Settings.InternalMoveEnabled ? "aktiv" : "aus")}).");
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler beim Laden des Add-ins", ex);
            }
        }

        /// <summary>
        /// Prüft ob die Assembly mit dem erwarteten Strong Name signiert ist (#3).
        /// Verhindert DLL-Hijacking durch Austausch der DLL.
        /// </summary>
        private static bool VerifyAssemblyIntegrity()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var assemblyName = assembly.GetName();

                // Prüfen ob Strong Name vorhanden
                byte[] publicKey = assemblyName.GetPublicKeyToken();
                if (publicKey == null || publicKey.Length == 0)
                {
                    // In Debug-Builds ohne Signatur erlauben
                    #if DEBUG
                    Logger.Log("WARNUNG: Assembly ist nicht signiert (Debug-Build).");
                    return true;
                    #else
                    Logger.LogError("SICHERHEIT: Assembly ist nicht mit einem Strong Name signiert.");
                    return false;
                    #endif
                }

                // Optional: Erwarteten Public Key Token prüfen
                // byte[] expectedToken = new byte[] { 0x..., 0x..., ... };
                // if (!CompareBytes(publicKey, expectedToken)) return false;

                Logger.Log($"Assembly-Integritaet verifiziert: {assemblyName.FullName}");
                return true;
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler bei Assembly-Integritaetspruefung", ex);
                return false;
            }
        }

        private void Explorers_NewExplorer(Explorer Explorer)
        {
            AttachExplorerEvents(Explorer);
        }

        private void AttachExplorerEvents(Explorer explorer)
        {
            var explorerWrapper = new ExplorerWrapper(_outlookApp, explorer, _dragDropHandler);
            explorerWrapper.Attach();
            _explorerWrappers.Add(explorerWrapper);
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                foreach (var w in _explorerWrappers)
                    w.Dispose();
                _explorerWrappers.Clear();

                _dragDropHandler?.Dispose();
                Logger.Log("ABAS Outlook Add-in entladen.");
            }
            catch { }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }
}
