using System;
using Microsoft.Win32;

namespace AbasOutlookAddin
{
    /// <summary>
    /// Einstellungen des Add-ins aus der Registry.
    ///
    /// Gelesen wird zuerst HKCU (benutzerspezifisch), dann HKLM (per GPO/Rollout
    /// vorgebbar). Fehlt beides, gilt der Standardwert – das Add-in laeuft also
    /// ohne jeden Registry-Eintrag wie bisher.
    ///
    ///   HKCU\Software\ABAS Outlook Addin  bzw.
    ///   HKLM\SOFTWARE\ABAS Outlook Addin
    ///     InternalMove (REG_DWORD)  1 = internes Verschieben aktiv (Standard)
    ///                               0 = aus: Quell-Mail bleibt immer erhalten
    /// </summary>
    internal static class Settings
    {
        private const string SubKey = @"Software\ABAS Outlook Addin";
        private const string ValueInternalMove = "InternalMove";

        private static bool? _internalMoveEnabled;

        /// <summary>
        /// Steuert, ob ein Drop innerhalb Outlooks die Quell-Mail entfernt (echtes
        /// Verschieben, v1.3.0) oder ob sie stehen bleibt (Kopie, Verhalten bis v1.2.0).
        ///
        /// STANDARD: AN (Verschieben) – das ist die ausdrueckliche Anforderung der
        /// Anwender. Technisch bedingt importiert Outlook die abgelegte .msg als neues
        /// Element und das Original wandert in "Geloeschte Elemente"; wer das nicht will,
        /// laesst den Papierkorb von Outlook beim Beenden automatisch leeren oder schaltet
        /// das Verschieben ueber InternalMove=0 ab (Registry, per Rollout/GPO).
        /// </summary>
        public static bool InternalMoveEnabled
        {
            get
            {
                if (!_internalMoveEnabled.HasValue)
                {
                    _internalMoveEnabled = ReadFlag(ValueInternalMove, true);
                    Logger.Log($"Einstellung InternalMove = {(_internalMoveEnabled.Value ? "1 (Verschieben aktiv)" : "0 (Verschieben aus)")}");
                }
                return _internalMoveEnabled.Value;
            }
        }

        private static bool ReadFlag(string valueName, bool defaultValue)
        {
            // Benutzer-Einstellung schlaegt Maschinen-Einstellung
            object value = ReadValue(Registry.CurrentUser, valueName)
                        ?? ReadValue(Registry.LocalMachine, valueName);

            if (value == null)
                return defaultValue;

            try
            {
                return Convert.ToInt32(value) != 0;
            }
            catch
            {
                return defaultValue;
            }
        }

        private static object ReadValue(RegistryKey hive, string valueName)
        {
            try
            {
                using (var key = hive.OpenSubKey(SubKey))
                {
                    return key?.GetValue(valueName);
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Registry-Einstellung '{valueName}' konnte nicht gelesen werden", ex);
                return null;
            }
        }
    }
}
