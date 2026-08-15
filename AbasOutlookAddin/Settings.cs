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
    ///     InternalMove   (REG_DWORD)  1 = internes Verschieben aktiv (Standard)
    ///                                 0 = aus: Quell-Mail bleibt immer erhalten
    ///     MaxAutoDelete  (REG_DWORD)  Obergrenze fuer automatisch entfernte Quell-Elemente
    ///                                 pro Drop (Standard 25, 0 = Verschieben faktisch aus)
    ///     PurgeFromTrash (REG_DWORD)  1 = Quell-Mail endgueltig aus dem Papierkorb entfernen
    ///                                 0 = im Papierkorb liegen lassen (Standard, wiederherstellbar)
    /// </summary>
    internal static class Settings
    {
        private const string SubKey = @"Software\ABAS Outlook Addin";
        private const string ValueInternalMove = "InternalMove";
        private const string ValueMaxAutoDelete = "MaxAutoDelete";
        private const string ValuePurgeFromTrash = "PurgeFromTrash";

        private static bool? _internalMoveEnabled;
        private static int? _maxAutoDelete;
        private static bool? _purgeFromTrashEnabled;

        /// <summary>
        /// Schadensbegrenzung: Wie viele Quell-Elemente darf EIN Drop hoechstens entfernen?
        ///
        /// Grund (Vorfall 2026-08): Ein einzelner Drop mit grosser Auswahl loeschte auf einen
        /// Schlag den halben Posteingang. Ein Anwender, der bewusst filet, zieht ein paar Mails –
        /// nicht hunderte. Wird die Grenze ueberschritten, bleibt ALLES erhalten; die Kopien im
        /// Zielordner sind dann zwar doppelt da, aber nichts ist verloren.
        /// </summary>
        public static int MaxAutoDelete
        {
            get
            {
                if (!_maxAutoDelete.HasValue)
                {
                    _maxAutoDelete = ReadInt(ValueMaxAutoDelete, 25);
                    if (_maxAutoDelete.Value < 0) _maxAutoDelete = 0;
                    Logger.Log($"Einstellung MaxAutoDelete = {_maxAutoDelete.Value}");
                }
                return _maxAutoDelete.Value;
            }
        }

        /// <summary>
        /// Soll die Quell-Mail nach dem Verschieben endgueltig aus „Geloeschte Elemente"
        /// verschwinden?
        ///
        /// STANDARD: NEIN. Bis v1.4.2 war das an; damit war jedes Fehlverhalten des Add-ins
        /// unwiderruflich. Der Papierkorb ist die letzte Rueckholmoeglichkeit und kostet nur
        /// Platz, bis Outlook ihn leert. Wer es anders will, setzt PurgeFromTrash=1.
        /// </summary>
        public static bool PurgeFromTrashEnabled
        {
            get
            {
                if (!_purgeFromTrashEnabled.HasValue)
                {
                    _purgeFromTrashEnabled = ReadFlag(ValuePurgeFromTrash, false);
                    Logger.Log($"Einstellung PurgeFromTrash = {(_purgeFromTrashEnabled.Value ? "1 (endgueltig loeschen)" : "0 (Papierkorb behalten)")}");
                }
                return _purgeFromTrashEnabled.Value;
            }
        }

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

        private static int ReadInt(string valueName, int defaultValue)
        {
            object value = ReadValue(Registry.CurrentUser, valueName)
                        ?? ReadValue(Registry.LocalMachine, valueName);
            if (value == null) return defaultValue;
            try { return Convert.ToInt32(value); }
            catch { return defaultValue; }
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
