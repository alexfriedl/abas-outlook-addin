using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace AbasOutlookAddin
{
    /// <summary>
    /// Kernlogik: Extrahiert Outlook-Elemente als temporäre Dateien
    /// und stellt sie als CF_HDROP für den ABAS Windows-Client bereit.
    /// </summary>
    public class DragDropHandler : IDisposable
    {
        private readonly Application _outlookApp;
        private readonly List<string> _tempFiles = new List<string>();
        private readonly string _tempDir;
        private bool _disposed;

        // Maximale Anzahl Temp-Dateien bevor erzwungenes Cleanup (#8)
        private const int MaxTempFiles = 200;

        // Sicheres, benutzerspezifisches Temp-Verzeichnis
        private static readonly string TempBasePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AbasOutlookAddin", "Temp");

        public DragDropHandler(Application outlookApp)
        {
            _outlookApp = outlookApp;

            // Alte Temp-Dateien aus vorherigen Sessions aufräumen (#2 - Crash-Recovery)
            CleanupStaleTempDirectories();

            _tempDir = Path.Combine(TempBasePath, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            // Restriktive ACLs setzen: nur aktueller User hat Zugriff (#2)
            SetRestrictiveAcl(_tempDir);

            Logger.Log($"Temp-Verzeichnis erstellt: {_tempDir}");
        }

        /// <summary>
        /// Setzt ACLs auf das Temp-Verzeichnis: nur der aktuelle Benutzer hat Zugriff.
        /// Verhindert, dass andere lokale Benutzer Temp-Dateien lesen können.
        /// </summary>
        private static void SetRestrictiveAcl(string directoryPath)
        {
            try
            {
                var dirInfo = new DirectoryInfo(directoryPath);
                var security = dirInfo.GetAccessControl();

                // Vererbung deaktivieren, bestehende Regeln entfernen
                security.SetAccessRuleProtection(true, false);

                // Nur aktueller Benutzer: Vollzugriff
                var currentUser = WindowsIdentity.GetCurrent().User;
                if (currentUser != null)
                {
                    security.AddAccessRule(new FileSystemAccessRule(
                        currentUser,
                        FileSystemRights.FullControl,
                        InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                        PropagationFlags.None,
                        AccessControlType.Allow));
                }

                dirInfo.SetAccessControl(security);
            }
            catch (System.Exception ex)
            {
                Logger.LogError("ACL konnte nicht gesetzt werden", ex);
            }
        }

        /// <summary>
        /// Räumt Temp-Verzeichnisse aus vorherigen Sessions auf,
        /// die durch Crashes liegen geblieben sein könnten (#2).
        /// </summary>
        private static void CleanupStaleTempDirectories()
        {
            try
            {
                if (!Directory.Exists(TempBasePath)) return;

                foreach (var dir in Directory.GetDirectories(TempBasePath))
                {
                    try
                    {
                        var dirInfo = new DirectoryInfo(dir);
                        // Älter als 1 Stunde = sicher verwaist
                        if (dirInfo.CreationTimeUtc < DateTime.UtcNow.AddHours(-1))
                        {
                            Directory.Delete(dir, true);
                            Logger.Log($"Verwaistes Temp-Verzeichnis gelöscht: {dir}");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Logger.LogError($"Cleanup fehlgeschlagen für: {dir}", ex);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Wird aufgerufen wenn der Benutzer mit dem Drag startet.
        /// Extrahiert die selektierten Outlook-Elemente als Dateien.
        /// </summary>
        public DataObject CreateDragData(Selection selection)
        {
            if (selection == null || selection.Count == 0)
                return null;

            // Temp-Wachstum begrenzen (#8)
            EnforceTempLimit();

            var files = new List<string>();

            foreach (object item in selection)
            {
                try
                {
                    string filePath = ExtractItem(item);
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        files.Add(filePath);
                        _tempFiles.Add(filePath);
                        Logger.Log($"Element extrahiert: {Path.GetFileName(filePath)}");
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.LogError("Fehler beim Extrahieren eines Elements", ex);
                }
                finally
                {
                    if (item != null && Marshal.IsComObject(item))
                        Marshal.ReleaseComObject(item);
                }
            }

            if (files.Count == 0)
                return null;

            // Standard Windows DataObject mit CF_HDROP erstellen
            var dataObject = new DataObject();
            var fileCollection = new System.Collections.Specialized.StringCollection();
            fileCollection.AddRange(files.ToArray());
            dataObject.SetFileDropList(fileCollection);

            return dataObject;
        }

        private string ExtractItem(object item)
        {
            string fileName = null;
            string tempPath = null;

            if (item is MailItem mail)
            {
                fileName = SanitizeFileName(mail.Subject) + ".msg";
                tempPath = GetUniquePath(_tempDir, fileName);
                mail.SaveAs(tempPath, OlSaveAsType.olMSG);
            }
            else if (item is AttachmentSelection)
            {
                return null;
            }
            else if (item is ContactItem contact)
            {
                fileName = SanitizeFileName(contact.FullName) + ".vcf";
                tempPath = GetUniquePath(_tempDir, fileName);
                contact.SaveAs(tempPath, OlSaveAsType.olVCard);
            }
            else if (item is AppointmentItem appointment)
            {
                fileName = SanitizeFileName(appointment.Subject) + ".ics";
                tempPath = GetUniquePath(_tempDir, fileName);
                appointment.SaveAs(tempPath, OlSaveAsType.olICal);
            }
            else if (item is TaskItem task)
            {
                fileName = SanitizeFileName(task.Subject) + ".msg";
                tempPath = GetUniquePath(_tempDir, fileName);
                task.SaveAs(tempPath, OlSaveAsType.olMSG);
            }

            // TOCTOU-Schutz: Sicherstellen dass die Datei noch im Temp-Dir liegt (#5)
            if (tempPath != null && !ValidateTempFilePath(tempPath))
            {
                Logger.LogError($"Sicherheitswarnung: Datei liegt ausserhalb des Temp-Verzeichnisses: {tempPath}");
                try { File.Delete(tempPath); } catch { }
                return null;
            }

            return tempPath;
        }

        /// <summary>
        /// Legt eine einzelne E-Mail als .msg UND zusätzlich alle echten Anhänge ab.
        /// Wird genutzt, wenn beim Ziehen die Umschalttaste (Shift) gehalten wird.
        /// Ohne Shift wird nur die .msg abgelegt (die Anhänge stecken darin ohnehin drin).
        /// </summary>
        public DataObject CreateDragDataWithAttachments(MailItem mail)
        {
            if (mail == null)
                return null;

            // Temp-Wachstum begrenzen (#8)
            EnforceTempLimit();

            var files = new List<string>();

            // 1) Die E-Mail selbst als .msg
            try
            {
                string msgPath = ExtractItem(mail);
                if (!string.IsNullOrEmpty(msgPath))
                {
                    files.Add(msgPath);
                    _tempFiles.Add(msgPath);
                    Logger.Log($"Element extrahiert: {Path.GetFileName(msgPath)}");
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler beim Extrahieren der E-Mail", ex);
            }

            // 2) Zusätzlich alle echten Anhänge (keine eingebetteten Bilder)
            files.AddRange(ExtractAttachments(mail));

            if (files.Count == 0)
                return null;

            var dataObject = new DataObject();
            var fileCollection = new System.Collections.Specialized.StringCollection();
            fileCollection.AddRange(files.ToArray());
            dataObject.SetFileDropList(fileCollection);

            return dataObject;
        }

        /// <summary>
        /// Speichert alle echten Anhänge (olByValue) einer E-Mail als Temp-Dateien
        /// und gibt deren Pfade zurück. Eingebettete Bilder werden übersprungen.
        /// </summary>
        private List<string> ExtractAttachments(MailItem mail)
        {
            var files = new List<string>();
            if (mail?.Attachments == null || mail.Attachments.Count == 0)
                return files;

            foreach (Attachment attachment in mail.Attachments)
            {
                try
                {
                    // Nur echte Anhänge, keine eingebetteten Bilder
                    if (attachment.Type == OlAttachmentType.olByValue)
                    {
                        string safeFileName = SanitizeFileName(attachment.FileName);
                        string tempPath = GetUniquePath(_tempDir, safeFileName);
                        attachment.SaveAsFile(tempPath);

                        // TOCTOU-Schutz (#5)
                        if (!ValidateTempFilePath(tempPath))
                        {
                            Logger.LogError($"Sicherheitswarnung: Anhang ausserhalb Temp-Dir: {tempPath}");
                            try { File.Delete(tempPath); } catch { }
                            continue;
                        }

                        files.Add(tempPath);
                        _tempFiles.Add(tempPath);
                        Logger.Log($"Anhang extrahiert: {Path.GetFileName(tempPath)}");
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.LogError($"Fehler beim Extrahieren des Anhangs", ex);
                }
            }
            return files;
        }

        /// <summary>
        /// Stellt sicher, dass eine Datei tatsächlich innerhalb des Temp-Verzeichnisses liegt.
        /// Verhindert Path Traversal und Symlink-Attacken (#1, #5).
        /// </summary>
        private bool ValidateTempFilePath(string filePath)
        {
            try
            {
                // Kanonischen Pfad auflösen (löst Symlinks, .., etc. auf)
                string canonicalPath = Path.GetFullPath(filePath);
                string canonicalTempDir = Path.GetFullPath(_tempDir);

                // Muss innerhalb des Temp-Verzeichnisses liegen
                if (!canonicalPath.StartsWith(canonicalTempDir + Path.DirectorySeparatorChar,
                    StringComparison.OrdinalIgnoreCase))
                    return false;

                // Prüfen ob die Datei ein Symlink/Reparse Point ist
                if (File.Exists(filePath))
                {
                    var fileInfo = new FileInfo(filePath);
                    if (fileInfo.Attributes.HasFlag(FileAttributes.ReparsePoint))
                    {
                        Logger.LogError($"Symlink erkannt und blockiert: {filePath}");
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Bereinigt Temp-Dateien nach dem Drop (verzögert, um ABAS Zeit zum Lesen zu geben).
        /// Cleanup erfolgt IMMER, auch bei abgebrochenen Drags (#8).
        /// </summary>
        public void ScheduleCleanup()
        {
            var timer = new Timer { Interval = 30000 }; // 30 Sekunden
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                CleanupTempFiles();
                timer.Dispose();
            };
            timer.Start();
        }

        /// <summary>
        /// Sofortiges Cleanup wenn zu viele Temp-Dateien existieren (#8).
        /// </summary>
        private void EnforceTempLimit()
        {
            if (_tempFiles.Count >= MaxTempFiles)
            {
                Logger.Log($"Temp-Limit ({MaxTempFiles}) erreicht, erzwinge Cleanup");
                CleanupTempFiles();
            }
        }

        private void CleanupTempFiles()
        {
            foreach (var file in _tempFiles)
            {
                try
                {
                    if (File.Exists(file))
                    {
                        // Vor dem Löschen nochmal validieren (#1 - Symlink-Schutz)
                        if (ValidateTempFilePath(file))
                        {
                            File.Delete(file);
                        }
                        else
                        {
                            Logger.LogError($"Cleanup übersprungen (ungültiger Pfad): {file}");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.LogError($"Konnte Temp-Datei nicht loeschen", ex);
                }
            }
            _tempFiles.Clear();
        }

        /// <summary>
        /// Bereinigt Dateinamen: entfernt ungültige Zeichen, Path-Traversal-Sequenzen,
        /// und begrenzt die Länge (#1).
        /// </summary>
        private static string SanitizeFileName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "Outlook_Element_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Ungültige Dateinamen-Zeichen entfernen
            foreach (char c in Path.GetInvalidFileNameChars())
                input = input.Replace(c, '_');

            // Path-Traversal-Sequenzen entfernen (#1)
            input = input.Replace("..", "_");
            input = input.Replace("/", "_");
            input = input.Replace("\\", "_");

            // Reservierte Windows-Dateinamen blockieren
            string nameUpper = input.Trim().ToUpperInvariant();
            string[] reserved = { "CON", "PRN", "AUX", "NUL",
                "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
                "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9" };
            foreach (var r in reserved)
            {
                if (nameUpper == r || nameUpper.StartsWith(r + "."))
                {
                    input = "_" + input;
                    break;
                }
            }

            // Länge begrenzen (Windows MAX_PATH)
            if (input.Length > 100) input = input.Substring(0, 100);

            return input.Trim();
        }

        private static string GetUniquePath(string dir, string fileName)
        {
            string path = Path.Combine(dir, fileName);

            // Sicherheitscheck: Ergebnis muss im Zielverzeichnis liegen (#1)
            string canonical = Path.GetFullPath(path);
            string canonicalDir = Path.GetFullPath(dir);
            if (!canonical.StartsWith(canonicalDir + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase))
            {
                // Fallback auf sicheren Namen
                fileName = "safe_" + Guid.NewGuid().ToString("N").Substring(0, 8) +
                           Path.GetExtension(fileName);
                path = Path.Combine(dir, fileName);
            }

            if (!File.Exists(path)) return path;

            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);
            int counter = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{nameWithoutExt}_{counter++}{ext}");
            }
            return path;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                CleanupTempFiles();
                try
                {
                    if (Directory.Exists(_tempDir))
                        Directory.Delete(_tempDir, true);
                }
                catch { }
                _disposed = true;
            }
        }
    }
}
