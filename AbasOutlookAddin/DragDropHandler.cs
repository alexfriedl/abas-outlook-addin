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

            // Body einmal holen: Damit wird geprüft, ob ein Anhang per Content-ID
            // im Text eingebettet ist (Signatur-Logos, Inline-Bilder).
            string htmlBody = null;
            try { htmlBody = mail.HTMLBody; } catch { }

            foreach (Attachment attachment in mail.Attachments)
            {
                try
                {
                    // Nur echte Anhänge, keine eingebetteten Bilder
                    if (attachment.Type != OlAttachmentType.olByValue)
                        continue;

                    // Signatur-Logos & Inline-Bilder aussortieren (Feedback Bastian):
                    // sie sind ebenfalls olByValue und landeten deshalb bisher mit im Drop.
                    if (IsInlineAttachment(attachment, htmlBody))
                    {
                        Logger.Log($"Anhang uebersprungen (im Text eingebettet): {GetAttachmentDisplayName(attachment)}");
                        continue;
                    }

                    string tempPath = ExtractAttachment(attachment);
                    if (!string.IsNullOrEmpty(tempPath))
                    {
                        files.Add(tempPath);
                        _tempFiles.Add(tempPath);
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.LogError($"Fehler beim Extrahieren des Anhangs", ex);
                }
            }
            return files;
        }

        // MAPI-Eigenschaften zum Erkennen eingebetteter Anhänge
        private const string PropAttachmentHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
        private const string PropAttachContentId = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";

        /// <summary>
        /// Erkennt Anhänge, die nicht als eigenständige Datei gemeint sind, sondern im
        /// Nachrichtentext stecken – typischerweise Signatur-Logos und Inline-Bilder.
        ///
        /// Zwei Kriterien, das erste das verlässlichere:
        ///   1) PR_ATTACHMENT_HIDDEN ist gesetzt (so markiert Outlook Inline-Anhänge),
        ///   2) der Anhang hat eine Content-ID, die im HTML-Body als "cid:" referenziert wird.
        ///
        /// Bewusst konservativ: Im Zweifel (Property nicht lesbar) gilt der Anhang als
        /// echter Anhang und wird mitgenommen – lieber eine Datei zu viel als eine fehlende.
        /// </summary>
        private static bool IsInlineAttachment(Attachment attachment, string htmlBody)
        {
            PropertyAccessor accessor = null;
            try
            {
                accessor = attachment.PropertyAccessor;
                if (accessor == null) return false;

                try
                {
                    object hidden = accessor.GetProperty(PropAttachmentHidden);
                    if (hidden is bool isHidden && isHidden)
                        return true;
                }
                catch { /* Property nicht vorhanden -> naechstes Kriterium */ }

                if (!string.IsNullOrEmpty(htmlBody))
                {
                    try
                    {
                        string contentId = accessor.GetProperty(PropAttachContentId) as string;
                        if (!string.IsNullOrEmpty(contentId) &&
                            htmlBody.IndexOf(contentId, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                    catch { }
                }
            }
            catch { }
            finally
            {
                if (accessor != null && Marshal.IsComObject(accessor))
                    Marshal.ReleaseComObject(accessor);
            }

            return false;
        }

        /// <summary>
        /// Legt die im Lesebereich bzw. in einer geöffneten E-Mail MARKIERTEN Anhänge
        /// als Temp-Dateien ab und stellt sie als CF_HDROP bereit (ab v1.4.0).
        /// Damit lassen sich Anhänge direkt – ohne Umweg über die .msg – ins DMS ziehen.
        /// </summary>
        public DataObject CreateDragDataFromAttachments(IList<Attachment> attachments)
        {
            if (attachments == null || attachments.Count == 0)
                return null;

            // Temp-Wachstum begrenzen (#8)
            EnforceTempLimit();

            var files = new List<string>();

            foreach (Attachment attachment in attachments)
            {
                try
                {
                    string tempPath = ExtractAttachment(attachment);
                    if (!string.IsNullOrEmpty(tempPath))
                    {
                        files.Add(tempPath);
                        _tempFiles.Add(tempPath);
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.LogError("Fehler beim Extrahieren eines markierten Anhangs", ex);
                }
                // Die COM-Referenzen gehören dem Aufrufer (ExplorerWrapper) und werden
                // dort freigegeben – hier nicht, sonst ist der Rückfall auf die
                // Live-Auswahl nicht mehr möglich.
            }

            if (files.Count == 0)
                return null;

            var dataObject = new DataObject();
            var fileCollection = new System.Collections.Specialized.StringCollection();
            fileCollection.AddRange(files.ToArray());
            dataObject.SetFileDropList(fileCollection);

            return dataObject;
        }

        /// <summary>
        /// Speichert einen einzelnen Anhang als Temp-Datei und gibt den Pfad zurück.
        /// Behandelt Sonderfälle: eingebettete E-Mails (olEmbeddeditem) haben oft keinen
        /// Dateinamen und werden als .msg abgelegt; OLE-Objekte lassen sich nicht als
        /// Datei speichern und werden übersprungen.
        /// </summary>
        private string ExtractAttachment(Attachment attachment)
        {
            if (attachment == null)
                return null;

            OlAttachmentType type;
            try { type = attachment.Type; }
            catch { type = OlAttachmentType.olByValue; }

            if (type == OlAttachmentType.olOLE)
            {
                Logger.Log($"Anhang uebersprungen (OLE-Objekt): {GetAttachmentDisplayName(attachment)}");
                return null;
            }

            string rawName = null;
            try { rawName = attachment.FileName; } catch { }

            if (string.IsNullOrWhiteSpace(rawName))
            {
                // Eingebettete E-Mails liefern haeufig keinen FileName -> Anzeigename + .msg
                rawName = GetAttachmentDisplayName(attachment);
                if (!rawName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                    rawName += ".msg";
            }

            string tempPath = GetUniquePath(_tempDir, SanitizeAttachmentName(rawName));
            attachment.SaveAsFile(tempPath);

            // TOCTOU-Schutz (#5)
            if (!ValidateTempFilePath(tempPath))
            {
                Logger.LogError($"Sicherheitswarnung: Anhang ausserhalb Temp-Dir: {tempPath}");
                try { File.Delete(tempPath); } catch { }
                return null;
            }

            Logger.Log($"Anhang extrahiert: {Path.GetFileName(tempPath)}");
            return tempPath;
        }

        private static string GetAttachmentDisplayName(Attachment attachment)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(attachment.DisplayName))
                    return attachment.DisplayName;
            }
            catch { }
            return "Anhang_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
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

        /// <summary>
        /// Wie <see cref="SanitizeFileName"/>, behält aber die Dateiendung.
        /// Wichtig für Anhänge: ohne Endung kann das DMS den Dateityp nicht zuordnen,
        /// und die 100-Zeichen-Grenze würde sie bei langen Namen abschneiden.
        /// </summary>
        private static string SanitizeAttachmentName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return SanitizeFileName(null);

            string ext;
            try { ext = Path.GetExtension(input) ?? string.Empty; }
            catch { ext = string.Empty; }

            // Unplausible "Endungen" (z. B. Punkt mitten im Namen) ignorieren
            if (ext.Length > 20) ext = string.Empty;

            foreach (char c in Path.GetInvalidFileNameChars())
                ext = ext.Replace(c, '_');

            string baseName = input.Substring(0, input.Length - ext.Length);
            baseName = SanitizeFileName(baseName);

            int maxBase = 100 - ext.Length;
            if (maxBase < 1) maxBase = 1;
            if (baseName.Length > maxBase)
                baseName = baseName.Substring(0, maxBase).Trim();
            if (baseName.Length == 0)
                baseName = "Anhang";

            return baseName + ext;
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

        /// <summary>
        /// Leichtgewichtige Referenz auf ein Outlook-Element (EntryID + StoreID),
        /// mit der sich das Element nach dem Drag zum Loeschen wiederfinden laesst.
        /// Haelt bewusst KEINE COM-Referenz.
        /// </summary>
        public sealed class ItemRef
        {
            public string EntryId;
            public string StoreId;
            public string Subject;
        }

        /// <summary>
        /// Sichert vor dem Drag die EntryIDs der selektierten Elemente, damit sie
        /// nach einem internen Outlook-Verschieben gezielt entfernt werden koennen.
        /// Gibt KEINE COM-Referenzen zurueck (alle werden hier freigegeben).
        /// </summary>
        public List<ItemRef> CaptureItemIds(Selection selection)
        {
            var refs = new List<ItemRef>();
            if (selection == null || selection.Count == 0)
                return refs;

            foreach (object item in selection)
            {
                try
                {
                    if (TryGetItemRef(item, out ItemRef reference))
                        refs.Add(reference);
                }
                catch (System.Exception ex)
                {
                    Logger.LogError("EntryID konnte nicht ermittelt werden", ex);
                }
                finally
                {
                    if (item != null && Marshal.IsComObject(item))
                        Marshal.ReleaseComObject(item);
                }
            }
            return refs;
        }

        private static bool TryGetItemRef(object item, out ItemRef reference)
        {
            reference = null;
            string entryId = null, storeId = null, subject = null;

            if (item is MailItem mail)
            {
                entryId = mail.EntryID; subject = mail.Subject; storeId = GetStoreId(mail.Parent);
            }
            else if (item is ContactItem contact)
            {
                entryId = contact.EntryID; subject = contact.FullName; storeId = GetStoreId(contact.Parent);
            }
            else if (item is AppointmentItem appointment)
            {
                entryId = appointment.EntryID; subject = appointment.Subject; storeId = GetStoreId(appointment.Parent);
            }
            else if (item is TaskItem task)
            {
                entryId = task.EntryID; subject = task.Subject; storeId = GetStoreId(task.Parent);
            }
            else
            {
                return false;
            }

            if (string.IsNullOrEmpty(entryId))
                return false;

            reference = new ItemRef { EntryId = entryId, StoreId = storeId, Subject = subject };
            return true;
        }

        /// <summary>Liest die StoreID aus dem Eltern-Ordner (fuer GetItemFromID bei Nicht-Standard-Stores).</summary>
        private static string GetStoreId(object parent)
        {
            try
            {
                if (parent is Folder folder)
                    return folder.StoreID;
            }
            catch { }
            finally
            {
                if (parent != null && Marshal.IsComObject(parent))
                    Marshal.ReleaseComObject(parent);
            }
            return null;
        }

        /// <summary>
        /// Entfernt die zuvor per <see cref="CaptureItemIds"/> gesicherten Quell-Elemente.
        /// Wird nur nach einem erkannten internen Outlook-Verschieben aufgerufen.
        /// Loeschen erfolgt via .Delete() -> Ordner "Geloeschte Elemente" (wiederherstellbar).
        /// </summary>
        public int DeleteItemsById(IList<ItemRef> refs)
        {
            if (refs == null || refs.Count == 0)
                return 0;

            int deleted = 0;
            NameSpace session = null;
            try
            {
                session = _outlookApp.Session;
                foreach (var r in refs)
                {
                    object item = null;
                    try
                    {
                        item = string.IsNullOrEmpty(r.StoreId)
                            ? session.GetItemFromID(r.EntryId)
                            : session.GetItemFromID(r.EntryId, r.StoreId);

                        if (item != null && DeleteOutlookItem(item))
                        {
                            deleted++;
                            Logger.Log($"Quell-Element nach Verschieben entfernt: {r.Subject}");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        // Kann fehlschlagen, wenn Outlook das Original bereits selbst verschoben hat -> ok.
                        Logger.LogError($"Quell-Element konnte nicht entfernt werden (evtl. schon verschoben): {r.Subject}", ex);
                    }
                    finally
                    {
                        if (item != null && Marshal.IsComObject(item))
                            Marshal.ReleaseComObject(item);
                    }
                }
            }
            finally
            {
                if (session != null && Marshal.IsComObject(session))
                    Marshal.ReleaseComObject(session);
            }
            return deleted;
        }

        private static bool DeleteOutlookItem(object item)
        {
            if (item is MailItem mail) { mail.Delete(); return true; }
            if (item is ContactItem contact) { contact.Delete(); return true; }
            if (item is AppointmentItem appointment) { appointment.Delete(); return true; }
            if (item is TaskItem task) { task.Delete(); return true; }
            return false;
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
