using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace AbasOutlookAddin
{
    /// <summary>
    /// COM-Standardinterface zum Ermitteln des Fensterhandles eines
    /// Office-Fensters. Das Outlook-Explorer-Objekt besitzt keine
    /// HWND-Eigenschaft, implementiert aber IOleWindow.
    /// </summary>
    [ComImport]
    [Guid("00000114-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleWindow
    {
        void GetWindow(out IntPtr phwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
    }

    /// <summary>
    /// Kapselt einen Outlook Explorer und installiert die Maus-Überwachung,
    /// um einen Drag &amp; Drop in den ABAS-Client zu starten.
    /// </summary>
    public class ExplorerWrapper : IDisposable
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _app;
        private readonly Explorer _explorer;
        private readonly DragDropHandler _handler;

        // Referenz halten, damit GC den Hook (und sein Delegate) nicht abräumt.
        private MouseDragWatcher _watcher;

        public ExplorerWrapper(Microsoft.Office.Interop.Outlook.Application app,
            Explorer explorer, DragDropHandler handler)
        {
            _app = app;
            _explorer = explorer;
            _handler = handler;
        }

        public void Attach()
        {
            try
            {
                IntPtr hwnd = IntPtr.Zero;
                if (_explorer is IOleWindow oleWindow)
                    oleWindow.GetWindow(out hwnd);

                _watcher = new MouseDragWatcher(_app, _explorer, _handler);
                _watcher.Install();

                Logger.Log($"Maus-Ueberwachung installiert (Explorer-HWND {hwnd}).");
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Maus-Ueberwachung konnte nicht installiert werden", ex);
            }
        }

        public void Dispose()
        {
            _watcher?.Dispose();
            _watcher = null;
        }
    }

    /// <summary>
    /// Erkennt einen Maus-Drag über einen THREAD-LOKALEN WH_MOUSE-Hook auf
    /// Outlooks UI-Thread. Kein globaler System-Hook, keine DLL-Injection –
    /// der Hook gilt ausschließlich für den aktuellen (Outlook-)Thread und
    /// sieht damit auch die Maus-Events der Kind-Fenster (z. B. der E-Mail-Liste),
    /// ohne dass Fensterklassen fest verdrahtet werden müssen.
    /// </summary>
    internal class MouseDragWatcher : IDisposable
    {
        private const int WH_MOUSE = 7;
        private const int HC_ACTION = 0;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONUP = 0x0202;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x; public int y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEHOOKSTRUCT
        {
            public POINT pt;
            public IntPtr hwnd;
            public uint wHitTestCode;
            public IntPtr dwExtraInfo;
        }

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        private const uint GA_ROOT = 2;

        // POSITIVLISTE: ABAS-Drag wird AUSSCHLIESSLICH gestartet, wenn der Klick in der
        // Outlook-Nachrichtenliste (die Übersicht mit "Heute/Gestern/Letzte Woche ...")
        // beginnt. Deren Fensterklasse ist "SUPERGRID". Ein Klick im Schreib-/Lesebereich,
        // im Ordnerbaum o. ä. startet damit KEINEN Drag.
        private static readonly string[] MessageListClasses =
        {
            "OutlookGrid",     // Nachrichtenliste in neueren Outlook-Versionen (M365)
            "SUPERGRID"        // Nachrichtenliste in älteren Outlook-Versionen
        };

        private static bool IsMessageList(IntPtr hwnd)
        {
            string cls = GetWindowClass(hwnd);
            foreach (var allowed in MessageListClasses)
                if (string.Equals(cls, allowed, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private static string GetWindowClass(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return string.Empty;
            var sb = new System.Text.StringBuilder(256);
            int n = GetClassName(hwnd, sb, sb.Capacity);
            return n > 0 ? sb.ToString() : string.Empty;
        }

        private readonly Microsoft.Office.Interop.Outlook.Application _app;
        private readonly Explorer _explorer;
        private readonly DragDropHandler _handler;
        private readonly HookProc _proc;   // Feld -> verhindert GC des Delegates
        private IntPtr _hookId = IntPtr.Zero;

        private bool _mouseDown;
        private POINT _downPoint;
        private IntPtr _downHwnd;      // Fenster, über dem die Maustaste gedrückt wurde
        private bool _dragInProgress; // Reentrancy-Schutz (#6)
        private bool _disposed;

        // Beim Mausklick gesicherte Anhang-Auswahl (siehe WM_LBUTTONDOWN).
        private System.Collections.Generic.List<Attachment> _capturedAttachments;
        private string _capturedSource;

        public MouseDragWatcher(Microsoft.Office.Interop.Outlook.Application app,
            Explorer explorer, DragDropHandler handler)
        {
            _app = app;
            _explorer = explorer;
            _handler = handler;
            _proc = HookCallback;
        }

        public void Install()
        {
            // Thread-lokaler Hook: hMod = 0, dwThreadId = aktueller (Outlook-)Thread.
            _hookId = SetWindowsHookEx(WH_MOUSE, _proc, IntPtr.Zero, GetCurrentThreadId());
            if (_hookId == IntPtr.Zero)
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(),
                    "SetWindowsHookEx (WH_MOUSE) fehlgeschlagen.");
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HC_ACTION && !_dragInProgress)
            {
                int msg = wParam.ToInt32();
                if (msg == WM_LBUTTONDOWN || msg == WM_MOUSEMOVE || msg == WM_LBUTTONUP)
                {
                    var hs = (MOUSEHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MOUSEHOOKSTRUCT));
                    switch (msg)
                    {
                        case WM_LBUTTONDOWN:
                            _mouseDown = true;
                            _downPoint = hs.pt;
                            _downHwnd = hs.hwnd;   // Fenster unter dem Cursor merken

                            // Anhang-Auswahl JETZT sichern: Der Hook laeuft noch VOR Outlooks
                            // eigener Klick-Verarbeitung, die eine Mehrfachauswahl auf den
                            // angeklickten Anhang zusammenklappt. Beim spaeteren Drag-Start
                            // waere davon nur noch ein Anhang uebrig (#Mehrfachauswahl).
                            ReleaseCapturedAttachments();
                            if (!IsMessageList(_downHwnd))
                                _capturedAttachments = CaptureAttachmentSelection(out _capturedSource);
                            break;

                        case WM_MOUSEMOVE:
                            if (_mouseDown && HasMovedEnough(hs.pt))
                            {
                                _mouseDown = false;

                                // Element-Drag NUR aus der Nachrichtenliste (SUPERGRID).
                                // Ausserhalb der Liste kommt ab v1.4.0 der Anhang-Drag zum Zug –
                                // aber nur, wenn Outlook tatsaechlich markierte Anhaenge meldet.
                                if (IsMessageList(_downHwnd))
                                {
                                    InitiateDrag();
                                }
                                else
                                {
                                    InitiateAttachmentDrag();
                                }
                            }
                            break;

                        case WM_LBUTTONUP:
                            _mouseDown = false;
                            break;
                    }
                }
            }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private bool HasMovedEnough(POINT current)
        {
            return Math.Abs(current.x - _downPoint.x) > SystemInformation.DragSize.Width ||
                   Math.Abs(current.y - _downPoint.y) > SystemInformation.DragSize.Height;
        }

        private void InitiateDrag()
        {
            if (_dragInProgress) return;
            _dragInProgress = true;

            try
            {
                Selection selection = null;
                try { selection = _explorer.Selection; } catch { /* kein Explorer-Kontext */ }
                if (selection == null || selection.Count == 0) return;

                Logger.Log($"Drag gestartet mit {selection.Count} Element(e)");

                // Strg-Zustand einmalig festhalten: dient sowohl der Anhang-Option als auch
                // (per Windows-Konvention "Strg = Kopieren") als Signal, dass ein interner
                // Outlook-Drop NICHT als Verschieben gewertet werden soll.
                bool ctrlHeld = (Control.ModifierKeys & Keys.Control) == Keys.Control;

                // Quell-EntryIDs VOR dem Extrahieren sichern (CreateDragData gibt die COM-Refs frei).
                // Nur damit laesst sich nach einem internen Verschieben das Original entfernen.
                var sourceRefs = _handler.CaptureItemIds(selection);

                DataObject dragData;

                // Standard: nur die .msg ablegen (die Anhänge stecken darin ohnehin drin).
                // Wird beim Losziehen die Strg-Taste (Ctrl) gehalten, wird eine einzelne
                // E-Mail zusätzlich mit allen Anhängen als separate Dateien abgelegt.
                if (ctrlHeld && selection.Count == 1
                    && selection[1] is MailItem mail && mail.Attachments.Count > 0)
                {
                    Logger.Log($"Strg gehalten: E-Mail als .msg + {mail.Attachments.Count} Anhang/Anhaengen ablegen");
                    dragData = _handler.CreateDragDataWithAttachments(mail);
                }
                else
                {
                    dragData = _handler.CreateDragData(selection);
                }

                if (dragData == null) return;

                // OLE Drag & Drop mit Copy (ABAS empfaengt CF_HDROP). Ein interner Outlook-Drop
                // meldet ebenfalls Copy; ob daraus ein Verschieben wird, entscheidet danach
                // TryCompleteInternalMove anhand des Ziel-Fensters (nicht anhand des Effekts).
                DragDropEffects result;
                using (var dragSource = new Control())
                {
                    result = dragSource.DoDragDrop(dragData, DragDropEffects.Copy);
                }

                Logger.Log($"Drag beendet, Ergebnis: {result}");

                // Internen Outlook-Drop erkennen und ggf. als echtes Verschieben abschliessen.
                TryCompleteInternalMove(result, ctrlHeld, sourceRefs);

                _handler.ScheduleCleanup();
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler beim Initiieren des Drags", ex);
                _handler.ScheduleCleanup();
            }
            finally
            {
                _dragInProgress = false;
            }
        }

        /// <summary>
        /// Startet einen Drag fuer die im Lesebereich bzw. in einer geoeffneten E-Mail
        /// MARKIERTEN Anhaenge (ab v1.4.0). Outlooks eigener Anhang-Drag liefert
        /// FileGroupDescriptor/FileContents – der ABAS-Client nimmt aber nur CF_HDROP an,
        /// deshalb legen wir die Anhaenge selbst als Temp-Dateien ab.
        ///
        /// Gestartet wird ausschliesslich, wenn Outlook eine nicht-leere AttachmentSelection
        /// meldet. Ist nichts markiert, verhaelt sich das Add-in wie bisher (Drag ignoriert)
        /// und Outlook macht seinen eigenen Drag.
        /// </summary>
        private void InitiateAttachmentDrag()
        {
            if (_dragInProgress) return;

            // Beim Klick gesicherte Auswahl (kann mehrere Anhaenge enthalten) ...
            var captured = _capturedAttachments;
            string capturedSource = _capturedSource;
            _capturedAttachments = null;   // ab hier gehoert die Liste dieser Methode

            // ... und die Auswahl, die Outlook JETZT meldet (nach dem Zusammenklappen
            // meist nur noch der angeklickte Anhang).
            var live = CaptureAttachmentSelection(out string liveSource);

            System.Collections.Generic.List<Attachment> attachments;
            string source;

            if (captured != null && captured.Count > (live?.Count ?? 0))
            {
                attachments = captured;
                source = capturedSource + ", Auswahl beim Klick";
                Logger.Log($"Mehrfachauswahl gerettet: beim Klick {captured.Count}, jetzt {live?.Count ?? 0} Anhang/Anhaenge.");
            }
            else
            {
                attachments = live;
                source = liveSource;
            }

            try
            {
                if (attachments == null || attachments.Count == 0)
                {
                    Logger.Log($"Drag ignoriert (kein Listen-Fenster, keine Anhang-Auswahl, Klasse='{GetWindowClass(_downHwnd)}').");
                    return;
                }

                _dragInProgress = true;
                Logger.Log($"Anhang-Drag gestartet mit {attachments.Count} Anhang/Anhaengen " +
                           $"(Quelle: {source}, Klasse='{GetWindowClass(_downHwnd)}').");

                DataObject dragData = _handler.CreateDragDataFromAttachments(attachments);

                // Rueckfall: Falls die beim Klick gesicherten COM-Referenzen inzwischen
                // veraltet sind, wenigstens den angeklickten Anhang liefern.
                if (dragData == null && !ReferenceEquals(attachments, live) && live != null && live.Count > 0)
                {
                    Logger.Log("Gesicherte Auswahl nicht mehr verwendbar – nutze aktuelle Auswahl.");
                    dragData = _handler.CreateDragDataFromAttachments(live);
                    attachments = live;
                }

                if (dragData == null)
                {
                    Logger.Log("Anhang-Drag abgebrochen: kein Anhang konnte als Datei abgelegt werden.");
                    return;
                }

                DragDropEffects result;
                using (var dragSource = new Control())
                {
                    result = dragSource.DoDragDrop(dragData, DragDropEffects.Copy);
                }

                Logger.Log($"Anhang-Drag beendet, Ergebnis: {result}");

                // Bewusst KEIN TryCompleteInternalMove: Anhaenge werden nie aus der
                // Quell-Mail entfernt, egal wohin sie gezogen werden.
                _handler.ScheduleCleanup();
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler beim Initiieren des Anhang-Drags", ex);
                _handler.ScheduleCleanup();
            }
            finally
            {
                // Beide Listen freigeben (die genutzte und die verworfene); der Rueckfall
                // oben braucht 'live' bis hierher, deshalb erst jetzt.
                ReleaseAttachments(captured);
                if (!ReferenceEquals(live, captured))
                    ReleaseAttachments(live);
                _dragInProgress = false;
            }
        }

        /// <summary>
        /// Liest die aktuell markierten Anhaenge aus und materialisiert sie als Liste.
        /// Materialisiert wird bewusst sofort: Die COM-Auswahl selbst kann sich aendern,
        /// die einzelnen Attachment-Objekte bleiben nutzbar.
        /// </summary>
        private System.Collections.Generic.List<Attachment> CaptureAttachmentSelection(out string source)
        {
            AttachmentSelection selection = GetAttachmentSelection(out source);
            var list = new System.Collections.Generic.List<Attachment>();
            if (selection == null)
                return list;

            try
            {
                foreach (object entry in selection)
                {
                    if (entry is Attachment attachment)
                        list.Add(attachment);
                    else
                        ReleaseIfCom(entry);
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Anhang-Auswahl konnte nicht gelesen werden", ex);
            }
            finally
            {
                ReleaseIfCom(selection);
            }

            return list;
        }

        private void ReleaseCapturedAttachments()
        {
            ReleaseAttachments(_capturedAttachments);
            _capturedAttachments = null;
            _capturedSource = null;
        }

        private static void ReleaseAttachments(System.Collections.Generic.IList<Attachment> attachments)
        {
            if (attachments == null) return;
            foreach (var attachment in attachments)
                ReleaseIfCom(attachment);
            attachments.Clear();
        }

        /// <summary>
        /// Holt die markierten Anhaenge – bevorzugt aus der Quelle, in der der Drag begann:
        /// gleiches Wurzelfenster wie der Explorer -> Lesebereich, sonst geoeffnete E-Mail
        /// (Inspector). Die jeweils andere Quelle dient als Rueckfallebene.
        /// </summary>
        private AttachmentSelection GetAttachmentSelection(out string source)
        {
            IntPtr downRoot = GetAncestor(_downHwnd, GA_ROOT);
            IntPtr explorerRoot = GetExplorerRootWindow();
            bool startedInExplorer = downRoot != IntPtr.Zero && downRoot == explorerRoot;

            if (startedInExplorer)
            {
                var fromExplorer = TryGetExplorerAttachments();
                if (fromExplorer != null && fromExplorer.Count > 0) { source = "Lesebereich"; return fromExplorer; }
                ReleaseIfCom(fromExplorer);

                var fromInspector = TryGetInspectorAttachments();
                if (fromInspector != null && fromInspector.Count > 0) { source = "geoeffnete E-Mail"; return fromInspector; }
                ReleaseIfCom(fromInspector);
            }
            else
            {
                var fromInspector = TryGetInspectorAttachments();
                if (fromInspector != null && fromInspector.Count > 0) { source = "geoeffnete E-Mail"; return fromInspector; }
                ReleaseIfCom(fromInspector);

                var fromExplorer = TryGetExplorerAttachments();
                if (fromExplorer != null && fromExplorer.Count > 0) { source = "Lesebereich"; return fromExplorer; }
                ReleaseIfCom(fromExplorer);
            }

            source = null;
            return null;
        }

        // Die Auswahl wird bei JEDEM Klick ausserhalb der Liste abgefragt – Fehler
        // deshalb nur einmal protokollieren, sonst laeuft das Log voll.
        private bool _explorerSelectionErrorLogged;
        private bool _inspectorSelectionErrorLogged;

        private AttachmentSelection TryGetExplorerAttachments()
        {
            try
            {
                return _explorer?.AttachmentSelection;
            }
            catch (System.Exception ex)
            {
                // Wirft z. B., wenn der Lesebereich aus ist oder die Outlook-Version
                // die Eigenschaft im aktuellen Kontext nicht bedient.
                if (!_explorerSelectionErrorLogged)
                {
                    _explorerSelectionErrorLogged = true;
                    Logger.LogError("Explorer.AttachmentSelection nicht verfuegbar (wird nur einmal protokolliert)", ex);
                }
                return null;
            }
        }

        private AttachmentSelection TryGetInspectorAttachments()
        {
            try
            {
                // Inspector bewusst NICHT freigeben – Outlook haelt das aktive
                // Inspector-Fenster ohnehin, und die Selection haengt daran.
                Inspector inspector = _app?.ActiveInspector();
                return inspector?.AttachmentSelection;
            }
            catch (System.Exception ex)
            {
                if (!_inspectorSelectionErrorLogged)
                {
                    _inspectorSelectionErrorLogged = true;
                    Logger.LogError("Inspector.AttachmentSelection nicht verfuegbar (wird nur einmal protokolliert)", ex);
                }
                return null;
            }
        }

        private static void ReleaseIfCom(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        /// <summary>
        /// Schliesst einen internen Outlook-Drop als echtes Verschieben ab.
        /// Entfernt die Quell-Elemente NUR, wenn alle Sicherheitsbedingungen erfuellt sind:
        ///   1) kein Strg (Strg = Kopieren),
        ///   2) der Drop wurde angenommen (Ergebnis != None),
        ///   3) das Ziel-Fenster gehoert zu Outlook selbst (gleiche Prozess-ID; ABAS ist ein
        ///      anderer Prozess und kann so NIE ein Loeschen ausloesen),
        ///   4) der Drop landete im Outlook-HAUPTFENSTER (gleiches Wurzelfenster wie der Explorer).
        ///      Ein Verfassen-/Inspector-Fenster ist ein eigenes Top-Level-Fenster und faellt damit
        ///      raus – so wird eine als Anhang gezogene Mail NICHT geloescht (Outlook meldet fuer
        ///      Ordner-Drops und Anhang-Drops gleichermassen 'Copy', deshalb reicht der Effekt nicht),
        ///   5) das Ziel ist NICHT die Nachrichtenliste selbst (Drop zurueck auf die Liste != Verschieben).
        /// Faellt eine Bedingung weg, bleibt es beim bisherigen Verhalten (Kopie) – kein Datenverlust.
        /// </summary>
        private void TryCompleteInternalMove(DragDropEffects result, bool ctrlHeld,
            System.Collections.Generic.IList<DragDropHandler.ItemRef> sourceRefs)
        {
            try
            {
                if (!Settings.InternalMoveEnabled) return;   // per Registry abgeschaltet
                if (ctrlHeld) return;                        // Strg = Kopieren
                if (result == DragDropEffects.None) return;  // Drop abgebrochen/abgelehnt
                if (sourceRefs == null || sourceRefs.Count == 0) return;

                IntPtr targetHwnd = GetDropTargetWindow();
                string cls = GetWindowClass(targetHwnd);
                GetWindowThreadProcessId(targetHwnd, out uint targetPid);
                uint ownPid = GetCurrentProcessId();

                IntPtr dropRoot = GetAncestor(targetHwnd, GA_ROOT);
                IntPtr explorerRoot = GetExplorerRootWindow();
                bool sameMainWindow = dropRoot != IntPtr.Zero && dropRoot == explorerRoot;

                Logger.Log($"Drop-Ziel: Klasse='{cls}', ZielPID={targetPid}, EigenePID={ownPid}, " +
                           $"Effekt={result}, DropRoot={dropRoot}, ExplorerRoot={explorerRoot}, Hauptfenster={sameMainWindow}");

                if (targetPid != ownPid) return;             // externes Ziel (z. B. ABAS) -> niemals loeschen
                if (!sameMainWindow) return;                 // eigenes Fenster (z. B. Verfassen) -> Anhang, nicht verschieben
                if (IsMessageList(targetHwnd)) return;       // zurueck auf die Nachrichtenliste -> kein Verschieben

                int deleted = _handler.DeleteItemsById(sourceRefs);
                Logger.Log($"Internes Verschieben abgeschlossen: {deleted} Quell-Element(e) entfernt.");

                // Damit im Papierkorb nichts liegen bleibt: verzoegert endgueltig entfernen,
                // aber nur mit Nachweis, dass die Mail im Zielordner angekommen ist.
                if (deleted > 0)
                    _handler.SchedulePermanentPurge(sourceRefs);
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler beim Abschliessen des internen Verschiebens", ex);
            }
        }

        /// <summary>Wurzelfenster (Top-Level) des Outlook-Explorers, ueber IOleWindow ermittelt.</summary>
        private IntPtr GetExplorerRootWindow()
        {
            try
            {
                if (_explorer is IOleWindow oleWindow)
                {
                    oleWindow.GetWindow(out IntPtr hwnd);
                    return GetAncestor(hwnd, GA_ROOT);
                }
            }
            catch { }
            return IntPtr.Zero;
        }

        /// <summary>Ermittelt das Fenster unter dem Mauszeiger (Drop-Zielpunkt).</summary>
        private static IntPtr GetDropTargetWindow()
        {
            if (!GetCursorPos(out POINT p))
                return IntPtr.Zero;
            return WindowFromPoint(p);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            ReleaseCapturedAttachments();
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }
    }
}
