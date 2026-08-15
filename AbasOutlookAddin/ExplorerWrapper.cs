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

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

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

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromPoint(POINT pt,
            [MarshalAs(UnmanagedType.Interface)] out Accessibility.IAccessible acc,
            [MarshalAs(UnmanagedType.Struct)] out object child);

        /// <summary>
        /// Liest ueber die Barrierefreiheits-Schnittstelle, WAS genau unter dem Mauszeiger liegt –
        /// das getroffene Element und seine uebergeordneten Elemente.
        ///
        /// Noetig, weil Anhangbereich und Nachrichtentext im selben Fenster liegen; ueber die
        /// Fensterklasse laesst sich "Anhang" nicht von "Text" unterscheiden, ueber den
        /// Elementnamen schon. Die Kette nach oben ist noetig, weil der direkte Treffer oft nur
        /// ein Teilstueck ist: Auf dem Anhang-Chip liefert er "169 bytes", erst das
        /// uebergeordnete Element heisst "Rechnung.pdf 169 bytes 1 of 1 attachments"
        /// (real gemessen an Outlook 365).
        /// </summary>
        private static System.Collections.Generic.List<string> GetAccessibleNamesAt(POINT pt)
        {
            var names = new System.Collections.Generic.List<string>();
            try
            {
                if (AccessibleObjectFromPoint(pt, out Accessibility.IAccessible acc, out object child) != 0
                    || acc == null)
                    return names;

                try { names.Add(acc.get_accName(child)); } catch { }

                object current = acc;
                for (int level = 0; level < 4; level++)
                {
                    if (!(current is Accessibility.IAccessible element)) break;
                    try { names.Add(element.get_accName(0)); } catch { }
                    try { current = element.accParent; } catch { break; }
                }
            }
            catch { }
            return names;
        }

        /// <summary>
        /// Taucht einer der markierten Anhaenge im Namen des Elements unter dem Zeiger (oder
        /// eines seiner uebergeordneten Elemente) auf? Nur dann steht der Zeiger wirklich auf
        /// einem Anhang – und nur dann darf das Add-in den Drag an sich reissen. Im Fliesstext
        /// liefert die Kette leere Namen bzw. den Dokumentnamen und faellt hier durch.
        /// </summary>
        private static bool PointerIsOnAttachment(System.Collections.Generic.IList<string> accNames,
            System.Collections.Generic.IList<Attachment> attachments)
        {
            if (accNames == null || accNames.Count == 0 || attachments == null || attachments.Count == 0)
                return false;

            foreach (var attachment in attachments)
            {
                foreach (var candidate in new[] { SafeFileName(attachment), SafeDisplayName(attachment) })
                {
                    if (string.IsNullOrWhiteSpace(candidate)) continue;
                    foreach (var name in accNames)
                    {
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        if (name.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                }
            }
            return false;
        }

        private static string SafeFileName(Attachment a)
        {
            try { return a.FileName; } catch { return null; }
        }

        private static string SafeDisplayName(Attachment a)
        {
            try { return a.DisplayName; } catch { return null; }
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
                            {
                                // Was liegt unter dem Zeiger? Wird beim Drag-Start gebraucht,
                                // um "Anhang angeklickt" von "Text markiert" zu unterscheiden.
                                _downAccName = string.Join(" | ", GetAccessibleNamesAt(hs.pt));   // nur fuers Log
                                _capturedAttachments = CaptureAttachmentSelection(out _capturedSource);
                            }
                            break;

                        case WM_MOUSEMOVE:
                            if (_mouseDown && HasMovedEnough(hs.pt))
                            {
                                _mouseDown = false;

                                // Ohne tatsaechlich gedrueckte Maustaste gibt es keinen Drag.
                                // Der Hook sieht nur Nachrichten des Outlook-UI-Threads; geht ein
                                // WM_LBUTTONUP anderswo verloren (modaler Dialog, fremdes Fenster),
                                // blieb _mouseDown bis v1.4.2 haengen. Die naechste Mausbewegung
                                // startete dann einen "Phantom-Drag", der sofort dort fallen liess,
                                // wo der Zeiger gerade stand – mit der GESAMTEN aktuellen Auswahl.
                                if (!IsLeftButtonPhysicallyDown())
                                {
                                    Logger.Log("Drag verworfen: Maustaste ist gar nicht gedrueckt (verlorenes WM_LBUTTONUP).");
                                    break;
                                }

                                // Element-Drag NUR aus der Nachrichtenliste (SUPERGRID).
                                // Ausserhalb der Liste kommt ab v1.4.0 der Anhang-Drag zum Zug –
                                // aber nur, wenn Outlook tatsaechlich markierte Anhaenge meldet.
                                if (IsMessageList(_downHwnd))
                                    InitiateDrag();
                                else
                                    InitiateAttachmentDrag();
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

        private const int VK_LBUTTON = 0x01;
        private const int VK_RBUTTON = 0x02;

        /// <summary>
        /// Fragt den echten Zustand der Maustasten ab – unabhaengig davon, welche Nachrichten
        /// der Hook gesehen hat. Beruecksichtigt vertauschte Maustasten (Linkshaender).
        /// </summary>
        private static bool IsLeftButtonPhysicallyDown()
        {
            int key = SystemInformation.MouseButtonsSwapped ? VK_RBUTTON : VK_LBUTTON;
            return (GetAsyncKeyState(key) & 0x8000) != 0;
        }

        // Name des Elements, auf das geklickt wurde (Barrierefreiheits-Schnittstelle).
        private string _downAccName;

        private bool _textDragLogged;

        private void LogTextDragOnce(System.Collections.Generic.IList<string> overNow, int selectedCount)
        {
            if (_textDragLogged) return;
            _textDragLogged = true;
            Logger.Log($"Anhang-Drag NICHT gestartet: der Zeiger steht nicht auf einem Anhang " +
                       $"(unter dem Zeiger: '{string.Join(" | ", overNow)}', markiert: {selectedCount} " +
                       $"Anhang/Anhaenge, Klasse='{GetWindowClass(_downHwnd)}'). Outlook behaelt das " +
                       $"Ziehen – z. B. zum Markieren von Text. Wird nur einmal protokolliert.");
        }

        private static POINT CursorPoint()
        {
            GetCursorPos(out POINT p);
            return p;
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

                // Startzeit merken: Nur was NACH diesem Zeitpunkt im Postfach entsteht, zaehlt
                // spaeter als Nachweis, dass der Drop die Mail wirklich importiert hat.
                DateTime dragStarted = DateTime.Now;

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
                TryCompleteInternalMove(result, ctrlHeld, sourceRefs, dragStarted);

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
        /// meldet UND der Klick nachweislich auf einem dieser Anhaenge lag. Die zweite
        /// Bedingung ist ab v1.5.0 neu und behebt das gemeldete Phaenomen, dass beim Markieren
        /// von Text stattdessen ein Anhang gezogen (und beim Loslassen angehaengt) wurde.
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

                // Der Zeiger muss JETZT auf einem der markierten Anhaenge stehen. Sonst zieht
                // der Anwender gerade Text – und Outlook meldet die Anhang-Auswahl nur, weil
                // vorher irgendwo ein Anhang angeklickt wurde.
                //
                // Bewusst die AKTUELLE Zeigerposition und nicht die beim Mausklick gemerkte:
                // Der thread-lokale Hook verpasst gelegentlich ein WM_LBUTTONDOWN (real
                // beobachtet beim Klick auf den Anhang-Chip). Dann zeigen _downHwnd und
                // _downAccName noch auf die VORIGE Geste – und genau daraus entstand das
                // gemeldete Verhalten "Text wird nicht markiert, stattdessen haengt der Anhang
                // an der anderen Mail". Ein echter Anhang-Drag beginnt mit wenigen Pixeln
                // Bewegung und steht dabei noch auf dem Anhang.
                var overNow = GetAccessibleNamesAt(CursorPoint());
                if (!PointerIsOnAttachment(overNow, attachments))
                {
                    LogTextDragOnce(overNow, attachments.Count);
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
        /// Holt die markierten Anhaenge AUSSCHLIESSLICH aus dem Fenster, in dem der Klick
        /// begann: gleiches Wurzelfenster wie der Explorer -> Lesebereich, sonst die geoeffnete
        /// E-Mail (Inspector).
        ///
        /// Bis v1.4.2 gab es hier einen Rueckfall auf die jeweils ANDERE Quelle. Genau daraus
        /// entstand das gemeldete Phaenomen: Ein Klick im Verfassen-Fenster von Mail 2 fand dort
        /// keine Anhang-Auswahl, griff dann auf den Lesebereich (Mail 1) zurueck – und haengte
        /// beim Ziehen den Anhang von Mail 1 an Mail 2. Ein Anhang aus einem anderen Fenster
        /// ist nie das, was der Anwender gerade zieht.
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
            }
            else if (IsAncestorOfActiveInspector(downRoot))
            {
                var fromInspector = TryGetInspectorAttachments();
                if (fromInspector != null && fromInspector.Count > 0) { source = "geoeffnete E-Mail"; return fromInspector; }
                ReleaseIfCom(fromInspector);
            }

            source = null;
            return null;
        }

        /// <summary>
        /// Stellt sicher, dass der Klick im AKTIVEN Inspector-Fenster stattfand. Sonst wuerde
        /// die Anhang-Auswahl eines anderen geoeffneten Fensters gezogen.
        /// </summary>
        private bool IsAncestorOfActiveInspector(IntPtr downRoot)
        {
            if (downRoot == IntPtr.Zero) return false;
            try
            {
                Inspector inspector = _app?.ActiveInspector();
                if (inspector is IOleWindow oleWindow)
                {
                    oleWindow.GetWindow(out IntPtr hwnd);
                    return GetAncestor(hwnd, GA_ROOT) == downRoot;
                }
            }
            catch { }
            return false;
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
        ///
        /// Diese fuenf Bedingungen beschreiben nur die GEOMETRIE des Drops – sie sagen nichts
        /// darueber, ob Outlook die .msg auch tatsaechlich irgendwo importiert hat. Genau daran
        /// scheiterte v1.4.2. Das eigentliche Loeschen entscheidet deshalb erst
        /// <see cref="DragDropHandler.ScheduleVerifiedMove"/> anhand eines Ankunftsnachweises.
        /// </summary>
        private void TryCompleteInternalMove(DragDropEffects result, bool ctrlHeld,
            System.Collections.Generic.IList<DragDropHandler.ItemRef> sourceRefs, DateTime dragStarted)
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

                _handler.ScheduleVerifiedMove(sourceRefs, dragStarted);
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
