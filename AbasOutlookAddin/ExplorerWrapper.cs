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
        private readonly Explorer _explorer;
        private readonly DragDropHandler _handler;

        // Referenz halten, damit GC den Hook (und sein Delegate) nicht abräumt.
        private MouseDragWatcher _watcher;

        public ExplorerWrapper(Explorer explorer, DragDropHandler handler)
        {
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

                _watcher = new MouseDragWatcher(_explorer, _handler);
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

        private readonly Explorer _explorer;
        private readonly DragDropHandler _handler;
        private readonly HookProc _proc;   // Feld -> verhindert GC des Delegates
        private IntPtr _hookId = IntPtr.Zero;

        private bool _mouseDown;
        private POINT _downPoint;
        private IntPtr _downHwnd;      // Fenster, über dem die Maustaste gedrückt wurde
        private bool _dragInProgress; // Reentrancy-Schutz (#6)
        private bool _disposed;

        public MouseDragWatcher(Explorer explorer, DragDropHandler handler)
        {
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
                            break;

                        case WM_MOUSEMOVE:
                            if (_mouseDown && HasMovedEnough(hs.pt))
                            {
                                _mouseDown = false;

                                // Drag NUR aus der Nachrichtenliste (SUPERGRID) starten.
                                if (IsMessageList(_downHwnd))
                                {
                                    InitiateDrag();
                                }
                                else
                                {
                                    Logger.Log($"Drag ignoriert (kein Listen-Fenster, Klasse='{GetWindowClass(_downHwnd)}').");
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
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }
    }
}
