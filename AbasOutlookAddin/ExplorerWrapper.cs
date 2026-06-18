using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace AbasOutlookAddin
{
    /// <summary>
    /// Kapselt einen Outlook Explorer und reagiert auf Maus-Events
    /// um den Drag & Drop Prozess zu starten.
    /// </summary>
    public class ExplorerWrapper
    {
        private readonly Explorer _explorer;
        private readonly DragDropHandler _handler;

        public ExplorerWrapper(Explorer explorer, DragDropHandler handler)
        {
            _explorer = explorer;
            _handler = handler;
        }

        public void Attach()
        {
            // Wir nutzen einen Low-Level-Hook auf das Explorer-Fenster
            // via NativeWindow Subclassing – kein API-Hooking auf Systemebene
            try
            {
                IntPtr hwnd = (IntPtr)_explorer.CurrentFolder.GetType()
                    .InvokeMember("HWND", System.Reflection.BindingFlags.GetProperty, null, _explorer, null);

                var nativeHandler = new OutlookNativeWindow(hwnd, _explorer, _handler);
                nativeHandler.AssignHandle(hwnd);

                Logger.Log($"Explorer-Fenster ueberwacht: {hwnd}");
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Explorer-Attach fehlgeschlagen, nutze Event-Fallback", ex);
                // Fallback: Nutze Outlook Selection-Change Event als Trigger
                AttachSelectionEvents();
            }
        }

        private void AttachSelectionEvents()
        {
            _explorer.SelectionChange += Explorer_SelectionChange;
        }

        private void Explorer_SelectionChange()
        {
            Logger.Log($"Selektion geaendert: {_explorer.Selection?.Count} Element(e)");
        }
    }

    /// <summary>
    /// Windows Message Subclassing für das Outlook Explorer Fenster.
    /// Fängt WM_MOUSEMOVE ab um den Drag zu erkennen – ohne globales Hooking.
    /// </summary>
    internal class OutlookNativeWindow : NativeWindow, IDisposable
    {
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONUP = 0x0202;

        private readonly Explorer _explorer;
        private readonly DragDropHandler _handler;
        private bool _mouseDown;
        private System.Drawing.Point _mouseDownPoint;
        private bool _disposed;

        // Verhindert Reentrancy während eines Drags (#6)
        private bool _dragInProgress;

        public OutlookNativeWindow(IntPtr hwnd, Explorer explorer, DragDropHandler handler)
        {
            _explorer = explorer;
            _handler = handler;
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_LBUTTONDOWN:
                    if (!_dragInProgress)
                    {
                        _mouseDown = true;
                        _mouseDownPoint = GetPoint(m.LParam);
                    }
                    break;

                case WM_MOUSEMOVE:
                    if (_mouseDown && !_dragInProgress && HasMovedEnough(GetPoint(m.LParam)))
                    {
                        _mouseDown = false;
                        InitiateDrag();
                    }
                    break;

                case WM_LBUTTONUP:
                    _mouseDown = false;
                    break;
            }

            base.WndProc(ref m);
        }

        private void InitiateDrag()
        {
            if (_dragInProgress) return;
            _dragInProgress = true;

            try
            {
                var selection = _explorer.Selection;
                if (selection == null || selection.Count == 0) return;

                Logger.Log($"Drag gestartet mit {selection.Count} Element(e)");

                DataObject dragData = null;

                // Prüfen ob einzelne E-Mail mit Anhängen selektiert
                if (selection.Count == 1 && selection[1] is MailItem mail && mail.Attachments.Count > 0)
                {
                    var choice = ShowDragChoiceDialog(mail);
                    if (choice == DragChoice.Cancel) return;

                    dragData = choice == DragChoice.Attachments
                        ? _handler.CreateDragDataFromAttachments(mail)
                        : _handler.CreateDragData(selection);
                }
                else
                {
                    dragData = _handler.CreateDragData(selection);
                }

                if (dragData == null) return;

                // Standard OLE Drag & Drop starten – ABAS empfängt CF_HDROP
                var result = System.Windows.Forms.DragDrop.DoDragDrop(dragData, DragDropEffects.Copy);

                Logger.Log($"Drag beendet, Ergebnis: {result}");

                // IMMER Cleanup planen, auch bei Abbruch (#8)
                _handler.ScheduleCleanup();
            }
            catch (System.Exception ex)
            {
                Logger.LogError("Fehler beim Initiieren des Drags", ex);
                // Auch bei Fehler aufräumen (#8)
                _handler.ScheduleCleanup();
            }
            finally
            {
                _dragInProgress = false;
            }
        }

        /// <summary>
        /// Auswahl-Dialog: E-Mail oder Anhänge?
        /// Nutzt TopMost um Manipulation durch andere Prozesse zu erschweren (#6).
        /// </summary>
        private DragChoice ShowDragChoiceDialog(MailItem mail)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "In ABAS ablegen";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterScreen;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.TopMost = true; // Verhindert, dass andere Fenster darüber liegen (#6)
                dialog.Size = new System.Drawing.Size(420, 200);

                var label = new Label
                {
                    Text = $"E-Mail: \"{TruncateForDisplay(mail.Subject, 50)}\"\n\n" +
                           $"Was soll in ABAS abgelegt werden?",
                    Location = new System.Drawing.Point(15, 15),
                    Size = new System.Drawing.Size(380, 60),
                    AutoSize = false
                };

                var btnEmail = new Button
                {
                    Text = $"E-Mail (.msg)",
                    DialogResult = DialogResult.Yes,
                    Location = new System.Drawing.Point(15, 90),
                    Size = new System.Drawing.Size(120, 35)
                };

                var btnAttachments = new Button
                {
                    Text = $"Anhänge ({mail.Attachments.Count})",
                    DialogResult = DialogResult.No,
                    Location = new System.Drawing.Point(145, 90),
                    Size = new System.Drawing.Size(120, 35)
                };

                var btnCancel = new Button
                {
                    Text = "Abbrechen",
                    DialogResult = DialogResult.Cancel,
                    Location = new System.Drawing.Point(275, 90),
                    Size = new System.Drawing.Size(120, 35)
                };

                dialog.Controls.AddRange(new Control[] { label, btnEmail, btnAttachments, btnCancel });
                dialog.AcceptButton = btnEmail;
                dialog.CancelButton = btnCancel;

                var result = dialog.ShowDialog();

                return result switch
                {
                    DialogResult.Yes => DragChoice.Email,
                    DialogResult.No => DragChoice.Attachments,
                    _ => DragChoice.Cancel
                };
            }
        }

        private static string TruncateForDisplay(string input, int maxLength)
        {
            if (string.IsNullOrEmpty(input)) return "(kein Betreff)";
            if (input.Length <= maxLength) return input;
            return input.Substring(0, maxLength) + "...";
        }

        private static System.Drawing.Point GetPoint(IntPtr lParam)
        {
            int x = (int)(lParam.ToInt64() & 0xFFFF);
            int y = (int)((lParam.ToInt64() >> 16) & 0xFFFF);
            return new System.Drawing.Point(x, y);
        }

        private bool HasMovedEnough(System.Drawing.Point current)
        {
            return Math.Abs(current.X - _mouseDownPoint.X) > SystemInformation.DragSize.Width ||
                   Math.Abs(current.Y - _mouseDownPoint.Y) > SystemInformation.DragSize.Height;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                ReleaseHandle();
                _disposed = true;
            }
        }
    }

    internal enum DragChoice { Email, Attachments, Cancel }
}
