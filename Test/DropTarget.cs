using System;
using System.Windows.Forms;
using System.Drawing;

namespace AbasDropTest
{
    /// <summary>
    /// ABAS-Ersatz zum Testen des Add-ins OHNE ABAS.
    /// Akzeptiert AUSSCHLIESSLICH CF_HDROP (echte Dateipfade / DataFormats.FileDrop).
    /// Outlooks natives Drag liefert virtuelle Dateien (FileGroupDescriptor/FileContents)
    /// und wird hier ABGELEHNT. Nur ein vom Add-in augmentierter Drag mit echten
    /// Dateipfaden wird angenommen und unten aufgelistet.
    /// </summary>
    public class DropForm : Form
    {
        private readonly ListBox _list = new ListBox();
        private readonly Label _hint = new Label();

        public DropForm()
        {
            Text = "ABAS-Ersatz – nur echte Dateipfade (CF_HDROP)";
            Size = new Size(640, 420);
            StartPosition = FormStartPosition.CenterScreen;
            AllowDrop = true;

            _hint.Dock = DockStyle.Top;
            _hint.Height = 60;
            _hint.TextAlign = ContentAlignment.MiddleCenter;
            _hint.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            _hint.Text = "E-Mail aus Outlook HIERHER ziehen.\n" +
                         "Akzeptiert nur echte Dateipfade (wie ABAS) – natives Outlook-Drag wird abgelehnt.";

            _list.Dock = DockStyle.Fill;
            _list.Font = new Font("Consolas", 9);

            Controls.Add(_list);
            Controls.Add(_hint);

            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            // NUR CF_HDROP akzeptieren – genau wie ein Ziel, das echte Pfade braucht.
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                _hint.BackColor = Color.FromArgb(200, 255, 200); // grün = akzeptiert
            }
            else
            {
                e.Effect = DragDropEffects.None;
                _hint.BackColor = Color.FromArgb(255, 210, 210); // rot = abgelehnt
                var formats = e.Data.GetFormats();
                _list.Items.Add("[ABGELEHNT] Kein CF_HDROP. Angebotene Formate: " + string.Join(", ", formats));
            }
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            _hint.BackColor = SystemColors.Control;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                _list.Items.Add("=== " + DateTime.Now.ToString("HH:mm:ss") + " : " + files.Length + " Datei(en) empfangen ===");
                foreach (var f in files)
                    _list.Items.Add("   " + f);
            }
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new DropForm());
        }
    }
}
