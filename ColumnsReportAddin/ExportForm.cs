using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ColumnsReportAddin
{
    public partial class ExportForm : Form
    {
        public string SelectedPath { get; private set; }

        public ExportForm()
        {
            InitializeComponent();
            WireEvents();
            BuildLinkedInIcon();
        }

        private void WireEvents()
        {
            btnBrowse.Click += (s, e) =>
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Select folder to save the Columns Report";
                    if (dlg.ShowDialog() == DialogResult.OK)
                        txtPath.Text = dlg.SelectedPath;
                }
            };

            btnExport.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtPath.Text))
                {
                    MessageBox.Show("Please select a folder first.", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                SelectedPath = txtPath.Text;
                DialogResult = DialogResult.OK;
                Close();
            };

            btnClose.Click += (s, e) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };
        }

        private void BuildLinkedInIcon()
        {
            Bitmap bmp = new Bitmap(32, 32);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

                using (var path = RoundedRect(new Rectangle(0, 0, 31, 31), 5))
                using (var fill = new SolidBrush(Color.FromArgb(0, 119, 181)))
                {
                    g.FillPath(fill, path);
                }

                using (Font f = new Font("Arial", 18, FontStyle.Bold))
                using (Brush b = new SolidBrush(Color.White))
                {
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    g.DrawString("in", f, b, new RectangleF(0, 0, 32, 32), sf);
                }
            }

            picLinkedIn.Image = bmp;
            picLinkedIn.Click += (s, e) =>
            {
                Process.Start("https://www.linkedin.com/in/hazem-sakr/");
            };
        }

        private GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var gp = new GraphicsPath();
            gp.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            gp.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            gp.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            gp.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            gp.CloseFigure();
            return gp;
        }
    }
}
