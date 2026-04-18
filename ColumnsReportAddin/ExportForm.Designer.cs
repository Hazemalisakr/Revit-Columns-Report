namespace ColumnsReportAddin
{
    partial class ExportForm
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.Label lblLocation;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.PictureBox picLinkedIn;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblLocation = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.picLinkedIn = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picLinkedIn)).BeginInit();
            this.SuspendLayout();

            // lblLocation
            this.lblLocation.AutoSize = true;
            this.lblLocation.Location = new System.Drawing.Point(20, 30);
            this.lblLocation.Name = "lblLocation";
            this.lblLocation.Size = new System.Drawing.Size(100, 13);
            this.lblLocation.Text = "Report Location:";

            // txtPath
            this.txtPath.Location = new System.Drawing.Point(140, 27);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(340, 20);
            this.txtPath.ReadOnly = true;

            // btnBrowse
            this.btnBrowse.Location = new System.Drawing.Point(495, 25);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;

            // btnExport
            this.btnExport.Location = new System.Drawing.Point(165, 140);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(100, 30);
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;

            // btnClose
            this.btnClose.Location = new System.Drawing.Point(320, 140);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 30);
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;

            // picLinkedIn
            this.picLinkedIn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.picLinkedIn.Location = new System.Drawing.Point(535, 140);
            this.picLinkedIn.Name = "picLinkedIn";
            this.picLinkedIn.Size = new System.Drawing.Size(32, 32);
            this.picLinkedIn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;

            // ExportForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 191);
            this.Controls.Add(this.lblLocation);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.picLinkedIn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Name = "ExportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Columns Report";
            ((System.ComponentModel.ISupportInitialize)(this.picLinkedIn)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
