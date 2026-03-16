namespace AttachmentPdfConverter
{
    partial class AttachmentPickerForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.CheckedListBox clbAttachments;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Button btnDeselectAll;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblHeader;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Panel panelTop;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblHeader = new System.Windows.Forms.Label();
            this.clbAttachments = new System.Windows.Forms.CheckedListBox();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnDeselectAll = new System.Windows.Forms.Button();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.panelTop = new System.Windows.Forms.Panel();
            this.panelBottom.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();

            // panelTop
            this.panelTop.Controls.Add(this.lblHeader);
            this.panelTop.Controls.Add(this.btnSelectAll);
            this.panelTop.Controls.Add(this.btnDeselectAll);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Size = new System.Drawing.Size(420, 60);
            this.panelTop.Padding = new System.Windows.Forms.Padding(12, 10, 12, 4);

            // lblHeader
            this.lblHeader.AutoSize = true;
            this.lblHeader.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblHeader.Location = new System.Drawing.Point(12, 10);
            this.lblHeader.Text = "Select attachments to convert to PDF:";

            // btnSelectAll
            this.btnSelectAll.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            this.btnSelectAll.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnSelectAll.Location = new System.Drawing.Point(235, 32);
            this.btnSelectAll.Size = new System.Drawing.Size(85, 25);
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);

            // btnDeselectAll
            this.btnDeselectAll.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            this.btnDeselectAll.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnDeselectAll.Location = new System.Drawing.Point(325, 32);
            this.btnDeselectAll.Size = new System.Drawing.Size(85, 25);
            this.btnDeselectAll.Text = "Deselect All";
            this.btnDeselectAll.Click += new System.EventHandler(this.btnDeselectAll_Click);

            // clbAttachments
            this.clbAttachments.CheckOnClick = true;
            this.clbAttachments.Dock = System.Windows.Forms.DockStyle.Fill;
            this.clbAttachments.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.clbAttachments.FormattingEnabled = true;
            this.clbAttachments.IntegralHeight = false;
            this.clbAttachments.Location = new System.Drawing.Point(0, 60);
            this.clbAttachments.Margin = new System.Windows.Forms.Padding(12);

            // panelBottom
            this.panelBottom.Controls.Add(this.btnConvert);
            this.panelBottom.Controls.Add(this.btnCancel);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(0, 270);
            this.panelBottom.Size = new System.Drawing.Size(420, 50);
            this.panelBottom.Padding = new System.Windows.Forms.Padding(12, 8, 12, 8);

            // btnConvert
            this.btnConvert.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            this.btnConvert.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnConvert.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnConvert.Location = new System.Drawing.Point(220, 10);
            this.btnConvert.Size = new System.Drawing.Size(95, 30);
            this.btnConvert.Text = "Convert";
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);

            // btnCancel
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCancel.Location = new System.Drawing.Point(320, 10);
            this.btnCancel.Size = new System.Drawing.Size(90, 30);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            // AttachmentPickerForm
            this.AcceptButton = this.btnConvert;
            this.CancelButton = this.btnCancel;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(420, 320);
            this.Controls.Add(this.clbAttachments);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.panelBottom);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Convert Attachments to PDF";
            this.panelBottom.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}
