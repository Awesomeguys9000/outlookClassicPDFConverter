using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace AttachmentPdfConverter
{
    public partial class AttachmentPickerForm : Form
    {
        public List<string> SelectedAttachments { get; private set; }

        public AttachmentPickerForm(List<string> attachmentNames)
        {
            InitializeComponent();
            SelectedAttachments = new List<string>();

            // Populate the checklist with attachment names
            foreach (var name in attachmentNames)
            {
                clbAttachments.Items.Add(name, true); // checked by default
            }
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbAttachments.Items.Count; i++)
                clbAttachments.SetItemChecked(i, true);
        }

        private void btnDeselectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbAttachments.Items.Count; i++)
                clbAttachments.SetItemChecked(i, false);
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            SelectedAttachments.Clear();
            foreach (var item in clbAttachments.CheckedItems)
            {
                SelectedAttachments.Add(item.ToString());
            }

            if (SelectedAttachments.Count == 0)
            {
                MessageBox.Show("Please select at least one attachment to convert.",
                    "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
