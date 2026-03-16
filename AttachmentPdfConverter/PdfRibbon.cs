using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AttachmentPdfConverter
{
    [ComVisible(true)]
    public class PdfRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        // Supported file extensions for conversion
        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            ".doc", ".docx", ".xlsx", ".csv"
        };

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            // Only show the button in the mail compose window
            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            {
                return GetResourceText("AttachmentPdfConverter.PdfRibbon.xml");
            }
            return null;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnConvertToPdf(IRibbonControl control)
        {
            try
            {
                // Get the active compose window
                var inspector = control.Context as Outlook.Inspector;
                if (inspector == null) return;

                var mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem == null)
                {
                    MessageBox.Show("This feature only works with email messages.",
                        "Convert to PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Collect convertible attachment names
                var convertible = new List<string>();
                for (int i = 1; i <= mailItem.Attachments.Count; i++)
                {
                    var att = mailItem.Attachments[i];
                    string ext = Path.GetExtension(att.FileName);
                    if (SupportedExtensions.Contains(ext))
                    {
                        convertible.Add(att.FileName);
                    }
                }

                if (convertible.Count == 0)
                {
                    MessageBox.Show(
                        "No convertible attachments found.\n\nSupported formats: .doc, .docx, .xlsx, .csv",
                        "Convert to PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Show the attachment picker dialog
                using (var picker = new AttachmentPickerForm(convertible))
                {
                    if (picker.ShowDialog() == DialogResult.OK && picker.SelectedAttachments.Count > 0)
                    {
                        ConvertSelectedAttachments(mailItem, picker.SelectedAttachments);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An unexpected error occurred:\n\n" + ex.Message,
                    "Convert to PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Conversion Logic

        private void ConvertSelectedAttachments(Outlook.MailItem mailItem, List<string> selectedNames)
        {
            string tempDir = Path.Combine(Path.GetTempPath(),
                "AttachmentPdfConverter_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            var cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            int successCount = 0;
            var errors = new List<string>();

            try
            {
                // Phase 1: Save selected attachments to temp and convert to PDF
                var conversions = new List<ConversionInfo>();

                for (int i = 1; i <= mailItem.Attachments.Count; i++)
                {
                    var att = mailItem.Attachments[i];
                    if (!selectedNames.Contains(att.FileName)) continue;

                    string tempPath = Path.Combine(tempDir, att.FileName);
                    att.SaveAsFile(tempPath);

                    try
                    {
                        string pdfPath = PdfConverter.ConvertToPdf(tempPath);
                        conversions.Add(new ConversionInfo
                        {
                            OriginalIndex = i,
                            OriginalName = att.FileName,
                            PdfPath = pdfPath
                        });
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{att.FileName}: {ex.Message}");
                    }
                }

                // Phase 2: Replace attachments (process in reverse order to preserve indices)
                conversions.Sort((a, b) => b.OriginalIndex.CompareTo(a.OriginalIndex));

                foreach (var conv in conversions)
                {
                    mailItem.Attachments[conv.OriginalIndex].Delete();
                    string pdfFileName = Path.ChangeExtension(conv.OriginalName, ".pdf");
                    mailItem.Attachments.Add(conv.PdfPath,
                        Outlook.OlAttachmentType.olByValue,
                        Type.Missing,
                        pdfFileName);
                    successCount++;
                }

                // Show results
                string message = $"Successfully converted {successCount} attachment(s) to PDF.";
                if (errors.Count > 0)
                {
                    message += "\n\nThe following files could not be converted:\n• " +
                               string.Join("\n• ", errors);
                }

                MessageBox.Show(message, "Convert to PDF", MessageBoxButtons.OK,
                    errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during conversion:\n\n" + ex.Message,
                    "Convert to PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = cursor;
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private class ConversionInfo
        {
            public int OriginalIndex;
            public string OriginalName;
            public string PdfPath;
        }

        #endregion
    }
}
