using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;

namespace AttachmentPdfConverter
{
    [ComVisible(true)]
    [Guid("F1A2B3C4-D5E6-4F78-9A0B-C1D2E3F4A5B6")]
    [ProgId("AttachmentPdfConverter.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Outlook.Application _outlookApp;
        private IRibbonUI _ribbon;

        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            ".doc", ".docx", ".xlsx", ".csv"
        };

        #region COM Registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            // Register as an Outlook add-in
            string keyName = @"Software\Microsoft\Office\Outlook\Addins\" + type.FullName;
            using (var key = Registry.CurrentUser.CreateSubKey(keyName))
            {
                key.SetValue("FriendlyName", "Attachment PDF Converter");
                key.SetValue("Description", "Converts email attachments to PDF using Microsoft Print to PDF");
                key.SetValue("LoadBehavior", 3); // Load at startup
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            string keyName = @"Software\Microsoft\Office\Outlook\Addins\" + type.FullName;
            Registry.CurrentUser.DeleteSubKey(keyName, false);
        }

        #endregion

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
            object AddInInst, ref Array custom)
        {
            _outlookApp = (Outlook.Application)Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            if (RibbonID == "Microsoft.Outlook.Mail.Compose")
            {
                return GetResourceText("AttachmentPdfConverter.PdfRibbon.xml");
            }
            return null;
        }

        #endregion

        #region Ribbon Callbacks (called by name from the XML)

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnConvertToPdf(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector == null) return;

                var mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem == null)
                {
                    MessageBox.Show("This feature only works with email messages.",
                        "Convert to PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

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
                // Phase 1: Save selected attachments and convert to PDF
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

                // Phase 2: Replace attachments (reverse order to preserve indices)
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
