using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttachmentPdfConverter
{
    /// <summary>
    /// Converts Office documents to PDF using "Microsoft Print to PDF" printer.
    /// This uses the Windows printing pipeline (not Office's built-in PDF export),
    /// which handles certain documents more reliably than other export methods.
    /// </summary>
    public static class PdfConverter
    {
        private const string PdfPrinterName = "Microsoft Print to PDF";
        private const int DialogTimeoutSeconds = 60;
        private const int FileWaitTimeoutSeconds = 30;

        /// <summary>
        /// Converts a document to PDF. Returns the path to the generated PDF file.
        /// </summary>
        public static string ConvertToPdf(string inputFilePath)
        {
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException("Input file not found.", inputFilePath);

            string ext = Path.GetExtension(inputFilePath).ToLowerInvariant();
            string pdfPath = Path.Combine(
                Path.GetDirectoryName(inputFilePath),
                Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf");

            // Delete existing PDF if present to avoid overwrite prompts
            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            switch (ext)
            {
                case ".doc":
                case ".docx":
                    ConvertWordToPdf(inputFilePath, pdfPath);
                    break;
                case ".xlsx":
                case ".csv":
                    ConvertExcelToPdf(inputFilePath, pdfPath);
                    break;
                default:
                    throw new NotSupportedException(
                        $"File type '{ext}' is not supported for PDF conversion.");
            }

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException(
                    $"PDF conversion failed — output file was not created for '{Path.GetFileName(inputFilePath)}'.");

            return pdfPath;
        }

        #region Word Conversion

        private static void ConvertWordToPdf(string inputPath, string outputPath)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                // Save and set the printer
                string previousPrinter = wordApp.ActivePrinter;
                wordApp.ActivePrinter = PdfPrinterName;

                // Open the document read-only
                doc = wordApp.Documents.Open(
                    inputPath,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);

                // Start a background task to handle the "Save Print Output As" dialog
                // from the Microsoft Print to PDF driver
                var cts = new CancellationTokenSource();
                var dialogTask = Task.Run(() =>
                    HandlePrintToFileDialog(outputPath, cts.Token));

                // Print the document. This call blocks until the dialog is dismissed.
                doc.PrintOut(Background: false);

                // Wait for the dialog handler to finish (it should be done by now)
                if (!dialogTask.Wait(TimeSpan.FromSeconds(DialogTimeoutSeconds)))
                {
                    cts.Cancel();
                    throw new TimeoutException(
                        "Timed out waiting for the PDF save dialog to appear.");
                }

                // Rethrow any exception from the dialog handler
                if (dialogTask.IsFaulted && dialogTask.Exception != null)
                    throw dialogTask.Exception.InnerException ?? dialogTask.Exception;

                // Wait for the PDF file to be fully written
                WaitForFile(outputPath);

                // Clean up
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                doc = null;

                wordApp.ActivePrinter = previousPrinter;
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    try { wordApp.Quit(false); } catch { }
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        #endregion

        #region Excel Conversion

        private static void ConvertExcelToPdf(string inputPath, string outputPath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Save and set the printer
                string previousPrinter = excelApp.ActivePrinter;
                excelApp.ActivePrinter = PdfPrinterName;

                // Open the workbook read-only
                workbook = excelApp.Workbooks.Open(
                    inputPath,
                    ReadOnly: true,
                    AddToMru: false);

                // Start background dialog handler
                var cts = new CancellationTokenSource();
                var dialogTask = Task.Run(() =>
                    HandlePrintToFileDialog(outputPath, cts.Token));

                // Print the entire workbook
                workbook.PrintOut(Preview: false);

                // Wait for dialog handler
                if (!dialogTask.Wait(TimeSpan.FromSeconds(DialogTimeoutSeconds)))
                {
                    cts.Cancel();
                    throw new TimeoutException(
                        "Timed out waiting for the PDF save dialog to appear.");
                }

                if (dialogTask.IsFaulted && dialogTask.Exception != null)
                    throw dialogTask.Exception.InnerException ?? dialogTask.Exception;

                WaitForFile(outputPath);

                workbook.Close(SaveChanges: false);
                workbook = null;

                excelApp.ActivePrinter = previousPrinter;
            }
            finally
            {
                if (workbook != null)
                {
                    try { workbook.Close(SaveChanges: false); } catch { }
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        #endregion

        #region UI Automation — Save Dialog Handler

        /// <summary>
        /// Watches for the "Save Print Output As" dialog from the Microsoft Print to PDF
        /// printer driver. When found, enters the output file path and clicks Save.
        /// This runs on a background thread while PrintOut blocks on the main thread.
        /// </summary>
        private static void HandlePrintToFileDialog(string outputPath, CancellationToken token)
        {
            AutomationElement dialog = null;
            var startTime = DateTime.UtcNow;

            // Poll for the Save dialog to appear
            while (!token.IsCancellationRequested &&
                   (DateTime.UtcNow - startTime).TotalSeconds < DialogTimeoutSeconds)
            {
                dialog = FindSaveDialog();
                if (dialog != null) break;
                Thread.Sleep(250);
            }

            if (dialog == null)
            {
                if (token.IsCancellationRequested) return;
                throw new TimeoutException(
                    "The 'Save Print Output As' dialog did not appear. " +
                    "Make sure 'Microsoft Print to PDF' is installed as a printer.");
            }

            // Small delay to let the dialog fully render
            Thread.Sleep(300);

            // Find the filename edit box and set the output path
            SetDialogFileName(dialog, outputPath);

            // Small delay then click Save
            Thread.Sleep(200);
            ClickSaveButton(dialog);
        }

        /// <summary>
        /// Searches for the "Save Print Output As" dialog window.
        /// </summary>
        private static AutomationElement FindSaveDialog()
        {
            try
            {
                var root = AutomationElement.RootElement;

                // The dialog from Microsoft Print to PDF is titled "Save Print Output As"
                var condition = new AndCondition(
                    new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Window),
                    new OrCondition(
                        new PropertyCondition(
                            AutomationElement.NameProperty, "Save Print Output As"),
                        new PropertyCondition(
                            AutomationElement.NameProperty, "Save As")));

                return root.FindFirst(TreeScope.Children, condition);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Sets the filename in the Save dialog's file name field.
        /// </summary>
        private static void SetDialogFileName(AutomationElement dialog, string filePath)
        {
            // Strategy 1: Find by AutomationId "1001" (standard Save As filename combo)
            AutomationElement fileNameBox = dialog.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "1001"));

            if (fileNameBox == null)
            {
                // Strategy 2: Find by name "File name:"
                fileNameBox = dialog.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, "File name:"));
            }

            if (fileNameBox != null)
            {
                // Try ValuePattern first
                object pattern;
                if (fileNameBox.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                {
                    ((ValuePattern)pattern).SetValue(filePath);
                    return;
                }

                // If it's a combo box, find the Edit child
                var editChild = fileNameBox.FindFirst(TreeScope.Children,
                    new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Edit));

                if (editChild != null &&
                    editChild.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                {
                    ((ValuePattern)pattern).SetValue(filePath);
                    return;
                }
            }

            // Strategy 3: Use SendKeys as last resort
            // Focus the dialog and use keyboard input
            try
            {
                dialog.SetFocus();
                Thread.Sleep(100);
                System.Windows.Forms.SendKeys.SendWait(filePath);
            }
            catch
            {
                throw new InvalidOperationException(
                    "Could not set the filename in the Save dialog. " +
                    "Please try again or check UI Automation permissions.");
            }
        }

        /// <summary>
        /// Clicks the Save button in the Save dialog.
        /// </summary>
        private static void ClickSaveButton(AutomationElement dialog)
        {
            // Find the Save button (Button ID "1" in standard Save dialog)
            var saveButton = dialog.FindFirst(TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "1")));

            if (saveButton == null)
            {
                // Fallback: find by name
                saveButton = dialog.FindFirst(TreeScope.Descendants,
                    new AndCondition(
                        new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.Button),
                        new OrCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Save"),
                            new PropertyCondition(AutomationElement.NameProperty, "&Save"))));
            }

            if (saveButton != null)
            {
                object pattern;
                if (saveButton.TryGetCurrentPattern(InvokePattern.Pattern, out pattern))
                {
                    ((InvokePattern)pattern).Invoke();
                    return;
                }
            }

            // Last resort: press Enter (Save button should be default)
            try
            {
                dialog.SetFocus();
                Thread.Sleep(100);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            }
            catch
            {
                throw new InvalidOperationException(
                    "Could not click the Save button in the Save dialog.");
            }
        }

        #endregion

        #region File Helpers

        /// <summary>
        /// Waits for a file to exist and not be locked (fully written).
        /// </summary>
        private static void WaitForFile(string filePath)
        {
            var startTime = DateTime.UtcNow;
            while ((DateTime.UtcNow - startTime).TotalSeconds < FileWaitTimeoutSeconds)
            {
                if (File.Exists(filePath))
                {
                    // Try to open the file to verify it's not locked
                    try
                    {
                        using (var fs = File.Open(filePath, FileMode.Open,
                            FileAccess.Read, FileShare.None))
                        {
                            if (fs.Length > 0)
                                return; // File exists, is readable, and has content
                        }
                    }
                    catch (IOException)
                    {
                        // File is still being written — keep waiting
                    }
                }
                Thread.Sleep(500);
            }

            // Final check
            if (!File.Exists(filePath))
                throw new TimeoutException(
                    $"PDF file was not created within {FileWaitTimeoutSeconds} seconds.");
        }

        #endregion
    }
}
