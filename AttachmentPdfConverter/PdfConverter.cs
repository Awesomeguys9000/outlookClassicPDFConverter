using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttachmentPdfConverter
{
    /// <summary>
    /// Converts Office documents to PDF using Word/Excel's built-in PDF export.
    /// Uses ExportAsFixedFormat which works reliably in headless/COM add-in contexts.
    /// </summary>
    public static class PdfConverter
    {
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

            // Delete existing PDF if present
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
                
                // CRITICAL: Prevent macros from running or prompting (which causes RPC_E_DISCONNECTED)
                wordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                // Open with all warnings and updates suppressed
                doc = wordApp.Documents.Open(
                    FileName: inputPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    PasswordDocument: "",
                    PasswordTemplate: "",
                    Revert: false,
                    WritePasswordDocument: "",
                    WritePasswordTemplate: "",
                    Format: Type.Missing,
                    Encoding: Type.Missing,
                    Visible: false,
                    OpenAndRepair: false,
                    DocumentDirection: Type.Missing,
                    NoEncodingDialog: true,
                    XMLTransform: Type.Missing);

                // Add a small delay to let Word fully load the layout before exporting
                Thread.Sleep(500);

                // Retry loop for the actual export, in case Word is temporarily busy rendering
                int maxRetries = 3;
                bool success = false;
                Exception lastEx = null;

                for (int i = 0; i < maxRetries; i++)
                {
                    try
                    {
                        doc.ExportAsFixedFormat(
                            outputPath,
                            Word.WdExportFormat.wdExportFormatPDF,
                            OpenAfterExport: false,
                            OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                            Range: Word.WdExportRange.wdExportAllDocument,
                            Item: Word.WdExportItem.wdExportDocumentContent,
                            IncludeDocProps: true,
                            KeepIRM: true,
                            CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                            DocStructureTags: true,
                            BitmapMissingFonts: true,
                            UseISO19005_1: false);
                        
                        success = true;
                        break;
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        Thread.Sleep(1000); // Wait 1 second before retry
                    }
                }

                if (!success && lastEx != null)
                {
                    throw new Exception($"Failed to export after {maxRetries} attempts. Last error: {lastEx.Message}", lastEx);
                }

                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                doc = null;
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

                workbook = excelApp.Workbooks.Open(
                    inputPath,
                    ReadOnly: true,
                    AddToMru: false);

                workbook.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    outputPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    OpenAfterPublish: false);

                workbook.Close(SaveChanges: false);
                workbook = null;
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
    }
}
