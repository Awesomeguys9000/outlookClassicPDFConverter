using System;

namespace AttachmentPdfConverter
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Add-in loaded — ribbon is created via CreateRibbonExtensibilityObject
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Clean up if needed
        }

        protected override object RequestService(Guid serviceGuid)
        {
            if (serviceGuid == typeof(Microsoft.Office.Core.IRibbonExtensibility).GUID)
            {
                return new PdfRibbon();
            }
            return base.RequestService(serviceGuid);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
