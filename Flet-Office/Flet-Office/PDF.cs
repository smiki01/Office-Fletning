using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace Flet_Office
{
    public class PDF
    {
        public PDF(Document skabelon, string FilePDF)
        {
            object paramMissing = Type.Missing;
            WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            WdExportOptimizeFor paramExportOptimizeFor =
                WdExportOptimizeFor.wdExportOptimizeForPrint;
            WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            WdExportCreateBookmarks paramCreateBookmarks =
                WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            skabelon.ExportAsFixedFormat(FilePDF, paramExportFormat, paramOpenAfterExport, paramExportOptimizeFor, paramExportRange,
                paramStartPage, paramEndPage, paramExportItem, paramIncludeDocProps, paramKeepIRM, paramCreateBookmarks,
                paramDocStructureTags, paramBitmapMissingFonts, paramUseISO19005_1, ref paramMissing);
        }
    }
}
