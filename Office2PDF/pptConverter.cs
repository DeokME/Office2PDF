using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Office2PDF
{
    class pptConverter
    {
        public static void convert(string originalFile, string outPutFile)
        {


            // Word Converter
            var powerPointApp = new PowerPoint.Application();
            //powerPointApp.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
            Microsoft.Office.Interop.PowerPoint.Presentations presentations = null;

            presentations = powerPointApp.Presentations;
            presentation = presentations.Open(originalFile, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue,
            Microsoft.Office.Core.MsoTriState.msoFalse);

            presentation.ExportAsFixedFormat(outPutFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen, Microsoft.Office.Core.MsoTriState.msoFalse,
                                                PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                                                Microsoft.Office.Core.MsoTriState.msoFalse, null, PowerPoint.PpPrintRangeType.ppPrintAll, string.Empty, false, true, true, true, false,
                                                Type.Missing);

            powerPointApp.Quit();
            return;
        }


    }
}
