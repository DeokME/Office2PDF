using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace Office2PDF
{
    class excelConverter
    {
        public static void convert(string originalFile, string outPutFile)
        {
            
            Excel.Workbook excelWorkbook = null;
            object unknownType = Type.Missing;


           
                var excelApp = new Excel.Application();
            //excelApp.Visible = false;
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            excelWorkbook = excelApp.Workbooks.Open(originalFile, unknownType, unknownType,
                unknownType, unknownType, unknownType,
                unknownType, unknownType, unknownType,
                unknownType, unknownType, unknownType,
                unknownType, unknownType, unknownType);
            
          

            excelWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outPutFile,
                                                    unknownType, unknownType, unknownType, unknownType, unknownType,
                                                    unknownType, unknownType);
                excelApp.Quit();
                return;
            

        }
    }
}
