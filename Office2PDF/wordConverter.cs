using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Office2PDF
{
    class wordConverter
    {
        public static void convert(string originalFile, string outPutFile)
        {

            // Word Converter
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            Word.Document doc = wordApp.Documents.Open(originalFile);
            doc.Activate();

            doc.SaveAs2(outPutFile, Word.WdSaveFormat.wdFormatPDF);
            doc.Close();
            wordApp.Quit();
            return;
        }

    }
}
