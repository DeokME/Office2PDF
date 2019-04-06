using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Office2PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length < 3)
            {
                Console.Write("Need 3 parameters - c:\\sample\\sample.doc c:\\sample sample.pdf");
                return;
            }


            string originalFile = args[0];
            string pdfPath = args[1];
            string pdfFileName = args[2];
            string outPutFile = pdfPath + "\\" + pdfFileName;

            string extension = Path.GetExtension(originalFile).Replace(".", "");
            //Console.Write(extension);
            switch (extension)
            {
                case "doc":
                    wordConverter.convert(originalFile, outPutFile);
                    break;
                case "ppt":
                case "pptx":
                    pptConverter.convert(originalFile, outPutFile);
                    break;
                case "xlsx":
                    excelConverter.convert(originalFile, outPutFile);
                    break;

            }




        }// end Main




    }
}
