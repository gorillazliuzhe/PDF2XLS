using SautinSoft;
using System;
using System.IO;

namespace PDF2XLS
{
    /// <summary>
    /// https://www.sautinsoft.com/products/pdf-focus/download.php
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            string pdfFile = @"c:\pdf\test.pdf";
            //PdfToExcelAsFiles(pdfFile);
            PdfToXMLAsFiles(pdfFile);
            Console.WriteLine("ok");
            Console.ReadKey();
        }

        public static void PdfToExcelAsFiles(string pdfFile)
        {
            try
            {
               
                string excelFile = Path.ChangeExtension(pdfFile, ".xls");

                PdfFocus f = new PdfFocus();
                // 'true' = Convert all data to spreadsheet (tabular and even textual).
                // 'false' = Skip textual data and convert only tabular (tables) data.
                f.ExcelOptions.ConvertNonTabularDataToSpreadsheet = false;

                // 'true'  = Preserve original page layout.
                // 'false' = Place tables before text.
                f.ExcelOptions.PreservePageLayout = true;

                f.OpenPdf(pdfFile);

                if (f.PageCount > 0)
                {
                    f.ToExcel(excelFile);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
           
        }

        public static void PdfToXMLAsFiles(string pdfFile)
        {

            string pathToXml = Path.ChangeExtension(pdfFile, ".xml");

            // Convert PDF file to XML file.
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();

            // This property is necessary only for registered version.
            //f.Serial = "XXXXXXXXXXX";

            // Let's convert only tables to XML and skip all textual data.
            f.XmlOptions.ConvertNonTabularDataToSpreadsheet = false;

            f.OpenPdf(pdfFile);

            if (f.PageCount > 0)
            {
                int result = f.ToXml(pathToXml);

                //Show XML document in browser
                if (result == 0)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pathToXml) { UseShellExecute = true });
                }
            }
        }
    }
}
