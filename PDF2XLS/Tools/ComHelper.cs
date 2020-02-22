﻿using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using PDF2XLS.Models;
using SautinSoft;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace PDF2XLS.Tools
{
    /// <summary>
    /// https://www.sautinsoft.com/products/pdf-focus/download.php
    /// </summary>
    public static class ComHelper
    {
        /// <summary>
        /// pdf生成xml
        /// </summary>
        /// <param name="pdfFile"></param>
        /// <returns></returns>
        public static bool PdfToXMLAsFiles(string pdfFile)
        {
            try
            {
                string pathToXml = Path.ChangeExtension(pdfFile, ".xml");

                // Convert PDF file to XML file.
                PdfFocus f = new PdfFocus();

                // This property is necessary only for registered version.
                // f.Serial = "XXXXXXXXXXX";

                // Let's convert only tables to XML and skip all textual data.
                f.XmlOptions.ConvertNonTabularDataToSpreadsheet = false;

                f.OpenPdf(pdfFile);

                if (f.PageCount > 0)
                {
                    int result = f.ToXml(pathToXml);
                    if (result == 0)
                    {
                        //Show XML document in browser 选择直接打开
                        // Process.Start(new ProcessStartInfo(pathToXml) { UseShellExecute = true });
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return false;
        }

        /// <summary>
        /// pdf生成excel
        /// </summary>
        /// <param name="pdfFile"></param>
        /// <returns></returns>
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

        
        public static async Task Export(string path, List<Bsjh> bsjhs)
        {
            IExporter exporter = new ExcelExporter();
            var result = await exporter.Export(path, bsjhs);
        }

    }
}
