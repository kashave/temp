using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

//using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.AcroForms;
using UglyToad.PdfPig.AcroForms.Fields;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Outline;

namespace OCRConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\kashave\FTS\data\External paper requests for GalaXe meeting\External paper requests for GalaXe meeting\Cardholder List Master.pdf";
            Console.WriteLine(ExtractTextFromPdf(path));
            
            Console.ReadLine();
        }

        


        private static string ExtractTextFromPdf(string fileName)
        {
            StringBuilder text = new StringBuilder();
            PdfReader pdfReader = new PdfReader(fileName);
            for (int page = 1; page <= pdfReader.NumberOfPages; page++)
            {
                ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                currentText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.UTF8.GetBytes(currentText)));
                text.Append(currentText);
            }
            pdfReader.Close();
           return text.ToString();
        }
            



        }
}
