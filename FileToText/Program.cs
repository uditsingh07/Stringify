using System;
using GemBox.Document;
using GemBox.Pdf;
using GemBox.Spreadsheet;
using GemBox.Presentation;
using System.Text;
using System.Linq;
using System.IO;

namespace FileToText
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("File Name: ");
            string path = Console.ReadLine();
            string strpath = System.IO.Path.GetExtension(path);


            //key for using GemBox
            GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            GemBox.Presentation.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            //Extracting data from .txt
            if (strpath == ".txt")
            {
                var data = System.IO.File.ReadAllText(@path);
                Console.WriteLine(data);
                Console.WriteLine("\n");
            }
            //Extracting data from .docx
            else if (strpath == ".docx" || strpath == "doc")
            {
                var doc_Data = DocumentModel.Load(@path);
                string doc_text = doc_Data.Content.ToString();
                Console.WriteLine(doc_text);
            }
            //Extracting data from .pdf
            else if (strpath == ".pdf")
                //using foreach to iterate through pages
                using (var document = PdfDocument.Load(@path))
                {
                    foreach (var page in document.Pages)
                    {
                        Console.WriteLine(page.Content.ToString());
                    }
                }


            //Extracting data from excel
            

            //Console.WriteLine(value.GetType().FullName);
            Console.ReadKey();
        }
    }
}
