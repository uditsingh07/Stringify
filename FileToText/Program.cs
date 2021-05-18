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
            Console.WriteLine("File path as(C:\\Users\\Desktop\\fileConverson\\<FileName>.<FileExtension>)");
            Console.WriteLine("Enter file path: ");
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
            else if (strpath == ".xlsx")
            {
                ExcelFile spreadsheet = ExcelFile.Load(@path);

                //Loop to move between multiple sheets
                foreach (ExcelWorksheet worksheet in spreadsheet.Worksheets)
                {
                    Console.WriteLine(worksheet.Name);

                    //Loop for rows
                    foreach (ExcelRow row in worksheet.Rows)
                    {
                        //Loop for cells in a row
                        foreach (ExcelCell cell in row.AllocatedCells)
                        {
                            //Reading data from cell
                            string value = cell.Value?.ToString() ?? "EMPTY";


                            Console.Write($"{value}".PadRight(20));

                        }
                        Console.WriteLine("\n");
                    }
                }

            }

            //Extracting data from ppt
            else if (strpath == ".pptx" || strpath == ".ppt")
            {
                var presentation = PresentationDocument.Load(@path);
                var sb = new StringBuilder();
                var i = 0;
                try
                {
                    while (presentation.Slides[i] != null)
                    {
                        var slide = presentation.Slides[i];
                        foreach (var shape in slide.Content.Drawings.OfType<Shape>())
                        {

                            sb.AppendLine();

                            foreach (var paragraph in shape.Text.Paragraphs)
                            {
                                foreach (var run in paragraph.Elements.OfType<TextRun>())
                                {
                                    var isBold = run.Format.Bold;
                                    var text = run.Text;

                                    sb.AppendFormat("{0}{1}{2}", isBold ? "<b>" : "", text, isBold ? "</b>" : "");
                                }

                                sb.AppendLine();
                            }


                        }

                        sb.ToString();

                        if (sb == null)
                        {
                            i = -1;
                        }
                        else
                        {
                            i++;
                        }
                    }

                }
                catch
                {
                    if (sb != null)
                    {
                        Console.WriteLine(sb);

                    }
                }

            }
            

            
            Console.ReadKey();
        }
    }
}
