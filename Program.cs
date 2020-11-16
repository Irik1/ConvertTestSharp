using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConvertTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //ReadExcelFileSAX("Таблицы печатных форм.xlsx");
            //ReadExcelFileDOM("Таблицы печатных форм.xlsx");

            var list = new List<DocumentFields>();

            //var list = new List<string>();
            using (var doc =
                SpreadsheetDocument.Open("Таблицы печатных форм.xlsx", false))
            {
                var worksheet =
                    (WorksheetPart)doc.WorkbookPart.GetPartById("rId1");

                var test = worksheet.Rows().SelectMany(row => row.Cells());
                //int i = 1;

                //Split<DocumentFields>(worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.ColumnId == "B")))

                //Запрос по B
                //foreach (var cell in worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.ColumnId == "B")))
                foreach (var cell in worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.ColumnId == "C" && n.SharedString == "1-opt_m")))
                {
                    var temp = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.ColumnId == "C"));

                    var ColB = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.Column == "B" + cell.Row)).FirstOrDefault();
                    var ColC = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.Column == "C" + cell.Row)).FirstOrDefault();
                    var ColD = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.Column == "D" + cell.Row)).FirstOrDefault();
                    var ColE = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.Column == "E" + cell.Row)).FirstOrDefault();
                    var ColF = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.Column == "F" + cell.Row)).FirstOrDefault();
                    var ColG = worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.Column == "G" + cell.Row)).FirstOrDefault();
                    //Console.WriteLine(cell.GetString());
                    list.Add(new DocumentFields
                    {
                        Name = ColB.GetString(),
                        Value = "temp",
                        Type = byte.Parse(ColE.GetString()),
                        Style = new TxtStyle
                        {
                            FontSize = float.Parse(ColF.GetString())
                        },
                        MaxTulpeCount = int.Parse(ColD.GetString())
                    });
                    //i++;
                }

                //Console.WriteLine("Запрос по A:");
                //Console.WriteLine();
                //foreach (var cell in worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.ColumnId == "A")))
                //{
                //    Console.WriteLine(cell.GetString());
                //    list.Add(cell.GetString());
                //}

                //File.WriteAllLines("file1.txt", list);

                //Console.WriteLine();
                //list.Clear();

                //Console.WriteLine("Запрос по A и B:");

                //foreach (var cell in worksheet.Rows().SelectMany(row => row.Cells().Where(n => n.ColumnId == "A" || n.ColumnId == "B")))
                //{
                //    Console.WriteLine(cell.GetString());
                //    list.Add(cell.GetString());
                //}
                //File.WriteAllLines("file2.txt", list);
                doc.Close();
            }

            
            ConvertToPDF convert = new ConvertToPDF(list, "test.docx", "pdf\\","pdf\\","test2");
            string file =  convert.FillPDF();
            //File.Open(file,FileMode.Open);
            Console.ReadKey();
        }

        //static void ReadExcelFileDOM(string fileName)
        //{
        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
        //    {
        //        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        //        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
        //        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        //        string text;
        //        foreach (Row r in sheetData.Elements<Row>())
        //        {
        //            foreach (Cell c in r.Elements<Cell>())
        //            {
        //                text = c.CellValue.Text;
        //                Console.Write(text + " ");
        //            }
        //        }
        //        Console.WriteLine();
        //        Console.ReadKey();
        //    }
        //}


        //// The SAX approach.
        //static void ReadExcelFileSAX(string fileName)
        //{
        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
        //    {
        //        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        //        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

        //        OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
        //        string text;
        //        while (reader.Read())
        //        {
        //            if (reader.ElementType == typeof(CellValue))
        //            {
        //                text = reader.GetText();
        //                Console.Write(text + " ");
        //            }
        //        }
        //        Console.WriteLine();
        //        Console.ReadKey();
        //    }
        //}

    }
}
