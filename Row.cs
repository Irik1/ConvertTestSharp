using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ConvertTest
{
    public class Row
    {
        public XElement RowElement { get; set; }
        public string RowId { get; set; }
        public string Spans { get; set; }
        public IEnumerable<Cell> Cells()
        {
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            SpreadsheetDocument doc = (SpreadsheetDocument)Parent.OpenXmlPackage;
            SharedStringTablePart sharedStringTable = doc.WorkbookPart.SharedStringTablePart;
            return
                from cell in this.RowElement.Elements(s + "c")
                let cellType = (string)cell.Attribute("t")
                let sharedString = cellType == "s" ?
                    sharedStringTable
                    .GetXDocument()
                    .Root
                    .Elements(s + "si")
                    .Skip((int)cell.Element(s + "v"))
                    .First()
                    .Descendants(s + "t")
                    .StringConcatenate(e => (string)e)
                    : null
                let column = (string)cell.Attribute("r")
                select new Cell(this)
                {
                    CellElement = cell,
                    Row = (string)RowElement.Attribute("r"),
                    Column = column,
                    ColumnId = column.Split('0', '1', '2', '3', '4', '5', '6', '7', '8', '9').First(),
                    Type = (string)cell.Attribute("t"),
                    Formula = (string)cell.Element(s + "f"),
                    Value = (string)cell.Element(s + "v"),
                    SharedString = sharedString
                };
        }
        public WorksheetPart Parent { get; set; }
        public Row(WorksheetPart parent) { Parent = parent; }

    }

    public class Cell
    {
        public XElement CellElement { get; set; }
        public string Row { get; set; }
        public string Column { get; set; }
        public string ColumnId { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }
        public string SharedString { get; set; }
        public Row Parent { get; set; }
        public Cell(Row parent) { Parent = parent; }

        public string GetString()
        {
            return SharedString ?? Value;
        }
    }

    public static class LocalExtensions
    {
        public static IEnumerable<Row> Rows(this WorksheetPart worksheetPart)
        {
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return
                from row in worksheetPart
                    .GetXDocument()
                    .Root
                    .Element(s + "sheetData")
                    .Elements(s + "row")
                select new Row(worksheetPart)
                {
                    RowElement = row,
                    RowId = (string)row.Attribute("r"),
                    Spans = (string)row.Attribute("spans")
                };
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source,
            Func<T, string> func)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item));
            return sb.ToString();
        }
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (XmlReader xr = XmlReader.Create(sr))
                xdoc = XDocument.Load(xr);
            part.AddAnnotation(xdoc);
            return xdoc;
        }
    }
}
