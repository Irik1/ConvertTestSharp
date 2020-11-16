using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using dotenv.net.Utilities;
//using Oblstat.Models;
using System.Diagnostics;
using System.Text.RegularExpressions;
//using System.Web.Mvc;



namespace ConvertTest
{

    public class TxtStyle
    {
        //Название шрифта
        private string fontName;
        public string FontName
        {
            get
            {
                return fontName;
            }

            set
            {
                fontName = value;
            }
        }
        //Размер шрифта
        private float fontSize;
        public float FontSize
        {
            get
            {
                return fontSize;
            }

            set
            {
                fontSize = value;
            }
        }
        //Уплотнение шрифта
        private float fontSealing;
        public float FontSealing
        {
            get
            {
                return fontSealing;
            }

            set
            {
                fontSealing = value;
            }
        }

        //Отступ
        private float paragraphIndent;
        public float ParagraphIndent
        {
            get
            {
                return paragraphIndent;
            }

            set
            {
                paragraphIndent= value;
            }
        }
        //Допустим ли перенос текста в ячейке
        bool IsMultiline;
        //Ширина ячейки
        private int cellWidth;
        public int CellWidth
        {
            get
            {
                return cellWidth;
            }

            set
            {
                cellWidth = value;
            }
        }
    }

    //Класс входящих данных
    public class DocumentFields : ICloneable
    {
        //Название поля. Совпадает с именем закладки в шаблоне
        public string Name { get; set; }
        //Обозначает тип поля, с которым имеем дело. 1 - обычный текст, 2 - таблица с вертикальными кортежами, 3 - с горизонтальными кортежами. В целом, можно разделить горизонтальные кортежа с расширяемыми ячкйками, либо же обрабатывать все одинаково - пока не знаю, какой из вариантов лучше
        private byte type;
        public byte Type
        {
            get
            {
                return type;
            }

            set
            {
                type = value;
            }
        }
        //Содержимое поле, если оно является обычным текстом закладки
        private string value;
        //Метод доступа к полю Value. Позволяет поменять значение только если поле ValueMaxLen заполнено или новое значение удовлетворяет ограничению ValueMaxLen
        public string Value
        {
            get
            {
                return value;
            }

            set
            {
                if (ValueMaxLen == 0 || value.Length <= ValueMaxLen)
                     this.value = value;
            }
        }
        //Перечень значений строки таблицы, к которой относится данная запись
        private List<List<string>> tableValue;
        public List<List<string>> TableValue
        {
            get
            {
                return tableValue;
            }

            set
            {
                tableValue = value;
            }
        }

        //Максимальный размер полей в таблице
        private List<int> tableValueMaxLen;
        public List<int> TableValueMaxLen
        {
            get
            {
                return tableValueMaxLen;
            }

            set
            {
                tableValueMaxLen = value;
            }
        }
        //Максимальный размер обычного поля
        private int valueMaxLen;
        public int ValueMaxLen
        {
            get
            {
                return valueMaxLen;
            }

            set
            {
                valueMaxLen = value;
            }
        }
        //Стиль для единичного значения
        private TxtStyle style;
        public TxtStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                style = value;
            }
        }

        //Стиль для табличных значений
        private List<TxtStyle> tableStyles;
        public List<TxtStyle> TableStyles
        {
            get
            {
                return tableStyles;
            }

            set
            {
                tableStyles = value;
            }
        }

        //Имена закладок табличных значений (зачем?)
        private List<string> tableNames;
        public List<string> TableNames
        {
            get
            {
                return tableNames;
            }

            set
            {
                tableNames = value;
            }
        }

        //Размер кортежа данных
        private int tupleSize;
        public int TupleSize
        {
            get
            {
                return tupleSize;
            }

            set
            {
                tupleSize = value;
            }
        }
        //Число кортежей
        private int tulpeCount;
        public int TulpeCount
        {
            get
            {
                return tulpeCount;
            }

            set
            {
                tulpeCount = value;
            }
        }
        //Максимальное число кортежей на одной странице
        private int maxTulpeCount;
        public int MaxTulpeCount
        {
            get
            {
                return maxTulpeCount;
            }

            set
            {
                maxTulpeCount = value;
            }
        }

        //Высота строк
        private List<float> rowHeight;
        public List<float> RowHeight
        {
            get
            {
                return rowHeight;
            }

            set
            {
                rowHeight = value;
            }
        }
        //Вычисление общей высоты таблицы
        public float CalcRowHeight()
        {
            float height=0;
            foreach (var elem in RowHeight)
            {
                height += elem;
            }
            return height;
        }
        //Вычисление высоты элемента массива с указанным индексом. Пока не реализована работа с таблицами - трогать смысла нет
        public float CalcRowHeight(int index)
        {
            float height = 0;
            foreach (var elem in RowHeight)
            {
                height += elem;
            }
            return height;
        }

        //Конструктор по умолчанию
        public DocumentFields() {
            Name = "Times New Roman";
            Type = 0;
            ValueMaxLen = 0;
            Value = "";
            TableValueMaxLen = null;
            TableValue = null;
            Style = new TxtStyle
            {
                FontName = "Times New Roman",
                FontSize = 10,
                FontSealing = 0
            };
        }
        //Конструктор с параметрами
        public DocumentFields(string Name, string Value, List<List<string>> TableValue, TxtStyle Style
            , List<int> TableValueMaxLen = null, int ValueMaxLen = 0, byte Type = 0)
        {
            this.Name = Name;
            this.Type = Type;
            this.ValueMaxLen = ValueMaxLen;
            this.Value = Value;
            this.TableValueMaxLen = TableValueMaxLen;
            this.TableValue = TableValue;
            this.Style = Style;
        }

        //Конструктор копирования
        public object Clone()
        {
            return new DocumentFields
            {
                Name = Name,
                Type = Type,
                Value = Value,
                TableValue = TableValue,
                Style = Style,
                TableValueMaxLen = TableValueMaxLen,
                ValueMaxLen = ValueMaxLen,  
            };
        }

        private bool disposed = false;
        //Деструктор
        ~DocumentFields()
        {
            Dispose(false);
        }
        //Освобождение памяти
        public void Dispose()
        {
            Dispose(true);
            // подавляем финализацию
            GC.SuppressFinalize(this);
        }
        //Освобождение памяти
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Освобождаем управляемые ресурсы
                }
                // освобождаем неуправляемые объекты
                disposed = true;
            }
        }

    }
    //Класс конвертации в PDF
    public class ConvertToPDF
    {
        private List<DocumentFields> fields;
        private string templatePath;
        private string docSavePath;
        private string PDFPath;
        private string fileName;

        //Конструктор с параметрами
        public ConvertToPDF(List<DocumentFields> documentFields, string templatePath, string PDFPath, string docSavePath, string fileName)
        {
            fields = documentFields;
            this.templatePath = templatePath;
            this.docSavePath = docSavePath;
            this.PDFPath = PDFPath;
            this.fileName = fileName;
        }

        //При отправке данных в функцию отправляем перечень всех полей, а также название шаблона документа
        //Возвращает путь
        public string FillPDF()
        {
            //Копируем шаблон. Как вариант, можно привязать в таблице БД полное название шаблона и его путь, в связи с чем пользователь отправляет название, а уже на сервере находится путь к шаблону. Либо же отправлять на сервер сразу название документа шаблона, а сервер уже сам достраивает путь к папке с данным шаблоном.
            //Открываем скопированный шаблон по типу того, что делается в ГКНТ
            string dat = DateTime.Now.ToString();
            dat = dat.Replace(" ", "_");
            dat = dat.Replace(":", "-");

            //string fileNameDoc = /*templatePath +*/ "/Temp/Doc" + "_" + dat + ".docx";
            fileName = fileName + "_" + dat;
            string fileNameDoc = docSavePath + fileName + ".docx";
            string fileNamePDF;
            fileNamePDF = PDFPath + "Doc_" + dat + ".pdf";

            //Копируем шаблон в конечную папку
            System.IO.File.Copy(templatePath, fileNameDoc);
            //Открываем файл
            using (var document = WordprocessingDocument.Open(fileNameDoc, true))
            {
                MainDocumentPart doc = document.MainDocumentPart;
                var bookmarks = doc.Document.Descendants<BookmarkStart>().ToList();
                BookmarkStart bookMarkToWriteAfter;
                int TableCount = 0;
                foreach (var elem in fields)
                {
                    //Если текущий элемент - обычная строка
                    //Пожалуй, правильнее будет в итоговом коде использовать switch вместо if
                    if (elem.Type == 1)
                    {
                        // Если это обычная запись - просто вставляем ее в заранее заготовленную закладку по типу того, что было в ГКНТ
                        bookMarkToWriteAfter = bookmarks.FirstOrDefault(bm => bm.Name == elem.Name.Trim());
                        if (bookMarkToWriteAfter != null)
                        {
                            // В целом, можно предусмотреть аварийный try на случай, если соответствующая закладка не будет обнаружена в шаблоне
                            InsertBookmarkText(bookMarkToWriteAfter, elem.Value, false);
                        }
                        else return "Ошибка";
                    }
                    //Если текущий элемент - таблица с вертикальными кортежами
                    else if (elem.Type == 2)
                    {
                        //обновляем количестао таблиц
                        TableCount++;
                        if (elem.TableValue.Count > 0)
                        {
                            int ColCount = elem.TableValue.Count;
                            List<List<string>> TableList = new List<List<string>>();
                            //Заполняем список списком данными
                            foreach (var el in elem.TableValue)
                            {
                                TableList.Add(el);
                            }
                            //Возможно придется как-то модифицировать функцию из ГКНТ, чтобы первая строка как-то выделялась
                            //Функция заполнения таблицы данными
                            AddToTableVertical(doc, TableCount, TableList);
                        }
                    }
                    else if (elem.Type == 3)
                    {
                        //обновляем количестао таблиц
                        TableCount++;
                        if (elem.TableValue.Count > 0)
                        {
                            int ColCount = elem.TableValue.Count;
                            List<List<string>> TableList = new List<List<string>>();
                            //Заполняем список списком данными
                            foreach (var el in elem.TableValue)
                            {
                                TableList.Add(el);
                            }
                            //Возможно придется как-то модифицировать функцию из ГКНТ, чтобы первая строка как-то выделялась
                            //Функция заполнения таблицы данными
                            AddToTableHorizontal(doc, TableCount, TableList);
                        }

                        //В целом, реализация такая же, как и для типа 2 с различием в виде вызываемой функции - в типе 2 вставляет данные слева направо, в типе 3 - сверху вниз
                    }
                }
                document.Close();
            }
            //Функция конвертации с помощью Libre office
            //В целом, функцию конвертации можно вывести в отдельный сервис, что будет работать с программой
            LibreOfficeConvertWordToPDF(fileNameDoc, PDFPath);
            //Возвращаем путь к сконвертированному документу
            return PDFPath +"converted\\" + fileName + ".pdf";
            //return new FilePathResult(fileSavePath + fileNamePDF, System.Net.Mime.MediaTypeNames.Application.Pdf);
        }

        //Вставка текста в закладку
        private void InsertBookmarkText(BookmarkStart bookmark, string value, bool splitByLine = false)
        {
            T getOrCreate<T, P>(P source, Func<P, T> getter, Action<T> considerNew)
            {
                var target = getter(source);
                if (target == null)
                {
                    target = Activator.CreateInstance<T>();
                    considerNew(target);
                }
                return target;
            }

            T getSiblingOrCreate<T, P>(P element)
                where T : OpenXmlElement
                where P : OpenXmlElement
                => getOrCreate(element, src => src.NextSibling<T>(), nEl => element.InsertAfterSelf(nEl));


            T getChildrenOrCreate<T, P>(P scope)
                where T : OpenXmlElement
                where P : OpenXmlElement
                => getOrCreate(scope, src => src.GetFirstChild<T>(), nEl => scope.AppendChild(nEl));

            Text getTextOf(Paragraph parent)
                => getChildrenOrCreate<Text, Run>(
                    getChildrenOrCreate<Run, Paragraph>(parent));

            Text getTextOf2(BookmarkStart bookmark1)
                => getChildrenOrCreate<Text, Run>(
                    getSiblingOrCreate<Run, BookmarkStart>(bookmark));

            if (!splitByLine)
            {
                getTextOf2(bookmark).Text = value;
            }
            else if (bookmark.Parent is Paragraph template)
            {
                value = value.Replace("\t", " ");
                var lineBreak = new Regex("\r?\n");

                var lines = lineBreak.Split(value).Select(l =>
                {
                    var it = (Paragraph)template.Clone();
                    getTextOf(it).Text = l;
                    return it;
                });
                var prev = template;
                foreach (var line in lines)
                {
                    prev.InsertAfterSelf(line);
                    prev = line;
                }
                template.Remove();
            }
        }

        //Вертикальная вставка данных в таблицу
        private static void AddToTableVertical(MainDocumentPart doc, int table_number, List<List<string>> data)
        {
            Table table = doc.Document.Body.Elements<Table>().ElementAt(table_number);
            // строка-образец
            TableRow ltr = table.Elements<TableRow>().Last();
            foreach (var item in data)
            {
                var tr = new TableRow();
                int j = 0;
                int i = 0;
                foreach (var val in item)
                {
                    // получаем форматирование очередной ячейки в строке-образце
                    TableCell template_cell = ltr.Elements<TableCell>().ElementAt(j);
                    j++;
                    Paragraph template_paragraph = template_cell.Elements<Paragraph>().First();
                    ParagraphMarkRunProperties template_run = template_paragraph.ParagraphProperties.Elements<ParagraphMarkRunProperties>().First();

                    //Создаем новую ячейку
                    var tc = new TableCell();
                    Paragraph new_paragraph = new Paragraph();
                    Run new_run = new Run();

                    // устанавливаем форматирование новой ячейки
                    new_paragraph.Append(template_paragraph.ParagraphProperties.CloneNode(true));
                    new_run.Append(template_run.CloneNode(true));
                    new_paragraph.Append(new_run);
                    //Вставляем текст в новую ячейку
                    tc.Append(new_paragraph);
                    InsertTableText(new_run, val);
                    tr.Append(tc);
                }
                i++;
                table.Append(tr);
            }
            table.RemoveChild(table.Elements<TableRow>().ElementAt(1));
        }

        //Горизонтальная вставка данных в таблицу, пока не отличается от вертикальной, потом исправлю
        private static void AddToTableHorizontal(MainDocumentPart doc, int table_number, List<List<string>> data)
        {
            Table table = doc.Document.Body.Elements<Table>().ElementAt(table_number);
            // строка-образец
            TableRow ltr = table.Elements<TableRow>().Last();
            foreach (var item in data)
            {
                var tr = new TableRow();
                int j = 0;
                int i = 0;
                foreach (var val in item)
                {
                    // получаем форматирование очередной ячейки в строке-образце
                    TableCell template_cell = ltr.Elements<TableCell>().ElementAt(j);
                    j++;
                    Paragraph template_paragraph = template_cell.Elements<Paragraph>().First();
                    ParagraphMarkRunProperties template_run = template_paragraph.ParagraphProperties.Elements<ParagraphMarkRunProperties>().First();

                    var tc = new TableCell();
                    Paragraph new_paragraph = new Paragraph();
                    Run new_run = new Run();

                    // устанавливаем форматирование новой ячейки
                    new_paragraph.Append(template_paragraph.ParagraphProperties.CloneNode(true));
                    new_run.Append(template_run.CloneNode(true));
                    new_paragraph.Append(new_run);
                    tc.Append(new_paragraph);
                    InsertTableText(new_run, val);
                    tr.Append(tc);
                }
                i++;
                table.Append(tr);
            }
            table.RemoveChild(table.Elements<TableRow>().ElementAt(1));
        }

        //Вставка текста в поле таблицы
        private static void InsertTableText(Run run, string value)
        {
            T getOrCreate<T, P>(P source, Func<P, T> getter, Action<T> considerNew)
            {
                var target = getter(source);
                if (target == null)
                {
                    target = Activator.CreateInstance<T>();
                    considerNew(target);
                }
                return target;
            }

            T getChildrenOrCreate<T, P>(P scope)
                where T : OpenXmlElement
                where P : OpenXmlElement
                => getOrCreate(scope, src => src.GetFirstChild<T>(), nEl => scope.AppendChild(nEl));

            Text getTextOf(Paragraph parent)
                => getChildrenOrCreate<Text, Run>(
                    getChildrenOrCreate<Run, Paragraph>(parent));

            if (run.Parent is Paragraph template)
            {
                value = value.Replace("\t", " ");
                var lineBreak = new Regex("\r?\n");

                var lines = lineBreak.Split(value).Select(l =>
                {
                    var it = (Paragraph)template.Clone();
                    getTextOf(it).Text = l;
                    return it;
                });
                var prev = template;
                foreach (var line in lines)
                {
                    prev.InsertAfterSelf(line);
                    prev = line;
                }
                template.Remove();
            }
        }

        //Конвертация WORD в PDF, пока что функцией. Потом будет отдельным сервисом
        private void LibreOfficeConvertWordToPDF(string fileNameDoc, string fileSavePath)
        {
            ////путь к исполняемому файлу LibreOffice
            //var envReader = new EnvReader();
            //if (!envReader.TryGetStringValue("LIBREOFFICE_PATH", out string libreOfficePath))
            //    libreOfficePath = @"D:\LibreOffice\App\libreoffice\program\soffice.exe";
            ////libreOfficePath = @"C:\LibreOfficePortable\App\libreoffice\program\soffice.exe";
            string libreOfficePath = @"D:\LibreOffice\App\libreoffice\program\soffice.exe";
            string folder = "converted";
            Process pdfProcess = new Process();
            pdfProcess.StartInfo.FileName = libreOfficePath;
            pdfProcess.StartInfo.Arguments = "-env:UserInstallation=file:///C:/test/NPP  --norestore --nofirststartwizard --headless --convert-to pdf \"" + fileNameDoc + "\" --outdir \"" + folder + "\"";
            pdfProcess.StartInfo.WorkingDirectory = fileSavePath;
            pdfProcess.Start();
            pdfProcess.WaitForExit();
        }
    }
}