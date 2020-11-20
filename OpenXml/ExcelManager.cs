using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;

namespace MSOfficeManager.OpenXml
{
    public static class ExcelManager
    {
        /// <summary>
        /// Создать пустой документ
        /// </summary>
        /// <param name="path"></param>
        public static void CreateDoc(string path, bool isCreateDefaultSheet)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            CreateDoc(doc, isCreateDefaultSheet);
        }

        /// <summary>
        /// Создать пустой документ
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="isCreateDefaultSheet"></param>
        public static void CreateDoc(Stream stream, bool isCreateDefaultSheet)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            CreateDoc(doc, isCreateDefaultSheet);
        }

        /// <summary>
        /// Добавить лист с таблицей
        /// </summary>
        /// <param name="path"></param>
        /// <param name="name"></param>
        /// <param name="columnsSize"></param>
        /// <param name="header"></param>
        /// <param name="content"></param>
        public static void AddTable(string path, string name, double[] columnsSize, string[] header, string[,] content)
        {
            ExcelDoc d = new ExcelDoc(path, false);
            AddTable(d, name, columnsSize, header, content);
        }

        /// <summary>
        /// Добавить лист с таблицей
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="name"></param>
        /// <param name="columnsSize"></param>
        /// <param name="header"></param>
        /// <param name="content"></param>
        public static void AddTable(Stream stream, string name, double[] columnsSize, string[] header, string[,] content)
        {
            ExcelDoc d = new ExcelDoc(stream);
            AddTable(d, name, columnsSize, header, content);
        }

        private static void CreateDoc(SpreadsheetDocument doc, bool isCreateDefaultSheet)
        {
            WorkbookPart wb = doc.AddWorkbookPart();
            wb.Workbook = new Workbook();

            if (isCreateDefaultSheet)
            {
                WorksheetPart ws = wb.AddNewPart<WorksheetPart>();
                ws.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = wb.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = doc.WorkbookPart.
                    GetIdOfPart(ws),
                    SheetId = 1,
                    Name = "Sheet"
                };
                sheets.Append(sheet);
            }

            wb.Workbook.Save();
            doc.Close();
        }

        private static void AddTable(ExcelDoc d, string name, double[] columnsSize, string[] header, string[,] content)
        {
            if (columnsSize == null || header == null || content == null || columnsSize.Length != header.Length || columnsSize.Length != content.GetLength(1))
                throw new Exception("Входные параметры имели не верный формат");
            if (columnsSize.Length == 0) return;

            ExcelSheet s = d.AddSheet(name);
            s.CreateColumns(columnsSize);
            for (int c = 0; c < columnsSize.Length; c++)
                if (!string.IsNullOrEmpty(header[c]) && header[c][0] == '=')
                    s.SetCellFormulaValue(0, c, header[c], false, ValueTypes.String, CellStyles.BoldBorderCenter);
                else
                    s.SetCellStringValue(0, c, header[c], ValueTypes.String, CellStyles.BoldBorderCenter);
            for (int r = 0; r < content.GetLength(0); r++)
                for (int c = 0; c < content.GetLength(1); c++)
                    if (!string.IsNullOrEmpty(content[r, c]) && content[r, c][0] == '=')
                        s.SetCellFormulaValue(r + 1, c, content[r, c], false, ValueTypes.String, CellStyles.NormalBorderLeft);
                    else
                        s.SetCellStringValue(r + 1, c, content[r, c], ValueTypes.String, CellStyles.NormalBorderLeft);
            d.Save();
            d.Close();
        }
    }
}