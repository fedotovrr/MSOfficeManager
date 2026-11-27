using Microsoft.Office.Interop.Excel;
using MSOfficeManager.Templates;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using static MSOfficeManager.Templates.Static;

namespace MSOfficeManager.API
{
    /// <summary>
    /// Document
    /// - Документ
    /// </summary>
    public class ExcelDoc : MSOfficeDocument
    {
        /// <summary>
        /// Document
        /// - Документ
        /// </summary>
        public Workbook Document => Doc as Workbook;

        public int WorksheetCount => Document == null ? 0 : Document.Worksheets.Count;

        /// <summary>
        /// Document
        /// - Документ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="app"></param>
        /// <param name="doc"></param>
        /// <param name="isReadOnly"></param>
        internal ExcelDoc(string path, ExcelApp app, Workbook doc, bool isReadOnly) : base(path, app, doc, isReadOnly)
        {

        }


        /// <summary>
        /// Convert to PDF
        /// - Конвертирование в PDF
        /// </summary>
        /// <param name="path">путь сохранения</param>
        [HandleProcessCorruptedStateExceptions]
        public void ToPDF(string path, int? WorksheetIndex = null, string WorksheetName = null)
        {
            Worksheet ws = null;
            if (WorksheetIndex != null)
            {
                foreach (Worksheet w in Document.Worksheets)
                    if (w.Index == WorksheetIndex && w.Visible == XlSheetVisibility.xlSheetVisible)
                    {
                        ws = w;
                        break;
                    }
                if (ws == null) throw new ArgumentOutOfRangeException($"Не удалось конвертировать файл {Path}: лист c индексом {WorksheetIndex} не найден или скрыт");
            }
            if (!string.IsNullOrEmpty(WorksheetName))
            {
                if (ws == null)
                {
                    foreach (Worksheet w in Document.Worksheets)
                        if (w.Name == WorksheetName && w.Visible == XlSheetVisibility.xlSheetVisible)
                        {
                            ws = w;
                            break;
                        }
                    if (ws == null) throw new ArgumentOutOfRangeException($"Не удалось конвертировать файл {Path}: лист c именем {WorksheetName} не найден или скрыт");
                }
                else if (ws.Name != WorksheetName) 
                    throw new ArgumentOutOfRangeException($"Не удалось конвертировать файл {Path}: лист {WorksheetName} не соответсвует индексу {WorksheetIndex}");
            }
            try
            {
                if (System.IO.Path.GetDirectoryName(path) is string dir && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                HidePublishing hp = new HidePublishing();
                hp.Start();
                if (ws == null) Document.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path);
                else ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path);
                hp.Stop();
            }
            catch (Exception e)
            {
                App.Close();
                throw new Exception($"Не удалось конвертировать файл {Path} путь сохранения {path}", e);
            }
        }

        /// <summary>
        /// Get all Tables by templates
        /// - Получить таблицы по заданному шаблону
        /// </summary>
        /// <param name="tables"></param>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public object[][] GetTable(IEnumerable<TemplateTable> tables)
        {
            try
            {
                if (tables == null) return null;
                List<object[]> rts = new List<object[]>();
                for (int t = 1; t <= Document.Sheets.Count; t++)
                {
                    Worksheet sheet = Document.Sheets[t] as Worksheet;
                    int rc = sheet.UsedRange.Rows.Count;
                    if (rc > 1)
                    {
                        int cc = sheet.UsedRange.Columns.Count;
                        string[] header = new string[cc];
                        for (int i = 1; i <= cc; i++)
                        {
                            try
                            {
                                header[i - 1] = ((sheet.Cells[1, i] as Range)?.Value as string)?.Trim(GetControls());
                            }
                            catch (Exception) { }
                        }
                        foreach (TemplateTable tt in tables)
                            if (tt.StartRow + 1 <= rc && tt.IsEqual(header))
                            {
                                List<object> rt = new List<object>();
                                for (int r = tt.StartRow + 1; r <= rc; r++)
                                {
                                    string[] row = new string[cc];
                                    for (int i = 1; i <= cc; i++)
                                    {
                                        try
                                        {
                                            row[i - 1] = ((sheet.Cells[r, i] as Range)?.Text as string).Trim(GetControls());
                                        }
                                        catch (Exception) { }
                                    }
                                    object item = tt.ToItem(row);
                                    if (item != null)
                                        rt.Add(item);
                                }
                                if (rt.Count > 0)
                                    rts.Add(rt.ToArray());
                            }
                    }
                }
                return rts.Count > 0 ? rts.ToArray() : null;
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }

        /// <summary>
        /// Add Sheet with table
        /// - Добавить лист с таблицей
        /// </summary>
        /// <param name="name"></param>
        /// <param name="columnsSize"></param>
        /// <param name="header"></param>
        /// <param name="content"></param>
        [HandleProcessCorruptedStateExceptions]
        public void AddTebleSheet(string name, int[] columnsSize, string[] header, string[,] content)
        {
            try
            {
                if (columnsSize == null || header == null || content == null || columnsSize.Length != header.Length || columnsSize.Length != content.GetLength(1))
                    throw new Exception("Входные параметры имели не верный формат");
                if (columnsSize.Length == 0) return;

                dynamic sheet = Document.Worksheets.Add();
                if (!string.IsNullOrEmpty(name))
                    sheet.Name = name;

                for (int c = 0; c < columnsSize.Length; c++)
                {
                    sheet.Columns[c + 1].ColumnWidth = columnsSize[c];
                    sheet.Cells[1, c + 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    sheet.Cells[1, c + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    sheet.Cells[1, c + 1].NumberFormat = "@";
                    sheet.Cells[1, c + 1].Value = header[c];
                    sheet.Cells[1, c + 1].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    sheet.Cells[1, c + 1].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    sheet.Cells[1, c + 1].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    sheet.Cells[1, c + 1].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    sheet.Cells[1, c + 1].WrapText = true;
                    sheet.Cells[1, c + 1].Font.Bold = true;
                }
                for (int r = 0; r < content.GetLength(0); r++)
                    for (int c = 0; c < content.GetLength(1); c++)
                    {
                        sheet.Cells[r + 2, c + 1].NumberFormat = "@";
                        sheet.Cells[r + 2, c + 1].Value = content[r, c];
                        sheet.Cells[r + 2, c + 1].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        sheet.Cells[r + 2, c + 1].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        sheet.Cells[r + 2, c + 1].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        sheet.Cells[r + 2, c + 1].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        sheet.Cells[r + 2, c + 1].WrapText = true;
                    }
                sheet.Rows[1].AutoFit();
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }
    }
}
