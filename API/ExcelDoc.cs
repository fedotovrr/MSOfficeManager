using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.ExceptionServices;

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
        public void ToPDF(string path)
        {
            try
            {
                HidePublishing hp = new HidePublishing();
                hp.Start();
                Document.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path);
                hp.Stop();
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
