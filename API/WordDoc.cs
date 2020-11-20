using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using MSOfficeManager.Templates;
using static MSOfficeManager.Templates.Static;

namespace MSOfficeManager.API
{
    /// <summary>
    /// Word document
    /// - Документ
    /// </summary>
    public class WordDoc : MSOfficeDocument
    {
        /// <summary>
        /// Word document
        /// - Документ
        /// </summary>
        public Document Document => Doc as Document;

        /// <summary>
        /// Word document
        /// - Документ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="app"></param>
        /// <param name="doc"></param>
        /// <param name="isReadOnly"></param>
        internal WordDoc(string path, WordApp app, Document doc, bool isReadOnly) : base(path, app, doc, isReadOnly)
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
                Document.ExportAsFixedFormat(path, WdExportFormat.wdExportFormatPDF, false);
                hp.Stop();
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
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
                for (int t = 1; t <= Document.Tables.Count; t++)
                {
                    Table table = Document.Tables[t];
                    int rc = table.Rows.Count;
                    if (rc > 1)
                    {
                        int cc = table.Columns.Count;
                        string[] header = new string[cc];
                        for (int i = 1; i <= cc; i++)
                        {
                            try
                            {
                                header[i - 1] = table.Cell(1, i).Range.Text?.Trim(GetControls());
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
                                            row[i - 1] = table.Cell(r, i).Range.Text?.Trim(GetControls());
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
        /// Получить значения Shapes по по имени
        /// </summary>
        /// <param name="identifyByName"></param>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public string[][] GetTextShapeValue(IEnumerable<Func<string, bool>> identifyByName)
        {
            try
            {
                if (identifyByName == null || Doc == null) return null;

                string[][] ret = GetTextShapeValue(identifyByName, Document.Shapes);
                for (int i = 1; i <= Document.Sections.Count; i++)
                {
                    string[][] v = GetTextShapeValue(identifyByName, Document.Sections[i].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes);
                    for (int j = 0; j < ret.Length; j++)
                        ret[j] = ret[j].Concat(v[j]).ToArray();
                    v = GetTextShapeValue(identifyByName, Document.Sections[i].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes);
                    for (int j = 0; j < ret.Length; j++)
                        ret[j] = ret[j].Concat(v[j]).ToArray();
                }
                return ret.Select(x => x.ToArray()).ToArray();
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }

        private string[][] GetTextShapeValue(IEnumerable<Func<string, bool>> identifyByName, Shapes shapes)
        {
            if (identifyByName == null || shapes == null || !identifyByName.Any()) return new string[0][];

            List<string>[] ret = identifyByName.Select(x => new List<string>()).ToArray();
            for (int t = 1; t <= shapes.Count; t++)
                if (shapes[t] != null && shapes[t].TextFrame != null && shapes[t].TextFrame.HasText == -1)
                {
                    string valueContent = shapes[t].TextFrame.TextRange.Text?.Trim(GetControls());
                    string valueName = shapes[t].Name;
                    int i = 0;
                    foreach (Func<string, bool> f in identifyByName)
                    {
                        if (f(valueName))
                            ret[i].Add(valueContent);
                        i++;
                    }
                }
            return ret.Select(x => x.ToArray()).ToArray();
        }


        /// <summary>
        /// Fill all Table with equal size
        /// - Заполнить все таблицы с указанной размерностью
        /// </summary>
        /// <param name="table"></param>
        [HandleProcessCorruptedStateExceptions]
        public void FillTable(string[,] table)
        {
            try
            {
                for (int t = 1; t <= Document.Tables.Count; t++)
                {
                    Table dtable = Document.Tables[t];
                    int rc = dtable.Rows.Count;
                    if (rc > 0 && rc == table.GetLongLength(0))
                    {
                        int cc = dtable.Columns.Count;
                        if (cc > 0 && cc == table.GetLongLength(1))
                            for (int r = 1; r <= rc; r++)
                                for (int c = 1; c <= cc; c++)
                                    if (!String.IsNullOrEmpty(table[r - 1, c - 1]))
                                    {
                                        try
                                        {
                                            dtable.Cell(r, c).Range.Text = table[r - 1, c - 1];
                                        }
                                        catch (Exception) { }
                                    }
                    }
                }
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }


        /// <summary>
        /// Fill By Template
        /// - Заполнение по шаблону
        /// </summary>
        /// <param name="template"></param>
        [HandleProcessCorruptedStateExceptions]
        public void FillByTemplate(TemplateData template)
        {
            try
            {
                if (template == null || Doc == null) return;
                FillTable(template.Tables);
                FillShape(template.FooterTextShapes, Document.Shapes);
                for (int i = 1; i <= Document.Sections.Count; i++)
                {
                    FillShape(template.FooterTextShapes, Document.Sections[i].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes);
                    FillShape(template.FooterTextShapes, Document.Sections[i].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes);
                }
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }

        private void FillShape(IEnumerable<TemplateTextShape> templateTextShapes, Shapes shapes)
        {
            if (templateTextShapes == null || shapes == null || !templateTextShapes.Any()) return;
            for (int t = 1; t <= shapes.Count; t++)
                if (shapes[t] != null && shapes[t].TextFrame != null && shapes[t].TextFrame.HasText == -1)
                {
                    string valueContent = shapes[t].TextFrame.TextRange.Text?.Trim(GetControls());
                    string valueName = shapes[t].Name;
                    foreach (TemplateTextShape ts in templateTextShapes)
                        if ((ts.IdentifyByContent != null && ts.IdentifyByContent(valueContent)) || (ts.IdentifyByName != null && ts.IdentifyByName(valueName)))
                            shapes[t].TextFrame.TextRange.Text = ts.GetValue();
                }
        }

        private void FillTable(IEnumerable<TemplateTable> tables)
        {
            if (tables == null || !tables.Any()) return;
            for (int t = 1; t <= Document.Tables.Count; t++)
            {
                Table table = Document.Tables[t];
                int rc = table.Rows.Count;
                if (rc > 1)
                {
                    int cc = table.Columns.Count;
                    string[] header = new string[cc];
                    for (int i = 1; i <= cc; i++)
                    {
                        try
                        {
                            header[i - 1] = table.Cell(1, i).Range.Text?.Trim(GetControls());
                        }
                        catch (Exception) { }
                    }
                    foreach (TemplateTable tt in tables)
                        if (tt.IsEqual(header) && tt.GetCells() is Templates.Cell[][] cells)
                        {
                            int ri = rc;
                            for (int r = 0; r < cells.Length; r++)
                                if (cells[r] != null)
                                {
                                    if (r < cells.Length - 1) table.Rows.Add();
                                    FormatRow(cells[r], true);
                                    for (int i = 0; i < cells[r].Length; i++)
                                    {
                                        Templates.Cell cell = cells[r][i];
                                        if (cell != null && cell.Column > -1 && cell.Column < cc)
                                        {
                                            int c = cell.Column + 1;
                                            int mc = cell.LastMergeColumn + 1;
                                            if (mc > cc) mc = cc;
                                            if (mc > c)
                                            {
                                                for (int j = c + 1; j <= mc; j++)
                                                {
                                                    try
                                                    {
                                                        table.Cell(ri, j).Range.Text = null;
                                                    }
                                                    catch (Exception) { }
                                                }
                                                try
                                                {
                                                    table.Cell(ri, c).Merge(table.Cell(ri, mc));
                                                }
                                                catch (Exception) { }
                                            }
                                            try
                                            {
                                                table.Cell(ri, c).Range.Text = cell.Value;
                                            }
                                            catch (Exception) { }
                                        }
                                    }
                                    ri++;
                                }
                        }
                }
            }
        }


        /// <summary>
        /// Add Paragraph with Table
        /// - Добавить параграф с таблицей
        /// </summary>
        /// <param name="columnsSize"></param>
        /// <param name="header"></param>
        /// <param name="content"></param>
        [HandleProcessCorruptedStateExceptions]
        public void AddTable(int[] columnsSize, string[] header, string[,] content)
        {
            try
            {
                if (columnsSize == null || header == null || content == null || columnsSize.Length != header.Length || columnsSize.Length != content.GetLength(1))
                    throw new Exception("Входные параметры имели не верный формат");
                if (columnsSize.Length == 0) return;

                object missing = System.Reflection.Missing.Value;
                Table table = Document.Tables.Add(Document.Range(0, 0), content.GetLength(0) + 1, content.GetLength(1), ref missing, ref missing);
                table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderVertical].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders[WdBorderType.wdBorderHorizontal].LineStyle = WdLineStyle.wdLineStyleSingle;
                for (int c = 0; c < columnsSize.Length; c++)
                {
                    table.Columns[c + 1].Width = columnsSize[c];
                    table.Columns[c + 1].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    table.Cell(1, c + 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    table.Cell(1, c + 1).Range.Font.Bold = 1;
                    table.Cell(1, c + 1).Range.Text = header[c];
                }
                for (int r = 0; r < content.GetLength(0); r++)
                    for (int c = 0; c < content.GetLength(1); c++)
                    {
                        table.Cell(r + 2, c + 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        table.Cell(r + 2, c + 1).Range.Text = content[r, c];
                    }
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }


        /// <summary>
        /// Получить структуру документа
        /// </summary>
        /// <returns></returns>
        public List<Heading> GetHeadings()
        {
            return GetHeadings(0);
        }

        /// <summary>
        /// Get Headings
        /// - Получить структуру документа
        /// </summary>
        /// <param name="PagesCount"></param>
        /// <returns></returns>
        public List<Heading> GetHeadings(int PagesCount)
        {
            try
            { 
                List<Heading> h = null;
                if (IsReadOnly)
                    h = GetHeadingsReadOnly();
                else
                    h = GetHeadingByTableOfContents();

                if (PagesCount > 0 && h != null && h.Count > 0)
                {
                    for (int i = 0; i < h.Count - 1; i++) h[i].PageOrPages = h[i + 1].PageOrPages - h[i].PageOrPages;
                    h[h.Count - 1].PageOrPages = 0;
                    h[h.Count - 1].PageOrPages = PagesCount - h.Sum(x => x.PageOrPages);
                }

                return h;
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }

        private List<Heading> GetHeadingsReadOnly()
        {
            List<Tuple<string, int, int>> h = new List<Tuple<string, int, int>>();
            int pc = Document.Paragraphs.Count;
            Parallel.For(1, pc, (i) => {
                Paragraph p = Document.Paragraphs[i];
                Style style = p.get_Style() as Style;
                if (style.ParagraphFormat.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    string header = p.Range.Text?.Trim(GetControls());
                    string sList = p.Range.ListFormat.ListString;
                    if (!string.IsNullOrEmpty(header))
                        h.Add(new Tuple<string, int, int>(sList + " " + header, (int)p.Range.Information[WdInformation.wdActiveEndAdjustedPageNumber], i));
                }
            });
            return h.OrderBy(x => x.Item3).Select(x => new Heading(x.Item1, x.Item2)).ToList();
        }

        private List<Heading> GetHeadingByTableOfContents()
        {
            //добавляем раздел
            Document.Sections[1].Range.InsertBreak(WdBreakType.wdSectionBreakNextPage);

            //добавляем содержание в документ
            try
            {
                Document.TablesOfContents.Add(Document.Sections[1].Range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {
                return new List<Heading>();
            }
            Range oRng = Document.TablesOfContents[Document.TablesOfContents.Count].Range;
            //разбиваем поле содержания
            oRng.Fields.Unlink();
            //преобразуем разбитое содержание в таблицу
            Table oTbl = oRng.ConvertToTable();

            //преобразование данных таблицы
            int columncount = oTbl.Columns.Count;
            List<string[]> Content = new List<string[]>();
            foreach (Row oRow in oTbl.Rows)
            {
                string[] str = new string[3];
                if (columncount > 0)
                {
                    switch (columncount)
                    {
                        case 1:
                            str[0] = oRow.Cells[1].Range.Text;
                            str[0] = str[0].Substring(0, str[0].Length - 2);
                            break;
                        case 2:
                            str[0] = oRow.Cells[1].Range.Text;
                            str[1] = oRow.Cells[2].Range.Text;
                            str[0] = str[0].Substring(0, str[0].Length - 2);
                            str[1] = str[1].Substring(0, str[1].Length - 2);
                            break;
                        case 3:
                            str[0] = oRow.Cells[1].Range.Text;
                            str[1] = oRow.Cells[2].Range.Text;
                            str[2] = oRow.Cells[3].Range.Text;
                            str[0] = str[0].Substring(0, str[0].Length - 2);
                            str[1] = str[1].Substring(0, str[1].Length - 2);
                            str[2] = str[2].Substring(0, str[2].Length - 2);
                            break;
                    }
                }
                Content.Add(str);
            }
            Document.Sections[1].Range.Delete();
            List<Heading> h = new List<Heading>();
            for (int i = 0; i < Content.Count; i++)
            {
                if (Content[i][2] == null || Content[i][2] == "")
                {
                    if (Content[i][0][0] != '(' || Content[i][0][Content[i][0].Length - 1] != ')')
                        h.Add(new Heading(Content[i][0], Convert.ToInt32(Content[i][1])));
                }
                else h.Add(new Heading(Content[i][0] + " " + Content[i][1], Convert.ToInt32(Content[i][2])));
            }
            return h;
        }


        /// <summary>
        /// Convert Shapes to PDF
        /// - Конвертировать Shapes документа в PDF
        /// </summary>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public List<PDFShape> ShapesToPDF()
        {
            try
            {
                List<PDFShape> r = new List<PDFShape> ();
                string tempPath = System.IO.Path.GetTempFileName();
                foreach (Shape s in Document.Shapes)
                    s.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                foreach (Shape s in Document.Shapes)
                {
                    s.Left = 0.1f;
                    s.Top = 0.1f;
                    s.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                    Document.PageSetup.TopMargin = 1f;
                    Document.PageSetup.BottomMargin = 1f;
                    Document.PageSetup.LeftMargin = 1f;
                    Document.PageSetup.RightMargin = 1f;
                    Document.PageSetup.HeaderDistance = 1f;
                    Document.PageSetup.FooterDistance = 1f;
                    Document.PageSetup.PageWidth = s.Width < 36 ? 36f + 2f : s.Width + 2f;
                    Document.PageSetup.PageHeight = s.Height < 5.2 ? 5.2f + 2f : s.Height + 2f;
                    ToPDF(tempPath);
                    r.Add(new PDFShape(s.Name, System.IO.File.ReadAllBytes(tempPath)));
                    s.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                }
                try
                {
                    System.IO.File.Delete(tempPath);
                }
                catch (Exception) { }
                return r;
            }
            catch (Exception e)
            {
                App.Close();
                throw e;
            }
        }
    }
}
