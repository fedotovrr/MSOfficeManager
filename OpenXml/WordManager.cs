using MSOfficeManager.Templates;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using static MSOfficeManager.Templates.Static;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.IO;

namespace MSOfficeManager.OpenXml
{
    /// <summary>
    /// Поддерживается только формат docx
    /// </summary>
    public static class WordManager
    {
        /// <summary>
        /// Создать пустой документ
        /// </summary>
        /// <param name="path"></param>
        public static void CreateDoc(string path)
        {
            WordprocessingDocument doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();
            doc.MainDocumentPart.Document = new Document(new Body());
            doc.Save();
            doc.Close();
        }

        /// <summary>
        /// Добавить параграф с таблицей
        /// </summary>
        /// <param name="path"></param>
        /// <param name="columnsSize"></param>
        /// <param name="header"></param>
        /// <param name="content"></param>
        public static void AddTable(string path, double[] columnsSize, string[] header, string[,] content)
        {
            WordprocessingDocument doc = OpenOrCreateDoc(path);
            AddTable(doc, columnsSize, header, content);
        }

        /// <summary>
        /// Добавить параграф с таблицей
        /// </summary>
        /// <param name="path"></param>
        /// <param name="columnsSize"></param>
        /// <param name="header"></param>
        /// <param name="content"></param>
        public static void AddTable(Stream stream, double[] columnsSize, string[] header, string[,] content)
        {
            WordprocessingDocument doc = OpenOrCreateDoc(stream);
            AddTable(doc, columnsSize, header, content);
        }

        /// <summary>
        /// Получить таблицы по заданному шаблону
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        public static object[][] GetTable(string path, IEnumerable<TemplateTable> tables)
        {
            if (Path.GetExtension(path).ToLower() != ".docx" && Path.GetExtension(path).ToLower() != ".docm") throw new Exception("Поддерживается только форматы .docx .docm");
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            object[][] ret;
            try
            {
                ret = GetTable(fs, tables);
            }
            catch (Exception e)
            {
                fs.Close();
                throw e;
            }
            fs.Close();
            return ret;
        }

        /// <summary>
        /// Получить таблицы по заданному шаблону
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        public static object[][] GetTable(Stream stream, IEnumerable<TemplateTable> tables)
        {
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, false);
            return GetTable(doc, tables);
        }

        /// <summary>
        /// Получить значения Shapes по имени
        /// </summary>
        /// <param name="path"></param>
        /// <param name="identifyByName"></param>
        /// <returns></returns>
        public static string[][] GetTextShapeValue(string path, IEnumerable<Func<string, bool>> identifyByName)
        {
            if (Path.GetExtension(path).ToLower() != ".docx" && Path.GetExtension(path).ToLower() != ".docm") throw new Exception("Поддерживается только форматы .docx .docm");
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            string[][] ret;
            try
            {
                ret = GetTextShapeValue(fs, identifyByName);
            }
            catch (Exception e)
            {
                fs.Close();
                throw e;
            }
            fs.Close();
            return ret;
        }

        /// <summary>
        /// Получить значения Shapes по имени
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="identifyByName"></param>
        /// <returns></returns>
        public static string[][] GetTextShapeValue(Stream stream, IEnumerable<Func<string, bool>> identifyByName)
        {
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, false);
            return GetTextShapeValue(doc, identifyByName);
        }

        /// <summary>
        /// Заполнить все таблицы с указанной размерностью
        /// </summary>
        /// <param name="path"></param>
        /// <param name="table"></param>
        public static void FillTable(string path, string[,] table)
        {
            if (Path.GetExtension(path).ToLower() != ".docx" && Path.GetExtension(path).ToLower() != ".docm") throw new Exception("Поддерживается только форматы .docx .docm");
            WordprocessingDocument doc = WordprocessingDocument.Open(path, true);
            FillTable(doc, table);
        }

        /// <summary>
        /// Заполнить все таблицы с указанной размерностью
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="table"></param>
        public static void FillTable(Stream stream, string[,] table)
        {
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, true);
            FillTable(doc, table);
        }

        /// <summary>
        /// Заполнение по шаблону
        /// </summary>
        /// <param name="path"></param>
        /// <param name="template"></param>
        public static void FillByTemplate(string path, TemplateData template)
        {
            if (Path.GetExtension(path).ToLower() != ".docx" && Path.GetExtension(path).ToLower() != ".docm") throw new Exception("Поддерживается только форматы .docx .docm");
            WordprocessingDocument doc = WordprocessingDocument.Open(path, true);
            FillByTemplate(doc, template);
        }

        /// <summary>
        /// Заполнение по шаблону
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="template"></param>
        public static void FillByTemplate(Stream stream, TemplateData template)
        {
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, true);
            FillByTemplate(doc, template);
        }

        private static void AddTable(WordprocessingDocument doc, double[] columnsSize, string[] header, string[,] content)
        {
            if (columnsSize == null || header == null || content == null || columnsSize.Length != header.Length || columnsSize.Length != content.GetLength(1))
                throw new Exception("Входные параметры имели не верный формат");
            if (columnsSize.Length == 0) return;

            Table table = new Table();
            TableProperties prop = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 }
                )
            );
            table.AppendChild(prop);

            TableRow head = new TableRow();
            table.AppendChild(head);
            for (int c = 0; c < columnsSize.Length; c++)
            {
                Bold bold = new Bold();
                bold.Val = OnOffValue.FromBoolean(true);

                Run r = new Run(new Text(header[c]));
                r.RunProperties = new RunProperties(bold);

                Paragraph p = new Paragraph(r);
                p.ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { Before = "0", After = "0" };

                TableCell tc = new TableCell(p);
                tc.AppendChild(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = columnsSize[c].ToString() + "pt" }));
                tc.TableCellProperties = new TableCellProperties(new TableCellMargin());
                tc.TableCellProperties.TableCellMargin.LeftMargin = new LeftMargin() { Width = "5pt" };
                tc.TableCellProperties.TableCellMargin.RightMargin = new RightMargin() { Width = "5pt" };
                tc.TableCellProperties.TableCellMargin.TopMargin = new TopMargin() { Width = "5pt" };
                tc.TableCellProperties.TableCellMargin.BottomMargin = new BottomMargin() { Width = "5pt" };
                tc.TableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                head.AppendChild(tc);
            }

            for (int r = 0; r < content.GetLength(0); r++)
            {
                TableRow tr = new TableRow();
                table.AppendChild(tr);
                for (int c = 0; c < content.GetLength(1); c++)
                {
                    Paragraph p = new Paragraph(new Run(new Text(content[r, c])));
                    p.ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Left });
                    p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { Before = "0", After = "0" };

                    TableCell tc = new TableCell(p);
                    tc.AppendChild(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = columnsSize[c].ToString() + "pt" }));
                    tc.TableCellProperties = new TableCellProperties(new TableCellMargin());
                    tc.TableCellProperties.TableCellMargin.LeftMargin = new LeftMargin() { Width = "5pt" };
                    tc.TableCellProperties.TableCellMargin.RightMargin = new RightMargin() { Width = "5pt" };
                    tc.TableCellProperties.TableCellMargin.TopMargin = new TopMargin() { Width = "5pt" };
                    tc.TableCellProperties.TableCellMargin.BottomMargin = new BottomMargin() { Width = "5pt" };
                    tc.TableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                    tr.AppendChild(tc);
                }
            }

            doc.MainDocumentPart.Document.Body.AppendChild(table);
            doc.Save();
            doc.Close();
        }

        private static object[][] GetTable(WordprocessingDocument doc, IEnumerable<TemplateTable> tables)
        {
            List<object[]> rts = new List<object[]>();
            try
            {
                if (tables != null && tables.Any())
                {
                    IEnumerable<Table> docTables = doc.MainDocumentPart.Document.Body.Descendants<Table>();
                    foreach (Table table in docTables)
                    {
                        TableRow[] rows = table.Elements<TableRow>()?.ToArray();
                        if (rows != null && rows.Length > 1)
                        {
                            string[] header = rows[0].Elements<TableCell>().Select(x => x.InnerText?.Trim(GetControls())).ToArray();
                            int cc = header.Length;
                            foreach (TemplateTable tt in tables)
                                if (tt.StartRow < rows.Length && tt.IsEqual(header))
                                {
                                    List<object> rt = new List<object>();
                                    for (int r = tt.StartRow; r < rows.Length; r++)
                                    {
                                        string[] row = rows[r].Elements<TableCell>().Select(x => x.InnerText?.Trim(GetControls())).ToArray();
                                        object item = tt.ToItem(row);
                                        if (item != null)
                                            rt.Add(item);
                                    }
                                    if (rt.Count > 0)
                                        rts.Add(rt.ToArray());
                                }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                doc.Close();
                throw e;
            }
            doc.Close();
            return rts.Count > 0 ? rts.ToArray() : null;
        }

        private static string[][] GetTextShapeValue(WordprocessingDocument doc, IEnumerable<Func<string, bool>> identifyByName)
        {
            string[][] ret = null;
            try
            {
                if (identifyByName != null && identifyByName.Any())
                {
                    ret = new string[identifyByName.Count()][];
                    int i = 0;
                    Body body = doc.MainDocumentPart.Document.Body;
                    foreach (Func<string, bool> identify in identifyByName)
                    {
                        IEnumerable<string> v = GetTextShapeValue(body.Descendants<WordprocessingShape>(), identify);
                        foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
                            v = v.Concat(GetTextShapeValue(headerPart.Header.Descendants<WordprocessingShape>(), identify));
                        foreach (FooterPart footerPart in doc.MainDocumentPart.FooterParts)
                            v = v.Concat(GetTextShapeValue(footerPart.Footer.Descendants<WordprocessingShape>(), identify));
                        ret[i] = v.ToArray();
                        i++;
                    }
                }
            }
            catch (Exception e)
            {
                doc.Close();
                throw e;
            }
            doc.Close();
            return ret;
        }

        private static IEnumerable<string> GetTextShapeValue(IEnumerable<WordprocessingShape> shapes, Func<string, bool> identifyByName)
        {
            List<string> ret = new List<string>();
            if (identifyByName == null || shapes == null || !shapes.Any()) return ret;
            foreach (WordprocessingShape shape in shapes)
                if (shape != null && !String.IsNullOrEmpty(shape.InnerText) && identifyByName(GetShapeName(shape)))
                    ret.Add(shape.InnerText);
            return ret;
        }

        private static void FillTable(WordprocessingDocument doc, string[,] table)
        {
            try
            {
                if (table != null)
                {
                    IEnumerable<Table> docTables = doc.MainDocumentPart.Document.Body.Descendants<Table>();
                    foreach (Table dtable in docTables)
                    {
                        TableRow[] rows = dtable.Elements<TableRow>()?.ToArray();
                        if (rows != null && rows.Length > 0 && rows.Length == table.GetLongLength(0))
                        {
                            int cc = rows.Select(x => x.Elements<TableCell>().Count()).Max();
                            if (cc > 0 && cc == table.GetLongLength(1))
                                for (int r = 0; r < rows.Length; r++)
                                {
                                    TableCell[] cells = rows[r].Elements<TableCell>().ToArray();
                                    for (int c = 0; c < cells.Length; c++)
                                        if (!String.IsNullOrEmpty(table[r, c]))
                                            ReplaceParagraph(cells[c], table[r, c]);
                                }
                        }
                    }
                }
                doc.Close();
            }
            catch (Exception e)
            {
                doc.Close();
                throw e;
            }
        }

        private static void FillByTemplate(WordprocessingDocument doc, TemplateData template)
        {
            try
            {
                if (template != null)
                {
                    Body body = doc.MainDocumentPart.Document.Body;

                    FillTable(body.Descendants<Table>(), template.Tables);
                    FillShape(body.Descendants<WordprocessingShape>(), template.TextShapes);
                    foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
                        FillShape(headerPart.Header.Descendants<WordprocessingShape>(), template.HeaderTextShape);
                    foreach (FooterPart footerPart in doc.MainDocumentPart.FooterParts)
                        FillShape(footerPart.Footer.Descendants<WordprocessingShape>(), template.FooterTextShapes);

                    doc.Save();
                }
                doc.Close();
            }
            catch (Exception e)
            {
                doc.Close();
                throw e;
            }
        }

        private static void FillShape(IEnumerable<WordprocessingShape> shapes, IEnumerable<TemplateTextShape> templateTextShapes)
        {
            if (templateTextShapes == null || shapes == null || !shapes.Any() || !templateTextShapes.Any()) return;
            foreach (WordprocessingShape shape in shapes)
                if (shape != null)
                    foreach (TemplateTextShape ts in templateTextShapes)
                        if ((ts.IdentifyByContent != null && ts.IdentifyByContent(shape.InnerText)) || (ts.IdentifyByName != null && ts.IdentifyByName(GetShapeName(shape))))
                        {
                            TextBoxInfo2 tb = shape.Elements<TextBoxInfo2>().Count() > 0 ? shape.Elements<TextBoxInfo2>().First() : null;
                            if (tb == null)
                            {
                                tb = new TextBoxInfo2();
                                shape.AppendChild(tb);
                            }
                            TextBoxContent tbc = tb.Elements<TextBoxContent>().Count() > 0 ? tb.Elements<TextBoxContent>().First() : null;
                            if (tb == null)
                            {
                                tbc = new TextBoxContent();
                                shape.AppendChild(tbc);
                            }
                            ReplaceParagraph(tbc, ts.GetValue());
                        }
        }

        private static void FillTable(IEnumerable<Table> docTables, IEnumerable<TemplateTable> tables)
        {
            if (docTables == null || tables == null || !docTables.Any() || !tables.Any()) return;
            foreach (Table table in docTables)
            {
                int rc = table.Elements<TableRow>().Count();
                if (rc > 1)
                {
                    string[] header = table.Elements<TableRow>().First().Elements<TableCell>().Select(x => x.InnerText?.Trim(GetControls())).ToArray();
                    int cc = header.Length;
                    foreach (TemplateTable tt in tables)
                        if (tt.IsEqual(header) && tt.GetCells() is Templates.Cell[][] cells)
                        {
                            for (int r = 0; r < cells.Length; r++)
                                if (cells[r] != null)
                                {
                                    TableRow nr = table.Elements<TableRow>().Last().Clone() as TableRow;
                                    TableCell[] tc = table.Elements<TableRow>().Last().Elements<TableCell>().ToArray();
                                    FormatRow(cells[r], false);
                                    for (int i = 0; i < cells[r].Length; i++)
                                    {
                                        Templates.Cell cell = cells[r][i];
                                        if (cell != null && cell.Column > -1 && cell.Column < cc)
                                        {
                                            int c = cell.Column;
                                            int mc = cell.LastMergeColumn;
                                            if (mc > cc - 1) mc = cc - 1;
                                            if (mc > c)
                                            {
                                                tc[c].Append(new TableCellProperties(new HorizontalMerge() { Val = MergedCellValues.Restart }));
                                                for (int p = c + 1; p <= mc; p++)
                                                    tc[p].Append(new TableCellProperties(new HorizontalMerge() { Val = MergedCellValues.Continue }));
                                            }
                                            ReplaceParagraph(tc[c], cell.Value);
                                        }
                                    }
                                    if (r < cells.Length - 1)
                                        table.AppendChild(nr);
                                }
                        }
                }
            }
        }

        private static void ReplaceParagraph(OpenXmlCompositeElement element, string value)
        {
            string[] vs = value?.Split('\n');
            if (vs == null) vs = new string[1];
            Paragraph sp = element.Elements<Paragraph>().Count() > 0 ? element.Elements<Paragraph>().First() : new Paragraph();
            element.RemoveAllChildren<Paragraph>();
            for (int i = 0; i < vs.Length; i++)
            {
                Paragraph p = sp.Clone() as Paragraph;
                element.AppendChild(p);
                Run r = p.Elements<Run>().Count() > 0 ? p.Elements<Run>().First() : new Run();
                p.RemoveAllChildren<Run>();
                p.AppendChild(r);
                r.RemoveAllChildren<Text>();
                r.AppendChild(new Text(vs[i]));
            }
        }

        private static string GetShapeName(WordprocessingShape shape)
        {
            DocProperties dp = ((shape.Parent as DocumentFormat.OpenXml.Drawing.GraphicData)?.Parent as DocumentFormat.OpenXml.Drawing.Graphic)?.Parent is Anchor a && a.Elements<DocProperties>().Count() > 0 ? a.Elements<DocProperties>().First() : null;
            return dp?.Name;
        }

        private static WordprocessingDocument OpenOrCreateDoc(string path)
        {
            if (File.Exists(path))
            {
                if (Path.GetExtension(path).ToLower() != ".docx") throw new Exception("Поддерживается только формат .docx");
                return WordprocessingDocument.Open(path, true);
            }
            else
            {
                WordprocessingDocument doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
                doc.AddMainDocumentPart();
                doc.MainDocumentPart.Document = new Document(new Body());
                return doc;
            }
        }

        private static WordprocessingDocument OpenOrCreateDoc(Stream stream)
        {
            if (stream.Length > 0)
                return WordprocessingDocument.Open(stream, true);
            else
            {
                WordprocessingDocument doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
                doc.AddMainDocumentPart();
                doc.MainDocumentPart.Document = new Document(new Body());
                return doc;
            }
        }



        private static void AddShape(OpenXmlElement[] shapeProperties, OpenXmlElement[] shapeStyle, OpenXmlElement[] textBodyProperties)
        {
            WordprocessingDocument docp = WordprocessingDocument.Open(@"C:\Users\Федотов РР\Desktop\test.docx", true);
            Body bodyp = docp.MainDocumentPart.Document.Body;
            Anchor ap = bodyp.Descendants<Paragraph>().First().Descendants<Run>().First().Descendants<AlternateContent>().First().Descendants<AlternateContentChoice>().First().Descendants<Drawing>().First().Descendants<Anchor>().First();
            WordprocessingShape wpsp = ap.Descendants<DocumentFormat.OpenXml.Drawing.Graphic>().First().Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().First().Descendants<WordprocessingShape>().First();

            WordprocessingDocument doc = WordprocessingDocument.Create(@"C:\Users\Федотов РР\Desktop\test2.docx", WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();

            Anchor a = new Anchor();
            doc.MainDocumentPart.Document = new Document(new Body(new Paragraph(new Run(bodyp.Descendants<Paragraph>().First().Descendants<Run>().First().Descendants<AlternateContent>().First().Clone() as AlternateContent))));
            a.SimplePosition = new SimplePosition() { X = 0, Y = 0 };
            a.HorizontalPosition = new HorizontalPosition(new PositionOffset("0"));
            a.VerticalPosition = new VerticalPosition(new PositionOffset("0"));
            WordprocessingShape wps = new WordprocessingShape();
            a.AppendChild(new DocumentFormat.OpenXml.Drawing.Graphic(new DocumentFormat.OpenXml.Drawing.GraphicData(wps)));
            wps.AppendChild(new NonVisualDrawingShapeProperties());
            ShapeProperties sp = new ShapeProperties();
            ShapeStyle ss = new ShapeStyle();
            TextBodyProperties tbp = new TextBodyProperties();
            for (int i = 0; i < shapeProperties.Length; i++)
                sp.AppendChild(shapeProperties[i].Clone() as OpenXmlElement);
            for (int i = 0; i < shapeStyle.Length; i++)
                ss.AppendChild(shapeStyle[i].Clone() as OpenXmlElement);
            //for (int i = 0; i < textBodyProperties.Length; i++)
            //    tbp.AppendChild(textBodyProperties[i].Clone() as OpenXmlElement);
            wps.AppendChild(sp);
            wps.AppendChild(ss);
            wps.AppendChild(tbp);

            doc.Save();
            doc.Close();
        }
    }
}