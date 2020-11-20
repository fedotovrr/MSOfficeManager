using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MSOfficeManager.OpenXml
{
    /// <summary>
    /// Документ Excel
    /// </summary>
    public class ExcelDoc : IDisposable
    {
        private SpreadsheetDocument SpreadSheet;
        private List<ExcelSheet> Sheets = new List<ExcelSheet>();

        public double DefaultFontSize
        {
            get
            {
                if (SpreadSheet.WorkbookPart.WorkbookStylesPart == null || SpreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts.Descendants<Font>().Count() == 0) return 0;
                return SpreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts.Descendants<Font>().First().FontSize.Val;
            }
            set
            {
                if (SpreadSheet.WorkbookPart.WorkbookStylesPart == null) return;
                SpreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts.Descendants<Font>().First().FontSize.Val = value;
            }
        }

        
        public ExcelDoc(string path, bool overwrite)
        {
            if (Path.GetExtension(path).ToLower() != ".xlsx") throw new Exception("Поддерживается только формат .xlsx");
            bool ex = File.Exists(path);
            if (ex && overwrite)
            {
                File.Delete(path);
                ex = false;
            }
            if (ex)
            {
                SpreadSheet = SpreadsheetDocument.Open(path, true);
                RefreshStyles();
                foreach (WorksheetPart sheet in SpreadSheet.WorkbookPart.WorksheetParts)
                    Sheets.Add(new ExcelSheet(SpreadSheet, sheet));
            }
            else
            {
                SpreadSheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
                WorkbookPart wbpart = SpreadSheet.AddWorkbookPart();
                wbpart.Workbook = new Workbook();
                RefreshStyles();
            }
        }

        public ExcelDoc(Stream stream)
        {
            if (stream.Length > 0)
            {
                SpreadSheet = SpreadsheetDocument.Open(stream, true);
                RefreshStyles();
                foreach (WorksheetPart sheet in SpreadSheet.WorkbookPart.WorksheetParts)
                    Sheets.Add(new ExcelSheet(SpreadSheet, sheet));
            }
            else
            {
                SpreadSheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                WorkbookPart wbpart = SpreadSheet.AddWorkbookPart();
                wbpart.Workbook = new Workbook();
                RefreshStyles();
            }
        }

        public ExcelSheet AddSheet(string name)
        {
            ExcelSheet r = new ExcelSheet(SpreadSheet, name);
            Sheets.Add(r);
            return r;
        }

        public void SetAllFontSize(double value)
        {
            if (SpreadSheet.WorkbookPart.WorkbookStylesPart == null) return;
            Font[] f = SpreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts.Descendants<Font>().ToArray();
            for (int i = 1; i < f.Length; i++)
                f[i].FontSize.Val = value;
        }


        public void Save()
        {
            SpreadSheet.Save();
            SpreadSheet.Close();
        }

        public void Close()
        {
            SpreadSheet = null;
            for (int i = 0; i < Sheets.Count; i++)
                Sheets[i].Dispose();
            Sheets.Clear();
        }

        public void Dispose()
        {
            Close();
        }


        private void RefreshStyles()
        {
            WorkbookStylesPart stylesPart = SpreadSheet.WorkbookPart.WorkbookStylesPart == null ? SpreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>() : SpreadSheet.WorkbookPart.WorkbookStylesPart;
            if (stylesPart.Stylesheet == null || stylesPart.Stylesheet.CellFormats == null || stylesPart.Stylesheet.CellFormats.Count() < 10)
                stylesPart.Stylesheet = GenerateStyleSheet();
        }

        private static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    // Index 0 - The default font.
                    new Font(
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    // Index 1 - The default font.
                    new Font(
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    // Index 2 - The bold font.
                    new Font(
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" })
            ),

            new Fills(
                // Index 0 - The default fill.
                new Fill(
                    new PatternFill() { PatternType = PatternValues.None }),
                // Index 1 - The default fill of gray 125 (required)
                new Fill(
                    new PatternFill() { PatternType = PatternValues.Gray125 }),
                // Index 2 - The yellow fill.
                new Fill(
                    new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }) { PatternType = PatternValues.Solid })
            ),

            new Borders(
                    // Index 0 - The default border.
                    new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()),
                // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                new Border(
                    new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
            ),

            new CellFormats(
                // Index 0 - normal noborder left
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 1 - normal noborder center
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 2 - bold noborder left
                new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 3 - bold noborder center
                new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 4 - bold border left
                new CellFormat() { FontId = 2, FillId = 0, BorderId = 1, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 5 - bold border center
                new CellFormat() { FontId = 2, FillId = 0, BorderId = 1, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 6 - normal border left
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 1, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 7 - normal border center
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 1, NumberFormatId = 0, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 8 - numbering # ##0,00 normal noborder center
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, NumberFormatId = 4, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } },
                // Index 9 - numbering # ##0,00 normal border center
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 1, NumberFormatId = 4, ApplyFont = true, ApplyNumberFormat = true, Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } })
            );
        }





        private static void testShape()
        {
            SpreadsheetDocument SpreadSheet = SpreadsheetDocument.Open(@"C:\Users\Федотов РР\Desktop\test.xlsx", true);

            WorksheetPart worksheetPart = SpreadSheet.WorkbookPart.WorksheetParts.First();
            OpenXmlElement[] childs = worksheetPart.DrawingsPart.WorksheetDrawing.ChildElements.ToArray();

            Shape s = childs.First().Descendants<Shape>().First();
            OpenXmlElement[] sp = s.Descendants<ShapeProperties>().First().ChildElements.ToArray();
            OpenXmlElement[] ss = s.Descendants<ShapeStyle>().First().ChildElements.ToArray();
            OpenXmlElement[] stb = s.Descendants<TextBody>().First().ChildElements.ToArray();
            //WordManager.AddShape(sp, ss, stb);

            SpreadSheet.Close();
            return;

            ExcelDoc d2 = new ExcelDoc(@"C:\Users\Федотов РР\Desktop\test2.xlsx", true);
            d2.AddSheet("test").CreateColumns(new double[] { 9.140625, 5, 4, 3 });
            WorksheetPart wsp = d2.SpreadSheet.WorkbookPart.WorksheetParts.First();
            if (wsp.DrawingsPart == null)
            {
                wsp.AddNewPart<DrawingsPart>("rId1");
                wsp.Worksheet.AppendChild(new Drawing() { Id = "rId1" });
                wsp.DrawingsPart.WorksheetDrawing = new WorksheetDrawing();
                for (int i = 0; i < childs.Length; i++)
                {
                    if (childs[i] is TwoCellAnchor)
                    {
                        TwoCellAnchor na = childs[i].Clone() as TwoCellAnchor;
                        na.FromMarker.ColumnId.Text = "1";
                        na.FromMarker.RowId.Text = "1";
                        na.FromMarker.ColumnOffset.Text = ((int)(2.5 * 66691.28)).ToString();
                        na.FromMarker.RowOffset.Text = ((int)(7.5 * 12700)).ToString();
                        na.ToMarker.ColumnId.Text = "4";
                        na.ToMarker.RowId.Text = "4";
                        na.ToMarker.ColumnOffset.Text = "0";
                        na.ToMarker.RowOffset.Text = "0";
                        wsp.DrawingsPart.WorksheetDrawing.AppendChild(na);
                    }
                }
            }
            d2.Save();
            d2.Close();
        }
    }
}
