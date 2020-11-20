using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace MSOfficeManager.OpenXml
{
    public class ExcelSheet : IDisposable
    {
        private SpreadsheetDocument SpreadSheet;
        private WorksheetPart WorksheetPart;
        private Sheet Sheet;

        public string Name => Sheet.Name;

        public SheetCells Cells { get; private set; }

     
        internal ExcelSheet(SpreadsheetDocument spreadSheet, string name)
        {
            SpreadSheet = spreadSheet;

            WorksheetPart = SpreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
            WorksheetPart.Worksheet = new Worksheet();
            WorksheetPart.Worksheet.AppendChild(new SheetData());

            Sheets sheets = SpreadSheet.WorkbookPart.Workbook.Descendants<Sheets>().Count() == 0 ? SpreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets()) : SpreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = SpreadSheet.WorkbookPart.GetIdOfPart(WorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            if (string.IsNullOrEmpty(name)) name = "Sheet";
            if (sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == name) != null)
            {
                int nc = 1;
                string tname = name + nc;
                while (sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == tname) != null)
                {
                    tname = name + nc;
                    nc++;
                }
                name = tname;
            }

            Sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = name };
            sheets.Append(Sheet);

            SetPageSetup();
            SetPageMargins();

            Cells = new SheetCells(this);
        }

        internal ExcelSheet(SpreadsheetDocument spreadSheet, WorksheetPart worksheetPart)
        {
            SpreadSheet = spreadSheet;
            WorksheetPart = worksheetPart;
            string IdOfPart = SpreadSheet.WorkbookPart.GetIdOfPart(WorksheetPart);
            Sheet = SpreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Descendants<Sheet>().FirstOrDefault(x => x.Id.Value == IdOfPart);

            Cells = new SheetCells(this);
        }

        public void Dispose()
        {
            SpreadSheet = null;
            WorksheetPart = null;
            Sheet = null;
            Cells.Dispose();
        }


        public void CreateColumns(double[] columnsSize)
        {
            Columns columns;
            if (WorksheetPart.Worksheet.Elements<Columns>().Count() > 0)
                columns = WorksheetPart.Worksheet.Elements<Columns>().First();
            else
            {
                columns = new Columns();
                if (WorksheetPart.Worksheet.Elements<SheetData>().Count() > 0)
                    WorksheetPart.Worksheet.InsertBefore(columns, WorksheetPart.Worksheet.Elements<SheetData>().First());
                else
                    WorksheetPart.Worksheet.AppendChild(columns);
            }

            columns.RemoveAllChildren();
            for (int c = 0; c < columnsSize.Length; c++)
                columns.Append(new Column() { Min = (UInt32)(c + 1), Max = (UInt32)(c + 1), Width = columnsSize[c] > 0 ? columnsSize[c] + 0.7109375 : columnsSize[c], CustomWidth = true });
        }

        public void SetCellStringValue(string cell1, string cell2, string value, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            SetCellStringValue(GetRowIndex(cell1), GetColumnIndex(cell1), GetRowIndex(cell2), GetColumnIndex(cell2), value, vtype, style);
        }

        public void SetCellStringValue(int row1, int column1, int row2, int column2, string value, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            int frow = Math.Min(row1, row2);
            int lrow = Math.Max(row1, row2);
            int fcolumn = Math.Min(column1, column2);
            int lcolumn = Math.Max(column1, column2);
            for (int r = frow; r <= lrow; r++)
                for (int c = fcolumn; c <= lcolumn; c++)
                    SetCellStringValue(r, c, value, vtype, style);
        }

        public void SetCellStringValue(string cell, string value, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            SetCellStringValue(GetRowIndex(cell), GetColumnIndex(cell), value, vtype, style);
        }

        public void SetCellStringValue(int row, int column, string value, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            Cell cell = InsertCellInWorksheet(row + 1, column + 1, WorksheetPart);
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>((CellValues)vtype);
            cell.StyleIndex = (uint)style;
        }

        public void SetCellFormulaValue(string cell1, string cell2, string value, bool isArray, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            SetCellFormulaValue(GetRowIndex(cell1), GetColumnIndex(cell1), GetRowIndex(cell2), GetColumnIndex(cell2), value, isArray, vtype, style);
        }

        public void SetCellFormulaValue(int row1, int column1, int row2, int column2, string value, bool isArray, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            int frow = Math.Min(row1, row2);
            int lrow = Math.Max(row1, row2);
            int fcolumn = Math.Min(column1, column2);
            int lcolumn = Math.Max(column1, column2);
            for (int r = frow; r <= lrow; r++)
                for (int c = fcolumn; c <= lcolumn; c++)
                    SetCellFormulaValue(r, c, value, isArray, vtype, style);
        }

        public void SetCellFormulaValue(string cell, string value, bool isArray, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            SetCellFormulaValue(GetRowIndex(cell), GetColumnIndex(cell), value, isArray, vtype, style);
        }

        public void SetCellFormulaValue(int row, int column, string value, bool isArray, ValueTypes vtype = ValueTypes.String, CellStyles style = CellStyles.NormalNoborderLeft)
        {
            Cell cell = InsertCellInWorksheet(row + 1, column + 1, WorksheetPart);
            cell.CellFormula = new CellFormula(value);
            if (isArray)
            {
                cell.CellFormula.FormulaType = CellFormulaValues.Array;
                cell.CellFormula.Reference = cell.CellReference;
            }
            cell.DataType = new EnumValue<CellValues>((CellValues)vtype);
            cell.StyleIndex = (uint)style;
        }

        public void MergeCell(string cell1, string cell2)
        {
            MergeCells mergeCells;
            if (WorksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
                mergeCells = WorksheetPart.Worksheet.Elements<MergeCells>().First();
            else
            {
                mergeCells = new MergeCells();
                if (WorksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                    WorksheetPart.Worksheet.InsertAfter(mergeCells, WorksheetPart.Worksheet.Elements<CustomSheetView>().First());
                else
                    WorksheetPart.Worksheet.InsertAfter(mergeCells, WorksheetPart.Worksheet.Elements<SheetData>().First());
            }

            MergeCell mergeCell = new MergeCell();
            mergeCell.Reference = new StringValue(cell1 + ":" + cell2);
            mergeCells.Append(mergeCell);
        }

        public void MergeCell(int row1, int column1, int row2, int column2)
        {
            MergeCell(GetCellReference(row1 + 1, column1 + 1), GetCellReference(row2 + 1, column2 + 1));
        }

        public void InsertHorizontalPageBreak(int rowIndex)
        {
            rowIndex++;
            RowBreaks rb = WorksheetPart.Worksheet.GetFirstChild<RowBreaks>();
            if (rb == null)
            {
                rb = new RowBreaks();
                rb.ManualBreakCount = (UInt32Value)0;
                rb.Count = (UInt32Value)0;
                WorksheetPart.Worksheet.Append(rb);
            }
            Break rowBreak = new Break() { Id = (UInt32Value)(uint)rowIndex, Max = (UInt32Value)16383U, ManualPageBreak = true };
            rb.Append(rowBreak);
            rb.ManualBreakCount++;
            rb.Count++;
        }

        public void InsertVerticalPageBreak(int columnIndex)
        {
            columnIndex++;
            ColumnBreaks cb = WorksheetPart.Worksheet.GetFirstChild<ColumnBreaks>();
            if (cb == null)
            {
                cb = new ColumnBreaks();
                cb.ManualBreakCount = (UInt32Value)0;
                cb.Count = (UInt32Value)0;
                WorksheetPart.Worksheet.Append(cb);
            }
            Break br = new Break() { Id = (UInt32Value)(uint)columnIndex, Max = (UInt32Value)1048575U, ManualPageBreak = true };
            cb.Append(br);
            cb.ManualBreakCount++;
            cb.Count++;
        }

        public void SetPageSetup(int paperSize = 9, bool isHorizontal = false, int fitToWidth = 0, int fitToHeight = 0,  int scale = 100)
        {
            PageSetup pageSetup;
            if (WorksheetPart.Worksheet.Elements<PageSetup>().Count() > 0)
                pageSetup = WorksheetPart.Worksheet.Elements<PageSetup>().First();
            else
            {
                pageSetup = new PageSetup();
                if (WorksheetPart.Worksheet.Elements<PageMargins>().Count() > 0)
                    WorksheetPart.Worksheet.InsertAfter(pageSetup, WorksheetPart.Worksheet.Elements<PageMargins>().First());
                else if (WorksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                    WorksheetPart.Worksheet.InsertAfter(pageSetup, WorksheetPart.Worksheet.Elements<CustomSheetView>().First());
                else
                    WorksheetPart.Worksheet.InsertAfter(pageSetup, WorksheetPart.Worksheet.Elements<SheetData>().First());
            }
            pageSetup.PaperSize = (uint)paperSize;
            pageSetup.Orientation = isHorizontal ? OrientationValues.Landscape : OrientationValues.Portrait;
            pageSetup.FitToWidth = (uint)fitToWidth;
            pageSetup.FitToHeight = (uint)fitToHeight;
            pageSetup.Scale = (uint)scale;
        }

        public void SetPageMargins(double left = 0.25, double right = 0.25, double top = 0.75, double bottom = 0.75, double header = 0.3, double footer = 0.3)
        {
            PageMargins margins;
            if (WorksheetPart.Worksheet.Elements<PageMargins>().Count() > 0)
                margins = WorksheetPart.Worksheet.Elements<PageMargins>().First();
            else
            {
                margins = new PageMargins();
                if (WorksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                    WorksheetPart.Worksheet.InsertAfter(margins, WorksheetPart.Worksheet.Elements<CustomSheetView>().First());
                else
                    WorksheetPart.Worksheet.InsertAfter(margins, WorksheetPart.Worksheet.Elements<SheetData>().First());
            }
            margins.Left = left;
            margins.Right = right;
            margins.Top = top;
            margins.Bottom = bottom;
            margins.Header = header;
            margins.Footer = footer;
        }



        private Cell GetCell(int indexRow, int indexColumn)
        {
            return InsertCellInWorksheet(indexRow + 1, indexColumn + 1, WorksheetPart);
        }

        private static Cell InsertCellInWorksheet(int indexRow, int indexColumn, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = GetCellReference(indexRow, indexColumn);

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == indexRow).Count() != 0)
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == indexRow).First();
            else
            {
                row = new Row() { RowIndex = (uint)indexRow };
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Count() > 0)
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                return newCell;
            }
        }

        private static int GetColumnIndex(string cellReference)
        {
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);
                mulitplier = mulitplier * 26;
            }

            return columnNumber;
        }

        private static int GetRowIndex(string cellReference)
        {
            string s = Regex.Replace(cellReference, @"[^0-9]", string.Empty);
            return Convert.ToInt32(s) - 1;
        }

        private static string GetCellReference(int row, int column)
        {
            if (row < 1 || column < 1) throw new Exception("Не корректная ссылка");
            string str = string.Empty;
            while (column > 0)
            {
                str = (char)(column % 26 + 64) + str;
                column /= 26;
            }
            return str + row;
        }


        public class SheetCells
        {
            private ExcelSheet Parent;

            public SheetCell this[int row, int column] => new SheetCell(Parent.GetCell(row, column));

            public SheetCell this[string cell] => new SheetCell(Parent.GetCell(GetRowIndex(cell), GetColumnIndex(cell)));

            internal SheetCells(ExcelSheet parent)
            {
                Parent = parent;
            }

            internal void Dispose()
            {
                Parent = null;
            }
        }

        public class SheetCell
        {
            private Cell Cell;

            public string Value { get => Cell.CellValue.InnerText; set => Cell.CellValue = new CellValue(value); }

            public string Formula { get => Cell.CellFormula.InnerText; set => Cell.CellFormula = new CellFormula(value); }

            public string FormulaArray
            {
                get
                {
                    return Cell.CellFormula.InnerText;
                }
                set
                {
                    Cell.CellFormula = new CellFormula(value);
                    Cell.CellFormula.FormulaType = CellFormulaValues.Array;
                    Cell.CellFormula.Reference = Cell.CellReference;
                }
            }

            public ValueTypes ValueType { set => Cell.DataType = new EnumValue<CellValues>((CellValues)value); }

            public CellStyles Style { set => Cell.StyleIndex = (uint)value; }


            internal SheetCell(Cell cell)
            {
                Cell = cell;
            }
        }
    }
}