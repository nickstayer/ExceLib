using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Drawing;
using System.Windows.Forms;


namespace ExceLib
{
    public enum Turbomode
    {
        off,
        on
    }

    public class ExcelApp : IDisposable
    {
        private readonly Application _app = null;
        private Workbooks _books = null;
        private Workbook _book = null;
        public Sheets Sheets = null;
        public Worksheet Sheet = null;
        private bool isProtectedSheet = false;

        public int LastColumn { get; private set; }
        public int LastRow { get; private set; }
        public int CurrentRow { get; set; } = 1;
        private int VisibleRowsFromStartRowcounter { get; set; }
        private HashSet<int> VisibleRows { get; set; }
        public int RemainsVisibleRowsCount { get; private set; }

        public ExcelApp(bool visible = false)
        {
            VisibleRows = new HashSet<int>();
            _app = new Application() { Visible = visible };
        }

        public bool OpenDoc(string fileFullName = null)
        {
            if (fileFullName != null && fileFullName.Contains("~"))
                return false;
            _books = _app.Workbooks;
            if (fileFullName == null)
            {
                _book = _books.Add(Missing.Value);
            }
            else
            {
                if (!File.Exists(fileFullName))
                {
                    _book = _books.Add(Missing.Value);
                    _book.SaveAs(fileFullName);
                }
                _book = _books.Open(fileFullName);
            }
            Sheets = _book.Sheets;
            Sheet = Sheets.Item[1] as Worksheet;
            try
            {
                LastColumn = GetLastCell().Column;
                LastRow = GetLastCell().Row;
            }
            catch
            {
                isProtectedSheet = true;
            }
            return true;
        }

        public void CalculateVisibleRows(IProgress<int> progress, int startRow = 1)
        {
            bool catchFirstRow = false;

            for (var i = 1; i < LastRow + 1; i++)
            {
                if (i == startRow)
                {
                    catchFirstRow = true;
                }
                if (catchFirstRow && !IsRowHidden(i))
                {
                    VisibleRows.Add(i);
                }
            }
            
            VisibleRowsFromStartRowcounter = VisibleRows.Count;
            RemainsVisibleRowsCount = VisibleRows.Count;
            progress.Report(VisibleRows.Count);
        }

        public bool IsRowHidden(int row)
        {
            ExcelRange range = Sheet.Rows[row] as ExcelRange;
            if (range == null)
            {
                return (bool)range.Hidden;
            }
            return false;
        }

        private ExcelRange GetLastCell()
        {
            var lastFilledCell = Sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            return lastFilledCell;
        }

        public int GetInt(int row, int column)
        {
            var value = (Sheet.Cells[row, column] as ExcelRange).Value?.ToString();
            int.TryParse(value, out int result);
            return result;
        }

        public int GetInt(string cell)
        {
            var value = Sheet.Range[cell].Value?.ToString();
            int.TryParse(value, out int result);
            return result;
        }

        public object GetValue(int row, int column)
        {
            return (Sheet.Cells[row, column] as ExcelRange).Value;
        }

        public object GetValue(string cell)
        {
            return Sheet.Range[cell].Value;
        }

        public object GetValue2(string cell)
        {
            return Sheet.Range[cell].Value2;
        }

        public string GetText(int row, int col)
        {
            if (col == 0)
                return string.Empty;
            var text = (Sheet.Cells[row, col] as ExcelRange).Text?.ToString();
            return text;
        }

        public string GetText(int col)
        {
            if (col == 0)
                return string.Empty;
            var text = (Sheet.Cells[CurrentRow, col] as ExcelRange).Text?.ToString();
            return text;
        }

        public string GetText(string cell)
        {
            var text = Sheet.Range[cell].Text?.ToString();
            return text;
        }

        public DateTime GetDate(int row, int col)
        {
            if (col == 0)
                return default;
            var obj = (Sheet.Cells[row, col] as ExcelRange).Text?.ToString();
            DateTime.TryParse(obj, out DateTime date);
            return date;
        }

        public DateTime GetDate(int col)
        {
            if (col == 0)
                return default;
            var obj = (Sheet.Cells[CurrentRow, col] as ExcelRange).Text.ToString();
            DateTime.TryParse(obj, out DateTime date);
            return date;
        }

        public DateTime GetDate(string cell)
        {
            var obj = Sheet.Range[cell].Text.ToString();
            DateTime.TryParse(obj, out DateTime date);
            return date;
        }

        public void SetValue(object value, int row, int col)
        {
            (Sheet.Cells[row, col] as ExcelRange).Value = value;
        }

        public void SearchReplace(object source, object target)
        {
            var range = Find(source);
            while(range != null)
            {
                range.Replace(source, target,
                    Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);
                range = Find(source);
            }
        }

        public ExcelRange GetFilledRange()
        {
            if (isProtectedSheet)
                throw new Exception("Защищенный лист");
            var cell1 = "A1";
            var cell2 = GetCellAddress(LastRow, LastColumn);
            var range = Sheet.Range[cell1, cell2];
            return range;
        }

        public void SetValue(object value, string cell)
        {
            Sheet.Range[cell].Value = value;
        }

        public void WriteNote(string str, int col = 0)
        {
            if (col == 0)
            {
                var freeColumn = GetFirstEmptyColumn();
                var targetCell = Sheet.Cells[CurrentRow, freeColumn] as ExcelRange;
                targetCell.Value = str; 
            }
            else
            {
                var targetCell = Sheet.Cells[CurrentRow, col] as ExcelRange;
                targetCell.Value = str;
            }
        }

        public void Set(object str, int col)
        {
            var targetCell = Sheet.Cells[CurrentRow, col] as ExcelRange;
            targetCell.Value = str;
        }

        private int GetFirstEmptyColumn()
        {
            if (isProtectedSheet)
                throw new Exception("Защищенный лист");
            for (var i = 0; ; i++)
            {
                var freeColumn = LastColumn + 1;
                var targetCell = Sheet.Cells[CurrentRow, freeColumn + i] as ExcelRange;
                if (targetCell.Text?.ToString() == "")
                {
                    return freeColumn + i;
                }
                else
                    continue;
            }
        }

        public void SaveBook()
        {
            _book.Save();
        }

        public void SaveBookAs(string file)
        {
            _book.SaveAs(file);
        }

        public void SaveSheetAsPdf(string file)
        {
            Sheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, file, 
                XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, false);
        }

        public void CloseBook()
        {
            try
            {
                if (_book != null)
                {
                    if (Sheet != null)
                        Marshal.ReleaseComObject(Sheet);
                    _book.Close(false);
                }
            }
            catch
            {
                return;
            }
        }

        public void Quit()
        {
            try
            {
                if (_app != null)
                {
                    _app.Quit();
                    ReleaseComObjects();
                }
            }
            catch
            {
                return;
            }
        }

        public ExcelRange Find(ExcelRange range, object value)
        {
            ExcelRange result = range.Find(value, Type.Missing, XlFindLookIn.xlValues,
                XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, 
                Type.Missing, Type.Missing, Type.Missing);
            return result;
        }

        public ExcelRange Find(object value)
        {
            ExcelRange result = Sheet.Cells.Find(value, Type.Missing, XlFindLookIn.xlValues,
                XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext,
                Type.Missing, Type.Missing, Type.Missing);
            return result;
        }

        public int GetRowCountInRange(ExcelRange range)
        {
            string address = range.get_Address(XlReferenceStyle.xlA1);
            int row = int.Parse(address.Split('$').Last());
            return row;
        }

        public void InsertRow(int row, XlInsertFormatOrigin formatOrigin)
        {
            ExcelRange cellRange = (ExcelRange)Sheet.Cells[row, 1];
            ExcelRange rowRange = cellRange.EntireRow;
            rowRange.Insert(XlInsertShiftDirection.xlShiftDown, formatOrigin);
        }

        public void SetAllBorders(int r1, int c1, int r2, int c2)
        {
            ExcelRange range = GetRange(r1, c1, r2, c2);
            Borders borders = range.Borders;
            borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            borders.Color = Color.Black;
        }

        // сомнительный
        public ExcelRange GetColumn(string columnLetter)
        {
            return Sheet.Range[$"{columnLetter}:{columnLetter}"];
        }

        // сомнительный
        private int CountOfNoteColumns()
        {
            if (isProtectedSheet)
                throw new Exception("Защищенный лист");
            var lastCell = GetLastCell();
            var offset = lastCell.Column - LastColumn;
            return offset;
        }

        public string GetCellAddress(int r, int c)
        {
            ExcelRange range = (ExcelRange)Sheet.Cells[r, c];
            string address = range.get_Address(XlReferenceStyle.xlA1).Replace("$", "");
            return address;
        }

        public string GetCellAddress(int column)
        {
            int row = 1;
            ExcelRange range = (ExcelRange)Sheet.Cells[row, column];
            string address = range.get_Address(XlReferenceStyle.xlA1)
                .Replace("$", "").Replace("1", "");
            return address;
        }

        private ExcelRange GetRange(int r1, int c1, int r2, int c2)
        {
            var cell1 = GetCellAddress(r1, c1);
            var cell2 = GetCellAddress(r2, c2);
            return Sheet.Range[cell1, cell2];
        }

        private ExcelRange GetRange(int r, int c)
        {
            var cell = GetCellAddress(r, c);
            return Sheet.Range[cell];
        }

        private ExcelRange GetRange(string cell)
        {
            return Sheet.Range[cell];
        }

        private ExcelRange GetRange(string cell1, string cell2)
        {
            return Sheet.Range[cell1, cell2];
        }

        public void ColorRange(int r1, int c1, int r2, int c2, Color color)
        {
            ExcelRange range = GetRange(r1, c1, r2, c2);
            range.Interior.Color = color;
        }

        public void ColorRange(int r, int c, Color color)
        {
            ExcelRange range = GetRange(r, c);
            range.Interior.Color = color;
        }

        public void ColorRange(string cell, Color color)
        {
            ExcelRange range = GetRange(cell);
            range.Interior.Color = color;
        }

        public void ColorRange(string cell1, string cell2, Color color)
        {
            ExcelRange range = GetRange(cell1, cell2);
            range.Interior.Color = color;
        }

        public void SpeedUp(Turbomode turbomode)
        {
            if (turbomode == Turbomode.on)
            {
                _app.ScreenUpdating = false;
                _app.Calculation = XlCalculation.xlCalculationManual;
                _app.DisplayAlerts = false;
            }
            else if(turbomode == Turbomode.off)
            {
                _app.ScreenUpdating = true;
                _app.Calculation = XlCalculation.xlCalculationAutomatic;
                _app.DisplayAlerts = true;
            }
        }

        public void SwitchTo(string sheetName)
        {
            Sheet = Sheets[sheetName] as Worksheet;
        }

        private void ReleaseComObjects()
        {
            if (Sheets != null)
                Marshal.ReleaseComObject(Sheets);
            if (_book != null)
                Marshal.ReleaseComObject(_book);
            if (_books != null)
                Marshal.ReleaseComObject(_books);
            if (_app != null)
                Marshal.ReleaseComObject(_app);
        }

        public void FormatNoteColumns()
        {
            if (isProtectedSheet)
                throw new Exception("Защищенный лист");
            var countOfNoteColumns = CountOfNoteColumns();
            var lastColWidth = GetLastColumnWidth();
            for (var i = 1; i <= countOfNoteColumns; i++)
            {
                var columnLetter = GetCellAddress(LastColumn + i);
                var rangeName = $"{columnLetter}:{columnLetter}";
                var range = Sheet.get_Range(rangeName);
                range.ColumnWidth = lastColWidth;
                range.WrapText = true;
            }
        }

        private double GetLastColumnWidth()
        {
            if (isProtectedSheet)
                throw new Exception("Защищенный лист");
            var lastColumnLetter = GetCellAddress(LastColumn);
            var lastRangeName = $"{lastColumnLetter}:{lastColumnLetter}";
            var lastRange = Sheet.get_Range(lastRangeName);
            double lastColWidth = (double)lastRange.ColumnWidth;
            return lastColWidth;
        }

        public void AutoFit(string firstCell, string lastCell)
        {
            var range = Sheet.get_Range(firstCell, lastCell);
            range.EntireColumn.AutoFit();
        }

        public void ChangeOrientation(string firstCell, string lastCell, int degree)
        {
            var range = Sheet.Range[firstCell, lastCell];
            range.Orientation = degree;
        }

        public void ChangeOrientation(int r1, int c1, int r2, int c2, int degree)
        {
            var cell1 = Sheet.Cells[r1, c1];
            var cell2 = Sheet.Cells[r2, c2];
            var range = Sheet.Range[cell1, cell2];
            range.Orientation = degree;
        }

        public void FreezePanes(string cell)
        {
            var range = Sheet.get_Range(cell);
            range.Select();
            _app.ActiveWindow.FreezePanes = true;
        }

        public static bool IsExcelFile(string pathToFile)
        {
            return pathToFile.Contains("xls");
        }

        public string GetFormula(ExcelRange r)
        {
            var str = r.Formula.ToString();
            return str;
        }

        public void IncreaseRow()
        {
            if (VisibleRows.Contains(CurrentRow) && RemainsVisibleRowsCount > 0)
                RemainsVisibleRowsCount--;
            CurrentRow++;
        }

        public void SetExcelWindowHalfScreenRight()
        {
            // Получение ширины и высоты экрана
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            // Задание ширины и высоты окна Excel
            int windowWidth = (screenWidth / 2);
            int windowHeight = screenHeight;

            // Задание позиции окна Excel
            int windowLeft = windowWidth;


            // Задание размеров и позиции окна
            _app.Width = windowWidth;
            _app.Height = windowHeight;
            _app.Left = windowLeft;
        }

        public void Dispose()
        {
            _book?.Close(false);
            Quit();
            try
            {
                var process = GetExcelProcess();
                process?.Kill();
            }
            catch { }
        }

        public (int,int) FindStringAndReturnRowColumn(string searchString, int maxColumn = 0, int maxRow = 0)
        {
            var lastColumn = maxColumn == 0 ? LastColumn : maxColumn;
            var lastRow = maxRow == 0 ? LastRow : maxRow;
            for (int row = 1; row < lastRow; row++)
            {
                for (int column = 1; column < lastColumn; column++)
                {
                    var value = GetText(row, column);
                    if (!string.IsNullOrEmpty(value) && value == searchString)
                    {
                        return (row, column);
                    }
                }
            }
            return(0,0);
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public Process GetExcelProcess()
        {
            GetWindowThreadProcessId(_app.Hwnd, out int id);
            return Process.GetProcessById(id);
        }
    }
}