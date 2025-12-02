using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Models;
using System.Collections.Generic;

namespace FluentNPOI
{
    public class FluentSheet : FluentSheetBase
    {
        public FluentSheet(IWorkbook workbook, ISheet sheet, Dictionary<string, ICellStyle> cellStylesCached)
            : base(workbook, sheet, cellStylesCached)
        {
        }

        public ISheet GetSheet()
        {
            return _sheet;
        }

        public FluentSheet SetColumnWidth(ExcelColumns col, int width)
        {
            _sheet.SetColumnWidth((int)col, width * 256);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        public FluentSheet SetColumnWidth(ExcelColumns startCol, ExcelColumns endCol, int width)
        {
            for (int i = (int)startCol; i <= (int)endCol; i++)
            {
                _sheet.SetColumnWidth(i, width * 256);
            }
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        public FluentSheet SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int row)
        {
            _sheet.SetExcelCellMerge(startCol, endCol, row);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        public FluentSheet SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int firstRow, int lastRow)
        {
            _sheet.SetExcelCellMerge(startCol, endCol, firstRow, lastRow);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        public FluentCell SetCellPosition(ExcelColumns col, int row)
        {
            if (_sheet == null) throw new System.InvalidOperationException("No active sheet. Call UseSheet(...) first.");

            var normalizedCol = NormalizeCol(col);
            var normalizedRow = NormalizeRow(row);

            var rowObj = _sheet.GetRow(normalizedRow) ?? _sheet.CreateRow(normalizedRow);
            var cell = rowObj.GetCell((int)normalizedCol) ?? rowObj.CreateCell((int)normalizedCol);
            return new FluentCell(_workbook, _sheet, cell, _cellStylesCached);
        }

        public FluentTable<T> SetTable<T>(System.Collections.Generic.IEnumerable<T> table, ExcelColumns startCol, int startRow)
        {
            return new FluentTable<T>(_workbook, _sheet, table, startCol, startRow, _cellStylesCached, new List<TableCellSet>(), new List<TableCellSet>());
        }

        /// <summary>
        /// 獲取指定位置單元格的值
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="row">行位置（1-based）</param>
        /// <returns>單元格的值</returns>
        public object GetCellValue(ExcelColumns col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell((int)col);
            return GetCellValue(cell);
        }

        /// <summary>
        /// 獲取指定位置單元格的值並轉換為指定類型
        /// </summary>
        /// <typeparam name="T">目標類型</typeparam>
        /// <param name="col">列位置</param>
        /// <param name="row">行位置（1-based）</param>
        /// <returns>轉換後的值</returns>
        public T GetCellValue<T>(ExcelColumns col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return default(T);

            var cell = rowObj.GetCell((int)col);
            return GetCellValue<T>(cell);
        }

        /// <summary>
        /// 獲取指定位置單元格的公式字符串
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="row">行位置（1-based）</param>
        /// <returns>公式字符串（不含 '=' 前綴）</returns>
        public string GetCellFormula(ExcelColumns col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell((int)col);
            return GetCellFormulaValue(cell);
        }

        /// <summary>
        /// 獲取指定位置的單元格對象（用於更高級的讀取操作）
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="row">行位置（1-based）</param>
        /// <returns>FluentCell 對象，可以鏈式調用讀取方法</returns>
        public FluentCell GetCellPosition(ExcelColumns col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell((int)col);
            if (cell == null) return null;

            return new FluentCell(_workbook, _sheet, cell, _cellStylesCached);
        }
    }
}

