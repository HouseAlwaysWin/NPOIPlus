using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Models;
using FluentNPOI.Streaming.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 表格操作類
    /// </summary>
    public class FluentSheet : FluentSheetBase
    {
        /// <summary>
        /// 初始化 FluentSheet 实例
        /// </summary>
        /// <param name="workbook">工作簿对象</param>
        /// <param name="sheet">工作表对象</param>
        /// <param name="cellStylesCached">样式缓存字典</param>
        public FluentSheet(IWorkbook workbook, ISheet sheet, Dictionary<string, ICellStyle> cellStylesCached)
            : base(workbook, sheet, cellStylesCached)
        {
        }

        /// <summary>
        /// 获取 NPOI 工作表对象
        /// </summary>
        /// <returns>ISheet 对象</returns>
        public ISheet GetSheet()
        {
            return _sheet;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="width">列宽（字符数）</param>
        /// <returns>FluentSheet 实例，支持链式调用</returns>
        public FluentSheet SetColumnWidth(ExcelCol col, int width)
        {
            _sheet.SetColumnWidth((int)col, width * 256);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// 批量设置列宽
        /// </summary>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="width">列宽（字符数）</param>
        /// <returns>FluentSheet 实例，支持链式调用</returns>
        public FluentSheet SetColumnWidth(ExcelCol startCol, ExcelCol endCol, int width)
        {
            for (int i = (int)startCol; i <= (int)endCol; i++)
            {
                _sheet.SetColumnWidth(i, width * 256);
            }
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// 合并单元格（横向合并）
        /// </summary>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="row">行号（1-based）</param>
        /// <returns>FluentSheet 实例，支持链式调用</returns>
        public FluentSheet SetExcelCellMerge(ExcelCol startCol, ExcelCol endCol, int row)
        {
            _sheet.SetExcelCellMerge(startCol, endCol, row);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// 合并单元格（区域合并）
        /// </summary>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="firstRow">起始行（1-based）</param>
        /// <param name="lastRow">结束行（1-based）</param>
        /// <returns>FluentSheet 实例，支持链式调用</returns>
        public FluentSheet SetExcelCellMerge(ExcelCol startCol, ExcelCol endCol, int firstRow, int lastRow)
        {
            _sheet.SetExcelCellMerge(startCol, endCol, firstRow, lastRow);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// 设置单元格位置并返回 FluentCell 实例
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="row">行号（1-based）</param>
        /// <returns>FluentCell 实例</returns>
        public FluentCell SetCellPosition(ExcelCol col, int row)
        {
            var cell = SetCellPositionInternal(col, row);
            return new FluentCell(_workbook, _sheet, cell, _cellStylesCached);
        }

        /// <summary>
        /// 使用 FluentMapping 設定表格資料（推薦使用）
        /// </summary>
        /// <typeparam name="T">資料型別</typeparam>
        /// <param name="table">資料集合</param>
        /// <param name="mapping">FluentMapping 設定</param>
        /// <param name="startRow">起始行（1-based），若未指定則使用 mapping.StartRow（預設為 1）</param>
        /// <returns>FluentTable 实例（已套用 mapping）</returns>
        public FluentTable<T> SetTable<T>(IEnumerable<T> table, FluentMapping<T> mapping, int? startRow = null) where T : new()
        {
            // 優先使用參數指定的 startRow，否則使用 mapping 中的 StartRow
            int actualStartRow = startRow ?? mapping.StartRow;
            var fluentTable = new FluentTable<T>(_workbook, _sheet, table, ExcelCol.A, actualStartRow, _cellStylesCached, new List<TableCellSet>(), new List<TableCellSet>());
            return fluentTable.WithMapping(mapping);
        }

        /// <summary>
        /// 使用 DataTableMapping 寫入 DataTable 資料
        /// </summary>
        /// <param name="dataTable">DataTable 資料</param>
        /// <param name="mapping">DataTableMapping 設定 (可為 null，將自動產生)</param>
        /// <param name="startRow">起始行（1-based），若未指定則使用 mapping.StartRow（預設為 1）</param>
        public FluentSheet WriteDataTable(System.Data.DataTable dataTable, DataTableMapping mapping = null, int? startRow = null)
        {
            var actualMapping = mapping ?? DataTableMapping.FromDataTable(dataTable);
            // 優先使用參數指定的 startRow，否則使用 mapping 中的 StartRow
            int actualStartRow = startRow ?? actualMapping.StartRow;
            var mappings = actualMapping.GetMappings().Where(m => m.ColumnIndex.HasValue).ToList();

            bool writeTitle = mappings.Any(m => !string.IsNullOrEmpty(m.Title));

            // 寫入標題行
            if (writeTitle)
            {
                var titleRow = GetOrCreateRow(actualStartRow - 1);
                foreach (var map in mappings)
                {
                    var cell = GetOrCreateCell(titleRow, (int)map.ColumnIndex.Value);
                    cell.SetCellValue(map.Title ?? map.ColumnName ?? "");
                    ApplyStyle(cell, map.TitleStyleKey);
                }
            }

            // 寫入資料行
            int dataRowStart = actualStartRow - 1 + (writeTitle ? 1 : 0);
            for (int rowIdx = 0; rowIdx < dataTable.Rows.Count; rowIdx++)
            {
                var dataRow = dataTable.Rows[rowIdx];
                var excelRow = GetOrCreateRow(dataRowStart + rowIdx);

                foreach (var map in mappings)
                {
                    var colIdx = (int)map.ColumnIndex.Value;
                    var cell = GetOrCreateCell(excelRow, colIdx);

                    if (map.FormulaFunc != null)
                        cell.SetCellFormula(map.FormulaFunc(dataRowStart + rowIdx + 1, (ExcelCol)colIdx));
                    else
                        SetCellValueInternal(cell, actualMapping.GetValue(map, dataRow, dataRowStart + rowIdx + 1, (ExcelCol)colIdx));

                    ApplyStyle(cell, map.StyleKey);
                }
            }

            return this;
        }

        private IRow GetOrCreateRow(int rowIndex) => _sheet.GetRow(rowIndex) ?? _sheet.CreateRow(rowIndex);

        private ICell GetOrCreateCell(IRow row, int colIndex) => row.GetCell(colIndex) ?? row.CreateCell(colIndex);

        private void ApplyStyle(ICell cell, string styleKey)
        {
            if (string.IsNullOrEmpty(styleKey)) styleKey = "global";

            if (!string.IsNullOrEmpty(styleKey) && _cellStylesCached.TryGetValue(styleKey, out var style))
                cell.CellStyle = style;
        }

        private void SetCellValueInternal(ICell cell, object value)
        {
            if (value == null || value == DBNull.Value) { cell.SetCellValue(""); return; }

            switch (value)
            {
                case string s: cell.SetCellValue(s); break;
                case DateTime dt: cell.SetCellValue(dt); break;
                case double d: cell.SetCellValue(d); break;
                case int i: cell.SetCellValue(i); break;
                case long l: cell.SetCellValue(l); break;
                case decimal dec: cell.SetCellValue((double)dec); break;
                case bool b: cell.SetCellValue(b); break;
                default: cell.SetCellValue(value.ToString()); break;
            }
        }

        /// <summary>
        /// 獲取單元格的值並轉換為指定類型
        /// </summary>
        private object GetCellValueForType(ICell cell, System.Type targetType)
        {
            if (cell == null)
                return null;

            try
            {
                // 使用泛型方法獲取值 - 找到受保護的 GetCellValue<T>(ICell) 方法
                var methods = typeof(FluentCellBase).GetMethods(
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                var method = methods.FirstOrDefault(m =>
                    m.Name == "GetCellValue" &&
                    m.IsGenericMethodDefinition &&
                    m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType == typeof(ICell));

                if (method != null)
                {
                    var genericMethod = method.MakeGenericMethod(targetType);
                    return genericMethod.Invoke(this, new object[] { cell });
                }
            }
            catch
            {
                // 如果反射失敗，使用備用方案
            }

            // 備用方案: 直接獲取值
            return GetCellValue(cell);
        }

        /// <summary>
        /// 獲取指定位置單元格的值
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="row">行位置（1-based）</param>
        /// <returns>單元格的值</returns>
        public object GetCellValue(ExcelCol col, int row)
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
        public T GetCellValue<T>(ExcelCol col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return default;

            var cell = rowObj.GetCell((int)col);
            return GetCellValue<T>(cell);
        }

        /// <summary>
        /// 獲取指定位置單元格的公式字符串
        /// </summary>
        /// <param name="col">列位置</param>
        /// <param name="row">行位置（1-based）</param>
        /// <returns>公式字符串（不含 '=' 前綴）</returns>
        public string GetCellFormula(ExcelCol col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell((int)col);
            return GetCellFormulaValue(cell);
        }

        // /// <summary>
        // /// 獲取指定位置的單元格對象（用於更高級的讀取操作）
        // /// </summary>
        // /// <param name="col">列位置</param>
        // /// <param name="row">行位置（1-based）</param>
        // /// <returns>FluentCell 對象，可以鏈式調用讀取方法</returns>
        // public FluentCell GetCellPosition(ExcelCol col, int row)
        // {
        //     var normalizedRow = NormalizeRow(row);
        //     var rowObj = _sheet.GetRow(normalizedRow);
        //     if (rowObj == null) return null;

        //     var cell = rowObj.GetCell((int)col);
        //     if (cell == null) return null;

        //     return new FluentCell(_workbook, _sheet, cell, col, normalizedRow, _cellStylesCached);
        // }

        /// <summary>
        /// 批量设置单元格范围样式（使用样式键名）
        /// </summary>
        /// <param name="cellStyleKey">样式缓存键名</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="startRow">起始行（1-based）</param>
        /// <param name="endRow">结束行（1-based）</param>
        /// <returns>FluentSheet 实例，支持链式调用</returns>
        public FluentSheet SetCellStyleRange(string cellStyleKey, ExcelCol startCol, ExcelCol endCol, int startRow, int endRow)
        {
            base.SetCellStyleRange(new CellStyleConfig(cellStyleKey, null), startCol, endCol, startRow, endRow);
            return this;
        }

        /// <summary>
        /// 批量设置单元格范围样式（使用样式配置）
        /// </summary>
        /// <param name="cellStyleConfig">样式配置对象</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="startRow">起始行（1-based）</param>
        /// <param name="endRow">结束行（1-based）</param>
        /// <returns>FluentSheet 实例，支持链式调用</returns>
        public new FluentSheet SetCellStyleRange(CellStyleConfig cellStyleConfig, ExcelCol startCol, ExcelCol endCol, int startRow, int endRow)
        {
            base.SetCellStyleRange(cellStyleConfig, startCol, endCol, startRow, endRow);
            return this;
        }

        /// <summary>
        /// 設定 Sheet 級別的全域樣式
        /// Setup sheet-level global style
        /// </summary>
        /// <param name="styles">樣式設定函數</param>
        /// <returns></returns>
        public FluentSheet SetupSheetGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            styles(_workbook, newCellStyle);
            string sheetGlobalKey = $"global_{_sheet.SheetName}";

            // Remove existing sheet global style if present
            if (_cellStylesCached.ContainsKey(sheetGlobalKey))
            {
                _cellStylesCached.Remove(sheetGlobalKey);
            }

            _cellStylesCached.Add(sheetGlobalKey, newCellStyle);
            return this;
        }


        /// <summary>
        /// 獲取指定列的最後一行（從指定起始行開始向下查找）
        /// </summary>
        /// <param name="col">要檢查的列</param>
        /// <param name="startRow">起始行（1-based）</param>
        /// <returns>最後一行的行號（1-based），如果沒有找到則返回起始行</returns>
        private int GetLastRowWithData(ExcelCol col, int startRow)
        {
            if (_sheet == null) return startRow;

            // LastRowNum 是 0-based，轉換為 1-based
            int lastRowNum = _sheet.LastRowNum + 1;

            // 從最後一行向上查找，找到第一個有數據的行
            for (int row = lastRowNum; row >= startRow; row--)
            {
                var normalizedRow = NormalizeRow(row);
                var rowObj = _sheet.GetRow(normalizedRow);

                if (rowObj != null)
                {
                    var cell = rowObj.GetCell((int)col);
                    if (cell != null && !IsCellEmpty(cell))
                    {
                        return row;
                    }
                }
            }

            // 如果沒有找到有數據的行，返回起始行
            return startRow;
        }

        /// <summary>
        /// 檢查單元格是否為空
        /// </summary>
        private bool IsCellEmpty(ICell cell)
        {
            if (cell == null) return true;

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return true;
                case CellType.String:
                    return string.IsNullOrWhiteSpace(cell.StringCellValue);
                case CellType.Numeric:
                    // 可以根據需要決定是否將 0 視為空
                    return false;
                case CellType.Boolean:
                    return false;
                case CellType.Formula:
                    return false;
                default:
                    return true;
            }
        }
    }
}

