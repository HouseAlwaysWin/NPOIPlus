using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Models;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 表格操作類
    /// </summary>
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

        public FluentTable<T> SetTable<T>(IEnumerable<T> table, ExcelColumns startCol, int startRow)
        {
            return new FluentTable<T>(_workbook, _sheet, table, startCol, startRow, _cellStylesCached, new List<TableCellSet>(), new List<TableCellSet>());
        }

        /// <summary>
        /// 從 Excel 讀取表格數據並轉換為物件集合
        /// </summary>
        /// <typeparam name="T">目標類型</typeparam>
        /// <param name="startCol">起始列</param>
        /// <param name="startRow">起始行（1-based）</param>
        /// <param name="endRow">結束行（1-based）</param>
        /// <returns>物件集合</returns>
        public List<T> GetTable<T>(ExcelColumns startCol, int startRow, int endRow)
        {
            var result = new List<T>();
            var type = typeof(T);

            // 獲取所有公開的屬性和欄位,按照定義順序排序
            var properties = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                .OrderBy(p => p.MetadataToken)
                .ToArray();
            var fields = type.GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                .OrderBy(f => f.MetadataToken)
                .ToArray();

            // 計算需要讀取的列數
            int columnCount = properties.Length + fields.Length;

            // 遍歷每一行
            for (int row = startRow; row <= endRow; row++)
            {
                var normalizedRow = NormalizeRow(row);
                var rowObj = _sheet.GetRow(normalizedRow);

                // 如果行不存在,跳過
                if (rowObj == null)
                    continue;

                // 創建新實例 (使用 FormatterServices 以支持沒有無參構造函數的類型)
                T instance = (T)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(type);
                int colOffset = 0;

                // 設置所有欄位的值
                foreach (var field in fields)
                {
                    var colIndex = (int)startCol + colOffset;
                    var cell = rowObj.GetCell(colIndex);

                    if (cell != null)
                    {
                        var value = GetCellValueForType(cell, field.FieldType);
                        field.SetValue(instance, value);
                    }

                    colOffset++;
                }

                // 設置所有屬性的值
                foreach (var prop in properties)
                {
                    if (!prop.CanWrite)
                        continue;

                    var colIndex = (int)startCol + colOffset;
                    var cell = rowObj.GetCell(colIndex);

                    if (cell != null)
                    {
                        var value = GetCellValueForType(cell, prop.PropertyType);
                        prop.SetValue(instance, value);
                    }

                    colOffset++;
                }

                result.Add(instance);
            }

            return result;
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

        public FluentSheet SetCellStyleRange(string cellStyleKey, ExcelColumns startCol, ExcelColumns endCol, int startRow, int endRow)
        {
            base.SetCellStyleRange(new CellStyleConfig(cellStyleKey, null), startCol, endCol, startRow, endRow);
            return this;
        }

        public FluentSheet SetCellStyleRange(CellStyleConfig cellStyleConfig, ExcelColumns startCol, ExcelColumns endCol, int startRow, int endRow)
        {
            base.SetCellStyleRange(cellStyleConfig, startCol, endCol, startRow, endRow);
            return this;
        }
    }
}

