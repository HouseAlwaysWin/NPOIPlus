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
    /// <typeparam name="T">表格數據類型</typeparam>
    public class FluentTable<T> : FluentSheetBase
    {
        private IEnumerable<T> _table;
        private ExcelCol _startCol;
        private int _startRow;
        private List<TableCellSet> _cellBodySets;
        private List<TableCellSet> _cellTitleSets;
        private IReadOnlyList<ColumnMapping> _columnMappings;

        public FluentTable(IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
            ExcelCol startCol, int startRow,
            Dictionary<string, ICellStyle> cellStylesCached, List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
            : base(workbook, sheet, cellStylesCached)
        {
            _table = table;
            // ExcelCol 已經是有效的列枚舉，不需要規範化
            _startCol = startCol;
            _startRow = NormalizeRow(startRow);
            _cellTitleSets = cellTitleSets;
            _cellBodySets = cellBodySets;
        }

        /// <summary>
        /// 使用 FluentMapping 設定欄位對應（可取代 BeginTitleSet/BeginBodySet）
        /// </summary>
        public FluentTable<T> WithMapping<TMapping>(FluentMapping<TMapping> mapping) where TMapping : new()
        {
            _columnMappings = mapping.GetMappings();
            return this;
        }

        private T GetItemAt(int index)
        {
            var items = _table as IList<T> ?? _table?.ToList() ?? new List<T>();
            if (index < 0 || index >= items.Count) return default;
            return items[index];
        }

        private void SetCellAction(List<TableCellSet> cellSets, IRow rowObj, int colIndex, int targetRowIndex, object item)
        {
            foreach (var cellset in cellSets)
            {
                var cell = rowObj.GetCell(colIndex) ?? rowObj.CreateCell(colIndex);

                // 優先使用 TableCellNameMap 中的 Value，如果沒有則從 item 中獲取
                Func<TableCellParams, object> setValueAction = cellset.SetValueAction;
                Func<TableCellParams, object> setFormulaValueAction = cellset.SetFormulaValueAction;

                TableCellParams cellParams = new TableCellParams
                {
                    ColNum = (ExcelCol)colIndex,
                    RowNum = targetRowIndex,
                    RowItem = item
                };
                object value = cellset.CellValue ?? GetTableCellValue(cellset.CellName, item);
                cellParams.CellValue = value;

                // 準備泛型參數（供泛型委派使用）
                var cellParamsT = new TableCellParams<T>
                {
                    ColNum = (ExcelCol)colIndex,
                    RowNum = targetRowIndex,
                    RowItem = item is T tItem ? tItem : default,
                    CellValue = value
                };

                TableCellStyleParams cellStyleParams =
                new TableCellStyleParams
                {
                    Workbook = _workbook,
                    ColNum = (ExcelCol)colIndex,
                    RowNum = targetRowIndex,
                    RowItem = item
                };
                SetCellStyle(cell, cellset, cellStyleParams);

                if (cellset.CellType == CellType.Formula)
                {
                    if (cellset.SetFormulaValueActionGeneric != null)
                    {
                        if (cellset.SetFormulaValueActionGeneric is Func<TableCellParams<T>, object> gFormula)
                        {
                            value = gFormula(cellParamsT);
                        }
                        else
                        {
                            value = cellset.SetFormulaValueActionGeneric.DynamicInvoke(cellParamsT);
                        }
                    }
                    else if (setFormulaValueAction != null)
                    {
                        value = setFormulaValueAction(cellParams);
                    }
                    SetFormulaValue(cell, value);
                }
                else
                {
                    if (cellset.SetValueActionGeneric != null)
                    {
                        if (cellset.SetValueActionGeneric is Func<TableCellParams<T>, object> gValue)
                        {
                            value = gValue(cellParamsT);
                        }
                        else
                        {
                            value = cellset.SetValueActionGeneric.DynamicInvoke(cellParamsT);
                        }
                    }
                    else if (setValueAction != null)
                    {
                        value = setValueAction(cellParams);
                    }
                    SetCellValue(cell, value, cellset.CellType);
                }

                colIndex++;
            }
        }

        private FluentTable<T> SetRow(int rowOffset = 0)
        {
            if (_cellBodySets == null || _cellBodySets.Count == 0) return this;

            var targetRowIndex = _startRow + rowOffset;

            var item = GetItemAt(rowOffset);

            int colIndex = (int)_startCol;
            if (_cellTitleSets != null && _cellTitleSets.Count > 0)
            {
                var titleRowObj = _sheet.GetRow(_startRow) ?? _sheet.CreateRow(_startRow);
                SetCellAction(_cellTitleSets, titleRowObj, colIndex, _startRow, item);
                targetRowIndex++;
            }

            var rowObj = _sheet.GetRow(targetRowIndex) ?? _sheet.CreateRow(targetRowIndex);
            SetCellAction(_cellBodySets, rowObj, colIndex, targetRowIndex, item);

            return this;
        }

        /// <summary>
        /// 使用 FluentMapping 寫入一行
        /// </summary>
        private FluentTable<T> SetRowWithMapping(int rowOffset, bool writeTitle)
        {
            if (_columnMappings == null || _columnMappings.Count == 0) return this;

            var item = GetItemAt(rowOffset);
            var targetRowIndex = _startRow + rowOffset + (writeTitle ? 1 : 0);

            // 寫入標題行（只在第一次）
            if (writeTitle && rowOffset == 0)
            {
                var titleRow = _sheet.GetRow(_startRow) ?? _sheet.CreateRow(_startRow);
                foreach (var map in _columnMappings.Where(m => m.ColumnIndex.HasValue))
                {
                    var colIdx = (int)map.ColumnIndex.Value;
                    var cell = titleRow.GetCell(colIdx) ?? titleRow.CreateCell(colIdx);
                    cell.SetCellValue(map.Title ?? map.Property?.Name ?? map.ColumnName ?? "");

                    // 套用標題樣式
                    if (map.TitleStyleRef != null)
                    {
                        string refKey = $"{_sheet.SheetName}_{map.TitleStyleRef.Column}{map.TitleStyleRef.Row}";
                        if (_cellStylesCached.TryGetValue(refKey, out var cachedStyle))
                        {
                            cell.CellStyle = cachedStyle;
                        }
                        else
                        {
                            var refRow = _sheet.GetRow(map.TitleStyleRef.Row);
                            var refCell = refRow?.GetCell((int)map.TitleStyleRef.Column);
                            if (refCell != null && refCell.CellStyle != null)
                            {
                                ICellStyle newCellStyle = _workbook.CreateCellStyle();
                                newCellStyle.CloneStyleFrom(refCell.CellStyle);
                                _cellStylesCached[refKey] = newCellStyle;
                                cell.CellStyle = newCellStyle;
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(map.TitleStyleKey) && _cellStylesCached.TryGetValue(map.TitleStyleKey, out var titleStyle))
                    {
                        cell.CellStyle = titleStyle;
                    }
                    else if (cell.CellStyle.Index == 0) // Try global style
                    {
                        string globalStyleKey = $"global_{_sheet.SheetName}";
                        if (_cellStylesCached.TryGetValue(globalStyleKey, out var globalStyle))
                        {
                            cell.CellStyle = globalStyle;
                        }
                    }
                }
            }

            // 寫入資料行
            var dataRow = _sheet.GetRow(targetRowIndex) ?? _sheet.CreateRow(targetRowIndex);
            foreach (var map in _columnMappings.Where(m => m.ColumnIndex.HasValue))
            {
                var colIdx = (int)map.ColumnIndex.Value;
                var cell = dataRow.GetCell(colIdx) ?? dataRow.CreateCell(colIdx);

                // 公式優先
                if (map.FormulaFunc != null)
                {
                    var formula = map.FormulaFunc(targetRowIndex + 1, colIdx); // Excel 是 1-based
                    cell.SetCellFormula(formula);
                }
                else
                {
                    // 計算值 (優先使用 ValueFunc，否則從屬性取值)
                    object value;
                    if (map.ValueFunc != null)
                    {
                        value = map.ValueFunc(item);
                    }
                    else if (map.Property != null)
                    {
                        value = map.Property.GetValue(item);
                    }
                    else
                    {
                        value = null;
                    }

                    SetCellValue(cell, value, map.CellType ?? CellType.Unknown);
                }

                // 套用資料樣式
                string styleKey = map.StyleKey;

                // 如果有動態樣式設定，優先使用
                if (map.DynamicStyleFunc != null)
                {
                    string dynamicKey = map.DynamicStyleFunc(item);
                    if (!string.IsNullOrEmpty(dynamicKey))
                    {
                        styleKey = dynamicKey;
                    }
                }

                if (!string.IsNullOrEmpty(styleKey) && _cellStylesCached.TryGetValue(styleKey, out var dataStyle))
                {
                    cell.CellStyle = dataStyle;
                }

                // 複製資料樣式
                if (map.DataStyleRef != null)
                {
                    string refKey = $"{_sheet.SheetName}_{map.DataStyleRef.Column}{map.DataStyleRef.Row}";
                    if (_cellStylesCached.TryGetValue(refKey, out var cachedStyle))
                    {
                        cell.CellStyle = cachedStyle;
                    }
                    else
                    {
                        var refRow = _sheet.GetRow(map.DataStyleRef.Row);
                        var refCell = refRow?.GetCell((int)map.DataStyleRef.Column);
                        if (refCell != null && refCell.CellStyle != null)
                        {
                            ICellStyle newCellStyle = _workbook.CreateCellStyle();
                            newCellStyle.CloneStyleFrom(refCell.CellStyle);
                            _cellStylesCached[refKey] = newCellStyle;
                            cell.CellStyle = newCellStyle;
                        }
                    }
                }

                // 如果沒有設定任何樣式，嘗試套用 Sheet 全域樣式
                if (cell.CellStyle.Index == 0) // Default style index is usually 0
                {
                    string globalStyleKey = $"global_{_sheet.SheetName}";
                    if (_cellStylesCached.TryGetValue(globalStyleKey, out var globalStyle))
                    {
                        cell.CellStyle = globalStyle;
                    }
                }
            }

            return this;
        }


        public FluentTable<T> BuildRows()
        {
            // 如果有 FluentMapping，使用 mapping 方式寫入
            if (_columnMappings != null)
            {
                bool writeTitle = _columnMappings.Any(m => !string.IsNullOrEmpty(m.Title));
                for (int i = 0; i < _table.Count(); i++)
                {
                    SetRowWithMapping(i, writeTitle);
                }
                return this;
            }

            // 否則使用原有方式
            for (int i = 0; i < _table.Count(); i++)
            {
                SetRow(i);
            }
            return this;
        }
    }
}
