using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Models;
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

        public FluentTable(IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
            ExcelCol startCol, int startRow,
            Dictionary<string, ICellStyle> cellStylesCached, List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
            : base(workbook, sheet, cellStylesCached)
        {
            _table = table;
            _startCol = NormalizeCol(startCol);
            _startRow = NormalizeRow(startRow);
            _cellTitleSets = cellTitleSets;
            _cellBodySets = cellBodySets;
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

        public FluentTableCell<T> BeginBodySet(string cellName)
        {
            _cellBodySets.Add(new TableCellSet { CellName = cellName });
            return new FluentTableCell<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
        }

        public FluentTableHeader<T> BeginTitleSet(string title)
        {
            _cellTitleSets.Add(new TableCellSet { CellName = $"{title}_TITLE", CellValue = title });
            return new FluentTableHeader<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, title, _cellTitleSets, _cellBodySets);
        }

        public FluentTable<T> BuildRows()
        {
            for (int i = 0; i < _table.Count(); i++)
            {
                SetRow(i);
            }
            return this;
        }
    }
}

