using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 表格標題操作類
    /// </summary>
    /// <typeparam name="T">表格數據類型</typeparam>
	public class FluentTableHeader<T> : FluentTableBase<T>
    {
        public FluentTableHeader(
            IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
            ExcelCol startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached,
            string title,
            List<TableCellSet> titleCellSets, List<TableCellSet> cellBodySets)
            : base(workbook, sheet, table, startCol, startRow, cellStylesCached, titleCellSets, cellBodySets,
                  titleCellSets.FirstOrDefault(c => c.CellName == $"{title}_TITLE"))
        {
            _currentCellSet.CellValue = _currentCellSet.CellValue;
        }

        public FluentTableHeader<T> SetValue(Func<TableCellParams, object> valueAction)
        {
            SetValueActionInternal(valueAction);
            return this;
        }

        public FluentTableHeader<T> SetValue(Func<TableCellParams<T>, object> valueAction)
        {
            SetValueActionGenericInternal(valueAction);
            return this;
        }

        public FluentTableHeader<T> SetFormulaValue(object value)
        {
            SetFormulaValueInternal(value);
            return this;
        }

        public FluentTableHeader<T> SetFormulaValue(Func<TableCellParams, object> valueAction)
        {
            SetFormulaValueActionInternal(valueAction);
            return this;
        }

        public FluentTableHeader<T> SetFormulaValue(Func<TableCellParams<T>, object> valueAction)
        {
            SetFormulaValueActionGenericInternal(valueAction);
            return this;
        }

        public FluentTableHeader<T> SetCellStyle(string cellStyleKey)
        {
            SetCellStyleInternal(cellStyleKey);
            return this;
        }

        public FluentTableHeader<T> SetCellStyle(Func<TableCellStyleParams, CellStyleConfig> cellStyleAction)
        {
            SetCellStyleInternal(cellStyleAction);
            return this;
        }

        public FluentTableHeader<T> SetCellType(CellType cellType)
        {
            SetCellTypeInternal(cellType);
            return this;
        }

        public FluentTableCell<T> BeginBodySet(string cellName)
        {
            _cellBodySets.Add(new TableCellSet { CellName = cellName });
            return new FluentTableCell<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
        }

        public FluentTableHeader<T> CopyStyleFromCell(ExcelCol col, int rowIndex)
        {
            CopyStyleFromCellInternal(col, rowIndex);
            return this;
        }
    }
}

