using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOIPlus
{
	public class FluentTableCell<T> : FluentTableBase<T>
	{
		public FluentTableCell(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow,
			Dictionary<string, ICellStyle> cellStylesCached,
			string cellName,
			List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
			: base(workbook, sheet, table, startCol, startRow, cellStylesCached, cellTitleSets, cellBodySets,
				  cellBodySets.First(c => c.CellName == cellName))
		{
		}

	public FluentTableCell<T> SetValue(object value)
	{
		SetValueInternal(value);
		return this;
	}

	public FluentTableCell<T> SetValue(Func<TableCellParams, object> valueAction)
	{
		SetValueActionInternal(valueAction);
		return this;
	}

	public FluentTableCell<T> SetValue(Func<TableCellParams<T>, object> valueAction)
	{
		SetValueActionGenericInternal(valueAction);
		return this;
	}

	public FluentTableCell<T> SetFormulaValue(object value)
	{
		SetFormulaValueInternal(value);
		return this;
	}

	public FluentTableCell<T> SetFormulaValue(Func<TableCellParams, object> valueAction)
	{
		SetFormulaValueActionInternal(valueAction);
		return this;
	}

	public FluentTableCell<T> SetFormulaValue(Func<TableCellParams<T>, object> valueAction)
	{
		SetFormulaValueActionGenericInternal(valueAction);
		return this;
	}

	public FluentTableCell<T> SetCellStyle(string cellStyleKey)
	{
		SetCellStyleInternal(cellStyleKey);
		return this;
	}

	public FluentTableCell<T> SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
	{
		SetCellStyleInternal(cellStyleKey, cellStyleAction);
		return this;
	}

	public FluentTableCell<T> SetCellType(CellType cellType)
	{
		SetCellTypeInternal(cellType);
		return this;
	}

		public FluentTableCell<T> CopyStyleFromCell(ExcelColumns col, int rowIndex)
		{
			CopyStyleFromCellInternal(col, rowIndex);
			return this;
		}

		public FluentTable<T> End()
		{
			return new FluentTable<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, _cellTitleSets, _cellBodySets);
		}
	}
}

