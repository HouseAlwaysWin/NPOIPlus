using NPOI.SS.UserModel;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Base
{
	public abstract class FluentTableBase<T> : FluentCellBase
	{
		protected List<TableCellSet> _cellBodySets;
		protected List<TableCellSet> _cellTitleSets;
		protected IEnumerable<T> _table;
		protected ExcelColumns _startCol;
		protected int _startRow;
		protected TableCellSet _currentCellSet;

		protected FluentTableBase(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached,
			List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets,
			TableCellSet currentCellSet)
			: base(workbook, sheet, cellStylesCached)
		{
			_table = table;
			_startCol = startCol;
			_startRow = startRow;
			_cellBodySets = cellBodySets;
			_cellTitleSets = cellTitleSets;
			_currentCellSet = currentCellSet;
		}

	protected void SetValueInternal(object value)
	{
		_currentCellSet.CellValue = value;
	}

	protected void SetValueActionInternal(Func<TableCellParams, object> valueAction)
	{
		_currentCellSet.SetValueAction = valueAction;
	}

	protected void SetValueActionGenericInternal(Func<TableCellParams<T>, object> valueAction)
	{
		_currentCellSet.SetValueActionGeneric = valueAction;
	}

	protected void SetFormulaValueInternal(object value)
	{
		_currentCellSet.CellValue = value;
		_currentCellSet.CellType = CellType.Formula;
	}

	protected void SetFormulaValueActionInternal(Func<TableCellParams, object> valueAction)
	{
		_currentCellSet.SetFormulaValueAction = valueAction;
		_currentCellSet.CellType = CellType.Formula;
	}

	protected void SetFormulaValueActionGenericInternal(Func<TableCellParams<T>, object> valueAction)
	{
		_currentCellSet.SetFormulaValueActionGeneric = valueAction;
		_currentCellSet.CellType = CellType.Formula;
	}

	protected void SetCellStyleInternal(string cellStyleKey)
	{
		_currentCellSet.CellStyleKey = cellStyleKey;
	}

	protected void SetCellStyleInternal(Func<TableCellStyleParams, CellStyleConfig> cellStyleAction)
	{
		_currentCellSet.SetCellStyleAction = cellStyleAction;
	}

	protected void SetCellTypeInternal(CellType cellType)
	{
		_currentCellSet.CellType = cellType;
	}

	protected void CopyStyleFromCellInternal(ExcelColumns col, int rowIndex)
	{
		string key = $"{_sheet.SheetName}_{col}{rowIndex}";
		ICell cell = _sheet.GetExcelCell(col, rowIndex);
		if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(key))
		{
			ICellStyle newCellStyle = _workbook.CreateCellStyle();
			newCellStyle.CloneStyleFrom(cell.CellStyle);
			_cellStylesCached.Add(key, newCellStyle);
			_currentCellSet.CellStyleKey = key;
		}
	}
	}
}

