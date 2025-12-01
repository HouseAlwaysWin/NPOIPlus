using NPOI.SS.UserModel;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;

namespace NPOIPlus.Base
{
	public abstract class FluentTableBase<T>
	{
		protected IWorkbook _workbook;
		protected List<TableCellSet> _cellBodySets;
		protected List<TableCellSet> _cellTitleSets;
		protected ISheet _sheet;
		protected IEnumerable<T> _table;
		protected ExcelColumns _startCol;
		protected int _startRow;
		protected Dictionary<string, ICellStyle> _cellStylesCached;
		protected TableCellSet _currentCellSet;

		protected FluentTableBase(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached,
			List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets,
			TableCellSet currentCellSet)
		{
			_workbook = workbook;
			_sheet = sheet;
			_table = table;
			_startCol = startCol;
			_startRow = startRow;
			_cellStylesCached = cellStylesCached;
			_cellBodySets = cellBodySets;
			_cellTitleSets = cellTitleSets;
			_currentCellSet = currentCellSet;
		}

		protected void SetCellStyleInternal(string cellStyleKey)
		{
			_currentCellSet.CellStyleKey = cellStyleKey;
		}

		protected void SetCellStyleInternal(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			_currentCellSet.CellStyleKey = cellStyleKey;
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

