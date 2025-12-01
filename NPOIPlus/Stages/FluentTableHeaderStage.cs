using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOIPlus
{
	public class FluentTableHeaderStage<T> : FluentTableBase<T>
	{
		public FluentTableHeaderStage(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached,
			string title,
			List<TableCellSet> titleCellSets, List<TableCellSet> cellBodySets)
			: base(workbook, sheet, table, startCol, startRow, cellStylesCached, titleCellSets, cellBodySets,
				  titleCellSets.FirstOrDefault(c => c.CellName == $"{title}_TITLE"))
		{
			_currentCellSet.CellValue = _currentCellSet.CellValue;
		}

		public FluentTableHeaderStage<T> SetValue(Func<TableCellParams, object> valueAction)
		{
			_currentCellSet.SetValueAction = valueAction;
			return this;
		}

		public FluentTableHeaderStage<T> SetValue(Func<TableCellParams<T>, object> valueAction)
		{
			_currentCellSet.SetValueActionGeneric = valueAction;
			return this;
		}

		public FluentTableHeaderStage<T> SetFormulaValue(object value)
		{
			_currentCellSet.CellValue = value;
			_currentCellSet.CellType = CellType.Formula;
			return this;
		}

		public FluentTableHeaderStage<T> SetFormulaValue(Func<TableCellParams, object> valueAction)
		{
			_currentCellSet.SetFormulaValueAction = valueAction;
			_currentCellSet.CellType = CellType.Formula;
			return this;
		}

		public FluentTableHeaderStage<T> SetFormulaValue(Func<TableCellParams<T>, object> valueAction)
		{
			_currentCellSet.SetFormulaValueActionGeneric = valueAction;
			_currentCellSet.CellType = CellType.Formula;
			return this;
		}

		public FluentTableHeaderStage<T> SetCellStyle(string cellStyleKey)
		{
			SetCellStyleInternal(cellStyleKey);
			return this;
		}

		public FluentTableHeaderStage<T> SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			SetCellStyleInternal(cellStyleKey, cellStyleAction);
			return this;
		}

		public FluentTableHeaderStage<T> SetCellType(CellType cellType)
		{
			SetCellTypeInternal(cellType);
			return this;
		}

		public FluentTableCellStage<T> BeginBodySet(string cellName)
		{
			_cellBodySets.Add(new TableCellSet { CellName = cellName });
			return new FluentTableCellStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
		}

		public FluentTableHeaderStage<T> CopyStyleFromCell(ExcelColumns col, int rowIndex)
		{
			CopyStyleFromCellInternal(col, rowIndex);
			return this;
		}
	}
}

