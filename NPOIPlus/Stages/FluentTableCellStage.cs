using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOIPlus
{
	public class FluentTableCellStage<T> : FluentTableBase<T>
	{
		public FluentTableCellStage(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow,
			Dictionary<string, ICellStyle> cellStylesCached,
			string cellName,
			List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
			: base(workbook, sheet, table, startCol, startRow, cellStylesCached, cellTitleSets, cellBodySets,
				  cellBodySets.First(c => c.CellName == cellName))
		{
		}

		public FluentTableCellStage<T> SetValue(object value)
		{
			_currentCellSet.CellValue = value;
			return this;
		}

		public FluentTableCellStage<T> SetValue(Func<TableCellParams, object> valueAction)
		{
			_currentCellSet.SetValueAction = valueAction;
			return this;
		}

		public FluentTableCellStage<T> SetValue(Func<TableCellParams<T>, object> valueAction)
		{
			_currentCellSet.SetValueActionGeneric = valueAction;
			return this;
		}

		public FluentTableCellStage<T> SetFormulaValue(object value)
		{
			_currentCellSet.CellValue = value;
			_currentCellSet.CellType = CellType.Formula;
			return this;
		}

		public FluentTableCellStage<T> SetFormulaValue(Func<TableCellParams, object> valueAction)
		{
			_currentCellSet.SetFormulaValueAction = valueAction;
			_currentCellSet.CellType = CellType.Formula;
			return this;
		}

		public FluentTableCellStage<T> SetFormulaValue(Func<TableCellParams<T>, object> valueAction)
		{
			_currentCellSet.SetFormulaValueActionGeneric = valueAction;
			_currentCellSet.CellType = CellType.Formula;
			return this;
		}

		public FluentTableCellStage<T> SetCellStyle(string cellStyleKey)
		{
			SetCellStyleInternal(cellStyleKey);
			return this;
		}

		public FluentTableCellStage<T> SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			SetCellStyleInternal(cellStyleKey, cellStyleAction);
			return this;
		}

		public FluentTableCellStage<T> SetCellType(CellType cellType)
		{
			SetCellTypeInternal(cellType);
			return this;
		}

		public FluentTableCellStage<T> CopyStyleFromCell(ExcelColumns col, int rowIndex)
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

