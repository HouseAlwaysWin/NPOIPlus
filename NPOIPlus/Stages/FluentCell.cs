using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Helpers;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;

namespace NPOIPlus
{
	public class FluentCell : FluentCellBase
	{
		private ISheet _sheet;
		private ICell _cell;

		public FluentCell(IWorkbook workbook, ISheet sheet, ICell cell, Dictionary<string, ICellStyle> cellStylesCached = null)
			: base(workbook, cellStylesCached ?? new Dictionary<string, ICellStyle>())
		{
			_sheet = sheet;
			_cell = cell;
		}

		public FluentCell SetValue<T>(T value)
		{
			if (_cell == null) return this;
			SetCellValue(_cell, value);
			return this;
		}

		public FluentCell SetFormulaValue(object value)
		{
			if (_cell == null) return this;
			SetFormulaValue(_cell, value);
			return this;
		}

		public FluentCell SetCellStyle(string cellStyleKey)
		{
			if (_cell == null) return this;
			
			if (!string.IsNullOrWhiteSpace(cellStyleKey) && _cellStylesCached.ContainsKey(cellStyleKey))
			{
				_cell.CellStyle = _cellStylesCached[cellStyleKey];
			}
			return this;
		}

		public FluentCell SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			if (_cell == null) return this;

			if (!string.IsNullOrWhiteSpace(cellStyleKey) && !_cellStylesCached.ContainsKey(cellStyleKey))
			{
				ICellStyle newCellStyle = _workbook.CreateCellStyle();
				var cellStyleParams = new TableCellStyleParams
				{
					Workbook = _workbook,
					ColNum = (ExcelColumns)_cell.ColumnIndex,
					RowNum = _cell.RowIndex
				};
				cellStyleAction(cellStyleParams, newCellStyle);
				_cellStylesCached.Add(cellStyleKey, newCellStyle);
			}

			if (_cellStylesCached.ContainsKey(cellStyleKey))
			{
				_cell.CellStyle = _cellStylesCached[cellStyleKey];
			}
			return this;
		}

		public FluentCell SetCellType(CellType cellType)
		{
			if (_cell == null) return this;
			_cell.SetCellType(cellType);
			return this;
		}
	}
}

