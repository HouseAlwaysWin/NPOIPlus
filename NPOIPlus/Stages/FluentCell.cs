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

	public FluentCell SetCellStyle(Func<TableCellStyleParams, CellStyleConfig> cellStyleAction)
	{
		if (_cell == null) return this;

		var cellStyleParams = new TableCellStyleParams
		{
			Workbook = _workbook,
			ColNum = (ExcelColumns)_cell.ColumnIndex,
			RowNum = _cell.RowIndex,
			RowItem = null
		};
		
		// ✅ 先調用函數獲取樣式配置
		var config = cellStyleAction(cellStyleParams);
		
		if (!string.IsNullOrWhiteSpace(config.Key))
		{
			// ✅ 先檢查緩存
			if (!_cellStylesCached.ContainsKey(config.Key))
			{
				// ✅ 只在不存在時才創建新樣式
				ICellStyle newCellStyle = _workbook.CreateCellStyle();
				config.StyleSetter(newCellStyle);
				_cellStylesCached.Add(config.Key, newCellStyle);
			}
			_cell.CellStyle = _cellStylesCached[config.Key];
		}
		else
		{
			// 如果沒有返回 key，創建臨時樣式（不緩存）
			ICellStyle tempStyle = _workbook.CreateCellStyle();
			config.StyleSetter(tempStyle);
			_cell.CellStyle = tempStyle;
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

