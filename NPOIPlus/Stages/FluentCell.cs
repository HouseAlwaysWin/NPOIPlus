using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Helpers;
using System;

namespace NPOIPlus
{
	public class FluentCell : FluentCellBase, ICellStage
	{
		private ISheet _sheet;
		private ICell _cell;

		public FluentCell(IWorkbook workbook, ISheet sheet, ICell cell)
			: base(workbook, null)
		{
			_sheet = sheet;
			_cell = cell;
		}

		public ICellStage SetValue<T>(T value)
		{
			if (_cell == null) return this;
			SetCellValue(_cell, value);
			return this;
		}
	}
}

