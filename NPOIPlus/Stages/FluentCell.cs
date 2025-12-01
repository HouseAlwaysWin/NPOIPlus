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

		public FluentMemoryStream Save()
		{
			var ms = new FluentMemoryStream();
			ms.AllowClose = false;
			_workbook.Write(ms);
			ms.Flush();
			ms.Seek(0, System.IO.SeekOrigin.Begin);
			ms.AllowClose = true;
			return ms;
		}

	public ICellStage SetValue<T>(T value)
	{
		if (_cell == null) return this;
		SetCellValue(_cell, value);
		return this;
	}
	}
}

