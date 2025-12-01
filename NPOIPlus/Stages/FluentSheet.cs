using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Models;
using System.Collections.Generic;

namespace NPOIPlus
{
	public class FluentSheet : FluentSheetBase, ISheetStage
	{
		private ISheet _sheet;

		public FluentSheet(IWorkbook workbook, ISheet sheet, Dictionary<string, ICellStyle> cellStylesCached)
			: base(workbook, cellStylesCached)
		{
			_sheet = sheet;
		}

		public ISheet GetSheet()
		{
			return _sheet;
		}

		public ISheetStage SetColumnWidth(ExcelColumns col, int width)
		{
			_sheet.SetColumnWidth((int)col, width * 256);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public ISheetStage SetColumnWidth(ExcelColumns startCol, ExcelColumns endCol, int width)
		{
			for (int i = (int)startCol; i <= (int)endCol; i++)
			{
				_sheet.SetColumnWidth(i, width * 256);
			}
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public ISheetStage SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int row)
		{
			_sheet.SetExcelCellMerge(startCol, endCol, row);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public ISheetStage SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int firstRow, int lastRow)
		{
			_sheet.SetExcelCellMerge(startCol, endCol, firstRow, lastRow);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public ICellStage SetCell(ExcelColumns col, int row)
		{
			if (_sheet == null) throw new System.InvalidOperationException("No active sheet. Call UseSheet(...) first.");

			var rowObj = _sheet.GetRow(row) ?? _sheet.CreateRow(row);
			var cell = rowObj.GetCell((int)col) ?? rowObj.CreateCell((int)col);
			return new FluentCell(_workbook, _sheet, cell);
		}

		public ITableStage<T> SetTable<T>(System.Collections.Generic.IEnumerable<T> table, ExcelColumns startCol, int startRow)
		{
			return new FluentTable<T>(_workbook, _sheet, table, startCol, startRow, _cellStylesCached, new List<TableCellSet>(), new List<TableCellSet>());
		}
	}
}

