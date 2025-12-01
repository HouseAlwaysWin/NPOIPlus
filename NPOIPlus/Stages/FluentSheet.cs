using NPOI.SS.UserModel;
using NPOIPlus.Base;
using NPOIPlus.Models;
using System.Collections.Generic;

namespace NPOIPlus
{
	public class FluentSheet : FluentSheetBase
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

		public FluentSheet SetColumnWidth(ExcelColumns col, int width)
		{
			_sheet.SetColumnWidth((int)col, width * 256);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public FluentSheet SetColumnWidth(ExcelColumns startCol, ExcelColumns endCol, int width)
		{
			for (int i = (int)startCol; i <= (int)endCol; i++)
			{
				_sheet.SetColumnWidth(i, width * 256);
			}
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public FluentSheet SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int row)
		{
			_sheet.SetExcelCellMerge(startCol, endCol, row);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public FluentSheet SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int firstRow, int lastRow)
		{
			_sheet.SetExcelCellMerge(startCol, endCol, firstRow, lastRow);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public FluentCell SetCellPosition(ExcelColumns col, int row)
		{
			if (_sheet == null) throw new System.InvalidOperationException("No active sheet. Call UseSheet(...) first.");

			var normalizedCol = NormalizeStartCol(col);
			var normalizedRow = NormalizeStartRow(row);

			var rowObj = _sheet.GetRow(normalizedRow) ?? _sheet.CreateRow(normalizedRow);
			var cell = rowObj.GetCell((int)normalizedCol) ?? rowObj.CreateCell((int)normalizedCol);
			return new FluentCell(_workbook, _sheet, cell, _cellStylesCached);
		}

		public FluentTable<T> SetTable<T>(System.Collections.Generic.IEnumerable<T> table, ExcelColumns startCol, int startRow)
		{
			return new FluentTable<T>(_workbook, _sheet, table, startCol, startRow, _cellStylesCached, new List<TableCellSet>(), new List<TableCellSet>());
		}
	}
}

