using NPOI.SS.UserModel;
using NPOIPlus.Models;
using System.Collections.Generic;

namespace NPOIPlus
{
	public interface ISheetStage
	{
		ISheet GetSheet();
		ITableStage<T> SetTable<T>(IEnumerable<T> table, ExcelColumns startCol, int startRow);
		FluentCell SetCellPosition(ExcelColumns startCol, int startRow);
		ISheetStage SetColumnWidth(ExcelColumns col, int width);
		ISheetStage SetColumnWidth(ExcelColumns startCol, ExcelColumns endCol, int width);
		ISheetStage SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int row);
	}
}

