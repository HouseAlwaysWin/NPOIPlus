using NPOI.SS.UserModel;
using System;

namespace NPOIPlus
{
	public interface IWorkbookStage
	{
		IWorkbook GetWorkbook();
		IWorkbookStage ReadExcelFile(string filePath);
		IWorkbookStage SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles);
		IWorkbookStage SetupCellStyle(string cellStyleKey, Action<IWorkbook, ICellStyle> styles);
		ISheetStage UseSheet(string sheetName, bool createIfMissing = true);
		ISheetStage UseSheet(ISheet sheet);
		ISheetStage UseSheetAt(int index, bool createIfMissing = false);
	}
}

