using NPOI.SS.UserModel;
using NPOIPlus.Helpers;

namespace NPOIPlus
{
	public interface ITableStage<T>
	{
		FluentTableCellStage<T> BeginCellSet(string cellName);
		FluentTableHeaderStage<T> BeginTitleSet(string title);
		FluentTable<T> BuildRows();
		FluentMemoryStream ToStream();
		IWorkbook SaveToPath(string filePath);
	}
}

