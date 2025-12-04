using NPOI.SS.UserModel;

namespace FluentNPOI.Models
{
	// 非泛型參數類型（相容舊 API）
	public class TableCellStyleParams
	{
		public IWorkbook Workbook { get; set; }
		public ExcelCol ColNum { get; set; }
		public int RowNum { get; set; }
		public object RowItem { get; set; }

		public T GetRowItem<T>()
		{
			return RowItem is T t ? t : default;
		}
	}

	public class TableCellStyleParams<T>
	{
		public IWorkbook Workbook { get; set; }
		public ExcelCol ColNum { get; set; }
		public int RowNum { get; set; }
		public T RowItem { get; set; }
	}
}

