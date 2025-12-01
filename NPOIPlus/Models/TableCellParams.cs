namespace NPOIPlus.Models
{
	// 非泛型參數類型（相容舊 API）
	public class TableCellParams
	{
		public object CellValue { get; set; }
		public ExcelColumns ColNum { get; set; }
		public int RowNum { get; set; }
		public object RowItem { get; set; }

		public T GetRowItem<T>()
		{
			return RowItem is T t ? t : default;
		}
	}

	public class TableCellParams<T>
	{
		public object CellValue { get; set; }
		public ExcelColumns ColNum { get; set; }
		public int RowNum { get; set; }
		public T RowItem { get; set; }
	}
}

