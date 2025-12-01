using NPOI.SS.UserModel;

namespace NPOIPlus.Models
{
	public class TableCellStyleParams
	{
		public IWorkbook Workbook { get; set; }
		public ExcelColumns ColNum { get; set; }
		public int RowNum { get; set; }
	}
}

