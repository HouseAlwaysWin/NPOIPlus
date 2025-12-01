using NPOI.SS.UserModel;
using System;

namespace NPOIPlus.Models
{
	public class TableCellSet
	{
		public TableCellSet TitleCellSet { get; set; }
		public string CellName { get; set; }
		public object CellValue { get; set; }
		public Func<TableCellParams, object> SetValueAction { get; set; }
		public Func<TableCellParams, object> SetFormulaValueAction { get; set; }
		// 泛型委派（供 ITable*Stage<T> 使用）
		public Delegate SetValueActionGeneric { get; set; }
		public Delegate SetFormulaValueActionGeneric { get; set; }
		public string CellStyleKey { get; set; }
		public Action<TableCellStyleParams, ICellStyle> SetCellStyleAction { get; set; }
		public CellType CellType { get; set; }
	}
}

