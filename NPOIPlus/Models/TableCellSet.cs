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
		// 返回樣式配置，StyleSetter 只在需要時才執行
		public Func<TableCellStyleParams, CellStyleConfig> SetCellStyleAction { get; set; }
		public CellType CellType { get; set; }
	}
}

