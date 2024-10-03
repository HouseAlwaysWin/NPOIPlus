using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public delegate T DefaultType<T>(T value);
	public delegate string CellValueActionType(ICell cell, object cellValue = null, ExcelColumns colnum = 0, int rownum = 0);

	public class ExcelCellParam
	{
		public readonly object CellValue;
		public readonly string ColumnName;
		public readonly Action<ICellStyle> CellStyle;
		public readonly CellValueActionType CellValueAction;
		public readonly bool? IsFormula;

		public ExcelCellParam(object valueOrColName, CellValueActionType cellValueAction = null, Action<ICellStyle> style = null, bool? isFormula = false)
		{
			CellValueAction = cellValueAction;
			ColumnName = valueOrColName.ToString();
			CellValue = valueOrColName;
			CellStyle = style;
			IsFormula = isFormula;
		}
	}
}
