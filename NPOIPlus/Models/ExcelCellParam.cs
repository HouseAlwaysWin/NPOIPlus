using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public delegate T DefaultType<T>(T value);
	public delegate string FormulaCellValueType(int rownum = 0, ExcelColumns colnum = 0, object cellValue = null);
	public delegate object CellValueActionType(ICell cell, object cellValue = null);

	public class ExcelCellParam
	{
		public object CellValue { get; set; }
		public string ColumnName { get; set; }
		public Action<ICellStyle> CellStyle { get; set; }
		public CellValueActionType CellValueAction { get; set; }
		public FormulaCellValueType FormulaCellValue { get; set; }
		public bool? IsFormula { get; set; }

		public ExcelCellParam(CellValueActionType cellValueAction, Action<ICellStyle> style = null)
		{
			CellValueAction = cellValueAction;
			CellStyle = style;
			IsFormula = false;
		}
		public ExcelCellParam(FormulaCellValueType cellValue, Action<ICellStyle> style = null)
		{
			FormulaCellValue = cellValue;
			CellStyle = style;
			IsFormula = true;
		}

		public ExcelCellParam(object cellValue, Action<ICellStyle> style = null, bool? isFormula = null)
		{
			CellValue = cellValue;
			CellStyle = style;
			IsFormula = isFormula;
		}
		public ExcelCellParam(string columnName, object cellValue = null, Action<ICellStyle> style = null, bool? isFormula = null, FormulaCellValueType formulaCellValue = null)
		{
			CellValue = cellValue;
			ColumnName = columnName;
			CellStyle = style;
			IsFormula = isFormula;
			FormulaCellValue = formulaCellValue;
		}
	}
}
