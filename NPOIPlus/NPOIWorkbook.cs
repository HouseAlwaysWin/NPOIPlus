using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Streaming.Values;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Data;

namespace NPOIPlus
{

	public class NPOIWorkbook
	{
		public IWorkbook Workbook { get; set; }

		public DefaultType<int> SetDefaultIntCellValue = (value) => value;
		public DefaultType<double> SetDefaultDoubleCellValue = (value) => value;
		public DefaultType<bool> SetDefaultBoolCellValue = (value) => value;
		public DefaultType<string> SetDefaultStringCellValue = (value) => value;
		public DefaultType<DateTime> SetDefaultDateTimeCellValue = (value) => value;


		public Action<ICellStyle> SetGlobalCellStyle = (style) => { };
		public Action<ICellStyle> SetDefaultIntCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultDoubleCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultBoolCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultStringCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultDateTimeCellStyle = (value) => { };

		private Dictionary<string, ICellStyle> _styles;

		public NPOIWorkbook(IWorkbook workbook)
		{
			Workbook = workbook;
		}

		private void SetCellStyleBasedOnType(object value, ICellStyle style)
		{
			switch (value)
			{
				case int i:
					SetDefaultIntCellStyle?.Invoke(style);
					break;
				case double d:
					SetDefaultDoubleCellStyle?.Invoke(style);
					break;
				case bool b:
					SetDefaultBoolCellStyle?.Invoke(style);
					break;
				case string s:
					SetDefaultStringCellStyle?.Invoke(style);
					break;
				case DateTime dt:
					SetDefaultDateTimeCellStyle?.Invoke(style);
					break;
				default:
					SetDefaultStringCellStyle?.Invoke(style);
					break;
			}
		}


		private void SetCellValueBasedOnType(ICell cell, object value, CellValueActionType valueAction = null)
		{
			value = valueAction?.Invoke(cell, value) ?? value;
			switch (value)
			{
				case int i:
					var intValue = SetDefaultIntCellValue(i);
					cell.SetCellValue(intValue);
					break;
				case double d:
					var doubleValue = SetDefaultDoubleCellValue(d);
					cell.SetCellValue(doubleValue);
					break;
				case bool b:
					var boolValue = SetDefaultBoolCellValue(b);
					cell.SetCellValue(boolValue);
					break;
				case string s:
					var stringValue = SetDefaultStringCellValue(s);
					cell.SetCellValue(stringValue);
					break;
				case DateTime dt:
					var dateValue = SetDefaultDateTimeCellValue(dt);
					cell.SetCellValue(dateValue);
					break;
				default:
					cell.SetCellValue(value?.ToString());
					break;
			}
		}


		private void SetCellStyle(ICell cell, object cellValue, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null)
		{
			ICellStyle newCellStyle = Workbook.CreateCellStyle();
			SetGlobalCellStyle(newCellStyle);
			SetCellStyleBasedOnType(cellValue, newCellStyle);
			rowStyle?.Invoke(newCellStyle);
			colStyle?.Invoke(newCellStyle);
			cell.CellStyle = newCellStyle;
		}

		// 檢查並創建樣式
		public ICellStyle GetOrCreateStyle(string styleKey)
		{
			if (_styles.ContainsKey(styleKey))
			{
				return _styles[styleKey];  // 如果樣式已存在，直接返回
			}

			// 創建新樣式
			ICellStyle newStyle = Workbook.CreateCellStyle();
			// 這裡可以根據需要設置樣式屬性
			newStyle.Alignment = HorizontalAlignment.Center;

			// 將新樣式存入字典
			_styles[styleKey] = newStyle;
			return newStyle;
		}

		/// <summary>
		/// For set single cell
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellValue"></param>
		/// <param name="colnum"></param>
		/// <param name="rownum"></param>
		/// <param name="param"></param>
		/// <exception cref="Exception"></exception>
		public void SetExcelCell(ISheet sheet, object cellValue, ExcelColumns colnum, int rownum, Action<ICellStyle> style = null)
		{
			if (rownum < 1) rownum = 1;
			IRow row = sheet.CreateRow(rownum - 1);
			ICell cell = row.CreateCell((int)colnum);
			SetCellStyle(cell, cellValue, style);
			SetCellValueBasedOnType(cell, cellValue);
		}

		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns column, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null, bool? isFormula = null)
		{
			SetExcelCell(sheet, dataTable, tableIndex, tableColName, column, rownum, cellValue, colStyle, null, null, isFormula, null);
		}

		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns column, int rownum = 1, CellValueActionType cellValueAction = null, Action<ICellStyle> colStyle = null, bool? isFormula = null)
		{
			SetExcelCell(sheet, dataTable, tableIndex, tableColName, column, rownum, null, colStyle, null, cellValueAction, isFormula, null);
		}


		/// <summary>
		/// For set single cell with datatable
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="dataTable"></param>
		/// <param name="tableIndex"></param>
		/// <param name="tableColName"></param>
		/// <param name="colnum"></param>
		/// <param name="rownum"></param>
		/// <param name="cellValue"></param>
		/// <exception cref="Exception"></exception>
		private void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns colnum, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = false, FormulaCellValueType formulaCellValueType = null)
		{
			if (rownum < 1) rownum = 1;
			int zeroBaseIndex = rownum - 1;
			IRow row = sheet.GetRow(zeroBaseIndex) ?? sheet.CreateRow(zeroBaseIndex);
			ICell cell = row.CreateCell((int)colnum);
			var newValue = formulaCellValueType ?? cellValue ?? dataTable.Rows[tableIndex][tableColName];
			SetCellStyle(cell, newValue, colStyle, rowStyle);
			if (isFormula.HasValue)
			{
				if (isFormula.Value)
				{
					var newCellValue = formulaCellValueType(rownum, colnum, cellValue);
					cell.SetCellFormula(newCellValue?.ToString());
					return;
				}
			}

			SetCellValueBasedOnType(cell, newValue, cellValueAction);
		}

		public void SetColExcelCells(ISheet sheet, DataTable dataTable, int tableIndex, List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			for (int colIndex = 0; colIndex < param.Count; colIndex++)
			{
				var colnum = colIndex + startColnum;
				var col = param[colIndex];
				var isFormulaValue = col.IsFormula.HasValue ? col.IsFormula : isFormula;
				SetExcelCell(sheet, dataTable, tableIndex, col.ColumnName, colnum, rownum, col.CellValue, col.CellStyle, rowStyle, col.CellValueAction, isFormulaValue, col.FormulaCellValue);
			}
		}

		public void SetRowExcelCells(ISheet sheet, DataTable dataTable, List<ExcelCellParam> param, ExcelColumns startColnum, int startRownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			if (startRownum < 1) startRownum = 1;

			for (int dtIndex = 0; dtIndex < dataTable.Rows.Count; dtIndex++)
			{
				var rownum = startRownum + dtIndex;
				SetColExcelCells(sheet, dataTable, dtIndex, param, startColnum, rownum, rowStyle, isFormula);
			}
		}


	}
}
