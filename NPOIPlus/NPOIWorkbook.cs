using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Streaming.Values;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

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

		private void SetCellStyleBasedOnType(object cellValue, ICellStyle style)
		{
			bool isInt = true;
			bool isDouble = true;
			bool isDateTime = true;
			if (cellValue == DBNull.Value) return;

			if (!int.TryParse(cellValue.ToString(), out int i)) isInt = false;

			if (!double.TryParse(cellValue.ToString(), out double d)) isDouble = false;

			if (!DateTime.TryParse(cellValue.ToString(), out DateTime dt)) isDateTime = false;

			// 動態調整型別
			if (isInt)
			{
				SetDefaultIntCellStyle?.Invoke(style);
			}
			else if (isDouble)
			{
				SetDefaultDoubleCellStyle?.Invoke(style);
			}
			else if (isDateTime)
			{
				SetDefaultDateTimeCellStyle?.Invoke(style);
			}
			else
			{
				SetDefaultStringCellStyle?.Invoke(style);
			}
		}


		private void SetCellValueBasedOnType(ICell cell, object cellValue, CellValueActionType valueAction = null)
		{
			cellValue = valueAction?.Invoke(cell, cellValue) ?? cellValue;
			bool isInt = true;
			bool isDouble = true;
			bool isDateTime = true;
			if (cellValue == DBNull.Value) return;

			if (!int.TryParse(cellValue.ToString(), out int i)) isInt = false;

			if (!double.TryParse(cellValue.ToString(), out double d)) isDouble = false;

			if (!DateTime.TryParse(cellValue.ToString(), out DateTime dt)) isDateTime = false;

			// 動態調整型別
			if (isInt)
			{
				var intValue = SetDefaultIntCellValue(i);
				cell.SetCellValue(intValue);
			}
			else if (isDouble)
			{
				var doubleValue = SetDefaultDoubleCellValue(d);
				cell.SetCellValue(doubleValue);
			}
			else if (isDateTime)
			{
				var dateValue = SetDefaultDateTimeCellValue(dt);
				cell.SetCellValue(dateValue);
			}
			else
			{
				var stringValue = SetDefaultStringCellValue(cellValue?.ToString());
				cell.SetCellValue(stringValue);
			}
		}


		private void SetCellStyle(ICell cell, object cellValue, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, ExcelColumns colnum = 0, int rownum = 1)
		{
			//ICellStyle newCellStyle = Workbook.CreateCellStyle();
			ICellStyle newCellStyle = GetOrCreateStyle($"GlobalStyle");
			SetGlobalCellStyle(newCellStyle);
			SetCellStyleBasedOnType(cellValue, newCellStyle);
			if (rowStyle != null)
			{
				newCellStyle = GetOrCreateStyle($"RowStyle_{rownum}");
			}
			rowStyle?.Invoke(newCellStyle);
			if (colStyle != null)
			{
				newCellStyle = GetOrCreateStyle($"ColStyle_{colnum}");
			}
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
		public void SetExcelCell<T>(ISheet sheet, T cellValue, ExcelColumns colnum, int rownum, Action<ICellStyle> style = null)
		{
			if (rownum < 1) rownum = 1;
			int zeroBaseIndex = rownum - 1;
			IRow row = sheet.GetRow(zeroBaseIndex) ?? sheet.CreateRow(zeroBaseIndex);
			ICell cell = row.CreateCell((int)colnum);
			SetCellStyle(cell, cellValue, style);
			SetCellValueBasedOnType(cell, cellValue);
		}

		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns column, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null, bool? isFormula = null)
		{
			SetExcelCell(sheet, dataTable, tableIndex, tableColName, column, rownum, cellValue, colStyle, null, null, isFormula);
		}

		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns column, CellValueActionType cellValueAction, int rownum = 1, Action<ICellStyle> colStyle = null, bool? isFormula = null)
		{
			SetExcelCell(sheet, dataTable, tableIndex, tableColName, column, rownum, null, colStyle, null, cellValueAction, isFormula);
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
		private void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns colnum, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = false)
		{
			if (rownum < 1) rownum = 1;
			int zeroBaseIndex = rownum - 1;
			IRow row = sheet.GetRow(zeroBaseIndex) ?? sheet.CreateRow(zeroBaseIndex);
			ICell cell = row.CreateCell((int)colnum);
			var newValue = cellValueAction ?? cellValue ?? dataTable.Rows[tableIndex][tableColName];
			SetCellStyle(cell, newValue, colStyle, rowStyle);
			if (isFormula.HasValue)
			{
				if (isFormula.Value)
				{
					string newCellValue = cellValueAction?.Invoke(cell, cellValue, rownum, colnum);
					cell.SetCellFormula(newCellValue);
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
				SetExcelCell(sheet, dataTable, tableIndex, col.ColumnName, colnum, rownum, col.CellValue, col.CellStyle, rowStyle, col.CellValueAction, isFormulaValue);
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

		public void RemovwRowRange(ISheet sheet, int startRow = 1, int endRow = 2)
		{
			if (startRow < 1) startRow = 1;
			if (endRow < 2) endRow = 2;
			startRow = startRow - 1;
			endRow = endRow - 1;
			for (int i = endRow; i >= startRow; i--)
			{
				IRow row = sheet.GetRow(i);
				if (row != null)
				{
					sheet.RemoveRow(row);
				}
			}
		}


	}

	public class NpoiMemoryStream : MemoryStream
	{
		public NpoiMemoryStream()
		{
			// We always want to close streams by default to
			// force the developer to make the conscious decision
			// to disable it.  Then, they're more apt to remember
			// to re-enable it.  The last thing you want is to
			// enable memory leaks by default.  ;-)
			AllowClose = true;
		}

		public bool AllowClose { get; set; }

		public override void Close()
		{
			if (AllowClose)
				base.Close();
		}
	}
}
