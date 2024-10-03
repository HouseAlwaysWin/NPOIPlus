using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Streaming.Values;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

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

		public List<ExcelStyleCached> _cellStylesCached = new List<ExcelStyleCached>();
		public Dictionary<string, ICellStyle> _globalCellStyleCached = new Dictionary<string, ICellStyle>();

		public NPOIWorkbook(IWorkbook workbook)
		{
			Workbook = workbook;
		}

		#region CellStyle

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

		private string SetGlobalStyleKeyBasedOnType(object cellValue, string key)
		{
			bool isInt = true;
			bool isDouble = true;
			bool isDateTime = true;

			if (!int.TryParse(cellValue.ToString(), out int i)) isInt = false;

			if (!double.TryParse(cellValue.ToString(), out double d)) isDouble = false;

			if (!DateTime.TryParse(cellValue.ToString(), out DateTime dt)) isDateTime = false;

			// 動態調整型別
			if (isInt)
			{
				key = $"Int_{key}";
			}
			else if (isDouble)
			{
				key = $"double_{key}";
			}
			else if (isDateTime)
			{
				key = $"date_{key}";
			}
			else
			{
				key = $"str_{key}";
			}
			return key;
		}


		private void SetCellStyle(string cachedKey, ICell cell, object cellValue, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, ExcelColumns colnum = 0, int rownum = 1)
		{
			// 根據列號和行號或全局樣式鍵生成樣式key
			var styleGroup = _cellStylesCached.FirstOrDefault(s => s.GroupName == cachedKey);
			var style = styleGroup?.CellStyles;

			string key = SetGlobalStyleKeyBasedOnType(cellValue, "GlobalStyle");

			if (colStyle != null)
			{
				key = $"OnlyCellStyle_{colnum}";
			}
			else if (colStyle == null && rowStyle != null)
			{
				key = SetGlobalStyleKeyBasedOnType(cellValue, "GlobalRowStyle");
			}
			else
			{
				style = _globalCellStyleCached;
			}

			// 檢查是否已有樣式
			if (style.ContainsKey(key))
			{
				cell.CellStyle = style[key];  // 直接使用已存在的樣式
			}
			else
			{
				ICellStyle newCellStyle = Workbook.CreateCellStyle(); ;
				SetGlobalCellStyle(newCellStyle);
				SetCellStyleBasedOnType(cellValue, newCellStyle);
				rowStyle?.Invoke(newCellStyle);
				colStyle?.Invoke(newCellStyle);
				cell.CellStyle = newCellStyle;

				style.Add(key, newCellStyle);
			}
		}

		#endregion

		private void SetCellValueBasedOnType(ICell cell, object cellValue, CellValueActionType valueAction = null, ExcelColumns colnum = 0, int rownum = 1)
		{
			cellValue = valueAction?.Invoke(cell, cellValue, colnum, rownum) ?? cellValue;
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

		private object GetPropertyValue<T>(T obj, string propertyName)
		{
			if (obj != null && typeof(T).IsClass)
			{
				var parameter = Expression.Parameter(typeof(T), "obj");

				// 嘗試取得屬性
				MemberInfo memberInfo = typeof(T).GetProperty(propertyName)
										?? (MemberInfo)typeof(T).GetField(propertyName);

				if (memberInfo == null)
				{
					return null;
				}

				// 根據是屬性還是字段來動態生成訪問的 Expression
				Expression memberAccess;
				if (memberInfo is PropertyInfo property)
				{
					memberAccess = Expression.Property(parameter, property);
				}
				else if (memberInfo is FieldInfo field)
				{
					memberAccess = Expression.Field(parameter, field);
				}
				else
				{
					return null;
				}

				// 將結果轉換為 object
				var convert = Expression.Convert(memberAccess, typeof(object));

				// 創建 Lambda 表達式：obj => (object)obj.Member
				var lambda = Expression.Lambda<Func<T, object>>(convert, parameter);

				// 編譯並返回委託，然後取得值
				var compiledLambda = lambda.Compile();
				return compiledLambda(obj);
			}
			return null;
		}

		public void SetExcelCell<T>(ISheet sheet, T cellValue, ExcelColumns colnum, int rownum, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = null)
		{
			if (rownum < 1) rownum = 1;
			var sheetName = sheet.SheetName;
			var key = $"SetCell{sheetName}_{colnum}{rownum}";
			if (colStyle != null && rowStyle != null)
			{
				if (_cellStylesCached.FirstOrDefault(s => s.GroupName == key) == null)
				{
					_cellStylesCached.Add(new ExcelStyleCached
					{
						GroupName = key,
						CellStyles = new Dictionary<string, ICellStyle>()
					});
				}
			}
			SetExcelCell(sheet, key, cellValue, colnum, rownum, colStyle, rowStyle, cellValueAction, isFormula);
		}

		private void SetExcelCell<T>(ISheet sheet, string groupKey, T cellValue, ExcelColumns colnum, int rownum, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = null)
		{
			int zeroBaseIndex = rownum - 1;
			IRow row = sheet.GetRow(zeroBaseIndex) ?? sheet.CreateRow(zeroBaseIndex);
			ICell cell = row.CreateCell((int)colnum);

			//var newValue = dataTable.Columns.Contains(tableColName) ? dataTable.Rows?[tableIndex]?[tableColName] : cellValue;
			SetCellStyle(groupKey, cell, cellValue, colStyle, rowStyle, colnum, rownum);
			if (isFormula.HasValue)
			{
				if (isFormula.Value)
				{
					string newCellValue = cellValueAction?.Invoke(cell, cellValue, colnum, rownum);
					cell.SetCellFormula(newCellValue);
					return;
				}
			}
			SetCellValueBasedOnType(cell, cellValue, cellValueAction, colnum, rownum);
		}

		private void SetExcelCell<T>(ISheet sheet, string groupKey, IEnumerable<T> dataTable, int tableIndex, string tableColName, ExcelColumns colnum, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = false)
		{
			if (rownum < 1) rownum = 1;
			IRow row = sheet.GetExcelRowOrCreate(rownum);
			ICell cell = row.CreateCell((int)colnum);
			var newValue = GetPropertyValue(dataTable.ElementAt(tableIndex), tableColName) ?? cellValue;
			SetCellStyle(groupKey, cell, newValue, colStyle, rowStyle, colnum, rownum);
			if (isFormula.HasValue)
			{
				if (isFormula.Value)
				{
					string newCellValue = cellValueAction?.Invoke(cell, cellValue, colnum, rownum);
					cell.SetCellFormula(newCellValue);
					return;
				}
			}

			SetCellValueBasedOnType(cell, newValue, cellValueAction, colnum, rownum);
		}


		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns column, int rownum = 1, Action<ICellStyle> colStyle = null, bool? isFormula = null)
		{
			SetExcelCell(sheet, dataTable, tableIndex, tableColName, column, rownum, null, colStyle, null, null, isFormula);
		}

		private void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns colnum, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = false)
		{
			var sheetName = sheet.SheetName;
			var key = $"SetCell{sheetName}_{colnum}{rownum}";
			if (colStyle != null || rowStyle != null)
			{
				if (_cellStylesCached.FirstOrDefault(s => s.GroupName == key) == null)
				{
					_cellStylesCached.Add(new ExcelStyleCached
					{
						GroupName = key,
						CellStyles = new Dictionary<string, ICellStyle>()
					});
				}
			}
			SetExcelCell(sheet, key, dataTable, tableIndex, tableColName, cellValue, colnum, rownum, colStyle, rowStyle, cellValueAction, isFormula);
		}


		private void SetExcelCell(ISheet sheet, string groupKey, DataTable dataTable, int tableIndex, string tableColName, object cellValue = null, ExcelColumns colnum = 0, int rownum = 1, Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = false)
		{
			if (rownum < 1) rownum = 1;
			IRow row = sheet.GetExcelRowOrCreate(rownum);
			ICell cell = row.CreateCell((int)colnum);
			var newValue = dataTable.Columns.Contains(tableColName) ? dataTable.Rows?[tableIndex]?[tableColName] : cellValue;
			SetCellStyle(groupKey, cell, newValue, colStyle, rowStyle, colnum, rownum);
			if (isFormula.HasValue)
			{
				if (isFormula.Value)
				{
					string newCellValue = cellValueAction?.Invoke(cell, cellValue, colnum, rownum);
					cell.SetCellFormula(newCellValue);
					return;
				}
			}

			SetCellValueBasedOnType(cell, newValue, cellValueAction, colnum, rownum);
		}


		public void SetColExcelCells(ISheet sheet, DataTable dataTable, int tableIndex, List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			var sheetName = sheet.SheetName;
			var key = $"SetCol{sheetName}_{startColnum}{rownum}";
			if (_cellStylesCached.FirstOrDefault(s => s.GroupName == key) == null)
			{
				_cellStylesCached.Add(new ExcelStyleCached
				{
					GroupName = key,
					CellStyles = new Dictionary<string, ICellStyle>()
				});
			}
			SetColExcelCells(sheet, key, dataTable, tableIndex, param, startColnum, rownum, rowStyle, isFormula);
		}

		private void SetColExcelCells(ISheet sheet, string groupKey, DataTable dataTable, int tableIndex, List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			for (int colIndex = 0; colIndex < param.Count; colIndex++)
			{
				var colnum = colIndex + startColnum;
				var col = param[colIndex];
				var isFormulaValue = col.IsFormula.HasValue ? col.IsFormula : isFormula;
				SetExcelCell(sheet, groupKey, dataTable, tableIndex, col.ColumnName, col.CellValue, colnum, rownum, col.CellStyle, rowStyle, col.CellValueAction, isFormulaValue);
			}
		}
		public void SetColExcelCells<T>(ISheet sheet, List<T> table, int tableIndex, List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			var sheetName = sheet.SheetName;
			var key = $"SetCol{sheetName}_{startColnum}{rownum}";
			if (_cellStylesCached.FirstOrDefault(s => s.GroupName == key) == null)
			{
				_cellStylesCached.Add(new ExcelStyleCached
				{
					GroupName = key,
					CellStyles = new Dictionary<string, ICellStyle>()
				});
			}
			for (int colIndex = 0; colIndex < param.Count; colIndex++)
			{
				var colnum = colIndex + startColnum;
				var col = param[colIndex];
				var isFormulaValue = col.IsFormula.HasValue ? col.IsFormula : isFormula;
				SetExcelCell(sheet, key, table, tableIndex, col.ColumnName, colnum, rownum, col.CellValue, col.CellStyle, rowStyle, col.CellValueAction, isFormulaValue);
			}
		}

		private void SetColExcelCells<T>(ISheet sheet, string groupKey, IEnumerable<T> table, int tableIndex, List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			for (int colIndex = 0; colIndex < param.Count; colIndex++)
			{
				var colnum = colIndex + startColnum;
				var col = param[colIndex];
				var isFormulaValue = col.IsFormula.HasValue ? col.IsFormula : isFormula;
				SetExcelCell(sheet, groupKey, table, tableIndex, col.ColumnName, colnum, rownum, col.CellValue, col.CellStyle, rowStyle, col.CellValueAction, isFormulaValue);
			}
		}


		public void SetRowExcelCells(ISheet sheet, DataTable dataTable, List<ExcelCellParam> param, ExcelColumns startColnum, int startRownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			if (startRownum < 1) startRownum = 1;
			var sheetName = sheet.SheetName;
			var key = $"SetRow_{sheetName}_{startColnum}{startRownum}";
			if (rowStyle != null || param.Any(p => p.CellStyle != null))
			{
				if (_cellStylesCached.FirstOrDefault(s => s.GroupName == key) == null)
				{
					_cellStylesCached.Add(new ExcelStyleCached
					{
						GroupName = key,
						CellStyles = new Dictionary<string, ICellStyle>()
					});
				}
			}

			for (int dtIndex = 0; dtIndex < dataTable.Rows.Count; dtIndex++)
			{
				var rownum = startRownum + dtIndex;
				SetColExcelCells(sheet, key, dataTable, dtIndex, param, startColnum, rownum, rowStyle, isFormula);
			}
		}

		public void SetRowExcelCells<T>(ISheet sheet, IEnumerable<T> table, List<ExcelCellParam> param, ExcelColumns startColnum, int startRownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null)
		{
			if (startRownum < 1) startRownum = 1;
			var sheetName = sheet.SheetName;
			var key = $"SetRow_{sheetName}_{startColnum}{startRownum}";
			if (rowStyle != null || param.Any(p => p.CellStyle != null))
			{
				if (_cellStylesCached.FirstOrDefault(s => s.GroupName == key) == null)
				{
					_cellStylesCached.Add(new ExcelStyleCached
					{
						GroupName = key,
						CellStyles = new Dictionary<string, ICellStyle>()
					});
				}
			}

			for (int tIndex = 0; tIndex < table.Count(); tIndex++)
			{
				var rownum = startRownum + tIndex;
				SetColExcelCells(sheet, key, table, tIndex, param, startColnum, rownum, rowStyle, isFormula);
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
