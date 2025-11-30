using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Streaming.Values;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace NPOIPlus
{

	public interface IWorkbookStage
	{
		IWorkbook GetWorkbook();
		IWorkbookStage ReadExcelFile(string filePath);
		IWorkbookStage SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles);
		IWorkbookStage SetupCellStyle(string cellStyleKey, Action<IWorkbook, ICellStyle> styles);
		ISheetStage UseSheet(string sheetName, bool createIfMissing = true);
		ISheetStage UseSheet(ISheet sheet);
		ISheetStage UseSheetAt(int index, bool createIfMissing = false);
	}

	public interface ISheetStage
	{
		ISheet GetSheet();
		ITableStage<T> SetTable<T>(IEnumerable<T> table, ExcelColumns startCol, int startRow);
		ICellStage SetCell(ExcelColumns startCol, int startRow);
		ISheetStage SetColumnWidth(ExcelColumns col, int width);
		ISheetStage SetColumnWidth(ExcelColumns startCol, ExcelColumns endCol, int width);
		ISheetStage SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int row);
	}


	public interface ITableStage
	{
		// ITableStage SetCell(string cellName, object value);
		ITableCellStage BeginCellSet(string cellName);
		ITableHeaderStage BeginTitleSet(string title);
		ITableStage SetRow();
		FluentMemoryStream ToStream();
		IWorkbook SaveToPath(string filePath);
	}


	public interface ITableHeaderStage
	{
		ITableCellStage BeginBodySet(string cellName);
		ITableHeaderStage SetValue(Func<TableCellParams, object> valueAction);
		ITableHeaderStage SetFormulaValue(Func<TableCellParams, object> valueAction);
		ITableHeaderStage SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction);
		ITableHeaderStage SetCellStyle(string cellStyleKey);
		ITableHeaderStage SetCellType(CellType cellType);
		ITableHeaderStage CopyStyleFromCell(ExcelColumns col, int rowIndex);
	}

	public interface ITableCellStage
	{
		ITableCellStage SetValue(object value);
		ITableCellStage SetValue(Func<TableCellParams, object> valueAction);
		ITableCellStage SetFormulaValue(Func<TableCellParams, object> valueAction);
		ITableCellStage SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction);
		ITableCellStage SetCellStyle(string cellStyleKey);
		ITableCellStage SetCellType(CellType cellType);
		ITableCellStage CopyStyleFromCell(ExcelColumns col, int rowIndex);
		ITableStage End();
	}

	// 泛型版鏈式介面（不破壞舊介面，並存）
	public interface ITableStage<T>
	{
		ITableCellStage<T> BeginCellSet(string cellName);
		ITableHeaderStage<T> BeginTitleSet(string title);
		ITableStage<T> BuildRows();
		FluentMemoryStream ToStream();
		IWorkbook SaveToPath(string filePath);
	}
	public interface ITableHeaderStage<T>
	{
		ITableCellStage<T> BeginBodySet(string cellName);
		ITableHeaderStage<T> SetValue(Func<TableCellParams<T>, object> valueAction);
		ITableHeaderStage<T> SetFormulaValue(Func<TableCellParams<T>, object> valueAction);
		ITableHeaderStage<T> SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction);
		ITableHeaderStage<T> SetCellStyle(string cellStyleKey);
		ITableHeaderStage<T> SetCellType(CellType cellType);
		ITableHeaderStage<T> CopyStyleFromCell(ExcelColumns col, int rowIndex);
	}
	public interface ITableCellStage<T>
	{
		ITableCellStage<T> SetValue(object value);
		ITableCellStage<T> SetValue(Func<TableCellParams<T>, object> valueAction);
		ITableCellStage<T> SetFormulaValue(Func<TableCellParams<T>, object> valueAction);
		ITableCellStage<T> SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction);
		ITableCellStage<T> SetCellStyle(string cellStyleKey);
		ITableCellStage<T> SetCellType(CellType cellType);
		ITableCellStage<T> CopyStyleFromCell(ExcelColumns col, int rowIndex);
		ITableStage<T> End();
	}

	public interface ICellStage
	{
		ICellStage SetValue<T>(T value);
		FluentMemoryStream Save();
	}



	public class FluentWorkbook : IWorkbookStage
	{
		private IWorkbook _workbook;
		private ISheet _currentSheet;
		private Dictionary<string, ICellStyle> _cellStylesCached = new Dictionary<string, ICellStyle>();
		public FluentWorkbook(IWorkbook workbook)
		{
			_workbook = workbook;
		}

		public IWorkbook GetWorkbook()
		{
			return _workbook;
		}

		public IWorkbookStage ReadExcelFile(string filePath)
		{
			if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
			if (!File.Exists(filePath)) throw new FileNotFoundException("Excel file not found.", filePath);

			_currentSheet = null;

			string ext = Path.GetExtension(filePath)?.ToLowerInvariant();

			// 以讀取模式開啟並立即讀入記憶體，讀完即釋放檔案鎖
			using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				if (ext == ".xls")
				{
					_workbook = new HSSFWorkbook(fs);
				}
				else
				{
					// 預設使用 XSSF，支援 .xlsx/.xlsm
					_workbook = new XSSFWorkbook(fs);
				}
			}

			// 預設選取第一個工作表
			if (_workbook.NumberOfSheets > 0)
			{
				_currentSheet = _workbook.GetSheetAt(0);
			}

			return this;
		}

		public IWorkbookStage SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
		{
			ICellStyle newCellStyle = _workbook.CreateCellStyle();
			styles(_workbook, newCellStyle);
			_cellStylesCached.Clear();
			_cellStylesCached.Add("global", newCellStyle);
			return this;
		}

		public IWorkbookStage SetupCellStyle(string cellStyleKey, Action<IWorkbook, ICellStyle> styles)
		{
			ICellStyle newCellStyle = _workbook.CreateCellStyle();
			styles(_workbook, newCellStyle);
			_cellStylesCached.Add(cellStyleKey, newCellStyle);
			return this;
		}

		public ISheetStage UseSheet(string sheetName, bool createIfMissing = true)
		{
			_currentSheet = _workbook.GetSheet(sheetName);
			if (_currentSheet == null && createIfMissing)
			{
				_currentSheet = _workbook.CreateSheet(sheetName);
			}
			return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
		}

		public ISheetStage UseSheet(ISheet sheet)
		{
			_currentSheet = sheet;
			return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
		}

		public ISheetStage UseSheetAt(int index, bool createIfMissing = false)
		{
			_currentSheet = _workbook.GetSheetAt(index);
			if (_currentSheet == null && createIfMissing)
			{
				_currentSheet = _workbook.CreateSheet();
			}
			return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
		}


	}

	public class FluentSheet : ISheetStage
	{
		private ISheet _sheet;
		private IWorkbook _workbook;
		private Dictionary<string, ICellStyle> _cellStylesCached;

		public FluentSheet(IWorkbook workbook, ISheet sheet, Dictionary<string, ICellStyle> cellStylesCached)
		{
			_workbook = workbook;
			_sheet = sheet;
			_cellStylesCached = cellStylesCached;
		}

		public ISheet GetSheet()
		{
			return _sheet;
		}

		public ISheetStage SetColumnWidth(ExcelColumns col, int width)
		{
			_sheet.SetColumnWidth((int)col, width * 256);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public ISheetStage SetColumnWidth(ExcelColumns startCol, ExcelColumns endCol, int width)
		{
			for (int i = (int)startCol; i <= (int)endCol; i++)
			{
				_sheet.SetColumnWidth(i, width * 256);
			}
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}


		public ISheetStage SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int row)
		{
			_sheet.SetExcelCellMerge(startCol, endCol, row);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}

		public ISheetStage SetExcelCellMerge(ExcelColumns startCol, ExcelColumns endCol, int firstRow, int lastRow)
		{
			_sheet.SetExcelCellMerge(startCol, endCol, firstRow, lastRow);
			return new FluentSheet(_workbook, _sheet, _cellStylesCached);
		}


		public ICellStage SetCell(ExcelColumns col, int row)
		{
			if (_sheet == null) throw new InvalidOperationException("No active sheet. Call UseSheet(...) first.");

			var rowObj = _sheet.GetRow(row) ?? _sheet.CreateRow(row);
			var cell = rowObj.GetCell((int)col) ?? rowObj.CreateCell((int)col);
			return new FluentCell(_workbook, _sheet, cell);
		}

		public ITableStage<T> SetTable<T>(IEnumerable<T> table, ExcelColumns startCol, int startRow)
		{
			return new FluentTable<T>(_workbook, _sheet, table, startCol, startRow, _cellStylesCached, new List<TableCellSet>(), new List<TableCellSet>());
		}
	}


	public class FluentTable<T> : ITableStage, ITableStage<T>
	{
		private ISheet _sheet;
		private IWorkbook _workbook;
		private IEnumerable<T> _table;
		private IList<T> _itemsCache;
		private ExcelColumns _startCol;
		private int _startRow;
		private List<TableCellSet> _cellBodySets;
		private List<TableCellSet> _cellTitleSets;
		private Dictionary<string, ICellStyle> _cellStylesCached;
		public FluentTable(IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
		ExcelColumns startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached, List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
		{
			_workbook = workbook;
			_sheet = sheet;
			_table = table;
			_startCol = NormalizeStartCol(startCol);
			_startRow = NormalizeStartRow(startRow);
			_cellStylesCached = cellStylesCached;
			_cellTitleSets = cellTitleSets;
			_cellBodySets = cellBodySets;
		}

		private ExcelColumns NormalizeStartCol(ExcelColumns col)
		{
			int idx = (int)col;
			if (idx < 0) idx = 0;
			return (ExcelColumns)idx;
		}

		private int NormalizeStartRow(int row)
		{
			// 將使用者常見的 1-based 列號轉為 0-based，並確保不為負數
			if (row < 1) return 0;
			return row - 1;
		}

		private IList<T> GetItems()
		{
			if (_itemsCache != null) return _itemsCache;
			_itemsCache = _table as IList<T> ?? _table?.ToList() ?? new List<T>();
			return _itemsCache;
		}

		private T GetItemAt(int index)
		{
			var items = GetItems();
			if (index < 0 || index >= items.Count) return default;
			return items[index];
		}

		private object GetTableCellValue(string cellName, object item)
		{
			if (string.IsNullOrWhiteSpace(cellName) || item == null) return default;

			object value = null;

			if (item is DataRow dr)
			{
				if (dr.Table != null && dr.Table.Columns.Contains(cellName))
					value = dr[cellName];
			}
			else if (item is IDictionary<string, object> dictObj)
			{
				dictObj.TryGetValue(cellName, out value);
			}
			else if (item is IDictionary<string, string> dictStr)
			{
				if (dictStr.TryGetValue(cellName, out var s))
					value = s;
			}
			else
			{
				var type = item.GetType();
				var prop = type.GetProperty(cellName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
				if (prop != null)
				{
					value = prop.GetValue(item);
				}
				else
				{
					var field = type.GetField(cellName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
					if (field != null)
						value = field.GetValue(item);
				}
			}

			if (value == null || value == DBNull.Value) return default;
			return value;
		}

		private void SetCellValue(ICell cell, object value)
		{
			if (value is bool b)
			{
				cell.SetCellValue(b);
				return;
			}
			if (value is DateTime dt)
			{
				cell.SetCellValue(dt);
				return;
			}
			if (value is int i)
			{
				cell.SetCellValue((double)i);
				return;
			}
			if (value is long l)
			{
				cell.SetCellValue((double)l);
				return;
			}
			if (value is float f)
			{
				cell.SetCellValue((double)f);
				return;
			}
			if (value is double d)
			{
				cell.SetCellValue(d);
				return;
			}
			if (value is decimal m)
			{
				cell.SetCellValue((double)m);
				return;
			}

			cell.SetCellValue(value.ToString());
		}

		private void SetCellValue(ICell cell, object value, CellType cellType)
		{
			if (cell == null)
				return;

			if (value == null || value == DBNull.Value)
			{
				cell.SetCellValue(string.Empty);
				return;
			}

			// 1) 先依據 value 的實際型別寫入
			SetCellValue(cell, value);

			// 2) 若指定了 CellType（且非 Unknown），以 CellType 覆寫
			if (cellType == CellType.Unknown) return;
			if (cellType == CellType.Formula)
			{
				SetFormulaValue(cell, value);
				return;
			}

			var text = value.ToString();
			switch (cellType)
			{
				case CellType.Boolean:
					{
						if (bool.TryParse(text, out var bv)) { cell.SetCellValue(bv); return; }
						if (int.TryParse(text, out var iv)) { cell.SetCellValue(iv != 0); return; }
						cell.SetCellValue(!string.IsNullOrEmpty(text));
						return;
					}
				case CellType.Numeric:
					{
						if (double.TryParse(text, out var dv)) { cell.SetCellValue(dv); return; }
						if (DateTime.TryParse(text, out var dtv)) { cell.SetCellValue(dtv); return; }
						// 若無法轉換為數值/日期則保留前一步的寫入結果
						return;
					}
				case CellType.String:
					{
						cell.SetCellValue(text);
						return;
					}
				case CellType.Blank:
					{
						cell.SetCellValue(string.Empty);
						return;
					}
				case CellType.Error:
					{
						// NPOI 錯誤型別無從 object 直接設定，退為字串呈現
						cell.SetCellValue(text);
						return;
					}
				default:
					return;
			}
		}

		private void SetFormulaValue(ICell cell, object value)
		{
			if (cell == null) return;
			if (value == null || value == DBNull.Value) return;

			var formula = value.ToString();
			if (string.IsNullOrWhiteSpace(formula)) return;

			// NPOI SetCellFormula 需要純公式字串（不含 '='）
			if (formula.StartsWith("=")) formula = formula.Substring(1);

			cell.SetCellFormula(formula);
		}

		private void SetCellStyle(ICell cell, TableCellSet cellNameMap, TableCellStyleParams cellStyleParams)
		{
			if (!string.IsNullOrWhiteSpace(cellNameMap.CellStyleKey) && _cellStylesCached.ContainsKey(cellNameMap.CellStyleKey))
			{
				cell.CellStyle = _cellStylesCached[cellNameMap.CellStyleKey];
			}
			else if (string.IsNullOrWhiteSpace(cellNameMap.CellStyleKey) &&
					 _cellStylesCached.ContainsKey("global") &&
					 cellNameMap.SetCellStyleAction == null)
			{
				cell.CellStyle = _cellStylesCached["global"];
			}
			else if (!string.IsNullOrWhiteSpace(cellNameMap.CellStyleKey) && cellNameMap.SetCellStyleAction != null)
			{
				ICellStyle newCellStyle = _workbook.CreateCellStyle();
				cellNameMap.SetCellStyleAction(cellStyleParams, newCellStyle);
				cellNameMap.CellStyleKey = cellNameMap.CellStyleKey;
				_cellStylesCached.Add(cellNameMap.CellStyleKey, newCellStyle);
				cell.CellStyle = newCellStyle;
			}
		}

		private ITableStage SetRow(int rowOffset = 0)
		{
			if (_cellBodySets == null || _cellBodySets.Count == 0) return this;

			var targetRowIndex = _startRow + rowOffset;

			var item = GetItemAt(rowOffset);

			int colIndex = (int)_startCol;
			if (_cellTitleSets != null && _cellTitleSets.Count > 0)
			{
				var titleRowObj = _sheet.GetRow(_startRow) ?? _sheet.CreateRow(_startRow);
				SetCellAction(_cellTitleSets, titleRowObj, colIndex, _startRow, item);
				targetRowIndex++;
			}

			var rowObj = _sheet.GetRow(targetRowIndex) ?? _sheet.CreateRow(targetRowIndex);
			SetCellAction(_cellBodySets, rowObj, colIndex, targetRowIndex, item);

			return this;
		}

		private void SetCellAction(List<TableCellSet> cellSets, IRow rowObj, int colIndex, int targetRowIndex, object item)
		{
			foreach (var cellset in cellSets)
			{
				var cell = rowObj.GetCell(colIndex) ?? rowObj.CreateCell(colIndex);

				// 優先使用 TableCellNameMap 中的 Value，如果沒有則從 item 中獲取
				Func<TableCellParams, object> setValueAction = cellset.SetValueAction;
				Func<TableCellParams, object> setFormulaValueAction = cellset.SetFormulaValueAction;

				TableCellParams cellParams = new TableCellParams
				{
					ColNum = (ExcelColumns)colIndex,
					RowNum = targetRowIndex,
					RowItem = item
				};
				object value = cellset.CellValue ?? GetTableCellValue(cellset.CellName, item);
				cellParams.CellValue = value;

				// 準備泛型參數（供泛型委派使用）
				var cellParamsT = new TableCellParams<T>
				{
					ColNum = (ExcelColumns)colIndex,
					RowNum = targetRowIndex,
					RowItem = item is T tItem ? tItem : default,
					CellValue = value
				};

				TableCellStyleParams cellStyleParams =
				new TableCellStyleParams
				{
					Workbook = _workbook,
					ColNum = (ExcelColumns)colIndex,
					RowNum = targetRowIndex,
				};
				SetCellStyle(cell, cellset, cellStyleParams);

				if (cellset.CellType == CellType.Formula)
				{
					if (cellset.SetFormulaValueActionGeneric != null)
					{
						if (cellset.SetFormulaValueActionGeneric is Func<TableCellParams<T>, object> gFormula)
						{
							value = gFormula(cellParamsT);
						}
						else
						{
							value = cellset.SetFormulaValueActionGeneric.DynamicInvoke(cellParamsT);
						}
					}
					else if (setFormulaValueAction != null)
					{
						value = setFormulaValueAction(cellParams);
					}
					SetFormulaValue(cell, value);
				}
				else
				{
					if (cellset.SetValueActionGeneric != null)
					{
						if (cellset.SetValueActionGeneric is Func<TableCellParams<T>, object> gValue)
						{
							value = gValue(cellParamsT);
						}
						else
						{
							value = cellset.SetValueActionGeneric.DynamicInvoke(cellParamsT);
						}
					}
					else if (setValueAction != null)
					{
						value = setValueAction(cellParams);
					}
					SetCellValue(cell, value, cellset.CellType);
				}


				colIndex++;
			}
		}

		public ITableCellStage BeginCellSet(string cellName)
		{
			_cellBodySets.Add(new TableCellSet { CellName = cellName });
			return new FluentTableCellStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
		}
		ITableCellStage<T> ITableStage<T>.BeginCellSet(string cellName)
		{
			_cellBodySets.Add(new TableCellSet { CellName = cellName });
			return new FluentTableCellStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
		}

		public ITableHeaderStage BeginTitleSet(string title)
		{
			_cellTitleSets.Add(new TableCellSet { CellName = $"{title}_TITLE", CellValue = title });
			return new FluentTableHeaderStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, title, _cellTitleSets, _cellBodySets);
		}
		ITableHeaderStage<T> ITableStage<T>.BeginTitleSet(string title)
		{
			_cellTitleSets.Add(new TableCellSet { CellName = $"{title}_TITLE", CellValue = title });
			return new FluentTableHeaderStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, title, _cellTitleSets, _cellBodySets);
		}

		public ITableStage SetRow()
		{
			for (int i = 0; i < _table.Count(); i++)
			{
				SetRow(i);
			}
			return this;
		}
		ITableStage<T> ITableStage<T>.BuildRows()
		{
			for (int i = 0; i < _table.Count(); i++)
			{
				SetRow(i);
			}
			return this;
		}

		public FluentMemoryStream ToStream()
		{
			var ms = new FluentMemoryStream();
			ms.AllowClose = false;
			_workbook.Write(ms);
			ms.Flush();
			ms.Seek(0, SeekOrigin.Begin);
			ms.AllowClose = true;
			return ms;
		}
		FluentMemoryStream ITableStage<T>.ToStream()
		{
			return ToStream();
		}

		public IWorkbook SaveToPath(string filePath)
		{
			using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
			{
				_workbook.Write(outFile);
			}
			return _workbook;
		}
		IWorkbook ITableStage<T>.SaveToPath(string filePath)
		{
			return SaveToPath(filePath);
		}
	}

	public class FluentTableHeaderStage<T> : ITableHeaderStage, ITableHeaderStage<T>
	{
		private List<TableCellSet> _cellBodySets;
		private List<TableCellSet> _cellTitleSets;
		private TableCellSet _cellTitleSet;
		private IWorkbook _workbook;
		private ISheet _sheet;
		private IEnumerable<T> _table;
		private ExcelColumns _startCol;
		private int _startRow;
		private Dictionary<string, ICellStyle> _cellStylesCached;

		public FluentTableHeaderStage(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached,
			string title,
			List<TableCellSet> titleCellSets, List<TableCellSet> cellBodySets)
		{
			_workbook = workbook;
			_sheet = sheet;
			_table = table;
			_startCol = startCol;
			_startRow = startRow;
			_cellStylesCached = cellStylesCached;
			_cellBodySets = cellBodySets;
			_cellTitleSets = titleCellSets;
			_cellTitleSet = titleCellSets.FirstOrDefault(c => c.CellName == $"{title}_TITLE");
			_cellTitleSet.CellValue = _cellTitleSet.CellValue;
		}
		public ITableHeaderStage SetValue(Func<TableCellParams, object> valueAction)
		{
			_cellTitleSet.SetValueAction = valueAction;
			return this;
		}
		ITableHeaderStage<T> ITableHeaderStage<T>.SetValue(Func<TableCellParams<T>, object> valueAction)
		{
			_cellTitleSet.SetValueActionGeneric = valueAction;
			return this;
		}


		public ITableHeaderStage SetFormulaValue(object value)
		{
			_cellTitleSet.CellValue = value;
			_cellTitleSet.CellType = CellType.Formula;
			return this;
		}


		public ITableHeaderStage SetFormulaValue(Func<TableCellParams, object> valueAction)
		{
			_cellTitleSet.SetFormulaValueAction = valueAction;
			_cellTitleSet.CellType = CellType.Formula;
			return this;
		}
		ITableHeaderStage<T> ITableHeaderStage<T>.SetFormulaValue(Func<TableCellParams<T>, object> valueAction)
		{
			_cellTitleSet.SetFormulaValueActionGeneric = valueAction;
			_cellTitleSet.CellType = CellType.Formula;
			return this;
		}

		public ITableHeaderStage SetCellStyle(string cellStyleKey)
		{
			_cellTitleSet.CellStyleKey = cellStyleKey;
			return this;
		}
		ITableHeaderStage<T> ITableHeaderStage<T>.SetCellStyle(string cellStyleKey)
		{
			_cellTitleSet.CellStyleKey = cellStyleKey;
			return this;
		}

		public ITableHeaderStage SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			_cellTitleSet.CellStyleKey = cellStyleKey;
			_cellTitleSet.SetCellStyleAction = cellStyleAction;
			return this;
		}
		ITableHeaderStage<T> ITableHeaderStage<T>.SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			_cellTitleSet.CellStyleKey = cellStyleKey;
			_cellTitleSet.SetCellStyleAction = cellStyleAction;
			return this;
		}
		public ITableHeaderStage SetCellType(CellType cellType)
		{
			_cellTitleSet.CellType = cellType;
			return this;
		}
		ITableHeaderStage<T> ITableHeaderStage<T>.SetCellType(CellType cellType)
		{
			_cellTitleSet.CellType = cellType;
			return this;
		}
		public ITableCellStage BeginBodySet(string cellName)
		{
			_cellBodySets.Add(new TableCellSet { CellName = cellName });
			return new FluentTableCellStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
		}
		ITableCellStage<T> ITableHeaderStage<T>.BeginBodySet(string cellName)
		{
			_cellBodySets.Add(new TableCellSet { CellName = cellName });
			return new FluentTableCellStage<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, cellName, _cellTitleSets, _cellBodySets);
		}

	public ITableHeaderStage CopyStyleFromCell(ExcelColumns col, int rowIndex)
	{
		string key = $"{_sheet.SheetName}_{col}{rowIndex}";
		ICell cell = _sheet.GetExcelCell(col, rowIndex);
		if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(key))
		{
			SetCellStyle(key, (styleParams, style) =>
			{
				style.CloneStyleFrom(cell.CellStyle);
			});
		}
		return this;
	}
	ITableHeaderStage<T> ITableHeaderStage<T>.CopyStyleFromCell(ExcelColumns col, int rowIndex)
	{
		string key = $"{_sheet.SheetName}_{col}{rowIndex}";
		ICell cell = _sheet.GetExcelCell(col, rowIndex);
		if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(key))
		{
			SetCellStyle(key, (styleParams, style) =>
			{
				style.CloneStyleFrom(cell.CellStyle);
			});
		}
		return this;
	}
}

	public class FluentTableCellStage<T> : ITableCellStage, ITableCellStage<T>
	{
		private List<TableCellSet> _cellTitleSets;
		private List<TableCellSet> _cellBodySets;
		private TableCellSet _cellSet;
		private IWorkbook _workbook;
		private ISheet _sheet;
		private IEnumerable<T> _table;
		private ExcelColumns _startCol;
		private int _startRow;
		private Dictionary<string, ICellStyle> _cellStylesCached;

		public FluentTableCellStage(
			IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
			ExcelColumns startCol, int startRow,
			Dictionary<string, ICellStyle> cellStylesCached,
			string cellName,
			List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
		{
			_workbook = workbook;
			_sheet = sheet;
			_table = table;
			_startCol = startCol;
			_startRow = startRow;
			_cellStylesCached = cellStylesCached;
			_cellTitleSets = cellTitleSets;
			_cellBodySets = cellBodySets;
			_cellSet = cellBodySets.First(c => c.CellName == cellName);
		}

		public ITableCellStage SetValue(object value)
		{
			_cellSet.CellValue = value;
			return this;
		}

		public ITableCellStage SetValue(Func<TableCellParams, object> valueAction)
		{
			_cellSet.SetValueAction = valueAction;
			return this;
		}
		ITableCellStage<T> ITableCellStage<T>.SetValue(object value)
		{
			_cellSet.CellValue = value;
			return this;
		}
		ITableCellStage<T> ITableCellStage<T>.SetValue(Func<TableCellParams<T>, object> valueAction)
		{
			_cellSet.SetValueActionGeneric = valueAction;
			return this;
		}


		public ITableCellStage SetFormulaValue(object value)
		{
			_cellSet.CellValue = value;
			_cellSet.CellType = CellType.Formula;
			return this;
		}


		public ITableCellStage SetFormulaValue(Func<TableCellParams, object> valueAction)
		{
			_cellSet.SetFormulaValueAction = valueAction;
			_cellSet.CellType = CellType.Formula;
			return this;
		}
		ITableCellStage<T> ITableCellStage<T>.SetFormulaValue(Func<TableCellParams<T>, object> valueAction)
		{
			_cellSet.SetFormulaValueActionGeneric = valueAction;
			_cellSet.CellType = CellType.Formula;
			return this;
		}

		public ITableCellStage SetCellStyle(string cellStyleKey)
		{
			_cellSet.CellStyleKey = cellStyleKey;
			return this;
		}

		public ITableCellStage SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			_cellSet.CellStyleKey = cellStyleKey;
			_cellSet.SetCellStyleAction = cellStyleAction;
			return this;
		}
		ITableCellStage<T> ITableCellStage<T>.SetCellStyle(string cellStyleKey)
		{
			_cellSet.CellStyleKey = cellStyleKey;
			return this;
		}
		ITableCellStage<T> ITableCellStage<T>.SetCellStyle(string cellStyleKey, Action<TableCellStyleParams, ICellStyle> cellStyleAction)
		{
			_cellSet.CellStyleKey = cellStyleKey;
			_cellSet.SetCellStyleAction = cellStyleAction;
			return this;
		}
	public ITableCellStage SetCellType(CellType cellType)
	{
		_cellSet.CellType = cellType;
		return this;
	}
	ITableCellStage<T> ITableCellStage<T>.SetCellType(CellType cellType)
	{
		_cellSet.CellType = cellType;
		return this;
	}

	public ITableCellStage CopyStyleFromCell(ExcelColumns col, int rowIndex)
	{
		string key = $"{_sheet.SheetName}_{col}{rowIndex}";
		ICell cell = _sheet.GetExcelCell(col, rowIndex);
		if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(key))
		{
			SetCellStyle(key, (styleParams, style) =>
			{
				style.CloneStyleFrom(cell.CellStyle);
			});
		}
		return this;
	}
	ITableCellStage<T> ITableCellStage<T>.CopyStyleFromCell(ExcelColumns col, int rowIndex)
	{
		string key = $"{_sheet.SheetName}_{col}{rowIndex}";
		ICell cell = _sheet.GetExcelCell(col, rowIndex);
		if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(key))
		{
			SetCellStyle(key, (styleParams, style) =>
			{
				style.CloneStyleFrom(cell.CellStyle);
			});
		}
		return this;
	}

	public ITableStage End()
		{
			return new FluentTable<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, _cellTitleSets, _cellBodySets);
		}
		ITableStage<T> ITableCellStage<T>.End()
		{
			return new FluentTable<T>(_workbook, _sheet, _table, _startCol, _startRow, _cellStylesCached, _cellTitleSets, _cellBodySets);
		}


	}

	public class FluentMemoryStream : MemoryStream
	{
		public FluentMemoryStream()
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

	public class TableCellStyleParams
	{
		public IWorkbook Workbook { get; set; }
		public ExcelColumns ColNum { get; set; }
		public int RowNum { get; set; }
	}

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


	public class FluentCell : ICellStage
	{
		private ISheet _sheet;
		private IWorkbook _workbook;
		private ICell _cell;
		public FluentCell(IWorkbook workbook, ISheet sheet, ICell cell)
		{
			_workbook = workbook;
			_sheet = sheet;
			_cell = cell;
		}

		public FluentMemoryStream Save()
		{
			var ms = new FluentMemoryStream();
			ms.AllowClose = false;
			_workbook.Write(ms);
			ms.Flush();
			ms.Seek(0, SeekOrigin.Begin);
			ms.AllowClose = true;
			return ms;
		}

		public ICellStage SetValue<T>(T value)
		{
			var typeEnum = DetermineCellType(value, typeof(T));
			if (_cell == null) return this;

			var stringValue = value?.ToString() ?? string.Empty;

			switch (typeEnum)
			{
				case CellTypeEnum.Int:
					if (int.TryParse(stringValue, out var intVal))
					{
						_cell.SetCellValue(intVal);
					}
					else
					{
						_cell.SetCellValue(stringValue);
					}
					break;
				case CellTypeEnum.Double:
					if (double.TryParse(stringValue, out var doubleVal))
					{
						_cell.SetCellValue(doubleVal);
					}
					else
					{
						_cell.SetCellValue(stringValue);
					}
					break;
				case CellTypeEnum.DateTime:
					if (DateTime.TryParse(stringValue, out var dateVal))
					{
						_cell.SetCellValue(dateVal);
					}
					else
					{
						_cell.SetCellValue(stringValue);
					}
					break;
				case CellTypeEnum.String:
				default:
					_cell.SetCellValue(stringValue);
					break;
			}
			return this;
		}

		private enum CellTypeEnum
		{
			Int,
			Double,
			DateTime,
			String
		}

		private CellTypeEnum DetermineCellType(object cellValue, Type cellType = null)
		{
			// 明確的類型參數優先級最高
			if (cellType != null)
			{
				if (cellType == typeof(int)) return CellTypeEnum.Int;
				if (cellType == typeof(double) || cellType == typeof(float)) return CellTypeEnum.Double;
				if (cellType == typeof(DateTime)) return CellTypeEnum.DateTime;
				if (cellType == typeof(string)) return CellTypeEnum.String;
			}

			// 如果無值，回傳字串型別
			if (cellValue == null || cellValue == DBNull.Value) return CellTypeEnum.String;

			var stringValue = cellValue.ToString();

			// 嘗試按優先級判斷型別
			if (int.TryParse(stringValue, out _)) return CellTypeEnum.Int;
			if (double.TryParse(stringValue, out _)) return CellTypeEnum.Double;
			if (DateTime.TryParse(stringValue, out _)) return CellTypeEnum.DateTime;

			return CellTypeEnum.String;
		}
	}
}
