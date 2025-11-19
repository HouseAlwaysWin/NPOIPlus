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
		IWorkbookStage ReadExcelFile(string filePath);
		ISheetStage UseSheet(string sheetName, bool createIfMissing = true);
		ISheetStage UseSheet(ISheet sheet);
		ISheetStage UseSheetAt(int index, bool createIfMissing = false);
	}

	public interface ISheetStage
	{
		ISheetStage SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles);
		ITableStage SetTable<T>(IEnumerable<T> table, ExcelColumns startCol, int startRow);
		ICellStage SetCell(ExcelColumns startCol, int startRow);
	}


	public interface ITableStage
	{
		// ITableStage SetCell(string cellName, object value);
		ITableStage MapCellByName(string cellName, Func<object, object> value = null);
		ITableStage SetRow();
		FluentMemoryStream Save();
		IWorkbook Save(string filePath);
	}

	public interface ITableCellStage<T>
	{
		ITableCellStage<T> SetValue(T value);
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
		private FileStream _fileStream;
		private Dictionary<string, ICellStyle> _cellStylesCached = new Dictionary<string, ICellStyle>();
		public FluentWorkbook(IWorkbook workbook)
		{
			_workbook = workbook;
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


		public ISheetStage UseSheet(string sheetName, bool createIfMissing = true)
		{
			_currentSheet = _workbook.GetSheet(sheetName);
			if (_currentSheet == null && createIfMissing)
			{
				_currentSheet = _workbook.CreateSheet(sheetName);
			}
			return new FluentSheet(_workbook, _currentSheet);
		}

		public ISheetStage UseSheet(ISheet sheet)
		{
			_currentSheet = sheet;
			return new FluentSheet(_workbook, _currentSheet);
		}

		public ISheetStage UseSheetAt(int index, bool createIfMissing = false)
		{
			_currentSheet = _workbook.GetSheetAt(index);
			if (_currentSheet == null && createIfMissing)
			{
				_currentSheet = _workbook.CreateSheet();
			}
			return new FluentSheet(_workbook, _currentSheet);
		}
	}

	public class FluentSheet : ISheetStage
	{
		private ISheet _sheet;
		private IWorkbook _workbook;
		private Dictionary<string, ICellStyle> _cellStylesCached = new Dictionary<string, ICellStyle>();
		private Action<IWorkbook, ICellStyle> _globalCellStylesAction;
		public FluentSheet(IWorkbook workbook, ISheet sheet)
		{
			_workbook = workbook;
			_sheet = sheet;
		}

		public ICellStage SetCell(ExcelColumns col, int row)
		{
			if (_sheet == null) throw new InvalidOperationException("No active sheet. Call UseSheet(...) first.");

			var rowObj = _sheet.GetRow(row) ?? _sheet.CreateRow(row);
			var cell = rowObj.GetCell((int)col) ?? rowObj.CreateCell((int)col);
			return new FluentCell(_workbook, _sheet, cell);
		}

		public ITableStage SetTable<T>(IEnumerable<T> table, ExcelColumns startCol, int startRow)
		{
			return new FluentTable<T>(_workbook, _sheet, table, startCol, startRow, _cellStylesCached);
		}

		public ISheetStage SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
		{
			ICellStyle newCellStyle = _workbook.CreateCellStyle();
			styles(_workbook, newCellStyle);
			_cellStylesCached.Add("global", newCellStyle);
			return this;
		}
	}


	public class FluentTable<T> : ITableStage
	{
		private ISheet _sheet;
		private IWorkbook _workbook;
		private IEnumerable<T> _table;
		private IList<T> _itemsCache;
		private ExcelColumns _startCol;
		private int _startRow;
		private List<TableCellNameMap> _cellNameMaps = new List<TableCellNameMap>();
		private Dictionary<string, ICellStyle> _cellStylesCached;
		public FluentTable(IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
		ExcelColumns startCol, int startRow, Dictionary<string, ICellStyle> cellStylesCached)
		{
			_workbook = workbook;
			_sheet = sheet;
			_table = table;
			_startCol = NormalizeStartCol(startCol);
			_startRow = NormalizeStartRow(startRow);
			_cellStylesCached = cellStylesCached;
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

		private void SetCellStyle(ICell cell, TableCellNameMap cellNameMap)
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
				cellNameMap.SetCellStyleAction(newCellStyle);
				cellNameMap.CellStyleKey = cellNameMap.CellStyleKey;
				_cellStylesCached.Add(cellNameMap.CellStyleKey, newCellStyle);
				cell.CellStyle = newCellStyle;
			}
		}

		private ITableStage SetRow(int rowOffset = 0)
		{
			if (_cellNameMaps == null || _cellNameMaps.Count == 0) return this;

			var targetRowIndex = _startRow + rowOffset;
			var rowObj = _sheet.GetRow(targetRowIndex) ?? _sheet.CreateRow(targetRowIndex);

			var item = GetItemAt(rowOffset);

			int colIndex = (int)_startCol;
			foreach (var cellNameMap in _cellNameMaps)
			{
				var cell = rowObj.GetCell(colIndex) ?? rowObj.CreateCell(colIndex);

				// 優先使用 TableCellNameMap 中的 Value，如果沒有則從 item 中獲取
				Func<object, object> setValueAction = cellNameMap.SetValueAction;
				object value = GetTableCellValue(cellNameMap.CellName, item);
				if (setValueAction != null)
				{
					value = setValueAction(value);
				}

				SetCellStyle(cell, cellNameMap);
				// write value
				SetCellValue(cell, value);
				colIndex++;
			}

			return this;
		}

		public ITableStage MapCellByName(string cellName, Func<object, object> value = null)
		{
			_cellNameMaps.Add(new TableCellNameMap { CellName = cellName, SetValueAction = value });
			return this;
		}

		public ITableStage SetRow()
		{
			for (int i = 0; i < _table.Count(); i++)
			{
				SetRow(i);
			}
			return this;
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

		public IWorkbook Save(string filePath)
		{
			using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
			{
				_workbook.Write(outFile);
			}
			return _workbook;
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

	public class TableCellNameMap
	{
		public string CellName { get; set; }
		public Func<object, object> SetValueAction { get; set; }
		public string CellStyleKey { get; set; }
		public Action<ICellStyle> SetCellStyleAction { get; set; }
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
