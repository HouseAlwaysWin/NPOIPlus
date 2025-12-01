using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace NPOIPlus
{
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
}

