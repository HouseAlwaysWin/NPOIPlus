using NPOI.SS.UserModel;
using NPOIPlus.Helpers;
using System;

namespace NPOIPlus
{
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
			ms.Seek(0, System.IO.SeekOrigin.Begin);
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

