using NPOI.SS.UserModel;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;

namespace NPOIPlus.Base
{
	public abstract class FluentCellBase : FluentWorkbookBase
	{
		protected Dictionary<string, ICellStyle> _cellStylesCached;

		protected FluentCellBase()
		{
		}

		protected FluentCellBase(IWorkbook workbook, Dictionary<string, ICellStyle> cellStylesCached)
			: base(workbook)
		{
			_cellStylesCached = cellStylesCached;
		}

		protected ExcelColumns NormalizeStartCol(ExcelColumns col)
		{
			int idx = (int)col;
			if (idx < 0) idx = 0;
			return (ExcelColumns)idx;
		}

		protected int NormalizeStartRow(int row)
		{
			// 將使用者常見的 1-based 列號轉為 0-based，並確保不為負數
			if (row < 1) return 0;
			return row - 1;
		}

	protected void SetCellStyle(ICell cell, TableCellSet cellNameMap, TableCellStyleParams cellStyleParams)
	{
		// 如果有動態樣式設置函數，優先使用
		if (cellNameMap.SetCellStyleAction != null)
		{
			// ✅ 先調用函數獲取樣式配置
			var config = cellNameMap.SetCellStyleAction(cellStyleParams);
			
			if (!string.IsNullOrWhiteSpace(config.Key))
			{
				// ✅ 先檢查緩存
				if (!_cellStylesCached.ContainsKey(config.Key))
				{
					// ✅ 只在不存在時才創建新樣式
					ICellStyle newCellStyle = _workbook.CreateCellStyle();
					config.StyleSetter(newCellStyle);
					_cellStylesCached.Add(config.Key, newCellStyle);
				}
				// 始終使用緩存的樣式
				cell.CellStyle = _cellStylesCached[config.Key];
			}
			else
			{
				// 如果沒有返回 key，創建臨時樣式（不緩存）
				ICellStyle tempStyle = _workbook.CreateCellStyle();
				config.StyleSetter(tempStyle);
				cell.CellStyle = tempStyle;
			}
		}
		// 如果有固定的樣式 key，使用緩存的樣式
		else if (!string.IsNullOrWhiteSpace(cellNameMap.CellStyleKey) && _cellStylesCached.ContainsKey(cellNameMap.CellStyleKey))
		{
			cell.CellStyle = _cellStylesCached[cellNameMap.CellStyleKey];
		}
		// 如果都沒有，使用全局樣式
		else if (_cellStylesCached.ContainsKey("global"))
		{
			cell.CellStyle = _cellStylesCached["global"];
		}
	}

		protected void SetCellValue(ICell cell, object value)
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

		protected void SetCellValue(ICell cell, object value, CellType cellType)
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

		protected void SetFormulaValue(ICell cell, object value)
		{
			if (cell == null) return;
			if (value == null || value == DBNull.Value) return;

			var formula = value.ToString();
			if (string.IsNullOrWhiteSpace(formula)) return;

			// NPOI SetCellFormula 需要純公式字串（不含 '='）
			if (formula.StartsWith("=")) formula = formula.Substring(1);

			cell.SetCellFormula(formula);
		}
	}
}

