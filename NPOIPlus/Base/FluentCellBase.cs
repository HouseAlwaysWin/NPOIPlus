using NPOI.SS.UserModel;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;

namespace NPOIPlus.Base
{
	public abstract class FluentCellBase
	{
		protected Dictionary<string, ICellStyle> _cellStylesCached;
		protected IWorkbook _workbook;

		protected FluentCellBase()
		{
		}

		protected FluentCellBase(IWorkbook workbook, Dictionary<string, ICellStyle> cellStylesCached)
		{
			_workbook = workbook;
			_cellStylesCached = cellStylesCached;
		}

		protected void SetCellStyle(ICell cell, TableCellSet cellNameMap, TableCellStyleParams cellStyleParams)
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

