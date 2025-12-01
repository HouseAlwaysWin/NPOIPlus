using NPOI.SS.UserModel;
using NPOIPlus.Helpers;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NPOIPlus.Base
{
	public abstract class FluentSheetBase : FluentCellBase
	{
		public FluentSheetBase(
			IWorkbook workbook,
			Dictionary<string, ICellStyle> cellStylesCached)
			: base(workbook, cellStylesCached)
		{
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

	protected object GetTableCellValue(string cellName, object item)
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

		public IWorkbook SaveToPath(string filePath)
		{
			using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
			{
				_workbook.Write(outFile);
			}
			return _workbook;
		}
	}
}

