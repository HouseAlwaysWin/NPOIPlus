using NPOI.SS.UserModel;
using FluentNPOI.Helpers;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FluentNPOI.Base
{
	public abstract class FluentSheetBase : FluentCellBase
	{
		public FluentSheetBase(
			IWorkbook workbook,
			ISheet sheet,
			Dictionary<string, ICellStyle> cellStylesCached)
			: base(workbook, sheet, cellStylesCached)
		{
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
}
}

