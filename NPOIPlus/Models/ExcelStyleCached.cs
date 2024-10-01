using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public class ExcelStyleCached
	{
		public string GroupName { get; set; }
		public Dictionary<string, ICellStyle> CellStyles = new Dictionary<string, ICellStyle>();

	}
}
