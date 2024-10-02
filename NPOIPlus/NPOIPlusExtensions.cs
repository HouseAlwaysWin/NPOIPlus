using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text;

namespace NPOIPlus
{
	public static class NPOIPlusExtensions
	{

		public static IRow GetExcelRowOrCreate(this ISheet sheet, int row = 1)
		{
			if (row < 1) row = 1;
			IRow newRow = sheet.GetRow(row - 1) ?? sheet.CreateRow(row - 1);
			return newRow;
		}

		/// <summary>
		/// Use rgb to set foreground color
		/// </summary>
		/// <param name="style"></param>
		/// <param name="red"></param>
		/// <param name="green"></param>
		/// <param name="blue"></param>
		public static void SetCellFillForegroundColor(this ICellStyle style, byte red, byte green, byte blue)
		{
			XSSFColor redColor = new XSSFColor(new byte[] { red, green, blue }); // 設定為紅色
			style.FillPattern = FillPattern.SolidForeground;
			((XSSFCellStyle)style).SetFillForegroundColor(redColor);
		}

		/// <summary>
		/// Use hex color to set foreground color
		/// </summary>
		/// <param name="style"></param>
		/// <param name="hexColor"></param>
		public static void SetCellFillForegroundColor(this ICellStyle style, string hexColor)
		{
			byte[] rgbColor = HexToRgb(hexColor);  // 轉換為 RGB 數組
			XSSFColor xssfColor = new XSSFColor(rgbColor);

			// 設定背景色
			style.FillPattern = FillPattern.SolidForeground;
			((XSSFCellStyle)style).SetFillForegroundColor(xssfColor);
		}

		/// <summary>
		/// Use hex color to set foreground color
		/// </summary>
		/// <param name="style"></param>
		/// <param name="hexColor"></param>
		public static void SetCellFillForegroundColor(this ICellStyle style, IndexedColors colors)
		{
			style.FillForegroundColor = colors.Index;
		}


		// 將十六進制色碼轉換為 RGB
		private static byte[] HexToRgb(string hexColor)
		{
			hexColor = hexColor.Replace("#", string.Empty);

			// 轉換為 RGB 三個 byte
			byte r = Convert.ToByte(hexColor.Substring(0, 2), 16);
			byte g = Convert.ToByte(hexColor.Substring(2, 2), 16);
			byte b = Convert.ToByte(hexColor.Substring(4, 2), 16);

			return new byte[] { r, g, b };
		}

		public static void SetDateFormat(this ICellStyle style, IWorkbook workbook, string format)
		{
			IDataFormat dataFormat = workbook.CreateDataFormat();
			style.DataFormat = dataFormat.GetFormat(format);
		}

		public static void SetFontInfo(this ICellStyle style, IWorkbook workbook, string fontFamily, double fontHeight)
		{
			IFont font = workbook.CreateFont();
			font.FontName = fontFamily;
			font.FontHeightInPoints = fontHeight;
			style.SetFont(font);
		}

		public static void SetBorderAllStyle(this ICellStyle style, BorderStyle all = BorderStyle.None)
		{
			style.BorderBottom = all;
			style.BorderLeft = all;
			style.BorderRight = all;
			style.BorderTop = all;
		}

		public static void SetBorderStyle(this ICellStyle style,
			BorderStyle top = BorderStyle.None, BorderStyle right = BorderStyle.None,
			BorderStyle bottom = BorderStyle.None, BorderStyle left = BorderStyle.None)
		{
			style.BorderBottom = bottom;
			style.BorderLeft = left;
			style.BorderRight = right;
			style.BorderTop = top;
		}

		public static void SetAligment(this ICellStyle style, HorizontalAlignment horizontal = HorizontalAlignment.General)
		{
			style.Alignment = horizontal;
		}
		public static void SetAligment(this ICellStyle style, VerticalAlignment vertical = VerticalAlignment.None)
		{
			style.VerticalAlignment = vertical;
		}
		public static void SetAligment(this ICellStyle style, VerticalAlignment vertical = VerticalAlignment.None, HorizontalAlignment horizontal = HorizontalAlignment.General)
		{
			style.Alignment = horizontal;
			style.VerticalAlignment = vertical;
		}

		public static void SetColumnWidthRange(this ISheet sheet, ExcelColumns start, ExcelColumns end, int size)
		{
			for (int i = (int)start; i <= (int)end; i++)
			{
				sheet.SetColumnWidth(i, size * 256);
			}
		}

		public static void SetColumnWidth(this ISheet sheet, ExcelColumns col, int size)
		{
			sheet.SetColumnWidth((int)col, size * 256);
		}

		public static void CreateFreezePane(this ISheet sheet, ExcelColumns col, int row)
		{
			sheet.CreateFreezePane((int)col + 1, row);
		}
	}
}
