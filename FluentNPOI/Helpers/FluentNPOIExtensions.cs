using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text;

namespace FluentNPOI
{
	public static class FluentNPOIExtensions
	{
		public static ICell GetExcelCell(this ISheet sheet, ExcelCol colIndex, int rowIndex)
		{
			IRow row = sheet.GetExcellRow(rowIndex);
			ICell cell = row.GetCell((int)colIndex);
			return cell;
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
			byte[] rgbColor = HexToRgb(hexColor); // 轉換為 RGB 數組
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

		public static void SetDataFormat(this ICellStyle style, IWorkbook workbook, string format)
		{
			IDataFormat dataFormat = workbook.CreateDataFormat();
			style.DataFormat = dataFormat.GetFormat(format);
		}

		public static void SetFontInfo(this ICellStyle style, IWorkbook workbook, string fontFamily = null, double? fontHeight = null, bool isBold = false, bool isItalic = false, bool isStrikeout = false, IndexedColors color = null)
		{
			IFont font = workbook.CreateFont();
			if (fontFamily != null) font.FontName = fontFamily;
			if (fontHeight != null) font.FontHeightInPoints = fontHeight.Value;
			if (color != null) font.Color = color.Index;
			font.IsBold = isBold;
			font.IsItalic = isItalic;
			font.IsStrikeout = isStrikeout;
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

		public static void SetAligment(this ICellStyle style,
			HorizontalAlignment horizontal = HorizontalAlignment.General,
			VerticalAlignment vertical = VerticalAlignment.Center)
		{
			style.Alignment = horizontal;
			style.VerticalAlignment = vertical;
		}

		public static void SetAligment(this ICellStyle style, VerticalAlignment vertical = VerticalAlignment.Center,
			HorizontalAlignment horizontal = HorizontalAlignment.General)
		{
			style.Alignment = horizontal;
			style.VerticalAlignment = vertical;
		}

		public static void SetExcelCellMerge(this ISheet sheet, ExcelCol firstCol, ExcelCol lastCol,
			int firstRow = 1, int lastRow = 1)
		{
			if (firstRow < 1) firstRow = 1;
			if (lastRow < 1) lastRow = 1;

			CellRangeAddress region = new CellRangeAddress(firstRow - 1, lastRow - 1, (int)firstCol, (int)lastCol);
			sheet.AddMergedRegion(region);
		}

		public static void SetExcelCellMerge(this ISheet sheet, ExcelCol firstCol, ExcelCol lastCol, int row)
		{
			if (row < 1) row = 1;

			CellRangeAddress region = new CellRangeAddress(row - 1, row - 1, (int)firstCol, (int)lastCol);
			sheet.AddMergedRegion(region);
		}

		public static void SetExcelCellMerge(this ISheet sheet, ExcelCol col, int startRow, int endRow)
		{
			if (startRow < 1) startRow = 1;
			if (endRow < 1) endRow = 1;

			CellRangeAddress region = new CellRangeAddress(startRow - 1, endRow - 1, (int)col, (int)col);
			sheet.AddMergedRegion(region);
		}

		public static void SetColumnWidth(this ISheet sheet, ExcelCol startCol, ExcelCol endCol, double width)
		{
			for (int i = (int)startCol; i <= (int)endCol; i++)
			{
				sheet.SetColumnWidth(i, width * 256);
			}
		}

		public static void SetColumnWidth(this ISheet sheet, ExcelCol col, double width)
		{
			sheet.SetColumnWidth((int)col, width * 256);
		}

		/// <summary>
		/// 取得特定欄位的值
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="colNum"></param>
		/// <param name="rowNum"></param>
		/// <returns></returns>
		public static ICell GetCellValue(this ISheet sheet, ExcelCol colNum, int rowNum = 1)
		{
			if (rowNum < 1) rowNum = 1;
			rowNum = rowNum - 1;

			// 逐行讀取資料
			for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
			{
				IRow row = sheet.GetRow(rowIndex);
				if (row == null) continue;

				for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
				{
					ICell cell = row.GetCell(colIndex);
					if (rowIndex == rowNum && (int)colNum == colIndex)
					{
						return cell;
					}
				}
			}

			return null;
		}

		public static IRow GetExcellRow(this ISheet sheet, int rowNum = 1)
		{
			if (rowNum < 1) rowNum = 1;
			rowNum = rowNum - 1;
			IRow row = sheet.GetRow(rowNum) ?? sheet.CreateRow(rowNum);
			return row;
		}
	}
}


