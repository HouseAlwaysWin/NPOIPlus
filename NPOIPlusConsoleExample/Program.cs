

using System.Data;
using System;
using System.IO;
using NPOIPlus;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using NPOIPlus.Models;

namespace NPOIPlusConsoleExample
{
	internal class Program
	{
		static void Main(string[] args)
		{
			try
			{
				var filePath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test.xlsx";
				// 打開 Excel 文件
				using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
				{
					NPOIWorkbook workbook = new NPOIWorkbook(new XSSFWorkbook(file));

					workbook.SetGlobalCellStyle = (style) =>
					{
						style.SetAligment(VerticalAlignment.Center);
						style.SetBorderAllStyle(BorderStyle.None);
						style.SetFontInfo(workbook.Workbook, "Calibri", 10);
					};

					workbook.SetDefaultDateTimeCellStyle = (style) =>
					{
						style.SetDateFormat(workbook.Workbook, "yyyy-MM-dd");
					};

					ISheet sheet1 = workbook.Workbook.GetSheet("Sheet1");


					// 1. 創建 DataTable
					DataTable dataTable = new DataTable("ExampleTable");

					// 2. 添加列 (列名稱與類型)
					dataTable.Columns.Add("ID", typeof(int));         // 整數類型的 ID 列
					dataTable.Columns.Add("Name", typeof(string));    // 字串類型的 Name 列
					dataTable.Columns.Add("DateOfBirth", typeof(DateTime)); // 日期類型的 DateOfBirth 列

					// 3. 添加數據行
					dataTable.Rows.Add(1, "Alice", new DateTime(1990, 1, 1));
					dataTable.Rows.Add(2, "Bob", new DateTime(1985, 5, 23));
					dataTable.Rows.Add(3, "Charlie", new DateTime(2000, 10, 15));

					workbook.SetExcelCell(sheet1, dataTable, 1, "ID", ExcelColumns.D, 12, (cell, value) =>
					{
						return $"{value}st";
					});

					workbook.SetColExcelCells(sheet1, dataTable, 1, new List<ExcelCellParam>
					{
						new("ID",null,(style)=>{
							//style.Alignment = HorizontalAlignment.Left;
							style.SetAligment(HorizontalAlignment.Left);
						}),
						new("Name"),
						new("DateOfBirth"),
					}, ExcelColumns.A, 1, (style) =>
					{
						style.Alignment = HorizontalAlignment.Center;
					});

					workbook.SetRowExcelCells(sheet1, dataTable, new List<ExcelCellParam>
					{
						new("ID" ,
						null,(style)=>{
							style.Alignment = HorizontalAlignment.Center;
							//style.BorderBottom = BorderStyle.Thick;
							style.SetBorderStyle(bottom:BorderStyle.Thick);
						}
						),
						new("Name",
						null,(style)=>{
							style.Alignment = HorizontalAlignment.Left;
							// 設定單元格背景色（RGB 顏色）
							style.FillPattern = FillPattern.SolidForeground;
							//style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
							style.SetCellFillForegroundColor(IndexedColors.Grey25Percent);
							//style.SetCellFillForegroundColor("#FF5733");
						}
						),
						new("DateOfBirth",
						null,(style)=>{
							style.Alignment = HorizontalAlignment.Right;
						}
						),
						new((row, col, value) =>
						{
							return $"{col-3}{row}:{col-3}{row+1}";
						})
					}, ExcelColumns.E, 1, (style) =>
					{
						style.BorderBottom = BorderStyle.Double;
						style.FillPattern = FillPattern.SolidForeground;
						style.SetCellFillForegroundColor(100, 100, 0);
					});





					// 将修改后的内容写回文件
					using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
					{
						workbook.Workbook.Write(outFile);
					}

				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}
	}

}