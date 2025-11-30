using System.Data;
using System;
using System.IO;
using NPOIPlus;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using NPOIPlus.Models;
using System.Linq;

namespace NPOIPlusConsoleExample
{
	internal class Program
	{
		static void Main(string[] args)
		{
			try
			{
				var testData = new List<ExampleData>(){
					new ExampleData(1, "Alice Chen", new DateTime(1994, 1, 1)),
					new ExampleData(2, "Bob Lee", new DateTime(1989, 5, 12)),
					new ExampleData(3, "Søren", new DateTime(1985, 7, 23)),
					new ExampleData(4, "王小明", new DateTime(2000, 2, 29)),
					new ExampleData(5, "This is a very very very long name to test wrapping and width", new DateTime(1980, 12, 31)),
					new ExampleData(6, "Élodie", new DateTime(1995, 5, 15)),
					new ExampleData(7, "O'Connor", new DateTime(1975, 7, 7)),
					new ExampleData(8, "李雷", new DateTime(2010, 10, 10)),
					new ExampleData(9, "山田太郎", new DateTime(1999, 3, 3)),
					new ExampleData(10, "Мария", new DateTime(1988, 8, 8)),
					new ExampleData(11, "محمد", new DateTime(1991, 9, 9)),
					new ExampleData(12, "김민준", new DateTime(2004, 4, 4)),
				};

				var filePath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test.xlsx";
				var outputPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test2.xlsx";

				var fluent = new FluentWorkbook(new XSSFWorkbook(filePath));

				fluent
				// .UseSheet("Sheet1")
				// .SetColumnWidth(ExcelColumns.C, 20)
				.SetupGlobalCachedCellStyles((workbook, style) =>
				{
					style.SetAligment(HorizontalAlignment.Center);
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.SetFontInfo(workbook, "Calibri", 10);
				})
				// 日期格式樣式
				.SetupCellStyle("DateOfBirth", (workbook, style) =>
				{
					style.SetDataFormat(workbook, "yyyy-MM-dd");
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightGreen);
				})
				// 藍底白字標題風格（示意背景色）
				.SetupCellStyle("HeaderBlue", (workbook, style) =>
				{
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
				})
				.SetupCellStyle("BodyGreen", (workbook, style) =>
				{
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightGreen);
				})
				// 金額格式
				.SetupCellStyle("AmountCurrency", (workbook, style) =>
				{
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.SetDataFormat(workbook, "#,##0.00");
					style.SetAligment(HorizontalAlignment.Right);
				})
				// 黃底高亮
				.SetupCellStyle("HighlightYellow", (workbook, style) =>
				{
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.Yellow);
				});


				// Sheet1：只放一個表（A 欄開始），並有抬頭（標題列），涵蓋多種欄位型別
				fluent.UseSheet("Sheet1")
				.SetColumnWidth(ExcelColumns.A, ExcelColumns.H, 20)
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginTitleSet("ID").SetCellStyle("HeaderBlue")
				.BeginBodySet("ID").SetCellStyle("BodyGreen").End()

				.BeginTitleSet("名稱").SetCellStyle("HeaderBlue")
				.BeginBodySet("Name").SetCellStyle("BodyGreen").End()

				.BeginTitleSet("生日").SetCellStyle("HeaderBlue")
				.BeginBodySet("DateOfBirth").SetCellStyle("DateOfBirth").End()

				.BeginTitleSet("是否活躍").SetCellStyle("HeaderBlue")
				.BeginBodySet("IsActive").SetCellType(CellType.Boolean).End()

				.BeginTitleSet("分數").SetCellStyle("HeaderBlue")
				.BeginBodySet("Score").SetCellType(CellType.Numeric).End()

				.BeginTitleSet("金額").SetCellStyle("HeaderBlue")
				.BeginBodySet("Amount").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()

				.BeginTitleSet("備註").SetCellStyle("HeaderBlue")
				.BeginBodySet("Notes").SetCellType(CellType.String).End()

				.BeginTitleSet("可能為空").SetCellStyle("HeaderBlue")
				.BeginBodySet("MaybeNull").End()
				.BuildRows();

				// 第二個分頁（Summary）：不同資料與樣式示範
				var sheet2Data = new List<Dictionary<string, object>>
				{
					new Dictionary<string, object> {
						{ "Title", "Total Users" }, { "Value", testData.Count }, { "AsOfDate", DateTime.Today }, { "IsOk", true }, { "FormulaVal", "SUM(B2:B4)" }
					},
					new Dictionary<string, object> {
						{ "Title", "Active Rate" }, { "Value", 0.8765 }, { "AsOfDate", DateTime.Today.AddDays(-1) }, { "IsOk", true }, { "FormulaVal", "B3*100" }
					},
					new Dictionary<string, object> {
						{ "Title", "Remarks" }, { "Value", 12345.6789m }, { "AsOfDate", DateTime.Today.AddMonths(-1) }, { "IsOk", false }, { "FormulaVal", "AVERAGE(B2:B4)" }
					},
				};

				var sheet2 = fluent
				.UseSheet("Summary", true);

				sheet2
				.SetTable(sheet2Data, ExcelColumns.A, 1)
				.BeginCellSet("Title").SetCellStyle("HeaderBlue").End()
				.BeginCellSet("Value").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()
				.BeginCellSet("AsOfDate").SetCellStyle("DateStyle").End()
				.BeginCellSet("IsOk").SetCellType(CellType.Boolean).End()
				.BeginCellSet("FormulaVal").SetCellType(CellType.Formula).End()
				.BuildRows();

				// Sheet3：展示 CopyStyleFromCell 功能
				// 從 Sheet1 的樣式複製到新的資料表
				var sheet3Data = new List<Dictionary<string, object>>
			{
				new Dictionary<string, object> {
					{ "Product", "筆記本" }, { "Quantity", 100 }, { "Price", 25.50m }, { "InStock", true }
				},
				new Dictionary<string, object> {
					{ "Product", "原子筆" }, { "Quantity", 500 }, { "Price", 5.00m }, { "InStock", true }
				},
				new Dictionary<string, object> {
					{ "Product", "計算機" }, { "Quantity", 50 }, { "Price", 120.00m }, { "InStock", false }
				},
			};

				fluent.UseSheet("CopyStyleExample", true)
				.SetColumnWidth(ExcelColumns.A, ExcelColumns.D, 20)
				.SetTable(sheet3Data, ExcelColumns.A, 1)
				// 從 Sheet1 的 A1 (ID標題) 複製樣式到 Product 標題
				.BeginTitleSet("Product").SetCellStyle("HeaderBlue")
				.BeginBodySet("Product")
				.SetCellStyle("ProductStyle", (styleParams, style) =>
				{
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightGreen);
					style.SetBorderAllStyle(BorderStyle.Thin);
				})
				.End()
				.BuildRows()

				// 從 Sheet1 的 E1 (分數標題) 複製樣式到 Quantity 標題
				.BeginTitleSet("Quantity")
				.BeginBodySet("Quantity").SetCellType(CellType.Numeric).End()

				// 從 Sheet1 的 F1 (金額標題) 複製樣式到 Price 標題
				.BeginTitleSet("Price").CopyStyleFromCell(ExcelColumns.A, 1)
				.BeginBodySet("Price").SetCellType(CellType.Numeric)
				.CopyStyleFromCell(ExcelColumns.A, 2)
				.End()

				// 從 Sheet1 的 D1 (是否活躍標題) 複製樣式到 InStock 標題
				.BeginTitleSet("InStock")
				.SetCellStyle("ProductStyle")
				.BeginBodySet("InStock").SetCellType(CellType.Boolean).End()
				.BuildRows()
				.SaveToPath(outputPath);

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}
	}

}
