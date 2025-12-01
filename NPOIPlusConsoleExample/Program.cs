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
				new ExampleData(1, "Alice Chen", new DateTime(1994, 1, 1), true, 95.5, 12500.75m, "優秀學生"),
				new ExampleData(2, "Bob Lee", new DateTime(1989, 5, 12), false, 78.0, 8900.50m, "需改進"),
				new ExampleData(3, "Søren", new DateTime(1985, 7, 23), true, 88.5, 15000.00m, "表現良好"),
				new ExampleData(4, "王小明", new DateTime(2000, 2, 29), true, 92.0, 11200.80m, "進步快速"),
				new ExampleData(5, "This is a very very very long name to test wrapping and width", new DateTime(1980, 12, 31), false, 65.5, 7500.25m, "LongName"),
				new ExampleData(6, "Élodie", new DateTime(1995, 5, 15), true, 85.0, 9800.00m, "穩定發揮"),
				new ExampleData(7, "O'Connor", new DateTime(1975, 7, 7), false, 72.5, 8200.50m, "待觀察"),
				new ExampleData(8, "李雷", new DateTime(2010, 10, 10), true, 90.0, 10500.75m, "潛力股"),
				new ExampleData(9, "山田太郎", new DateTime(1999, 3, 3), true, 87.5, 9500.00m, "穩健型"),
				new ExampleData(10, "Мария", new DateTime(1988, 8, 8), false, 70.0, 8000.25m, "需加強"),
				new ExampleData(11, "محمد", new DateTime(1991, 9, 9), true, 93.5, 12000.00m, "頂尖"),
				new ExampleData(12, "김민준", new DateTime(2004, 4, 4), true, 89.0, 10200.50m, "均衡發展"),
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

				// 第二個分頁（Summary）：使用同一份 testData 展示不同欄位組合
				fluent.UseSheet("Summary", true)
				.SetColumnWidth(ExcelColumns.A, ExcelColumns.E, 20)
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginTitleSet("姓名").SetCellStyle("HeaderBlue")
				.BeginBodySet("Name").End()

				.BeginTitleSet("分數").SetCellStyle("HeaderBlue")
				.BeginBodySet("Score").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()

				.BeginTitleSet("日期").SetCellStyle("HeaderBlue")
				.BeginBodySet("DateOfBirth").SetCellStyle("DateOfBirth").End()

				.BeginTitleSet("狀態").SetCellStyle("HeaderBlue")
				.BeginBodySet("IsActive").SetCellType(CellType.Boolean).End()

				.BeginTitleSet("備註").SetCellStyle("HeaderBlue")
				.BeginBodySet("Notes").SetCellStyle("HighlightYellow").End()
				.BuildRows();

				// Sheet3：展示 CopyStyleFromCell 功能，使用同一份 testData
				fluent.UseSheet("CopyStyleExample", true)
				.SetColumnWidth(ExcelColumns.A, ExcelColumns.D, 20)
				.SetTable(testData, ExcelColumns.A, 1)
			// 展示 ID 欄位並套用自訂樣式
			.BeginTitleSet("編號").SetCellStyle("HeaderBlue")
			.BeginBodySet("ID")
			.SetCellStyle((styleParams) =>
			{
				return new CellStyleConfig("IDStyle", style =>
				{
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightGreen);
					style.SetBorderAllStyle(BorderStyle.Thin);
				});
			})
			.End()

				// 從 Sheet1 的樣式複製
				.BeginTitleSet("姓名").CopyStyleFromCell(ExcelColumns.A, 1)
				.BeginBodySet("Name").SetCellType(CellType.String).End()

				// 展示金額欄位並套用金額格式
				.BeginTitleSet("金額").SetCellStyle("HeaderBlue")
				.BeginBodySet("Amount").SetCellType(CellType.Numeric)
				.SetCellStyle("AmountCurrency")
				.End()

			// 展示活躍狀態欄位（根據數據動態變化樣式）
			.BeginTitleSet("活躍").SetCellStyle("HeaderBlue")
			.BeginBodySet("IsActive").SetCellType(CellType.Boolean)
			.SetCellStyle((styleParams) =>
			{
				// ✅ 根據資料決定返回哪個樣式
				if (styleParams.GetRowItem<ExampleData>().IsActive)
				{
					return new("IsActiveStyle1", style =>
					{
						style.SetAligment(HorizontalAlignment.Center);
						style.FillPattern = FillPattern.SolidForeground;
						style.SetCellFillForegroundColor(IndexedColors.LightGreen);
						style.SetBorderAllStyle(BorderStyle.Thin);
					});
				}
				return new("IsActiveStyle2", style =>
				{
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightYellow);
					style.SetBorderAllStyle(BorderStyle.Thin);
				});
			})
			.End()

				.BuildRows();

				fluent.UseSheet("SetCellValueExample", true)
				.SetCellPosition(ExcelColumns.A, 1)
				.SetValue("Hello, World!");

				fluent.SaveToPath(outputPath);

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}
	}

}
