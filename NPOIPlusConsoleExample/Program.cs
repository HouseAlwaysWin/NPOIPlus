

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
					new ExampleData(1, "John", new DateTime(1994, 1, 1)),
					new ExampleData(2, "Jane", new DateTime(1993, 1, 1)),
					new ExampleData(3, "Jim", new DateTime(1992, 1, 1)),
					new ExampleData(4, "王小明", new DateTime(2000, 2, 29)),
					new ExampleData(5, "A very very long name to test wrapping", new DateTime(1980, 12, 31)),
					new ExampleData(6, "Élodie", new DateTime(1995, 5, 15)),
					new ExampleData(7, "O'Connor", new DateTime(1975, 7, 7)),
					new ExampleData(8, "李雷", new DateTime(2010, 10, 10)),
					new ExampleData(9, "山田太郎", new DateTime(1999, 3, 3)),
					new ExampleData(10, "Мария", new DateTime(1988, 8, 8)),
					new ExampleData(11, "محمد", new DateTime(1991, 9, 9)),
					new ExampleData(12, "Zoë", new DateTime(2004, 4, 4)),
				};

				// 混合型別資料（測試 bool/decimal/double/null/公式字串等）
				var mixedData = new List<Dictionary<string, object>>()
				{
					new Dictionary<string, object> {
						{ "ID", 1001 }, { "Name", "Alice" }, { "DateOfBirth", new DateTime(1990, 1, 1) },
						{ "IsActive", true }, { "Score", 98.6 }, { "Amount", 123.45m },
						{ "Notes", "Hello" }, { "MaybeNull", DBNull.Value }, { "FormulaVal", "SUM(A2:A4)" }
					},
					new Dictionary<string, object> {
						{ "ID", 1002 }, { "Name", "Bob" }, { "DateOfBirth", new DateTime(1985, 5, 23) },
						{ "IsActive", false }, { "Score", 75.25 }, { "Amount", 0m },
						{ "Notes", "World" }, { "MaybeNull", "not null" }, { "FormulaVal", "A2*2" }
					},
					new Dictionary<string, object> {
						{ "ID", 1003 }, { "Name", "Charlie" }, { "DateOfBirth", new DateTime(2000, 10, 15) },
						{ "IsActive", true }, { "Score", 0.0 }, { "Amount", -456.78m },
						{ "Notes", "测试中文" }, { "MaybeNull", DBNull.Value }, { "FormulaVal", "AVERAGE(B2:B4)" }
					},
					new Dictionary<string, object> {
						{ "ID", 1004 }, { "Name", "Diana" }, { "DateOfBirth", new DateTime(1999, 12, 31) },
						{ "IsActive", true }, { "Score", 100d }, { "Amount", 9999999.99m },
						{ "Notes", "🚀 emoji test" }, { "MaybeNull", DBNull.Value }, { "FormulaVal", "MAX(C2:C4)" }
					},
				};
				var filePath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test.xlsx";
				var outputPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test2.xlsx";

				var fluent = new FluentWorkbook(new XSSFWorkbook(filePath));

				var sheet = fluent
				.UseSheet("Sheet1")
				.SetupGlobalCachedCellStyles((workbook, style) =>
				{
					style.SetAligment(HorizontalAlignment.Center);
					style.SetBorderAllStyle(BorderStyle.None);
					style.SetFontInfo(workbook, "Calibri", 10);
				})
				// 日期格式樣式
				.SetupCellStyle("DateOfBirth", (workbook, style) => { style.SetDataFormat(workbook, "yyyy-MM-dd"); })
				// 藍底白字標題風格（示意背景色）
				.SetupCellStyle("HeaderBlue", (workbook, style) =>
				{
					style.SetBorderAllStyle(BorderStyle.Thin);
					style.SetAligment(HorizontalAlignment.Center);
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
				})
				// 金額格式
				.SetupCellStyle("AmountCurrency", (workbook, style) =>
				{
					style.SetDataFormat(workbook, "#,##0.00");
					style.SetAligment(HorizontalAlignment.Right);
				})
				// 黃底高亮
				.SetupCellStyle("HighlightYellow", (workbook, style) =>
				{
					style.FillPattern = FillPattern.SolidForeground;
					style.SetCellFillForegroundColor(IndexedColors.Yellow);
				});

				// 第一組：ExampleData（A 欄開始）
				sheet
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginMapCell("ID").SetValue((value) => $"ID: {value.CellValue}, Col: {value.ColNum}, Row: {value.RowNum}").SetCellStyle("HeaderBlue").End()
				.BeginMapCell("Name").SetCellStyle("HeaderBlue").End()
				.BeginMapCell("DateOfBirth").SetCellStyle("DateOfBirth").End()
				.BeginMapCell("Test").SetValue("C1:C3").SetCellType(CellType.Formula).End()
				.SetRow();

				// 在 F 欄開始放入混合型別資料，測試更多型別與公式
				sheet
				.SetTable(mixedData, ExcelColumns.F, 1)
				.BeginMapCell("ID").End()
				.BeginMapCell("Name").SetCellStyle("HighlightYellow").End()
				.BeginMapCell("DateOfBirth").SetCellStyle("DateOfBirth").End()
				.BeginMapCell("IsActive").SetCellType(CellType.Boolean).End()
				.BeginMapCell("Score").SetCellType(CellType.Numeric).End()
				.BeginMapCell("Amount").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()
				.BeginMapCell("Notes").SetCellType(CellType.String).End()
				// 單次內嵌樣式（橘底），使用唯一鍵避免覆寫快取
				.BeginMapCell("FormulaVal").SetCellType(CellType.Formula).SetCellStyle("InlineOrangeFormula", (p, s) =>
				{
					s.FillPattern = FillPattern.SolidForeground;
					s.SetCellFillForegroundColor(IndexedColors.Orange);
				}).End()
				.SetRow();

				// 第三組：DataTable 測試（K 欄開始）
				var dataTable = new DataTable("DtSample");
				dataTable.Columns.Add("ID", typeof(int));
				dataTable.Columns.Add("Name", typeof(string));
				dataTable.Columns.Add("DateOfBirth", typeof(DateTime));
				dataTable.Columns.Add("IsActive", typeof(bool));
				dataTable.Columns.Add("Score", typeof(double));
				dataTable.Columns.Add("Amount", typeof(decimal));
				dataTable.Columns.Add("MaybeNull", typeof(object));
				dataTable.Columns.Add("FormulaVal", typeof(string));

				dataTable.Rows.Add(2001, "DT-Alice", new DateTime(1981, 1, 2), true, 88.5, 321.09m, DBNull.Value, "SUM(K2:K4)");
				dataTable.Rows.Add(2002, "DT-Bob", new DateTime(1979, 6, 30), false, 77.75, 0m, "not null", "L2*2");
				dataTable.Rows.Add(2003, "DT-Carol", new DateTime(2002, 11, 5), true, 0.0, -999.99m, DBNull.Value, "AVERAGE(M2:M4)");

				var dtList = dataTable.AsEnumerable().Select(r => new Dictionary<string, object>
				{
					{ "ID", r["ID"] },
					{ "Name", r["Name"] },
					{ "DateOfBirth", r["DateOfBirth"] },
					{ "IsActive", r["IsActive"] },
					{ "Score", r["Score"] },
					{ "Amount", r["Amount"] },
					{ "MaybeNull", r["MaybeNull"] },
					{ "FormulaVal", r["FormulaVal"] },
				}).ToList();

				sheet
				.SetTable(dtList, ExcelColumns.K, 1)
				.BeginMapCell("ID").SetCellStyle("HeaderBlue").End()
				.BeginMapCell("Name").SetCellStyle("HighlightYellow").End()
				.BeginMapCell("DateOfBirth").SetCellStyle("DateOfBirth").End()
				.BeginMapCell("IsActive").SetCellType(CellType.Boolean).End()
				.BeginMapCell("Score").SetCellType(CellType.Numeric).End()
				.BeginMapCell("Amount").SetCellType(CellType.Numeric).SetCellStyle("AmountCurrency").End()
				.BeginMapCell("MaybeNull").End()
				.BeginMapCell("FormulaVal").SetCellType(CellType.Formula).SetCellStyle("InlineOrangeFormulaDT", (p, s) =>
				{
					s.FillPattern = FillPattern.SolidForeground;
					s.SetCellFillForegroundColor(IndexedColors.LightOrange);
				}).End()
				.SetRow()
				.Save(outputPath);

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}
	}

}