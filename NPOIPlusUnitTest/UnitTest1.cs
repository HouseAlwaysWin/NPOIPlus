using Xunit;
using NPOIPlus;
using NPOIPlus.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace NPOIPlusUnitTest
{
	public class BasicTests
	{
		[Fact]
		public void CreateWorkbook_ShouldReturnValidWorkbook()
		{
			// Arrange & Act
			var fluentWorkbook = new FluentWorkbook(new XSSFWorkbook());
			var workbook = fluentWorkbook.GetWorkbook();

			// Assert
			Assert.NotNull(workbook);
			Assert.IsAssignableFrom<IWorkbook>(workbook);
		}

		[Fact]
		public void UseSheet_ShouldCreateNewSheet()
		{
			// Arrange
			var fluentWorkbook = new FluentWorkbook(new XSSFWorkbook());

			// Act
			var sheet = fluentWorkbook.UseSheet("TestSheet");

			// Assert
			Assert.NotNull(sheet);
			Assert.NotNull(sheet.GetSheet());
			Assert.Equal("TestSheet", sheet.GetSheet().SheetName);
		}

		[Fact]
		public void SetCellPosition_ShouldSetValue()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetCellPosition(ExcelColumns.A, 1)
				.SetValue("Hello World");

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var cell = sheet.GetRow(0)?.GetCell(0);
			Assert.NotNull(cell);
			Assert.Equal("Hello World", cell.StringCellValue);
		}

		[Fact]
		public void SetColumnWidth_ShouldSetWidth()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetColumnWidth(ExcelColumns.A, 30);

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var width = sheet.GetColumnWidth(0);
			Assert.Equal(30 * 256, width);
		}

		[Fact]
		public void SetCellPosition_WithNumericValue_ShouldSetNumber()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetCellPosition(ExcelColumns.A, 1)
				.SetValue(123);

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var cell = sheet.GetRow(0)?.GetCell(0);
			Assert.NotNull(cell);
			Assert.Equal(123.0, cell.NumericCellValue);
		}
	}

	public class TableTests
	{
		private class TestData
		{
			public int ID { get; set; }
			public string Name { get; set; } = string.Empty;
			public bool IsActive { get; set; }
		}

		[Fact]
		public void SetTable_ShouldCreateTableWithHeaders()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);
			var testData = new List<TestData>
			{
				new TestData { ID = 1, Name = "Alice", IsActive = true },
				new TestData { ID = 2, Name = "Bob", IsActive = false }
			};

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginTitleSet("編號")
				.BeginBodySet("ID").End()
				.BeginTitleSet("姓名")
				.BeginBodySet("Name").End()
				.BuildRows();

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var titleRow = sheet.GetRow(0);
			Assert.NotNull(titleRow);
			Assert.Equal("編號", titleRow.GetCell(0)?.StringCellValue);
			Assert.Equal("姓名", titleRow.GetCell(1)?.StringCellValue);
		}

		[Fact]
		public void SetTable_WithMultipleRows_ShouldCreateAllRows()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);
			var testData = new List<TestData>
			{
				new TestData { ID = 1, Name = "Alice", IsActive = true },
				new TestData { ID = 2, Name = "Bob", IsActive = false },
				new TestData { ID = 3, Name = "Charlie", IsActive = true }
			};

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginTitleSet("姓名")
				.BeginBodySet("Name").End()
				.BuildRows();

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			Assert.Equal(4, sheet.PhysicalNumberOfRows); // 1 標題 + 3 數據行
		}
	}

	public class CellStyleConfigTests
	{
		[Fact]
		public void Constructor_ShouldSetProperties()
		{
			// Arrange
			Action<ICellStyle> setter = (style) => { };

			// Act
			var config = new CellStyleConfig("TestKey", setter);

			// Assert
			Assert.Equal("TestKey", config.Key);
			Assert.Equal(setter, config.StyleSetter);
		}

		[Fact]
		public void Deconstruct_ShouldReturnKeyAndSetter()
		{
			// Arrange
			Action<ICellStyle> setter = (style) => { };
			var config = new CellStyleConfig("TestKey", setter);

			// Act
			var (key, styleSetter) = config;

			// Assert
			Assert.Equal("TestKey", key);
			Assert.Equal(setter, styleSetter);
		}
	}

	public class StyleCachingTests
	{
		private class TestData
		{
			public int ID { get; set; }
			public bool IsActive { get; set; }
		}

		[Fact]
		public void DynamicStyle_ShouldCacheByKey()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);
			var testData = new List<TestData>
			{
				new TestData { ID = 1, IsActive = true },
				new TestData { ID = 2, IsActive = false },
				new TestData { ID = 3, IsActive = true },
				new TestData { ID = 4, IsActive = false }
			};

			int styleSetterCallCount = 0;

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginTitleSet("狀態")
				.BeginBodySet("IsActive")
				.SetCellStyle((styleParams) =>
				{
					if (styleParams.GetRowItem<TestData>().IsActive)
					{
						return new CellStyleConfig("ActiveStyle", style =>
						{
							styleSetterCallCount++;
							style.FillPattern = FillPattern.SolidForeground;
							style.FillForegroundColor = IndexedColors.Green.Index;
						});
					}
					return new CellStyleConfig("InactiveStyle", style =>
					{
						styleSetterCallCount++;
						style.FillPattern = FillPattern.SolidForeground;
						style.FillForegroundColor = IndexedColors.Yellow.Index;
					});
				})
				.End()
				.BuildRows();

			// Assert - 4 行數據，但只有 2 種樣式，所以 StyleSetter 應該只被調用 2 次
			Assert.Equal(2, styleSetterCallCount);
		}

		[Fact]
		public void SetupCellStyle_ShouldApplyStyle()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);
			var testData = new List<TestData>
			{
				new TestData { ID = 1, IsActive = true },
				new TestData { ID = 2, IsActive = true }
			};

			// Act
			fluentWorkbook
				.SetupCellStyle("FixedStyle", (wb, style) =>
				{
					style.FillPattern = FillPattern.SolidForeground;
					style.FillForegroundColor = IndexedColors.Blue.Index;
				})
				.UseSheet("Sheet1")
				.SetTable(testData, ExcelColumns.A, 1)
				.BeginTitleSet("ID")
				.BeginBodySet("ID").SetCellStyle("FixedStyle").End()
				.BuildRows();

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var cell1 = sheet.GetRow(1)?.GetCell(0);
			var cell2 = sheet.GetRow(2)?.GetCell(0);

			// 驗證樣式已套用
			Assert.NotNull(cell1?.CellStyle);
			Assert.NotNull(cell2?.CellStyle);
			Assert.Equal(FillPattern.SolidForeground, cell1.CellStyle.FillPattern);
			Assert.Equal(FillPattern.SolidForeground, cell2.CellStyle.FillPattern);
			Assert.Equal(IndexedColors.Blue.Index, cell1.CellStyle.FillForegroundColor);
			Assert.Equal(IndexedColors.Blue.Index, cell2.CellStyle.FillForegroundColor);
		}
	}

	public class PositionNormalizationTests
	{
		[Fact]
		public void SetCellPosition_WithOneBased_ShouldConvertToZeroBased()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);

			// Act - 使用 1-based 行號
			fluentWorkbook.UseSheet("Sheet1")
				.SetCellPosition(ExcelColumns.A, 1)
				.SetValue("Row 1");

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var cell = sheet.GetRow(0)?.GetCell(0);
			Assert.NotNull(cell);
			Assert.Equal("Row 1", cell.StringCellValue);
		}

		[Fact]
		public void SetCellPosition_WithNegativeRow_ShouldUseZero()
		{
			// Arrange
			var workbook = new XSSFWorkbook();
			var fluentWorkbook = new FluentWorkbook(workbook);

			// Act
			fluentWorkbook.UseSheet("Sheet1")
				.SetCellPosition(ExcelColumns.A, -1)
				.SetValue("Test");

			// Assert
			var sheet = workbook.GetSheet("Sheet1");
			var cell = sheet.GetRow(0)?.GetCell(0);
			Assert.NotNull(cell);
		}
	}
}
