using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using FluentNPOI.Stages;

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
        [Fact]
        public void GetCellValue_ShouldReturnCorrectValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 為日期單元格設置格式
            fluentWorkbook.SetupCellStyle("DateFormat", (wb, style) =>
            {
                style.SetDataFormat(wb, "yyyy-MM-dd");
            });

            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // 設置不同類型的值
            sheet.SetCellPosition(ExcelColumns.A, 1).SetValue("Hello");
            sheet.SetCellPosition(ExcelColumns.B, 1).SetValue(123);
            sheet.SetCellPosition(ExcelColumns.C, 1).SetValue(45.67);
            sheet.SetCellPosition(ExcelColumns.D, 1).SetValue(true);

            // 設置日期值並套用日期格式
            sheet.SetCellPosition(ExcelColumns.E, 1)
                .SetValue(new DateTime(2024, 1, 15))
                .SetCellStyle("DateFormat");

            // Act & Assert
            var stringValue = sheet.GetCellValue<string>(ExcelColumns.A, 1);
            Assert.Equal("Hello", stringValue);

            var intValue = sheet.GetCellValue<int>(ExcelColumns.B, 1);
            Assert.Equal(123, intValue);

            var doubleValue = sheet.GetCellValue<double>(ExcelColumns.C, 1);
            Assert.Equal(45.67, doubleValue, 2);

            var boolValue = sheet.GetCellValue<bool>(ExcelColumns.D, 1);
            Assert.True(boolValue);

            var dateValue = sheet.GetCellValue<DateTime>(ExcelColumns.E, 1);
            Assert.Equal(new DateTime(2024, 1, 15), dateValue);
        }

        [Fact]
        public void GetCellValue_NonExistentCell_ShouldReturnDefault()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // Act
            var value = sheet.GetCellValue<string>(ExcelColumns.Z, 100);

            // Assert
            Assert.Null(value);
        }

        [Fact]
        public void FluentCell_GetValue_ShouldReturnCorrectValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // Act
            sheet.SetCellPosition(ExcelColumns.A, 1).SetValue("Test Value");
            var cell = sheet.GetCellPosition(ExcelColumns.A, 1);
            var value = cell.GetValue<string>();

            // Assert
            Assert.Equal("Test Value", value);
        }

        [Fact]
        public void SetAndGetFormula_ShouldWork()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // 設置一些數值
            sheet.SetCellPosition(ExcelColumns.A, 1).SetValue(10);
            sheet.SetCellPosition(ExcelColumns.B, 1).SetValue(20);

            // Act - 設置公式
            sheet.SetCellPosition(ExcelColumns.C, 1).SetFormulaValue("=A1+B1");

            // 讀取公式
            var formula = sheet.GetCellFormula(ExcelColumns.C, 1);

            // Assert
            Assert.Equal("A1+B1", formula);
        }

        [Fact]
        public void GetCellValue_WithObject_ShouldReturnCorrectType()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            sheet.SetCellPosition(ExcelColumns.A, 1).SetValue(123);
            sheet.SetCellPosition(ExcelColumns.B, 1).SetValue("Text");
            sheet.SetCellPosition(ExcelColumns.C, 1).SetValue(true);

            // Act
            var numValue = sheet.GetCellValue(ExcelColumns.A, 1);
            var textValue = sheet.GetCellValue(ExcelColumns.B, 1);
            var boolValue = sheet.GetCellValue(ExcelColumns.C, 1);

            // Assert
            Assert.IsType<double>(numValue);
            Assert.IsType<string>(textValue);
            Assert.IsType<bool>(boolValue);
        }
    }

    public class CellStyleRangeTests
    {
        [Fact]
        public void SetCellStyleRange_WithStringKey_ShouldApplyStyleToRange()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 設置一個預定義樣式
            fluentWorkbook.SetupCellStyle("TestRangeStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightBlue.Index;
                style.SetBorderAllStyle(BorderStyle.Thin);
            });

            // Act
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange("TestRangeStyle", ExcelColumns.A, ExcelColumns.C, 1, 3);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // 驗證範圍內的每個單元格都有正確的樣式
            for (int rowIndex = 0; rowIndex <= 2; rowIndex++) // 1-based 轉 0-based: 1-3 -> 0-2
            {
                var row = npoiSheet.GetRow(rowIndex);
                Assert.NotNull(row);

                for (int colIndex = 0; colIndex <= 2; colIndex++) // A-C -> 0-2
                {
                    var cell = row.GetCell(colIndex);
                    Assert.NotNull(cell);
                    Assert.NotNull(cell.CellStyle);
                    Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                    Assert.Equal(IndexedColors.LightBlue.Index, cell.CellStyle.FillForegroundColor);
                    Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderTop);
                }
            }
        }

        [Fact]
        public void SetCellStyleRange_WithCellStyleConfig_ShouldApplyStyleToRange()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var styleConfig = new CellStyleConfig("DynamicRangeStyle", style =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Yellow.Index;
                style.Alignment = HorizontalAlignment.Center;
            });

            // Act
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange(styleConfig, ExcelColumns.B, ExcelColumns.D, 2, 4);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // 驗證範圍內的每個單元格都有正確的樣式
            for (int rowIndex = 1; rowIndex <= 3; rowIndex++) // 2-4 -> 1-3
            {
                var row = npoiSheet.GetRow(rowIndex);
                Assert.NotNull(row);

                for (int colIndex = 1; colIndex <= 3; colIndex++) // B-D -> 1-3
                {
                    var cell = row.GetCell(colIndex);
                    Assert.NotNull(cell);
                    Assert.NotNull(cell.CellStyle);
                    Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                    Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
                    Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);
                }
            }
        }

        [Fact]
        public void SetCellStyleRange_SingleCell_ShouldApplyStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.SetupCellStyle("SingleCellStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Green.Index;
            });

            // Act - 只設置單個單元格 (A1)
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange("SingleCellStyle", ExcelColumns.A, ExcelColumns.A, 1, 1);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Green.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetCellStyleRange_WithExistingValues_ShouldPreserveValues()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.SetupCellStyle("PreserveStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Orange.Index;
            });

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // 先設置一些值
            sheet.SetCellPosition(ExcelColumns.A, 1).SetValue("Test1");
            sheet.SetCellPosition(ExcelColumns.B, 1).SetValue(123);
            sheet.SetCellPosition(ExcelColumns.C, 1).SetValue(true);

            // Act - 對這些已有值的單元格套用樣式
            sheet.SetCellStyleRange("PreserveStyle", ExcelColumns.A, ExcelColumns.C, 1, 1);

            // Assert - 驗證值被保留且樣式已套用
            var npoiSheet = workbook.GetSheet("Sheet1");
            var row = npoiSheet.GetRow(0);

            var cellA = row.GetCell(0);
            Assert.Equal("Test1", cellA.StringCellValue);
            Assert.Equal(IndexedColors.Orange.Index, cellA.CellStyle.FillForegroundColor);

            var cellB = row.GetCell(1);
            Assert.Equal(123.0, cellB.NumericCellValue);
            Assert.Equal(IndexedColors.Orange.Index, cellB.CellStyle.FillForegroundColor);

            var cellC = row.GetCell(2);
            Assert.True(cellC.BooleanCellValue);
            Assert.Equal(IndexedColors.Orange.Index, cellC.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetCellStyleRange_LargeRange_ShouldApplyToAllCells()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var styleConfig = new CellStyleConfig("LargeRangeStyle", style =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            });

            // Act - 設置較大範圍 (A1:E10)
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange(styleConfig, ExcelColumns.A, ExcelColumns.E, 1, 10);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // 驗證範圍的四個角落和中心點
            var testPoints = new[]
            {
                (row: 0, col: 0), // A1
				(row: 0, col: 4), // E1
				(row: 9, col: 0), // A10
				(row: 9, col: 4), // E10
				(row: 5, col: 2)  // C6 (中心)
			};

            foreach (var (row, col) in testPoints)
            {
                var cell = npoiSheet.GetRow(row)?.GetCell(col);
                Assert.NotNull(cell);
                Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                Assert.Equal(IndexedColors.Grey25Percent.Index, cell.CellStyle.FillForegroundColor);
            }

            // 驗證總共創建了 10 行（1-10）
            Assert.Equal(10, npoiSheet.PhysicalNumberOfRows);
        }
    }

    public class CopyStyleFromSheetCellTests
    {
        [Fact]
        public void CopyStyleFromSheetCell_ShouldCacheStyleAtWorkbookLevel()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 創建模板工作表並設置樣式
            var templateSheet = fluentWorkbook
                .SetupCellStyle("templateStyle", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Aqua.Index;
                    style.Alignment = HorizontalAlignment.Right;
                })
                .UseSheet("Template");

            templateSheet.SetCellPosition(ExcelColumns.B, 3)
                .SetCellStyle("templateStyle")
                .SetValue("Template Cell");

            // Act - 從模板工作表複製樣式到工作簿層級
            var templateSheetRef = templateSheet.GetSheet();
            fluentWorkbook.CopyStyleFromSheetCell("copiedFromTemplate", templateSheetRef, ExcelColumns.B, 3);

            // 在另一個工作表中使用該樣式
            var dataSheet = fluentWorkbook.UseSheet("Data");
            dataSheet.SetCellPosition(ExcelColumns.A, 1)
                .SetCellStyle("copiedFromTemplate")
                .SetValue("Using Copied Style");

            // Assert
            var npoiDataSheet = workbook.GetSheet("Data");
            var cell = npoiDataSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal("Using Copied Style", cell.StringCellValue);
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Aqua.Index, cell.CellStyle.FillForegroundColor);
            Assert.Equal(HorizontalAlignment.Right, cell.CellStyle.Alignment);
        }

        [Fact]
        public void CopyStyleFromSheetCell_ShouldWorkAcrossMultipleSheets()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 在第一個工作表創建樣式
            var sheet1 = fluentWorkbook
                .SetupCellStyle("originalStyle", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.LightGreen.Index;
                    var font = wb.CreateFont();
                    font.IsBold = true;
                    style.SetFont(font);
                })
                .UseSheet("Sheet1");

            sheet1.SetCellPosition(ExcelColumns.A, 1)
                .SetCellStyle("originalStyle")
                .SetValue("Source");

            var sheet1Ref = sheet1.GetSheet();

            // Act - 複製樣式並在多個工作表中使用
            fluentWorkbook.CopyStyleFromSheetCell("sharedStyle", sheet1Ref, ExcelColumns.A, 1);

            var sheet2 = fluentWorkbook.UseSheet("Sheet2");
            sheet2.SetCellPosition(ExcelColumns.A, 1).SetCellStyle("sharedStyle").SetValue("Sheet2 Data");

            var sheet3 = fluentWorkbook.UseSheet("Sheet3");
            sheet3.SetCellPosition(ExcelColumns.A, 1).SetCellStyle("sharedStyle").SetValue("Sheet3 Data");

            // Assert - 驗證所有工作表都使用了相同的樣式
            var npoiSheet2 = workbook.GetSheet("Sheet2");
            var npoiSheet3 = workbook.GetSheet("Sheet3");

            var cell2 = npoiSheet2.GetRow(0)?.GetCell(0);
            var cell3 = npoiSheet3.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
            Assert.Equal(IndexedColors.LightGreen.Index, cell2.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.LightGreen.Index, cell3.CellStyle.FillForegroundColor);
            Assert.True(cell2.CellStyle.GetFont(workbook).IsBold);
            Assert.True(cell3.CellStyle.GetFont(workbook).IsBold);
        }

        [Fact]
        public void CopyStyleFromSheetCell_WithNonExistentCell_ShouldNotCache()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var emptySheet = fluentWorkbook.UseSheet("Empty").GetSheet();

            // Act - 嘗試從不存在的單元格複製樣式
            fluentWorkbook.CopyStyleFromSheetCell("nonExistent", emptySheet, ExcelColumns.Z, 999);

            // 嘗試使用該樣式 - 應該不會有特殊樣式
            var testSheet = fluentWorkbook.UseSheet("Test");
            testSheet.SetCellPosition(ExcelColumns.A, 1)
                .SetCellStyle("nonExistent")
                .SetValue("Test");

            // Assert
            var npoiSheet = workbook.GetSheet("Test");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            // 因為樣式未被緩存，單元格應該使用默認樣式
            Assert.Equal(FillPattern.NoFill, cell.CellStyle.FillPattern);
        }

        [Fact]
        public void CopyStyleFromSheetCell_SameCacheKeyTwice_ShouldNotOverwrite()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 創建兩個不同樣式的工作表
            var blueSheet = fluentWorkbook
                .SetupCellStyle("blue", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Blue.Index;
                })
                .UseSheet("BluePalette");
            blueSheet.SetCellPosition(ExcelColumns.A, 1).SetCellStyle("blue").SetValue("Blue");

            var redSheet = fluentWorkbook
                .SetupCellStyle("red", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Red.Index;
                })
                .UseSheet("RedPalette");
            redSheet.SetCellPosition(ExcelColumns.A, 1).SetCellStyle("red").SetValue("Red");

            var blueSheetRef = blueSheet.GetSheet();
            var redSheetRef = redSheet.GetSheet();

            // Act - 第一次複製藍色樣式
            fluentWorkbook.CopyStyleFromSheetCell("myColor", blueSheetRef, ExcelColumns.A, 1);

            // 第二次嘗試用相同的鍵複製紅色樣式 - 應該被忽略
            fluentWorkbook.CopyStyleFromSheetCell("myColor", redSheetRef, ExcelColumns.A, 1);

            // 使用該樣式
            var testSheet = fluentWorkbook.UseSheet("Test");
            testSheet.SetCellPosition(ExcelColumns.A, 1)
                .SetCellStyle("myColor")
                .SetValue("Test");

            // Assert - 應該仍然是第一次複製的藍色
            var npoiSheet = workbook.GetSheet("Test");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(IndexedColors.Blue.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void CopyStyleFromSheetCell_ChainedCalls_ShouldWork()
        {
            // Arrange \u0026 Act
            var workbook = new XSSFWorkbook();

            // 使用鏈式調用複製多個樣式
            var fluentWorkbook = new FluentWorkbook(workbook);

            var paletteSheet = fluentWorkbook.UseSheet("Palette");

            // 設置多個樣式單元格
            fluentWorkbook
                .SetupCellStyle("style1", (wb, s) => s.FillForegroundColor = IndexedColors.Yellow.Index)
                .SetupCellStyle("style2", (wb, s) => s.FillForegroundColor = IndexedColors.Green.Index)
                .SetupCellStyle("style3", (wb, s) => s.FillForegroundColor = IndexedColors.Orange.Index);

            paletteSheet.SetCellPosition(ExcelColumns.A, 1).SetCellStyle("style1").SetValue("Style1");
            paletteSheet.SetCellPosition(ExcelColumns.A, 2).SetCellStyle("style2").SetValue("Style2");
            paletteSheet.SetCellPosition(ExcelColumns.A, 3).SetCellStyle("style3").SetValue("Style3");

            var paletteSheetRef = paletteSheet.GetSheet();

            // 鏈式調用複製所有樣式
            fluentWorkbook
                .CopyStyleFromSheetCell("copied1", paletteSheetRef, ExcelColumns.A, 1)
                .CopyStyleFromSheetCell("copied2", paletteSheetRef, ExcelColumns.A, 2)
                .CopyStyleFromSheetCell("copied3", paletteSheetRef, ExcelColumns.A, 3);

            // 使用複製的樣式
            var dataSheet = fluentWorkbook.UseSheet("Data");
            dataSheet.SetCellPosition(ExcelColumns.B, 1).SetCellStyle("copied1").SetValue("Copy1");
            dataSheet.SetCellPosition(ExcelColumns.B, 2).SetCellStyle("copied2").SetValue("Copy2");
            dataSheet.SetCellPosition(ExcelColumns.B, 3).SetCellStyle("copied3").SetValue("Copy3");

            // Assert
            var npoiDataSheet = workbook.GetSheet("Data");

            var cell1 = npoiDataSheet.GetRow(0)?.GetCell(1);
            var cell2 = npoiDataSheet.GetRow(1)?.GetCell(1);
            var cell3 = npoiDataSheet.GetRow(2)?.GetCell(1);

            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);

            Assert.Equal(IndexedColors.Yellow.Index, cell1.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Green.Index, cell2.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Orange.Index, cell3.CellStyle.FillForegroundColor);
        }
    }
}
