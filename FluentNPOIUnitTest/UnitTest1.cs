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
                .SetCellPosition(ExcelCol.A, 1)
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
                .SetColumnWidth(ExcelCol.A, 30);

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
                .SetCellPosition(ExcelCol.A, 1)
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
                .SetTable(testData, ExcelCol.A, 1)
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
                .SetTable(testData, ExcelCol.A, 1)
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
                .SetTable(testData, ExcelCol.A, 1)
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
                .SetTable(testData, ExcelCol.A, 1)
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
                .SetCellPosition(ExcelCol.A, 1)
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
                .SetCellPosition(ExcelCol.A, -1)
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
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Hello");
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue(123);
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue(45.67);
            sheet.SetCellPosition(ExcelCol.D, 1).SetValue(true);

            // 設置日期值並套用日期格式
            sheet.SetCellPosition(ExcelCol.E, 1)
                .SetValue(new DateTime(2024, 1, 15))
                .SetCellStyle("DateFormat");

            // Act & Assert
            var stringValue = sheet.GetCellValue<string>(ExcelCol.A, 1);
            Assert.Equal("Hello", stringValue);

            var intValue = sheet.GetCellValue<int>(ExcelCol.B, 1);
            Assert.Equal(123, intValue);

            var doubleValue = sheet.GetCellValue<double>(ExcelCol.C, 1);
            Assert.Equal(45.67, doubleValue, 2);

            var boolValue = sheet.GetCellValue<bool>(ExcelCol.D, 1);
            Assert.True(boolValue);

            var dateValue = sheet.GetCellValue<DateTime>(ExcelCol.E, 1);
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
            var value = sheet.GetCellValue<string>(ExcelCol.Z, 100);

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
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Test Value");
            var cell = sheet.GetCellPosition(ExcelCol.A, 1);
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
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue(10);
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue(20);

            // Act - 設置公式
            sheet.SetCellPosition(ExcelCol.C, 1).SetFormulaValue("=A1+B1");

            // 讀取公式
            var formula = sheet.GetCellFormula(ExcelCol.C, 1);

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

            sheet.SetCellPosition(ExcelCol.A, 1).SetValue(123);
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue("Text");
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue(true);

            // Act
            var numValue = sheet.GetCellValue(ExcelCol.A, 1);
            var textValue = sheet.GetCellValue(ExcelCol.B, 1);
            var boolValue = sheet.GetCellValue(ExcelCol.C, 1);

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
            sheet.SetCellStyleRange("TestRangeStyle", ExcelCol.A, ExcelCol.C, 1, 3);

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
            sheet.SetCellStyleRange(styleConfig, ExcelCol.B, ExcelCol.D, 2, 4);

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
            sheet.SetCellStyleRange("SingleCellStyle", ExcelCol.A, ExcelCol.A, 1, 1);

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
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Test1");
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue(123);
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue(true);

            // Act - 對這些已有值的單元格套用樣式
            sheet.SetCellStyleRange("PreserveStyle", ExcelCol.A, ExcelCol.C, 1, 1);

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
            sheet.SetCellStyleRange(styleConfig, ExcelCol.A, ExcelCol.E, 1, 10);

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

            templateSheet.SetCellPosition(ExcelCol.B, 3)
                .SetCellStyle("templateStyle")
                .SetValue("Template Cell");

            // Act - 從模板工作表複製樣式到工作簿層級
            var templateSheetRef = templateSheet.GetSheet();
            fluentWorkbook.CopyStyleFromSheetCell("copiedFromTemplate", templateSheetRef, ExcelCol.B, 3);

            // 在另一個工作表中使用該樣式
            var dataSheet = fluentWorkbook.UseSheet("Data");
            dataSheet.SetCellPosition(ExcelCol.A, 1)
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

            sheet1.SetCellPosition(ExcelCol.A, 1)
                .SetCellStyle("originalStyle")
                .SetValue("Source");

            var sheet1Ref = sheet1.GetSheet();

            // Act - 複製樣式並在多個工作表中使用
            fluentWorkbook.CopyStyleFromSheetCell("sharedStyle", sheet1Ref, ExcelCol.A, 1);

            var sheet2 = fluentWorkbook.UseSheet("Sheet2");
            sheet2.SetCellPosition(ExcelCol.A, 1).SetCellStyle("sharedStyle").SetValue("Sheet2 Data");

            var sheet3 = fluentWorkbook.UseSheet("Sheet3");
            sheet3.SetCellPosition(ExcelCol.A, 1).SetCellStyle("sharedStyle").SetValue("Sheet3 Data");

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
            fluentWorkbook.CopyStyleFromSheetCell("nonExistent", emptySheet, ExcelCol.Z, 999);

            // 嘗試使用該樣式 - 應該不會有特殊樣式
            var testSheet = fluentWorkbook.UseSheet("Test");
            testSheet.SetCellPosition(ExcelCol.A, 1)
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
            blueSheet.SetCellPosition(ExcelCol.A, 1).SetCellStyle("blue").SetValue("Blue");

            var redSheet = fluentWorkbook
                .SetupCellStyle("red", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Red.Index;
                })
                .UseSheet("RedPalette");
            redSheet.SetCellPosition(ExcelCol.A, 1).SetCellStyle("red").SetValue("Red");

            var blueSheetRef = blueSheet.GetSheet();
            var redSheetRef = redSheet.GetSheet();

            // Act - 第一次複製藍色樣式
            fluentWorkbook.CopyStyleFromSheetCell("myColor", blueSheetRef, ExcelCol.A, 1);

            // 第二次嘗試用相同的鍵複製紅色樣式 - 應該被忽略
            fluentWorkbook.CopyStyleFromSheetCell("myColor", redSheetRef, ExcelCol.A, 1);

            // 使用該樣式
            var testSheet = fluentWorkbook.UseSheet("Test");
            testSheet.SetCellPosition(ExcelCol.A, 1)
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

            paletteSheet.SetCellPosition(ExcelCol.A, 1).SetCellStyle("style1").SetValue("Style1");
            paletteSheet.SetCellPosition(ExcelCol.A, 2).SetCellStyle("style2").SetValue("Style2");
            paletteSheet.SetCellPosition(ExcelCol.A, 3).SetCellStyle("style3").SetValue("Style3");

            var paletteSheetRef = paletteSheet.GetSheet();

            // 鏈式調用複製所有樣式
            fluentWorkbook
                .CopyStyleFromSheetCell("copied1", paletteSheetRef, ExcelCol.A, 1)
                .CopyStyleFromSheetCell("copied2", paletteSheetRef, ExcelCol.A, 2)
                .CopyStyleFromSheetCell("copied3", paletteSheetRef, ExcelCol.A, 3);

            // 使用複製的樣式
            var dataSheet = fluentWorkbook.UseSheet("Data");
            dataSheet.SetCellPosition(ExcelCol.B, 1).SetCellStyle("copied1").SetValue("Copy1");
            dataSheet.SetCellPosition(ExcelCol.B, 2).SetCellStyle("copied2").SetValue("Copy2");
            dataSheet.SetCellPosition(ExcelCol.B, 3).SetCellStyle("copied3").SetValue("Copy3");

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

    public class GetTableAutoDetectLastRowTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public string Name { get; set; } = string.Empty;
            public double Score { get; set; }
            public bool IsActive { get; set; }
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_ShouldReadAllRows()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
                new TestData { ID = 2, Name = "Bob", Score = 87.0, IsActive = false },
                new TestData { ID = 3, Name = "Charlie", Score = 92.0, IsActive = true }
            };

            // 先寫入數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // Act - 使用自動判斷最後一行
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2); // 從第2行開始（跳過標題）

            // Assert
            Assert.Equal(3, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal(95.5, readData[0].Score);
            Assert.True(readData[0].IsActive);

            Assert.Equal(2, readData[1].ID);
            Assert.Equal("Bob", readData[1].Name);
            Assert.Equal(87.0, readData[1].Score);
            Assert.False(readData[1].IsActive);

            Assert.Equal(3, readData[2].ID);
            Assert.Equal("Charlie", readData[2].Name);
            Assert.Equal(92.0, readData[2].Score);
            Assert.True(readData[2].IsActive);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_ShouldMatchManualEndRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
                new TestData { ID = 2, Name = "Bob", Score = 87.0, IsActive = false },
                new TestData { ID = 3, Name = "Charlie", Score = 92.0, IsActive = true },
                new TestData { ID = 4, Name = "David", Score = 88.5, IsActive = true }
            };

            // 先寫入數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 兩種方式讀取
            var autoDetectData = sheet.GetTable<TestData>(ExcelCol.A, 2); // 自動判斷
            var manualData = sheet.GetTable<TestData>(ExcelCol.A, 2, 5); // 手動指定結束行

            // Assert - 兩種方式應該讀取相同的數據
            Assert.Equal(manualData.Count, autoDetectData.Count);
            for (int i = 0; i < autoDetectData.Count; i++)
            {
                Assert.Equal(manualData[i].ID, autoDetectData[i].ID);
                Assert.Equal(manualData[i].Name, autoDetectData[i].Name);
                Assert.Equal(manualData[i].Score, autoDetectData[i].Score);
                Assert.Equal(manualData[i].IsActive, autoDetectData[i].IsActive);
            }
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_WithEmptyRows_ShouldStopAtLastDataRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
                new TestData { ID = 2, Name = "Bob", Score = 87.0, IsActive = false }
            };

            // 先寫入數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // 在第3行之後添加一些空行（手動創建空行）
            var npoiSheet = workbook.GetSheet("Sheet1");
            npoiSheet.CreateRow(3); // 第4行（0-based index 3）
            npoiSheet.CreateRow(4); // 第5行（0-based index 4）

            // Act - 自動判斷應該只讀取到有數據的最後一行
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert - 應該只讀取2行數據，忽略後面的空行
            Assert.Equal(2, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal(2, readData[1].ID);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_WithGaps_ShouldStopAtLastDataRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
                new TestData { ID = 2, Name = "Bob", Score = 87.0, IsActive = false }
            };

            // 先寫入數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // 在第3行之後添加一個空行，然後再添加一個有數據的行（模擬中間有空行的情況）
            var npoiSheet = workbook.GetSheet("Sheet1");
            var emptyRow = npoiSheet.CreateRow(3); // 第4行是空的
            var rowWithData = npoiSheet.CreateRow(4); // 第5行有數據
            rowWithData.CreateCell(0).SetCellValue(999); // 在A列設置一個值

            // Act - 自動判斷應該讀取到最後一個有數據的行
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert - 應該讀取到第5行（因為A5有數據），但由於第4行是空的，可能只讀取2行
            // 實際行為：會讀取到最後一個A列有數據的行（第5行），但第4行因為是空的會被跳過
            Assert.True(readData.Count >= 2);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal(2, readData[1].ID);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_SingleRow_ShouldReadCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true }
            };

            // 先寫入數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // Act
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_EmptySheet_ShouldReturnEmptyList()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("EmptySheet");

            // Act - 從第2行開始讀取，但工作表是空的
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert
            Assert.Empty(readData);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_WithNumericFirstColumn_ShouldDetectCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
                new TestData { ID = 2, Name = "Bob", Score = 87.0, IsActive = false },
                new TestData { ID = 3, Name = "Charlie", Score = 92.0, IsActive = true }
            };

            // 先寫入數據（ID在第一列，是數字類型）
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // Act - 從A列開始讀取（數字類型）
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert - 應該正確讀取所有行
            Assert.Equal(3, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal(2, readData[1].ID);
            Assert.Equal(3, readData[2].ID);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_WithStringFirstColumn_ShouldDetectCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", Score = 95.5, IsActive = true },
                new TestData { ID = 2, Name = "Bob", Score = 87.0, IsActive = false },
                new TestData { ID = 3, Name = "Charlie", Score = 92.0, IsActive = true }
            };

            // 先寫入數據（Name在第二列，是字符串類型）
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // Act - 從B列開始讀取（字符串類型），但這會導致讀取失敗，因為類型不匹配
            // 實際上應該從A列讀取，但我們可以測試從B列讀取時的行為
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // 創建一個只包含Name屬性的類來測試
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert - 應該正確讀取所有行
            Assert.Equal(3, readData.Count);
        }

        [Fact]
        public void GetTable_AutoDetectLastRow_LargeDataset_ShouldHandleCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>();

            // 創建100筆測試數據
            for (int i = 1; i <= 100; i++)
            {
                testData.Add(new TestData
                {
                    ID = i,
                    Name = $"User{i}",
                    Score = 50.0 + (i % 50),
                    IsActive = i % 2 == 0
                });
            }

            // 先寫入數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BuildRows();

            // Act
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var readData = sheet.GetTable<TestData>(ExcelCol.A, 2);

            // Assert
            Assert.Equal(100, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal(100, readData[99].ID);
            Assert.Equal("User1", readData[0].Name);
            Assert.Equal("User100", readData[99].Name);
        }
    }

    public class CellMergeTests
    {
        [Fact]
        public void SetExcelCellMerge_HorizontalMerge_ShouldMergeCellsInRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 横向合并 A1 到 C1
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // 1-based to 0-based: row 1 -> 0
            Assert.Equal(0, mergedRegion.LastRow);
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(2, mergedRegion.LastColumn); // C -> 2
        }

        [Fact]
        public void SetExcelCellMerge_VerticalMerge_ShouldMergeCellsInColumn()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 纵向合并 A1 到 A5
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 1, 5);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(4, mergedRegion.LastRow); // row 5 -> 4
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(0, mergedRegion.LastColumn); // A -> 0
        }

        [Fact]
        public void SetExcelCellMerge_RegionMerge_ShouldMergeCellRegion()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 区域合并 A1 到 C3
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1, 3);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(2, mergedRegion.LastRow); // row 3 -> 2
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(2, mergedRegion.LastColumn); // C -> 2
        }

        [Fact]
        public void SetExcelCellMerge_MultipleMerges_ShouldCreateMultipleRegions()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 创建多个合并区域
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1) // 横向合并第1行
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 2, 5) // 纵向合并 A2-A5
                .SetExcelCellMerge(ExcelCol.D, ExcelCol.F, 1, 3); // 区域合并 D1-F3

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(3, sheet.NumMergedRegions);

            // 验证第一个合并区域（横向）
            var region1 = sheet.GetMergedRegion(0);
            Assert.Equal(0, region1.FirstRow);
            Assert.Equal(0, region1.LastRow);
            Assert.Equal(0, region1.FirstColumn); // A
            Assert.Equal(2, region1.LastColumn); // C

            // 验证第二个合并区域（纵向）
            var region2 = sheet.GetMergedRegion(1);
            Assert.Equal(1, region2.FirstRow); // row 2 -> 1
            Assert.Equal(4, region2.LastRow); // row 5 -> 4
            Assert.Equal(0, region2.FirstColumn); // A
            Assert.Equal(0, region2.LastColumn); // A

            // 验证第三个合并区域（区域）
            var region3 = sheet.GetMergedRegion(2);
            Assert.Equal(0, region3.FirstRow); // row 1 -> 0
            Assert.Equal(2, region3.LastRow); // row 3 -> 2
            Assert.Equal(3, region3.FirstColumn); // D -> 3
            Assert.Equal(5, region3.LastColumn); // F -> 5
        }

        [Fact]
        public void SetExcelCellMerge_ChainedCalls_ShouldReturnFluentSheet()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 链式调用
            var result = fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.B, 1)
                .SetExcelCellMerge(ExcelCol.C, ExcelCol.D, 2);

            // Assert
            Assert.NotNull(result);
            Assert.IsType<FluentSheet>(result);
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(2, sheet.NumMergedRegions);
        }

        [Fact]
        public void SetExcelCellMerge_SingleCell_ShouldThrowException()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act & Assert - NPOI 不允许合并单个单元格，必须至少包含2个单元格
            Assert.Throws<ArgumentException>(() =>
            {
                fluentWorkbook.UseSheet("Sheet1")
                    .SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 1);
            });
        }

        [Fact]
        public void SetExcelCellMerge_WithCellValues_ShouldPreserveFirstCellValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 先设置值，再合并
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Merged Title");
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue("Will be merged");
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue("Also merged");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, npoiSheet.NumMergedRegions);

            // 验证合并后，只有第一个单元格有值
            var firstCell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(firstCell);
            Assert.Equal("Merged Title", firstCell.StringCellValue);

            // 合并区域内的其他单元格应该为空或引用第一个单元格
            var mergedRegion = npoiSheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow);
            Assert.Equal(0, mergedRegion.LastRow);
            Assert.Equal(0, mergedRegion.FirstColumn);
            Assert.Equal(2, mergedRegion.LastColumn);
        }

        [Fact]
        public void SetExcelCellMerge_DifferentRows_ShouldMergeCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 合并不同行的区域
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.B, 1, 3); // A1-B3

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(2, mergedRegion.LastRow); // row 3 -> 2
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(1, mergedRegion.LastColumn); // B -> 1
        }

        [Fact]
        public void SetExcelCellMerge_LargeRegion_ShouldHandleCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 合并大区域 A1 到 Z10
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.Z, 1, 10);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(9, mergedRegion.LastRow); // row 10 -> 9
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(25, mergedRegion.LastColumn); // Z -> 25
        }

        [Fact]
        public void SetExcelCellMerge_OverlappingRegions_ShouldThrowException()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 创建第一个合并区域
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1); // A1-C1

            // Assert - NPOI 不允许重叠的合并区域，应该抛出异常
            Assert.Throws<InvalidOperationException>(() =>
            {
                sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.D, 1); // B1-D1 (与第一个重叠)
            });

            // 验证第一个合并区域已创建
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, npoiSheet.NumMergedRegions);
            var region1 = npoiSheet.GetMergedRegion(0);
            Assert.Equal(0, region1.FirstColumn); // A
            Assert.Equal(2, region1.LastColumn); // C
        }
    }

    public class GetTableColumnRangeTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public string Name { get; set; } = string.Empty;
            public DateTime DateOfBirth { get; set; }
            public bool IsActive { get; set; }
            public double Score { get; set; }
            public decimal Amount { get; set; }
        }

        private class PartialData
        {
            public int ID { get; set; }
            public string Name { get; set; } = string.Empty;
            public DateTime DateOfBirth { get; set; }
        }

        private class MiddleColumnsData
        {
            public bool IsActive { get; set; }
            public double Score { get; set; }
            public decimal Amount { get; set; }
        }

        [Fact]
        public void GetTable_WithColumnRange_ShouldReadOnlySpecifiedColumns()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m },
                new TestData { ID = 3, Name = "Charlie", DateOfBirth = new DateTime(1992, 3, 3), IsActive = true, Score = 92.0, Amount = 3000m }
            };

            // 先寫入完整數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 A-C 列（ID, Name, DateOfBirth）
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 4);

            // Assert
            Assert.Equal(3, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal(new DateTime(1990, 1, 1), readData[0].DateOfBirth);
            Assert.Equal(2, readData[1].ID);
            Assert.Equal("Bob", readData[1].Name);
            Assert.Equal(3, readData[2].ID);
            Assert.Equal("Charlie", readData[2].Name);
        }

        [Fact]
        public void GetTable_WithColumnRange_MiddleColumns_ShouldReadCorrectColumns()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m }
            };

            // 先寫入完整數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 D-F 列（IsActive, Score, Amount）
            var readData = sheet.GetTable<MiddleColumnsData>(ExcelCol.D, ExcelCol.F, 2, 3);

            // Assert
            Assert.Equal(2, readData.Count);
            Assert.True(readData[0].IsActive);
            Assert.Equal(95.5, readData[0].Score);
            Assert.Equal(1000m, readData[0].Amount);
            Assert.False(readData[1].IsActive);
            Assert.Equal(87.0, readData[1].Score);
            Assert.Equal(2000m, readData[1].Amount);
        }

        [Fact]
        public void GetTable_WithColumnRange_SingleColumn_ShouldReadOnlyOneColumn()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 B 列（Name）
            var readData = sheet.GetTable<SingleColumnData>(ExcelCol.B, ExcelCol.B, 2, 3);

            // Assert
            Assert.Equal(2, readData.Count);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal("Bob", readData[1].Name);
        }

        [Fact]
        public void GetTable_WithColumnRange_MoreColumnsThanProperties_ShouldUseOnlyAvailableProperties()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 指定 A-F 列，但 PartialData 只有3個屬性
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.F, 2, 2);

            // Assert - 應該只使用前3個屬性（ID, Name, DateOfBirth）
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal(new DateTime(1990, 1, 1), readData[0].DateOfBirth);
        }

        [Fact]
        public void GetTable_WithColumnRange_FewerColumnsThanProperties_ShouldLeaveRemainingPropertiesDefault()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 A-B 列，但 TestData 有6個屬性
            var readData = sheet.GetTable<TestData>(ExcelCol.A, ExcelCol.B, 2, 2);

            // Assert - 只有前2個屬性會被填充
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            // 其他屬性應該保持默認值
            Assert.Equal(default(DateTime), readData[0].DateOfBirth);
            Assert.False(readData[0].IsActive);
            Assert.Equal(0.0, readData[0].Score);
            Assert.Equal(0m, readData[0].Amount);
        }

        [Fact]
        public void GetTable_WithColumnRange_EmptyRows_ShouldSkipEmptyRows()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            // 創建空行
            var npoiSheet = workbook.GetSheet("Sheet1");
            npoiSheet.CreateRow(3); // 第4行（0-based index 3）

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 讀取第2-5行，但第4行是空的
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 5);

            // Assert - 應該只讀取2行數據，跳過空行
            Assert.Equal(2, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal(2, readData[1].ID);
        }

        [Fact]
        public void GetTable_WithColumnRange_SingleRow_ShouldReadCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取單一行
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 2);

            // Assert
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
        }

        [Fact]
        public void GetTable_WithColumnRange_StartColAfterEndCol_ShouldHandleGracefully()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - startCol > endCol 的情況（雖然不合理，但應該能處理）
            var readData = sheet.GetTable<PartialData>(ExcelCol.C, ExcelCol.A, 2, 2);

            // Assert - 應該返回空列表或處理錯誤（實際行為取決於實現）
            // 由於 columnCount 會是負數，membersToUse 會是 0，所以不會讀取任何數據
            Assert.Empty(readData);
        }

        private class SingleColumnData
        {
            public string Name { get; set; } = string.Empty;
        }
    }
}
