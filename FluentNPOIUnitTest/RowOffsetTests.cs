using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Streaming.Mapping;
using FluentNPOI.Stages;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace FluentNPOIUnitTest
{
    public class RowOffsetTests
    {
        private class TestData
        {
            public string Name { get; set; } = "";
            public int Score { get; set; }
            public string Note { get; set; } = "";
        }

        [Fact]
        public void WithRowOffset_ShouldOffsetTitleRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數").WithRowOffset(1);

            var testData = new List<TestData> { new TestData { Name = "Alice", Score = 95 } };

            // Act
            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Test");
            
            // A 欄標題在 row 0 (Excel row 1)
            Assert.Equal("姓名", sheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            
            // B 欄標題因為有 offset=1，應該在 row 1 (Excel row 2)
            Assert.Equal("分數", sheet.GetRow(1)?.GetCell(1)?.StringCellValue);
        }

        [Fact]
        public void WithRowOffset_ShouldOffsetDataRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數").WithRowOffset(1);

            var testData = new List<TestData> 
            { 
                new TestData { Name = "Alice", Score = 95 },
                new TestData { Name = "Bob", Score = 78 }
            };

            // Act
            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Test");
            
            // A 欄資料從 row 1 開始（row 0 是標題）
            Assert.Equal("Alice", sheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Bob", sheet.GetRow(2)?.GetCell(0)?.StringCellValue);
            
            // B 欄資料因為 offset=1，從 row 2 開始
            Assert.Equal(95.0, sheet.GetRow(2)?.GetCell(1)?.NumericCellValue);
            Assert.Equal(78.0, sheet.GetRow(3)?.GetCell(1)?.NumericCellValue);
        }

        [Fact]
        public void WithRowOffset_MultipleColumns_DifferentOffsets()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名");              // offset = 0
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數").WithRowOffset(1);   // offset = 1
            mapping.Map(x => x.Note).ToColumn(ExcelCol.C).WithTitle("備註").WithRowOffset(2);    // offset = 2

            var testData = new List<TestData> { new TestData { Name = "Alice", Score = 95, Note = "Good" } };

            // Act
            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Test");
            
            // 標題行
            Assert.Equal("姓名", sheet.GetRow(0)?.GetCell(0)?.StringCellValue);  // A1
            Assert.Equal("分數", sheet.GetRow(1)?.GetCell(1)?.StringCellValue);  // B2
            Assert.Equal("備註", sheet.GetRow(2)?.GetCell(2)?.StringCellValue);  // C3
            
            // 資料行
            Assert.Equal("Alice", sheet.GetRow(1)?.GetCell(0)?.StringCellValue); // A2
            Assert.Equal(95.0, sheet.GetRow(2)?.GetCell(1)?.NumericCellValue);   // B3
            Assert.Equal("Good", sheet.GetRow(3)?.GetCell(2)?.StringCellValue);  // C4
        }

        [Fact]
        public void WithRowOffset_ZeroOffset_ShouldWorkNormally()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名").WithRowOffset(0);
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數").WithRowOffset(0);

            var testData = new List<TestData> { new TestData { Name = "Alice", Score = 95 } };

            // Act
            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Test");
            
            // 相同列
            Assert.Equal("姓名", sheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("分數", sheet.GetRow(0)?.GetCell(1)?.StringCellValue);
            Assert.Equal("Alice", sheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal(95.0, sheet.GetRow(1)?.GetCell(1)?.NumericCellValue);
        }

        [Fact]
        public void WithStartRow_CombinedWithRowOffset()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數").WithRowOffset(1);
            mapping.WithStartRow(3); // 從第 3 列開始

            var testData = new List<TestData> { new TestData { Name = "Alice", Score = 95 } };

            // Act
            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Test");
            
            // StartRow=3 (0-based = 2) + RowOffset
            Assert.Equal("姓名", sheet.GetRow(2)?.GetCell(0)?.StringCellValue);  // A3 (startRow=3)
            Assert.Equal("分數", sheet.GetRow(3)?.GetCell(1)?.StringCellValue);  // B4 (startRow=3 + offset=1)
            
            // 資料
            Assert.Equal("Alice", sheet.GetRow(3)?.GetCell(0)?.StringCellValue); // A4
            Assert.Equal(95.0, sheet.GetRow(4)?.GetCell(1)?.NumericCellValue);   // B5
        }
    }
}
