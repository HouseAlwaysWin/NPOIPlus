using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload.Widgets;
using FluentNPOI.HotReload.Context;
using FluentNPOI.HotReload.Styling;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI.HotReload Phase 1.5: Flexbox Layout and Styling
    /// </summary>
    public class FlexboxStylingTests
    {
        #region FluentStyle Tests

        [Fact]
        public void FluentStyle_SetFontColor_ShouldStoreColor()
        {
            // Arrange & Act
            var style = new FluentStyle().SetFontColor(IndexedColors.Red);

            // Assert
            Assert.Equal(IndexedColors.Red.Index, style.FontColor?.Index);
        }

        [Fact]
        public void FluentStyle_SetBackgroundColor_ShouldStoreColor()
        {
            // Arrange & Act
            var style = new FluentStyle().SetBackgroundColor(IndexedColors.Yellow);

            // Assert
            Assert.Equal(IndexedColors.Yellow.Index, style.BackgroundColor?.Index);
        }

        [Fact]
        public void FluentStyle_ChainedMethods_ShouldSetAllProperties()
        {
            // Arrange & Act
            var style = new FluentStyle()
                .SetFontColor(IndexedColors.Blue)
                .SetBackgroundColor(IndexedColors.LightGreen)
                .SetBold()
                .SetItalic()
                .SetBorder(BorderStyle.Thin)
                .SetFontSize(14)
                .SetFontName("Arial");

            // Assert
            Assert.Equal(IndexedColors.Blue.Index, style.FontColor?.Index);
            Assert.Equal(IndexedColors.LightGreen.Index, style.BackgroundColor?.Index);
            Assert.True(style.IsBold);
            Assert.True(style.IsItalic);
            Assert.Equal(BorderStyle.Thin, style.Border);
            Assert.Equal(14, style.FontSize);
            Assert.Equal("Arial", style.FontName);
        }

        [Fact]
        public void FluentStyle_GetCacheKey_ShouldGenerateUniqueKeys()
        {
            // Arrange
            var style1 = new FluentStyle().SetBold();
            var style2 = new FluentStyle().SetBold().SetItalic();
            var style3 = new FluentStyle().SetBold();

            // Act
            var key1 = style1.GetCacheKey();
            var key2 = style2.GetCacheKey();
            var key3 = style3.GetCacheKey();

            // Assert
            Assert.NotEqual(key1, key2);
            Assert.Equal(key1, key3);
        }

        [Fact]
        public void FluentStyle_HasAnyStyle_ShouldReturnTrueWhenSet()
        {
            // Arrange
            var emptyStyle = new FluentStyle();
            var boldStyle = new FluentStyle().SetBold();

            // Assert
            Assert.False(emptyStyle.HasAnyStyle());
            Assert.True(boldStyle.HasAnyStyle());
        }

        #endregion

        #region StyleManager Tests

        [Fact]
        public void StyleManager_GetOrCreateStyle_ShouldCacheStyles()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var styleManager = new StyleManager();
            styleManager.Reset(workbook);

            var style = new FluentStyle().SetBold();

            // Act
            var cellStyle1 = styleManager.GetOrCreateStyle(style);
            var cellStyle2 = styleManager.GetOrCreateStyle(style);

            // Assert
            Assert.Same(cellStyle1, cellStyle2);
            Assert.Equal(1, styleManager.CachedStyleCount);
        }

        [Fact]
        public void StyleManager_GetOrCreateStyle_ShouldCacheDifferentStyles()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var styleManager = new StyleManager();
            styleManager.Reset(workbook);

            var boldStyle = new FluentStyle().SetBold();
            var italicStyle = new FluentStyle().SetItalic();

            // Act
            var cellStyle1 = styleManager.GetOrCreateStyle(boldStyle);
            var cellStyle2 = styleManager.GetOrCreateStyle(italicStyle);

            // Assert
            Assert.NotSame(cellStyle1, cellStyle2);
            Assert.Equal(2, styleManager.CachedStyleCount);
        }

        [Fact]
        public void StyleManager_Reset_ShouldClearCache()
        {
            // Arrange
            var workbook1 = new XSSFWorkbook();
            var workbook2 = new XSSFWorkbook();
            var styleManager = new StyleManager();

            styleManager.Reset(workbook1);
            styleManager.GetOrCreateStyle(new FluentStyle().SetBold());
            Assert.Equal(1, styleManager.CachedStyleCount);

            // Act
            styleManager.Reset(workbook2);

            // Assert
            Assert.Equal(0, styleManager.CachedStyleCount);
        }

        [Fact]
        public void StyleManager_ApplyStyle_ShouldSetCellStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);

            var styleManager = new StyleManager();
            styleManager.Reset(workbook);

            var style = new FluentStyle()
                .SetBold()
                .SetBackgroundColor(IndexedColors.Yellow);

            // Act
            styleManager.ApplyStyle(cell, style);

            // Assert
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
            var font = workbook.GetFontAt(cell.CellStyle.FontIndex);
            Assert.True(font.IsBold);
        }

        #endregion

        #region FlexibleRow Tests

        [Fact]
        public void FlexibleRow_Build_ShouldArrangeChildrenHorizontally()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            var row = new FlexibleRow(
                new Label("Col A"),
                new Label("Col B"),
                new Label("Col C")
            );

            // Act
            row.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Col A", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Col B", npoiSheet.GetRow(0)?.GetCell(1)?.StringCellValue);
            Assert.Equal("Col C", npoiSheet.GetRow(0)?.GetCell(2)?.StringCellValue);
        }

        [Fact]
        public void FlexibleRow_Build_ShouldCalculateColumnWidths()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Create a row with different weights
            var label1 = new Label("Wide Column");
            label1.Weight = 2;
            var label2 = new Label("Narrow");
            label2.Weight = 1;

            var row = new FlexibleRow(label1, label2).SetTotalWidth(90);

            // Act
            row.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            // Weight 2 out of 3 = 60 chars, Weight 1 out of 3 = 30 chars
            var col0Width = npoiSheet.GetColumnWidth(0) / 256; // Convert to characters
            var col1Width = npoiSheet.GetColumnWidth(1) / 256;

            Assert.True(col0Width > col1Width, $"Col0 ({col0Width}) should be wider than Col1 ({col1Width})");
        }

        [Fact]
        public void FlexibleRow_WithDefaultWeight_ShouldUseEqualWidths()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            var row = new FlexibleRow(
                new Label("A"),
                new Label("B"),
                new Label("C")
            ).SetTotalWidth(90);

            // Act
            row.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var col0Width = npoiSheet.GetColumnWidth(0);
            var col1Width = npoiSheet.GetColumnWidth(1);
            var col2Width = npoiSheet.GetColumnWidth(2);

            Assert.Equal(col0Width, col1Width);
            Assert.Equal(col1Width, col2Width);
        }

        #endregion

        #region Widget Style Integration Tests

        [Fact]
        public void ExcelWidget_WithStyle_ShouldStoreStyle()
        {
            // Arrange
            var style = new FluentStyle().SetBold();
            var label = new Label("Test");

            // Act
            label.WithStyle(style);

            // Assert
            Assert.Same(style, label.Style);
        }

        [Fact]
        public void ExcelWidget_WithWeight_ShouldStoreWeight()
        {
            // Arrange
            var label = new Label("Test");

            // Act
            label.WithWeight(3);

            // Assert
            Assert.Equal(3, label.Weight);
        }

        [Fact]
        public void ExcelContext_ApplyStyle_WithStyleManager_ShouldWork()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var styleManager = new StyleManager();
            styleManager.Reset(workbook);

            var ctx = new ExcelContext(sheet, styleManager);
            var style = new FluentStyle().SetBackgroundColor(IndexedColors.LightBlue);

            // Act
            ctx.SetValue("Styled Cell");
            ctx.ApplyStyle(style);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal(IndexedColors.LightBlue.Index, cell.CellStyle.FillForegroundColor);
        }

        #endregion

        #region Complex Layout Tests

        [Fact]
        public void ComplexLayout_WithFlexibleRowAndStyles_ShouldRenderCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var styleManager = new StyleManager();
            styleManager.Reset(workbook);
            var ctx = new ExcelContext(sheet, styleManager);

            // Create a styled table header with flexible widths
            var headerStyle = new FluentStyle()
                .SetBold()
                .SetBackgroundColor(IndexedColors.Grey25Percent);

            var header1 = new Label("產品名稱");
            header1.Weight = 2;
            header1.Style = headerStyle;

            var header2 = new Label("單價");
            header2.Weight = 1;
            header2.Style = headerStyle;

            var header3 = new Label("數量");
            header3.Weight = 1;
            header3.Style = headerStyle;

            var layout = new Column(
                new FlexibleRow(header1, header2, header3).SetTotalWidth(80),
                new FlexibleRow(
                    new Label("蘋果").WithWeight(2),
                    new Label("$30").WithWeight(1),
                    new Label("10").WithWeight(1)
                ).SetTotalWidth(80)
            );

            // Act
            layout.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // Header row
            Assert.Equal("產品名稱", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("單價", npoiSheet.GetRow(0)?.GetCell(1)?.StringCellValue);
            Assert.Equal("數量", npoiSheet.GetRow(0)?.GetCell(2)?.StringCellValue);

            // Data row
            Assert.Equal("蘋果", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal("$30", npoiSheet.GetRow(1)?.GetCell(1)?.StringCellValue);
            Assert.Equal("10", npoiSheet.GetRow(1)?.GetCell(2)?.StringCellValue);

            // Check column widths (first column should be wider)
            var col0Width = npoiSheet.GetColumnWidth(0);
            var col1Width = npoiSheet.GetColumnWidth(1);
            Assert.True(col0Width > col1Width);
        }

        #endregion
    }
}
