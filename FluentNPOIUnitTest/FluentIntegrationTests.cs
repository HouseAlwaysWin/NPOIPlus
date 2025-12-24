using System;
using System.IO;
using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload;
using FluentNPOI.HotReload.Widgets;
using FluentNPOI.HotReload.Context;
using NPOI.XSSF.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI and HotReload integration.
    /// Verifies that existing FluentNPOI code can work with the hot reload system.
    /// </summary>
    public class FluentIntegrationTests
    {
        #region FluentBridge Tests

        [Fact]
        public void FluentBridge_ShouldExecuteSheetAction()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            var bridge = new FluentBridge(s =>
            {
                s.SetCellPosition(ExcelCol.A, 1).SetValue("Hello");
                s.SetCellPosition(ExcelCol.B, 1).SetValue("World");
            });

            // Act
            bridge.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Hello", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("World", npoiSheet.GetRow(0)?.GetCell(1)?.StringCellValue);
        }

        [Fact]
        public void FluentBridge_WithPosition_ShouldReceiveCurrentPosition()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet, startRow: 5, startCol: ExcelCol.C);

            int receivedRow = 0;
            int receivedCol = 0;

            var bridge = new FluentBridge((s, row, col) =>
            {
                receivedRow = row;
                receivedCol = col;
            });

            // Act
            bridge.Build(ctx);

            // Assert
            Assert.Equal(5, receivedRow);
            Assert.Equal((int)ExcelCol.C, receivedCol);
        }

        [Fact]
        public void FluentBridge_InsideColumn_ShouldWorkWithWidgets()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Mix widgets with FluentBridge
            var layout = new Column(
                new Header("Report Title"),
                new FluentBridge(s =>
                {
                    // Traditional FluentNPOI code
                    s.SetCellPosition(ExcelCol.A, 2).SetValue("From FluentBridge");
                }),
                new Label("End of Report")
            );

            // Act
            layout.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Report Title", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("From FluentBridge", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal("End of Report", npoiSheet.GetRow(2)?.GetCell(0)?.StringCellValue);
        }

        #endregion

        #region FluentSheet.BuildWidget Extension Tests

        [Fact]
        public void BuildWidget_ShouldBuildWidgetIntoSheet()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            var widget = new Column(
                new Label("A1"),
                new Label("A2"),
                new Label("A3")
            );

            // Act
            sheet.BuildWidget(widget);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("A1", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("A2", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal("A3", npoiSheet.GetRow(2)?.GetCell(0)?.StringCellValue);
        }

        [Fact]
        public void BuildWidget_WithStartPosition_ShouldBuildAtOffset()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            var widget = new Label("Offset Cell");

            // Act
            sheet.BuildWidget(widget, startRow: 5, startCol: ExcelCol.D);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(4)?.GetCell(3); // Row 5 = index 4, Col D = index 3
            Assert.NotNull(cell);
            Assert.Equal("Offset Cell", cell.StringCellValue);
        }

        #endregion

        #region FluentHotReloadSession Tests

        [Fact]
        public void FluentHotReloadSession_ShouldBuildWithFluentCode()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                Action<FluentWorkbook> buildAction = wb =>
                {
                    wb.UseSheet("Sheet1")
                        .SetCellPosition(ExcelCol.A, 1)
                        .SetValue("FluentNPOI Works!")
                        .SetCellPosition(ExcelCol.B, 1)
                        .SetValue(123);
                };

                using var session = new FluentHotReloadSession(tempPath, buildAction);

                // Act
                session.Refresh();

                // Assert
                Assert.True(File.Exists(tempPath));

                using var fs = File.OpenRead(tempPath);
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheet("Sheet1");
                Assert.Equal("FluentNPOI Works!", sheet.GetRow(0)?.GetCell(0)?.StringCellValue);
                Assert.Equal(123.0, sheet.GetRow(0)?.GetCell(1)?.NumericCellValue);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_RefreshCount_ShouldIncrement()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(ExcelCol.A, 1).SetValue("Test");
                });

                // Act
                session.Refresh();
                session.Refresh();
                session.Refresh();

                // Assert
                Assert.Equal(3, session.RefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentLivePreview_CreateSession_ShouldReturnSession()
        {
            // Arrange & Act
            using var session = FluentLivePreview.CreateSession("test.xlsx", wb =>
            {
                wb.UseSheet("Sheet1").SetCellPosition(ExcelCol.A, 1).SetValue("Test");
            });

            // Assert
            Assert.NotNull(session);
            Assert.Equal("test.xlsx", session.OutputPath);
        }

        #endregion

        #region Hybrid Integration Tests

        [Fact]
        public void HybridApproach_WidgetAndFluentCode_ShouldWorkTogether()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                Action<FluentWorkbook> buildAction = wb =>
                {
                    var sheet = wb.UseSheet("Sheet1");

                    // Use FluentNPOI for some parts
                    sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Header from FluentNPOI");

                    // Use widgets for complex layouts
                    var widget = new Column(
                        new Label("Row 2 from Widget"),
                        new Row(
                            new Label("Cell 1"),
                            new Label("Cell 2"),
                            new Label("Cell 3")
                        )
                    );

                    // Build widget starting at row 2
                    sheet.BuildWidget(widget, startRow: 2);
                };

                using var session = new FluentHotReloadSession(tempPath, buildAction);
                session.Refresh();

                // Assert
                using var fs = File.OpenRead(tempPath);
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheet("Sheet1");

                Assert.Equal("Header from FluentNPOI", sheet.GetRow(0)?.GetCell(0)?.StringCellValue);
                Assert.Equal("Row 2 from Widget", sheet.GetRow(1)?.GetCell(0)?.StringCellValue);
                Assert.Equal("Cell 1", sheet.GetRow(2)?.GetCell(0)?.StringCellValue);
                Assert.Equal("Cell 2", sheet.GetRow(2)?.GetCell(1)?.StringCellValue);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        #endregion
    }
}
