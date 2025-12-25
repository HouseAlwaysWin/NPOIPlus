using System;
using System.IO;
using Xunit;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload;
using FluentNPOI.HotReload.HotReload;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI.HotReload Phase 2: Hot Reload Infrastructure
    /// </summary>
    public class HotReloadTests
    {
        #region HotReloadHandler Tests

        [Fact]
        public void HotReloadHandler_IsActive_ShouldBeFalseWhenNoSubscribers()
        {
            // Assert
            Assert.False(HotReloadHandler.IsActive);
        }

        [Fact]
        public void HotReloadHandler_RefreshRequested_ShouldBeInvoked()
        {
            // Arrange
            bool wasInvoked = false;
            Type[]? receivedTypes = null;

            void Handler(Type[]? types)
            {
                wasInvoked = true;
                receivedTypes = types;
            }

            HotReloadHandler.RefreshRequested += Handler;

            try
            {
                // Act
                HotReloadHandler.TriggerRefresh();

                // Assert
                Assert.True(wasInvoked);
                Assert.Null(receivedTypes); // TriggerRefresh passes null
            }
            finally
            {
                HotReloadHandler.RefreshRequested -= Handler;
            }
        }

        [Fact]
        public void HotReloadHandler_IsActive_ShouldBeTrueWithSubscribers()
        {
            // Arrange
            void Handler(Type[]? _) { }
            HotReloadHandler.RefreshRequested += Handler;

            try
            {
                // Assert
                Assert.True(HotReloadHandler.IsActive);
            }
            finally
            {
                HotReloadHandler.RefreshRequested -= Handler;
            }
        }

        #endregion

        #region FluentHotReloadSession Tests

        [Fact]
        public void FluentHotReloadSession_Constructor_ShouldStoreOutputPath()
        {
            // Arrange & Act
            using var session = new FluentHotReloadSession("test/output.xlsx", wb => { });

            // Assert
            Assert.Equal("test/output.xlsx", session.OutputPath);
        }

        [Fact]
        public void FluentHotReloadSession_Refresh_ShouldIncrementRefreshCount()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });

                // Act
                session.Refresh();
                session.Refresh();

                // Assert
                Assert.Equal(2, session.RefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_Refresh_ShouldCreateExcelFile()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(FluentNPOI.Models.ExcelCol.A, 1).SetValue("Test");
                });

                // Act
                session.Refresh();

                // Assert
                Assert.True(File.Exists(tempPath));
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_RefreshCompleted_ShouldBeInvoked()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            int? completedRefreshCount = null;

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });
                session.RefreshCompleted += (count) => completedRefreshCount = count;

                // Act
                session.Refresh();

                // Assert
                Assert.Equal(1, completedRefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_RefreshError_ShouldBeInvokedOnFailure()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            Exception? capturedError = null;

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => throw new InvalidOperationException("Test error"));
                session.RefreshError += (ex) => capturedError = ex;

                // Act
                session.Refresh();

                // Assert
                Assert.NotNull(capturedError);
                Assert.IsType<InvalidOperationException>(capturedError);
                Assert.Equal("Test error", capturedError.Message);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_Start_ShouldPerformInitialRefresh()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });

                // Act
                session.Start();
                session.Stop();

                // Assert
                Assert.Equal(1, session.RefreshCount);
                Assert.True(File.Exists(tempPath));
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        #endregion

        #region RunFromTemplate Tests

        [Fact]
        public void FluentHotReloadSession_WithFactory_ShouldLoadFromTemplate()
        {
            // Arrange
            var templatePath = Path.Combine(Path.GetTempPath(), $"template_{Guid.NewGuid()}.xlsx");
            var outputPath = Path.Combine(Path.GetTempPath(), $"output_{Guid.NewGuid()}.xlsx");

            try
            {
                // 1. Create a template file
                using (var fs = File.Create(templatePath))
                {
                    var wb = new XSSFWorkbook();
                    wb.CreateSheet("TemplateSheet").CreateRow(0).CreateCell(0).SetCellValue("Original");
                    wb.Write(fs);
                }

                // 2. Create session with factory that loads the template
                using var session = FluentLivePreview.CreateSession(outputPath, wb =>
                {
                    // Modify the loaded template
                    wb.UseSheet("TemplateSheet")
                      .SetCellPosition(FluentNPOI.Models.ExcelCol.B, 1)
                      .SetValue("Modified");
                },
                workbookFactory: () =>
                {
                    using var file = File.OpenRead(templatePath);
                    return WorkbookFactory.Create(file);
                });

                // Act
                session.Refresh();

                // Assert
                Assert.True(File.Exists(outputPath));

                using (var fs = File.OpenRead(outputPath))
                {
                    var resultWb = new XSSFWorkbook(fs);
                    var sheet = resultWb.GetSheet("TemplateSheet");
                    Assert.NotNull(sheet);
                    Assert.Equal("Original", sheet.GetRow(0).GetCell(0).StringCellValue);
                    Assert.Equal("Modified", sheet.GetRow(0).GetCell(1).StringCellValue);
                }
            }
            finally
            {
                if (File.Exists(templatePath)) File.Delete(templatePath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_WithStreamFactory_ShouldSupportMultipleRefreshes()
        {
            // This test simulates the logic inside RunFromTemplate(Stream) where we cache bytes

            // Arrange
            var templatePath = Path.Combine(Path.GetTempPath(), $"stream_template_{Guid.NewGuid()}.xlsx");
            var outputPath = Path.Combine(Path.GetTempPath(), $"stream_output_{Guid.NewGuid()}.xlsx");

            try
            {
                // 1. Create a template
                using (var fs = File.Create(templatePath))
                {
                    var wb = new XSSFWorkbook();
                    wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0).SetCellValue("StreamOriginal");
                    wb.Write(fs);
                }

                // 2. Read to memory (simulating RunFromTemplate logic)
                byte[] templateBytes;
                using (var fs = File.OpenRead(templatePath))
                using (var ms = new MemoryStream())
                {
                    fs.CopyTo(ms);
                    templateBytes = ms.ToArray();
                }

                // 3. Create session with factory using cached bytes
                using var session = FluentLivePreview.CreateSession(outputPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(FluentNPOI.Models.ExcelCol.A, 2).SetValue("Pass " + wb.GetWorkbook().GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
                },
                workbookFactory: () =>
                {
                    return WorkbookFactory.Create(new MemoryStream(templateBytes));
                });

                // Act - Refresh 1
                session.Refresh();

                // Assert 1
                using (var fs = File.OpenRead(outputPath))
                {
                    var wb = new XSSFWorkbook(fs);
                    Assert.Equal("Pass StreamOriginal", wb.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue);
                }

                // Act - Refresh 2 (Verify stream can be reused/recreated)
                session.Refresh();

                // Assert 2
                using (var fs = File.OpenRead(outputPath))
                {
                    var wb = new XSSFWorkbook(fs);
                    Assert.Equal("Pass StreamOriginal", wb.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue);
                }
            }
            finally
            {
                if (File.Exists(templatePath)) File.Delete(templatePath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        #endregion

        #region Integration Tests

        [Fact]
        public void FluentHotReloadSession_WithHotReloadHandler_ShouldRespondToEvents()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });
                session.Start();

                var initialCount = session.RefreshCount;

                // Act - simulate hot reload
                HotReloadHandler.TriggerRefresh();

                // Assert
                Assert.Equal(initialCount + 1, session.RefreshCount);

                session.Stop();
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_UpdateLogic_ShouldReflectInOutput()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            var currentValue = "Initial";

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(FluentNPOI.Models.ExcelCol.A, 1).SetValue(currentValue);
                });

                // Act - First refresh
                session.Refresh();

                // Change value in memory (simulating code change logic)
                currentValue = "Updated";
                session.Refresh();

                // Assert - Read the file and verify updated value
                using var fs = File.OpenRead(tempPath);
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheetAt(0);
                var cell = sheet.GetRow(0)?.GetCell(0);

                Assert.Equal("Updated", cell?.StringCellValue);
                Assert.Equal(2, session.RefreshCount);
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
