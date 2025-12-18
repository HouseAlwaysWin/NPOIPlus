using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace FluentNPOIUnitTest
{
    public class ComparisonTests
    {
        private readonly ITestOutputHelper _output;

        public ComparisonTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void Compare_Test2_Files()
        {
            var basePath = @"d:\Projects\FluentNPOI\FluentNPOIConsoleExample\bin\Debug\net6.0\Resources";
            var file1 = Path.Combine(basePath, "Test2_v2.xlsx");
            var file2 = Path.Combine(basePath, "Test2_old.xlsx");

            if (!File.Exists(file1) || !File.Exists(file2))
            {
                _output.WriteLine("Files not found for comparison.");
                return;
            }

            string result = ExcelComparer.Compare(file2, file1, "DataTableExample"); // A=Old, B=New

            _output.WriteLine(result);

            // Fail if there are differences (heuristic: result contains brackets like [CELL ...])
            if (result.Contains("[NAME]") || result.Contains("[CELL ") || result.Contains("[SHEET COUNT") || result.Contains("[ROW "))
            {
                // Assert.Fail is valid in xUnit 2.5+, but for compatibility let's use True(false)
                Assert.True(false, "Differences found:\n" + result);
            }
        }
    }
}
