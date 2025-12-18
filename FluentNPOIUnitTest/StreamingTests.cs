using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentNPOI.Streaming;
using FluentNPOI.Streaming.Abstractions;
using FluentNPOI.Streaming.Mapping;
using FluentNPOI.Streaming.Pipeline;
using FluentNPOI.Streaming.Rows;
using NPOI.XSSF.UserModel;
using Xunit;

namespace FluentNPOIUnitTest
{
    public class StreamingTests
    {
        #region Test Models

        public class Person
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
            public string Email { get; set; } = string.Empty;
        }

        #endregion

        #region StreamingRow Tests

        [Fact]
        public void StreamingRow_GetValue_ReturnsCorrectValue()
        {
            var values = new object[] { "John", 30, "john@example.com" };
            var row = new StreamingRow(0, values);

            Assert.Equal("John", row.GetValue(0));
            Assert.Equal(30, row.GetValue(1));
            Assert.Equal("john@example.com", row.GetValue(2));
        }

        [Fact]
        public void StreamingRow_GetValueT_ConvertsType()
        {
            var values = new object[] { "John", 30.5, "100" };
            var row = new StreamingRow(0, values);

            Assert.Equal("John", row.GetValue<string>(0));
            Assert.Equal(30, row.GetValue<int>(1));
        }

        [Fact]
        public void StreamingRow_GetValue_OutOfRange_ReturnsNull()
        {
            var values = new object[] { "John" };
            var row = new StreamingRow(0, values);

            Assert.Null(row.GetValue(10));
        }

        [Fact]
        public void StreamingRow_IsNull_DetectsEmptyValues()
        {
            var values = new object?[] { "John", null, "", "  " };
            var row = new StreamingRow(0, values!);

            Assert.False(row.IsNull(0));
            Assert.True(row.IsNull(1));
            Assert.True(row.IsNull(2));
            Assert.True(row.IsNull(3));
        }

        #endregion

        #region FluentMapping Tests

        [Fact]
        public void FluentMapping_Map_SetsColumnIndex()
        {
            var mapping = new FluentMapping<Person>();
            mapping.Map(x => x.Name).ToColumn(0);
            mapping.Map(x => x.Age).ToColumn(1);
            mapping.Map(x => x.Email).ToColumn(2);

            var mappings = mapping.GetMappings();
            Assert.Equal(3, mappings.Count);
        }

        [Fact]
        public void FluentMapping_Map_MapsRowToDto()
        {
            var mapping = new FluentMapping<Person>();
            mapping.Map(x => x.Name).ToColumn(0);
            mapping.Map(x => x.Age).ToColumn(1);
            mapping.Map(x => x.Email).ToColumn(2);

            var values = new object[] { "Alice", 25, "alice@example.com" };
            var row = new StreamingRow(0, values);

            var person = mapping.Map(row);

            Assert.Equal("Alice", person.Name);
            Assert.Equal(25, person.Age);
            Assert.Equal("alice@example.com", person.Email);
        }

        [Fact]
        public void FluentMapping_Map_HandlesNullValues()
        {
            var mapping = new FluentMapping<Person>();
            mapping.Map(x => x.Name).ToColumn(0);
            mapping.Map(x => x.Age).ToColumn(1);

            var values = new object?[] { null, null };
            var row = new StreamingRow(0, values!);

            var person = mapping.Map(row);

            // Name defaults to empty string in Person class, and null input doesn't override it
            Assert.True(string.IsNullOrEmpty(person.Name));
            Assert.Equal(0, person.Age);
        }

        #endregion

        #region StreamingPipeline Tests

        [Fact]
        public void StreamingPipeline_Skip_SkipsRows()
        {
            var reader = new MockStreamingReader(new[]
            {
                new object[] { "Header1", "Header2" },
                new object[] { "Value1", 1 },
                new object[] { "Value2", 2 }
            });

            var mapping = new FluentMapping<Person>();
            mapping.Map(x => x.Name).ToColumn(0);
            mapping.Map(x => x.Age).ToColumn(1);

            var pipeline = StreamingPipelineBuilder.CreatePipeline<Person>(reader, mapping)
                .Skip(1);

            var results = pipeline.ToList();

            Assert.Equal(2, results.Count);
            Assert.Equal("Value1", results[0].Name);
        }

        [Fact]
        public void StreamingPipeline_Where_FiltersRows()
        {
            var reader = new MockStreamingReader(new[]
            {
                new object[] { "Alice", 25 },
                new object[] { "Bob", 17 },
                new object[] { "Charlie", 30 }
            });

            var mapping = new FluentMapping<Person>();
            mapping.Map(x => x.Name).ToColumn(0);
            mapping.Map(x => x.Age).ToColumn(1);

            var pipeline = StreamingPipelineBuilder.CreatePipeline<Person>(reader, mapping)
                .Where(row => row.GetValue<int>(1) >= 18);

            var results = pipeline.ToList();

            Assert.Equal(2, results.Count);
            Assert.Equal("Alice", results[0].Name);
            Assert.Equal("Charlie", results[1].Name);
        }

        #endregion

        #region Integration Tests

        [Fact]
        public void FluentExcelReader_Read_WithMapping_ReadsExcelFile()
        {
            var testFilePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            try
            {
                CreateTestExcelFile(testFilePath, new[]
                {
                    new[] { "Name", "Age", "Email" },
                    new[] { "Alice", "25", "alice@example.com" },
                    new[] { "Bob", "30", "bob@example.com" }
                });

                var mapping = new FluentMapping<Person>();
                mapping.Map(x => x.Name).ToColumn(0);
                mapping.Map(x => x.Age).ToColumn(1);
                mapping.Map(x => x.Email).ToColumn(2);

                var results = FluentExcelReader.Read<Person>(testFilePath, mapping).ToList();

                Assert.Equal(2, results.Count);
                Assert.Equal("Alice", results[0].Name);
                Assert.Equal(25, results[0].Age);
                Assert.Equal("Bob", results[1].Name);
            }
            finally
            {
                if (File.Exists(testFilePath))
                    File.Delete(testFilePath);
            }
        }

        #endregion

        #region Helpers

        private void CreateTestExcelFile(string filePath, string[][] data)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");

            for (int i = 0; i < data.Length; i++)
            {
                var row = sheet.CreateRow(i);
                for (int j = 0; j < data[i].Length; j++)
                {
                    var cell = row.CreateCell(j);
                    if (int.TryParse(data[i][j], out int intValue))
                        cell.SetCellValue(intValue);
                    else
                        cell.SetCellValue(data[i][j]);
                }
            }

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        #endregion
    }

    #region Mock Classes

    public class MockStreamingReader : IStreamingReader
    {
        private readonly object[][] _data;
        private int _currentIndex;

        public MockStreamingReader(object[][] data)
        {
            _data = data;
        }

        public IReadOnlyList<string> SheetNames => new[] { "Sheet1" };

        public bool SelectSheet(string sheetName) => true;
        public bool SelectSheet(int sheetIndex) => true;

        public IEnumerable<IStreamingRow> ReadRows()
        {
            for (int i = _currentIndex; i < _data.Length; i++)
            {
                yield return new StreamingRow(i, _data[i]);
            }
        }

        public string[] ReadHeader()
        {
            if (_data.Length == 0)
                return Array.Empty<string>();

            _currentIndex = 1;
            return _data[0].Select(x => x?.ToString() ?? "").ToArray();
        }

        public void Dispose() { }
    }

    #endregion
}
