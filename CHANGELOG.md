# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Initial release of NPOIPlus
- Fluent API for Excel operations
- Support for writing data to Excel
- Support for reading data from Excel
- Style caching mechanism to avoid 64k style limit
- Dynamic styling based on cell data
- Support for `List<T>` and `DataTable` data sources
- Formula support (read and write)
- Extension methods for common operations
- Comprehensive unit tests
- Console application examples

### Features

#### Core Classes
- `FluentWorkbook` - Main entry point for fluent API
- `FluentSheet` - Sheet-level operations
- `FluentCell` - Cell-level operations
- `FluentTable<T>` - Table data binding
- `FluentTableHeader<T>` - Table header configuration
- `FluentTableCell<T>` - Table cell configuration

#### Base Classes
- `FluentWorkbookBase` - Common workbook operations
- `FluentCellBase` - Cell value and style operations
- `FluentSheetBase` - Sheet base operations
- `FluentTableBase<T>` - Table base operations

#### Models
- `ExcelColumns` - Column enumeration (A-ZZ)
- `CellStyleConfig` - Style configuration with caching support
- `TableCellSet` - Cell configuration model
- `TableCellParams` - Cell value parameters
- `TableCellStyleParams` - Style parameters

#### Helpers
- `FluentMemoryStream` - Memory stream wrapper
- `NPOIPlusExtensions` - Rich extension methods

#### Key Features
- **Style Management**
  - Global style configuration
  - Named style registry
  - Dynamic style with data-based conditions
  - Style caching to prevent Excel limit issues

- **Data Operations**
  - Read/Write single cells
  - Batch table data binding
  - Support for multiple data types (string, number, date, boolean)
  - Formula support

- **Excel Operations**
  - Create and modify workbooks
  - Multiple sheet management
  - Column width adjustment
  - Cell merging
  - Read existing Excel files

### Examples
- Basic read/write operations
- Table data binding with `List<T>`
- DataTable support
- Dynamic styling examples
- Conditional formatting
- Multi-sheet operations

### Documentation
- Comprehensive README (Chinese and English)
- API reference
- Code examples
- Best practices guide

## [1.0.0] - 2024-12-01

### Added
- Initial stable release

---

## Version History

### Planned Features (Future Releases)

#### v1.1.0
- [ ] Support for .xls (HSSF) format
- [ ] Image insertion support
- [ ] Chart creation support
- [ ] Data validation support
- [ ] Conditional formatting presets

#### v1.2.0
- [ ] Template support
- [ ] Batch file processing
- [ ] Performance optimizations
- [ ] Async/await support

#### v2.0.0
- [ ] Complete API redesign
- [ ] Plugin system
- [ ] Custom formula functions
- [ ] Advanced formatting options

---

## Migration Guide

### From Direct NPOI Usage

Before (Direct NPOI):
```csharp
var workbook = new XSSFWorkbook();
var sheet = workbook.CreateSheet("Sheet1");
var row = sheet.CreateRow(0);
var cell = row.CreateCell(0);
cell.SetCellValue("Hello");

var style = workbook.CreateCellStyle();
style.Alignment = HorizontalAlignment.Center;
cell.CellStyle = style;

using (var fs = new FileStream("output.xlsx", FileMode.Create))
{
    workbook.Write(fs);
}
```

After (NPOIPlus):
```csharp
var fluent = new FluentWorkbook(new XSSFWorkbook());

fluent.SetupGlobalCachedCellStyles((wb, style) =>
{
    style.SetAligment(HorizontalAlignment.Center);
});

fluent.UseSheet("Sheet1")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("Hello");

fluent.SaveToPath("output.xlsx");
```

---

## Breaking Changes

None in v1.0.0 (Initial Release)

---

## Known Issues

1. Large file performance - Processing files with 100k+ rows may require optimization
2. Style limit - While we implement caching, developers must use keys properly to avoid limits

---

## Contributors

Thanks to all contributors who helped build NPOIPlus!

---

## Support

- Report bugs: [GitHub Issues](../../issues)
- Ask questions: [GitHub Discussions](../../discussions)
- Email: [martinwang7963@gmail.com]


