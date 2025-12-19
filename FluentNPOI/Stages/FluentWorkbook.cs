using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using FluentNPOI.Base;
using System;
using System.Collections.Generic;
using System.IO;
using FluentNPOI.Models;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// Workbook operation class
    /// </summary>
    public class FluentWorkbook : FluentWorkbookBase, IDisposable
    {
        private ISheet _currentSheet;

        /// <summary>
        /// Initialize FluentWorkbook instance
        /// </summary>
        /// <param name="workbook">NPOI Workbook object</param>
        public FluentWorkbook(IWorkbook workbook)
            : base(workbook, new Dictionary<string, ICellStyle>())
        {
        }

        /// <summary>
        /// Read Excel file
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook ReadExcelFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("Excel file not found.", filePath);

            // Open with Read mode
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return ReadExcelStream(fs);
            }
        }

        /// <summary>
        /// Read Excel from stream
        /// </summary>
        /// <param name="stream">Excel file stream</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        /// <summary>
        /// Read Excel from stream
        /// </summary>
        /// <param name="stream">Excel file stream</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook ReadExcelStream(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            // Clean up existing workbook if any
            // NOTE: We only close it if we are loading a NEW one into the same wrapper.
            // This prevents resource leaks when reusing FluentWorkbook instance (though typically not recommended).
            _workbook?.Close();

            _currentSheet = null;

            // Use WorkbookFactory to automatically detect format
            _workbook = WorkbookFactory.Create(stream);

            // Select first sheet by default
            if (_workbook.NumberOfSheets > 0)
            {
                _currentSheet = _workbook.GetSheetAt(0);
            }

            return this;
        }

        /// <summary>
        /// Create a Streaming Pipeline Builder (Read-Modify-Write)
        /// </summary>
        /// <typeparam name="T">Data Model Type</typeparam>
        /// <param name="filePath">Input file path</param>
        /// <returns>StreamingBuilder</returns>
        /// <example>
        /// <code>
        /// FluentWorkbook.Stream&lt;MyData&gt;("data.xlsx")
        ///     .SaveAs("output.xlsx");
        /// </code>
        /// </example>
        public static StreamingBuilder<T> Stream<T>(string filePath) where T : new()
        {
            return new StreamingBuilder<T>(filePath);
        }

        /// <summary>
        /// Create a Streaming Pipeline Builder (Read-Modify-Write)
        /// </summary>
        /// <typeparam name="T">Data Model Type</typeparam>
        /// <param name="stream">Input stream</param>
        /// <returns>StreamingBuilder</returns>
        public static StreamingBuilder<T> Stream<T>(System.IO.Stream stream) where T : new()
        {
            return new StreamingBuilder<T>(stream);
        }

        /// <summary>
        /// Copy style from cell in specified sheet
        /// </summary>
        /// <param name="cellStyleKey">Style cache key</param>
        /// <param name="sheet">Source sheet</param>
        /// <param name="col">Column position</param>
        /// <param name="rowIndex">Row index (1-based)</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook CopyStyleFromSheetCell(string cellStyleKey, ISheet sheet, ExcelCol col, int rowIndex)
        {
            ICell cell = sheet.GetExcelCell(col, rowIndex);
            if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(cellStyleKey))
            {
                ICellStyle newCellStyle = _workbook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(cell.CellStyle);
                _cellStylesCached.Add(cellStyleKey, newCellStyle);
            }
            return this;
        }

        /// <summary>
        /// Set global cached cell style (clears all existing styles)
        /// </summary>
        /// <param name="styles">Style configuration action</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            styles(_workbook, newCellStyle);
            _cellStylesCached.Clear();
            _cellStylesCached.Add("global", newCellStyle);
            return this;
        }

        /// <summary>
        /// Set and cache cell style
        /// </summary>
        /// <param name="cellStyleKey">Style cache key</param>
        /// <param name="styles">Style configuration action</param>
        /// <param name="inheritFrom">Optional, inherited parent style key. If specified, copies all properties from parent first, then applies custom changes</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetupCellStyle(string cellStyleKey, Action<IWorkbook, ICellStyle> styles, string inheritFrom = null)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();

            // If inheritance is specified, copy all properties from parent style first
            if (!string.IsNullOrEmpty(inheritFrom) && _cellStylesCached.TryGetValue(inheritFrom, out var parentStyle))
            {
                newCellStyle.CloneStyleFrom(parentStyle);
            }

            // Apply custom changes (will override properties from parent style)
            styles(_workbook, newCellStyle);
            _cellStylesCached.Add(cellStyleKey, newCellStyle);
            return this;
        }


        /// <summary>
        /// Use specified sheet
        /// </summary>
        /// <param name="index">Sheet index (0-based)</param>
        /// <param name="createIfMissing">If true, creates a new sheet if index is out of range (append to end)</param>
        /// <returns>FluentSheet instance</returns>
        public FluentSheet UseSheet(int index, bool createIfMissing = false)
        {
            if (index >= 0 && index < _workbook.NumberOfSheets)
            {
                _currentSheet = _workbook.GetSheetAt(index);
            }
            else if (createIfMissing)
            {
                _currentSheet = _workbook.CreateSheet();
            }
            else
            {
                // If not found and not creating, _currentSheet might be null or keep previous?
                // Standard behavior: return null sheet wrapper or throw?
                // FluentSheet constructor allows null sheet, but methods might fail.
                // Let's try to get it, usually returns null if out of range, but NPOI throws ArgumentException for GetSheetAt if invalid.
                // So we must check range.
                // If we are here, index is invalid and createIfMissing is false.
                _currentSheet = null;
            }

            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// Use specified sheet
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <param name="createIfMissing">Create if missing</param>
        /// <returns></returns>
        public FluentSheet UseSheet(string sheetName, bool createIfMissing = true)
        {
            _currentSheet = _workbook.GetSheet(sheetName);
            if (_currentSheet == null && createIfMissing)
            {
                _currentSheet = _workbook.CreateSheet(sheetName);
            }
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// Use specified sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public FluentSheet UseSheet(ISheet sheet)
        {
            _currentSheet = sheet;
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }


        /// <summary>
        /// Save to file
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SaveToFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            // Ensure directory exists
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(fs);
            }
            return this;
        }

        /// <summary>
        /// Save to stream
        /// </summary>
        /// <param name="stream">Target stream</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SaveToStream(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            _workbook.Write(stream);
            return this;
        }

        /// <summary>
        /// Get all sheet names
        /// </summary>
        /// <returns>List of sheet names</returns>
        public List<string> GetSheetNames()
        {
            var names = new List<string>();
            for (int i = 0; i < _workbook.NumberOfSheets; i++)
            {
                names.Add(_workbook.GetSheetName(i));
            }
            return names;
        }

        /// <summary>
        /// Save current sheet as HTML file
        /// </summary>
        /// <param name="filePath">Output path</param>
        /// <param name="fullHtml">Generate full HTML (html/body tags)</param>
        /// <returns>FluentWorkbook instance</returns>
        public FluentWorkbook SaveAsHtml(string filePath, bool fullHtml = true)
        {
            if (_currentSheet == null) throw new InvalidOperationException("No active sheet selected.");

            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var html = FluentNPOI.Html.HtmlConverter.ConvertSheetToHtml(_currentSheet, fullHtml);
            System.IO.File.WriteAllText(filePath, html, System.Text.Encoding.UTF8);
            return this;
        }

        /// <summary>
        /// Export current sheet to HTML string
        /// </summary>
        /// <param name="fullHtml">Generate full HTML</param>
        /// <returns>HTML string</returns>
        public string ToHtmlString(bool fullHtml = true)
        {
            if (_currentSheet == null) throw new InvalidOperationException("No active sheet selected.");
            return FluentNPOI.Html.HtmlConverter.ConvertSheetToHtml(_currentSheet, fullHtml);
        }

        /// <summary>
        /// Export current sheet to PDF and save to file
        /// </summary>
        public FluentWorkbook SaveAsPdf(string filePath)
        {
            if (_currentSheet == null) throw new InvalidOperationException("No active sheet selected.");

            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            FluentNPOI.Pdf.PdfConverter.ConvertSheetToPdf(_currentSheet, _workbook, filePath);
            return this;
        }

        /// <summary>
        /// Export current sheet to PDF bytes
        /// </summary>
        public byte[] ToPdfBytes()
        {
            if (_currentSheet == null) throw new InvalidOperationException("No active sheet selected.");
            return FluentNPOI.Pdf.PdfConverter.ConvertSheetToPdfBytes(_currentSheet, _workbook);
        }

        /// <summary>
        /// Get sheet count
        /// </summary>
        public int SheetCount => _workbook.NumberOfSheets;

        /// <summary>
        /// Delete sheet by name
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook DeleteSheet(string sheetName)
        {
            int index = _workbook.GetSheetIndex(sheetName);
            if (index >= 0)
            {
                _workbook.RemoveSheetAt(index);
            }
            return this;
        }

        /// <summary>
        /// Delete sheet by index
        /// </summary>
        /// <param name="index">Sheet index (0-based)</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook DeleteSheetAt(int index)
        {
            if (index >= 0 && index < _workbook.NumberOfSheets)
            {
                _workbook.RemoveSheetAt(index);
            }
            return this;
        }

        /// <summary>
        /// Clone sheet
        /// </summary>
        /// <param name="sourceSheetName">Source sheet name</param>
        /// <param name="newSheetName">New sheet name</param>
        /// <returns>New FluentSheet instance</returns>
        public FluentSheet CloneSheet(string sourceSheetName, string newSheetName)
        {
            int sourceIndex = _workbook.GetSheetIndex(sourceSheetName);
            if (sourceIndex < 0)
                throw new ArgumentException($"Sheet '{sourceSheetName}' not found.", nameof(sourceSheetName));

            ISheet clonedSheet = _workbook.CloneSheet(sourceIndex);
            int clonedIndex = _workbook.GetSheetIndex(clonedSheet);
            _workbook.SetSheetName(clonedIndex, newSheetName);
            _currentSheet = clonedSheet;
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// Rename sheet
        /// </summary>
        /// <param name="oldName">Old name</param>
        /// <param name="newName">New name</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook RenameSheet(string oldName, string newName)
        {
            int index = _workbook.GetSheetIndex(oldName);
            if (index >= 0)
            {
                _workbook.SetSheetName(index, newName);
            }
            return this;
        }

        /// <summary>
        /// Set active sheet (sheet shown when opening Excel)
        /// </summary>
        /// <param name="index">Sheet index (0-based)</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetActiveSheet(int index)
        {
            if (index >= 0 && index < _workbook.NumberOfSheets)
            {
                _workbook.SetActiveSheet(index);
            }
            return this;
        }

        /// <summary>
        /// Set active sheet (sheet shown when opening Excel)
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetActiveSheet(string sheetName)
        {
            int index = _workbook.GetSheetIndex(sheetName);
            if (index >= 0)
            {
                _workbook.SetActiveSheet(index);
            }
            return this;
        }

        /// <summary>
        /// Close and release workbook resources
        /// </summary>
        public void Close()
        {
            _workbook?.Close();
        }

        /// <summary>
        /// Dispose workbook resources
        /// </summary>
        public void Dispose()
        {
            Close();
        }
    }
}

