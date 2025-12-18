using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader;
using FluentNPOI.Streaming.Abstractions;
using FluentNPOI.Streaming.Rows;

namespace FluentNPOI.Streaming.Readers
{
    /// <summary>
    /// ExcelDataReader 串流讀取器實作
    /// </summary>
    public class ExcelDataReaderAdapter : IStreamingReader
    {
        private readonly Stream _stream;
        private readonly IExcelDataReader _reader;
        private readonly bool _ownsStream;
        private readonly List<string> _sheetNames;
        private bool _disposed;

        static ExcelDataReaderAdapter()
        {
            // 註冊編碼提供者 (ExcelDataReader 需要)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// 從檔案路徑建立讀取器
        /// </summary>
        public ExcelDataReaderAdapter(string filePath)
            : this(File.OpenRead(filePath), true)
        {
        }

        /// <summary>
        /// 從 Stream 建立讀取器
        /// </summary>
        public ExcelDataReaderAdapter(Stream stream, bool ownsStream = false)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _ownsStream = ownsStream;
            _reader = ExcelReaderFactory.CreateReader(_stream);
            _sheetNames = new List<string>();

            // 取得所有工作表名稱
            do
            {
                _sheetNames.Add(_reader.Name);
            } while (_reader.NextResult());

            // 重置到第一個工作表
            _reader.Reset();
        }

        /// <inheritdoc/>
        public IReadOnlyList<string> SheetNames => _sheetNames;

        /// <inheritdoc/>
        public bool SelectSheet(string sheetName)
        {
            var index = _sheetNames.IndexOf(sheetName);
            return index >= 0 && SelectSheet(index);
        }

        /// <inheritdoc/>
        public bool SelectSheet(int sheetIndex)
        {
            if (sheetIndex < 0 || sheetIndex >= _sheetNames.Count)
                return false;

            _reader.Reset();
            for (int i = 0; i < sheetIndex; i++)
            {
                _reader.NextResult();
            }
            return true;
        }

        /// <inheritdoc/>
        public IEnumerable<IStreamingRow> ReadRows()
        {
            int rowIndex = 0;
            while (_reader.Read())
            {
                var values = new object[_reader.FieldCount];
                for (int i = 0; i < _reader.FieldCount; i++)
                {
                    values[i] = _reader.GetValue(i);
                }
                yield return new StreamingRow(rowIndex++, values);
            }
        }

        /// <inheritdoc/>
        public string[] ReadHeader()
        {
            if (!_reader.Read())
                return Array.Empty<string>();

            var headers = new string[_reader.FieldCount];
            for (int i = 0; i < _reader.FieldCount; i++)
            {
                headers[i] = _reader.GetValue(i)?.ToString() ?? string.Empty;
            }
            return headers;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;
            _reader?.Dispose();
            if (_ownsStream)
                _stream?.Dispose();
        }
    }
}
