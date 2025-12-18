using System;
using FluentNPOI.Streaming.Abstractions;

namespace FluentNPOI.Streaming.Rows
{
    /// <summary>
    /// 串流行資料實作
    /// </summary>
    public class StreamingRow : IStreamingRow
    {
        private readonly object[] _values;

        /// <summary>
        /// 建立串流行
        /// </summary>
        /// <param name="rowIndex">行號</param>
        /// <param name="values">欄位值陣列</param>
        public StreamingRow(int rowIndex, object[] values)
        {
            RowIndex = rowIndex;
            _values = values ?? Array.Empty<object>();
        }

        /// <inheritdoc/>
        public int RowIndex { get; }

        /// <inheritdoc/>
        public int ColumnCount => _values.Length;

        /// <inheritdoc/>
        public object GetValue(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= _values.Length)
                return null;
            return _values[columnIndex];
        }

        /// <inheritdoc/>
        public T GetValue<T>(int columnIndex)
        {
            var value = GetValue(columnIndex);
            if (value == null)
                return default;

            var targetType = typeof(T);
            var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (value is T typedValue)
                    return typedValue;

                if (value is IConvertible)
                    return (T)Convert.ChangeType(value, underlyingType);

                return default;
            }
            catch
            {
                return default;
            }
        }

        /// <inheritdoc/>
        public bool IsNull(int columnIndex)
        {
            var value = GetValue(columnIndex);
            if (value == null)
                return true;
            if (value is string str)
                return string.IsNullOrWhiteSpace(str);
            return false;
        }
    }
}
