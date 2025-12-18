using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using FluentNPOI.Streaming.Abstractions;

namespace FluentNPOI.Streaming.Mapping
{
    /// <summary>
    /// Fluent Mapping 配置，用於定義 Excel 欄位與屬性的對應
    /// </summary>
    /// <typeparam name="T">目標型別</typeparam>
    public class FluentMapping<T> : IRowMapper<T> where T : new()
    {
        private readonly List<ColumnMapping> _mappings = new List<ColumnMapping>();

        /// <summary>
        /// 開始設定屬性對應
        /// </summary>
        public FluentColumnBuilder<T> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            var propertyInfo = GetPropertyInfo(expression);
            var mapping = new ColumnMapping { Property = propertyInfo };
            _mappings.Add(mapping);
            return new FluentColumnBuilder<T>(this, mapping);
        }

        /// <summary>
        /// 取得所有 Mapping 設定
        /// </summary>
        public IReadOnlyList<ColumnMapping> GetMappings() => _mappings;

        /// <summary>
        /// 將串流行轉換為 DTO
        /// </summary>
        public T Map(IStreamingRow row)
        {
            var instance = new T();

            foreach (var mapping in _mappings)
            {
                if (!mapping.ColumnIndex.HasValue)
                    continue;

                var value = row.GetValue(mapping.ColumnIndex.Value);
                if (value == null)
                    continue;

                try
                {
                    var targetType = mapping.Property.PropertyType;
                    var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

                    object convertedValue;
                    if (value.GetType() == underlyingType)
                    {
                        convertedValue = value;
                    }
                    else if (value is IConvertible)
                    {
                        convertedValue = Convert.ChangeType(value, underlyingType);
                    }
                    else
                    {
                        continue;
                    }

                    mapping.Property.SetValue(instance, convertedValue);
                }
                catch
                {
                    // 轉換失敗，跳過
                }
            }

            return instance;
        }

        private PropertyInfo GetPropertyInfo<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            if (expression.Body is MemberExpression member)
                return member.Member as PropertyInfo;
            if (expression.Body is UnaryExpression unary && unary.Operand is MemberExpression unaryMember)
                return unaryMember.Member as PropertyInfo;
            throw new ArgumentException("Expression must be a property selector");
        }
    }

    /// <summary>
    /// Fluent 欄位設定建構器
    /// </summary>
    public class FluentColumnBuilder<T> where T : new()
    {
        private readonly FluentMapping<T> _parent;
        private readonly ColumnMapping _mapping;

        internal FluentColumnBuilder(FluentMapping<T> parent, ColumnMapping mapping)
        {
            _parent = parent;
            _mapping = mapping;
        }

        /// <summary>
        /// 設定對應的 Excel 欄位索引
        /// </summary>
        public FluentColumnBuilder<T> ToColumn(int columnIndex)
        {
            _mapping.ColumnIndex = columnIndex;
            return this;
        }

        /// <summary>
        /// 設定標題 (用於寫入時)
        /// </summary>
        public FluentColumnBuilder<T> WithTitle(string title)
        {
            _mapping.Title = title;
            return this;
        }

        /// <summary>
        /// 設定格式 (用於寫入時)
        /// </summary>
        public FluentColumnBuilder<T> WithFormat(string format)
        {
            _mapping.Format = format;
            return this;
        }

        /// <summary>
        /// 繼續設定下一個屬性
        /// </summary>
        public FluentColumnBuilder<T> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            return _parent.Map(expression);
        }
    }

    /// <summary>
    /// 欄位對應設定
    /// </summary>
    public class ColumnMapping
    {
        public PropertyInfo Property { get; set; }
        public int? ColumnIndex { get; set; }
        public string Title { get; set; }
        public string Format { get; set; }
    }
}
