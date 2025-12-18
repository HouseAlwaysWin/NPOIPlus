using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using FluentNPOI.Models;
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

                var value = row.GetValue((int)mapping.ColumnIndex.Value);
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
        /// 設定對應的 Excel 欄位
        /// </summary>
        public FluentColumnBuilder<T> ToColumn(ExcelCol column)
        {
            _mapping.ColumnIndex = column;
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
        /// 設定自訂值計算 (用於寫入時)
        /// </summary>
        public FluentColumnBuilder<T> WithValue(Func<T, object> valueFunc)
        {
            _mapping.ValueFunc = obj => valueFunc((T)obj);
            return this;
        }

        /// <summary>
        /// 設定公式 (用於寫入時)
        /// </summary>
        /// <param name="formulaFunc">公式函數，參數為 (row, col)，回傳公式字串 (不含 =)</param>
        public FluentColumnBuilder<T> WithFormula(Func<int, int, string> formulaFunc)
        {
            _mapping.FormulaFunc = formulaFunc;
            return this;
        }

        /// <summary>
        /// 設定資料儲存格樣式 (使用樣式 Key)
        /// </summary>
        public FluentColumnBuilder<T> WithStyle(string styleKey)
        {
            _mapping.StyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// 設定標題儲存格樣式 (使用樣式 Key)
        /// </summary>
        public FluentColumnBuilder<T> WithTitleStyle(string styleKey)
        {
            _mapping.TitleStyleKey = styleKey;
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
        public ExcelCol? ColumnIndex { get; set; }
        public string Title { get; set; }
        public string Format { get; set; }

        // 寫入時使用
        public Func<object, object> ValueFunc { get; set; }
        public Func<int, int, string> FormulaFunc { get; set; }
        public string StyleKey { get; set; }
        public string TitleStyleKey { get; set; }

        // 欄位名稱 (for DataTable)
        public string ColumnName { get; set; }
    }
}
