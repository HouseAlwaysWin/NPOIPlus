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
        /// 預設起始列（1-based），預設為 1
        /// </summary>
        public int StartRow { get; private set; } = 1;

        /// <summary>
        /// 設定表格預設起始列（1-based）
        /// </summary>
        /// <param name="row">起始列（1-based，第 1 列 = 1）</param>
        public FluentMapping<T> WithStartRow(int row)
        {
            StartRow = row < 1 ? 1 : row;
            return this;
        }

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
        /// 設定静態標題 (用於寫入時)
        /// </summary>
        public FluentColumnBuilder<T> WithTitle(string title)
        {
            _mapping.Title = title;
            return this;
        }

        /// <summary>
        /// 設定動態標題（完整版）(用於寫入時)
        /// </summary>
        /// <param name="titleFunc">標題函數，參數為 (row, col)，回傳標題字串</param>
        public FluentColumnBuilder<T> WithTitle(Func<int, ExcelCol, string> titleFunc)
        {
            _mapping.TitleFunc = titleFunc;
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
        /// 設定靜態值（所有列都使用相同值）(用於寫入時)
        /// </summary>
        /// <param name="value">靜態值</param>
        public FluentColumnBuilder<T> WithValue(object value)
        {
            _mapping.ValueFunc = (obj, row, col) => value;
            return this;
        }

        /// <summary>
        /// 設定自訂值計算（簡單版，僅需資料物件）(用於寫入時)
        /// </summary>
        /// <param name="valueFunc">值計算函數，僅接收資料物件參數</param>
        public FluentColumnBuilder<T> WithValue(Func<T, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((T)obj);
            return this;
        }

        /// <summary>
        /// 設定自訂值計算（完整版）(用於寫入時)
        /// </summary>
        /// <param name="valueFunc">值計算函數，參數為 (item, row, col)，row 為 Excel 1-based 行號，col 為 ExcelCol 欄位</param>
        public FluentColumnBuilder<T> WithValue(Func<T, int, ExcelCol, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((T)obj, row, col);
            return this;
        }

        /// <summary>
        /// 設定静態公式（簡單版）(用於寫入時)
        /// </summary>
        /// <param name="formula">公式字串（不含 =）</param>
        public FluentColumnBuilder<T> WithFormula(string formula)
        {
            _mapping.FormulaFunc = (row, col) => formula;
            return this;
        }

        /// <summary>
        /// 設定動態公式（完整版）(用於寫入時)
        /// </summary>
        /// <param name="formulaFunc">公式函數，參數為 (row, col)，回傳公式字串 (不含 =)</param>
        public FluentColumnBuilder<T> WithFormula(Func<int, ExcelCol, string> formulaFunc)
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
            _mapping.TitleStyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// 從指定儲存格複製標題樣式
        /// </summary>
        /// <param name="row">來源列（1-based，第 1 列 = 1）</param>
        /// <param name="col">來源欄</param>
        public FluentColumnBuilder<T> WithTitleStyleFrom(int row, ExcelCol col)
        {
            _mapping.TitleStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// 從指定儲存格複製資料樣式
        /// </summary>
        /// <param name="row">來源列（1-based，第 1 列 = 1）</param>
        /// <param name="col">來源欄</param>
        public FluentColumnBuilder<T> WithStyleFrom(int row, ExcelCol col)
        {
            _mapping.DataStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// 設定儲存格類型
        /// </summary>
        public FluentColumnBuilder<T> WithCellType(NPOI.SS.UserModel.CellType cellType)
        {
            _mapping.CellType = cellType;
            return this;
        }

        /// <summary>
        /// 設定動態樣式 (根據資料決定樣式 Key)
        /// </summary>
        /// <param name="styleFunc">接收資料物件，回傳樣式 Key</param>
        public FluentColumnBuilder<T> WithDynamicStyle(Func<T, string> styleFunc)
        {
            _mapping.DynamicStyleFunc = obj => styleFunc((T)obj);
            return this;
        }

        /// <summary>
        /// 設定欄位列偏移（此欄位相對於表格起始列往下偏移）
        /// </summary>
        /// <param name="offset">偏移量（正數表示往下偏移，預設 0）</param>
        public FluentColumnBuilder<T> WithRowOffset(int offset)
        {
            _mapping.RowOffset = offset;
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
        public Func<int, ExcelCol, string> TitleFunc { get; set; }
        public string Format { get; set; }

        // 寫入時使用
        public Func<object, int, ExcelCol, object> ValueFunc { get; set; }
        public Func<int, ExcelCol, string> FormulaFunc { get; set; }
        public string StyleKey { get; set; }
        public string TitleStyleKey { get; set; }
        public NPOI.SS.UserModel.CellType? CellType { get; set; }
        public Func<object, string> DynamicStyleFunc { get; set; }

        // 欄位名稱 (for DataTable)
        public string ColumnName { get; set; }

        // 樣式參考
        public StyleReference TitleStyleRef { get; set; }
        public StyleReference DataStyleRef { get; set; }

        /// <summary>
        /// 欄位列偏移（預設 0，正數表示往下偏移）
        /// </summary>
        public int RowOffset { get; set; } = 0;
    }

    public class StyleReference
    {
        public int Row { get; set; }
        public ExcelCol Column { get; set; }

        /// <summary>
        /// 從使用者輸入建立 StyleReference（自動將 1-based row 轉換為 0-based）
        /// </summary>
        /// <param name="row">使用者輸入的列號（1-based，第 1 列 = 1）</param>
        /// <param name="col">欄位</param>
        public static StyleReference FromUserInput(int row, ExcelCol col)
        {
            return new StyleReference
            {
                Row = row < 1 ? 0 : row - 1,  // 將 1-based 轉換為 0-based
                Column = col
            };
        }
    }
}
