using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using FluentNPOI.Models;

namespace FluentNPOI.Streaming.Mapping
{
    /// <summary>
    /// DataTable 專用的 Mapping 設定
    /// </summary>
    public class DataTableMapping
    {
        private readonly List<ColumnMapping> _mappings = new List<ColumnMapping>();

        /// <summary>
        /// 開始設定欄位對應
        /// </summary>
        public DataTableColumnBuilder Map(string columnName)
        {
            var mapping = new ColumnMapping { ColumnName = columnName };
            _mappings.Add(mapping);
            return new DataTableColumnBuilder(this, mapping);
        }

        /// <summary>
        /// 取得所有 Mapping 設定
        /// </summary>
        public IReadOnlyList<ColumnMapping> GetMappings() => _mappings;

        /// <summary>
        /// 從 DataTable 自動產生 Mapping
        /// </summary>
        public static DataTableMapping FromDataTable(DataTable dt)
        {
            var mapping = new DataTableMapping();
            var col = ExcelCol.A;

            foreach (DataColumn column in dt.Columns)
            {
                mapping.Map(column.ColumnName)
                    .ToColumn(col)
                    .WithTitle(column.ColumnName);
                col++;
            }

            return mapping;
        }

        /// <summary>
        /// 從 DataRow 取值
        /// </summary>
        public object GetValue(ColumnMapping map, DataRow row)
        {
            if (map.ValueFunc != null)
            {
                return map.ValueFunc(row);
            }

            if (!string.IsNullOrEmpty(map.ColumnName) && row.Table.Columns.Contains(map.ColumnName))
            {
                return row[map.ColumnName];
            }

            return null;
        }
    }

    /// <summary>
    /// DataTable 欄位設定建構器
    /// </summary>
    public class DataTableColumnBuilder
    {
        private readonly DataTableMapping _parent;
        private readonly ColumnMapping _mapping;

        internal DataTableColumnBuilder(DataTableMapping parent, ColumnMapping mapping)
        {
            _parent = parent;
            _mapping = mapping;
        }

        /// <summary>
        /// 設定對應的 Excel 欄位
        /// </summary>
        public DataTableColumnBuilder ToColumn(ExcelCol column)
        {
            _mapping.ColumnIndex = column;
            return this;
        }

        /// <summary>
        /// 設定標題
        /// </summary>
        public DataTableColumnBuilder WithTitle(string title)
        {
            _mapping.Title = title;
            return this;
        }

        /// <summary>
        /// 設定自訂值計算
        /// </summary>
        public DataTableColumnBuilder WithValue(Func<DataRow, object> valueFunc)
        {
            _mapping.ValueFunc = obj => valueFunc((DataRow)obj);
            return this;
        }

        /// <summary>
        /// 設定公式
        /// </summary>
        public DataTableColumnBuilder WithFormula(Func<int, int, string> formulaFunc)
        {
            _mapping.FormulaFunc = formulaFunc;
            return this;
        }

        /// <summary>
        /// 設定資料樣式
        /// </summary>
        public DataTableColumnBuilder WithStyle(string styleKey)
        {
            _mapping.StyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// 設定標題樣式
        /// </summary>
        public DataTableColumnBuilder WithTitleStyle(string styleKey)
        {
            _mapping.TitleStyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// 繼續設定下一個欄位
        /// </summary>
        public DataTableColumnBuilder Map(string columnName)
        {
            return _parent.Map(columnName);
        }
    }
}
