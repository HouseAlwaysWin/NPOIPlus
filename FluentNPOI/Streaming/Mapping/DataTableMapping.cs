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
        /// 預設起始列（1-based），預設為 1
        /// </summary>
        public int StartRow { get; private set; } = 1;

        /// <summary>
        /// 設定表格預設起始列（1-based）
        /// </summary>
        /// <param name="row">起始列（1-based，第 1 列 = 1）</param>
        public DataTableMapping WithStartRow(int row)
        {
            StartRow = row < 1 ? 1 : row;
            return this;
        }

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
        /// <param name="map">欄位對應設定</param>
        /// <param name="row">DataRow 資料</param>
        /// <param name="rowIndex">Excel 1-based 行號</param>
        /// <param name="colIndex">ExcelCol 欄位</param>
        public object GetValue(ColumnMapping map, DataRow row, int rowIndex, ExcelCol colIndex)
        {
            if (map.ValueFunc != null)
            {
                return map.ValueFunc(row, rowIndex, colIndex);
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
        /// 設定靜態標題
        /// </summary>
        public DataTableColumnBuilder WithTitle(string title)
        {
            _mapping.Title = title;
            return this;
        }

        /// <summary>
        /// 設定動態標題（完整版）
        /// </summary>
        /// <param name="titleFunc">標題函數，參數為 (row, col)，回傳標題字串</param>
        public DataTableColumnBuilder WithTitle(Func<int, ExcelCol, string> titleFunc)
        {
            _mapping.TitleFunc = titleFunc;
            return this;
        }

        /// <summary>
        /// 設定靜態值（所有列都使用相同值）
        /// </summary>
        /// <param name="value">靜態值</param>
        public DataTableColumnBuilder WithValue(object value)
        {
            _mapping.ValueFunc = (obj, row, col) => value;
            return this;
        }

        /// <summary>
        /// 設定自訂值計算（簡單版，僅需 DataRow）
        /// </summary>
        /// <param name="valueFunc">值計算函數，僅接收 DataRow 參數</param>
        public DataTableColumnBuilder WithValue(Func<DataRow, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((DataRow)obj);
            return this;
        }

        /// <summary>
        /// 設定自訂值計算（完整版）
        /// </summary>
        /// <param name="valueFunc">值計算函數，參數為 (row, excelRow, col)，excelRow 為 Excel 1-based 行號，col 為 ExcelCol 欄位</param>
        public DataTableColumnBuilder WithValue(Func<DataRow, int, ExcelCol, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((DataRow)obj, row, col);
            return this;
        }

        /// <summary>
        /// 設定靜態公式（簡單版）
        /// </summary>
        /// <param name="formula">公式字串（不含 =）</param>
        public DataTableColumnBuilder WithFormula(string formula)
        {
            _mapping.FormulaFunc = (row, col) => formula;
            return this;
        }

        /// <summary>
        /// 設定動態公式（完整版）
        /// </summary>
        /// <param name="formulaFunc">公式函數，參數為 (row, col)，回傳公式字串 (不含 =)</param>
        public DataTableColumnBuilder WithFormula(Func<int, ExcelCol, string> formulaFunc)
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
        /// 設定儲存格類型
        /// </summary>
        public DataTableColumnBuilder WithCellType(NPOI.SS.UserModel.CellType cellType)
        {
            _mapping.CellType = cellType;
            return this;
        }

        /// <summary>
        /// 從指定儲存格複製標題樣式
        /// </summary>
        /// <param name="row">來源列（1-based，第 1 列 = 1）</param>
        /// <param name="col">來源欄</param>
        public DataTableColumnBuilder WithTitleStyleFrom(int row, ExcelCol col)
        {
            _mapping.TitleStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// 從指定儲存格複製資料樣式
        /// </summary>
        /// <param name="row">來源列（1-based，第 1 列 = 1）</param>
        /// <param name="col">來源欄</param>
        public DataTableColumnBuilder WithStyleFrom(int row, ExcelCol col)
        {
            _mapping.DataStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// 設定欄位列偏移（此欄位相對於表格起始列往下偏移）
        /// </summary>
        /// <param name="offset">偏移量（正數表示往下偏移，預設 0）</param>
        public DataTableColumnBuilder WithRowOffset(int offset)
        {
            _mapping.RowOffset = offset;
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
