using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Helpers;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 單元格操作類
    /// </summary>
    public class FluentCell : FluentCellBase
    {
        private ICell _cell;

        public FluentCell(IWorkbook workbook, ISheet sheet, ICell cell, Dictionary<string, ICellStyle> cellStylesCached = null)
            : base(workbook, sheet, cellStylesCached ?? new Dictionary<string, ICellStyle>())
        {
            _cell = cell;
        }

        public FluentCell SetValue<T>(T value)
        {
            if (_cell == null) return this;
            SetCellValue(_cell, value);
            return this;
        }

        public FluentCell SetFormulaValue(object value)
        {
            if (_cell == null) return this;
            SetFormulaValue(_cell, value);
            return this;
        }

        public FluentCell SetCellStyle(string cellStyleKey)
        {
            if (_cell == null) return this;

            if (!string.IsNullOrWhiteSpace(cellStyleKey) && _cellStylesCached.ContainsKey(cellStyleKey))
            {
                _cell.CellStyle = _cellStylesCached[cellStyleKey];
            }
            return this;
        }

        public FluentCell SetCellStyle(Func<TableCellStyleParams, CellStyleConfig> cellStyleAction)
        {
            if (_cell == null) return this;

            var cellStyleParams = new TableCellStyleParams
            {
                Workbook = _workbook,
                ColNum = (ExcelColumns)_cell.ColumnIndex,
                RowNum = _cell.RowIndex,
                RowItem = null
            };

            // ✅ 先調用函數獲取樣式配置
            var config = cellStyleAction(cellStyleParams);

            if (!string.IsNullOrWhiteSpace(config.Key))
            {
                // ✅ 先檢查緩存
                if (!_cellStylesCached.ContainsKey(config.Key))
                {
                    // ✅ 只在不存在時才創建新樣式
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    config.StyleSetter(newCellStyle);
                    _cellStylesCached.Add(config.Key, newCellStyle);
                }
                _cell.CellStyle = _cellStylesCached[config.Key];
            }
            else
            {
                // 如果沒有返回 key，創建臨時樣式（不緩存）
                ICellStyle tempStyle = _workbook.CreateCellStyle();
                config.StyleSetter(tempStyle);
                _cell.CellStyle = tempStyle;
            }

            return this;
        }

        public FluentCell SetCellType(CellType cellType)
        {
            if (_cell == null) return this;
            _cell.SetCellType(cellType);
            return this;
        }

        /// <summary>
        /// 獲取當前單元格的值
        /// </summary>
        /// <returns>單元格的值（根據類型返回 bool, DateTime, double, string 或 null）</returns>
        public object GetValue()
        {
            return GetCellValue(_cell);
        }

        /// <summary>
        /// 獲取當前單元格的值並轉換為指定類型
        /// </summary>
        /// <typeparam name="T">目標類型</typeparam>
        /// <returns>轉換後的值</returns>
        public T GetValue<T>()
        {
            return GetCellValue<T>(_cell);
        }

        /// <summary>
        /// 獲取當前單元格的公式字符串（如果是公式單元格）
        /// </summary>
        /// <returns>公式字符串（不含 '=' 前綴），如果不是公式則返回 null</returns>
        public string GetFormula()
        {
            return GetCellFormulaValue(_cell);
        }

        /// <summary>
        /// 獲取當前單元格對象
        /// </summary>
        /// <returns>NPOI ICell 對象</returns>
        public ICell GetCell()
        {
            return _cell;
        }
    }
}

