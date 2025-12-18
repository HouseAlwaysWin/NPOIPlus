using System;
using System.Collections.Generic;

namespace FluentNPOI.Streaming.Abstractions
{
    /// <summary>
    /// 串流 Excel 讀取器介面
    /// </summary>
    public interface IStreamingReader : IDisposable
    {
        /// <summary>
        /// 取得所有工作表名稱
        /// </summary>
        IReadOnlyList<string> SheetNames { get; }

        /// <summary>
        /// 選擇工作表 (依名稱)
        /// </summary>
        /// <param name="sheetName">工作表名稱</param>
        /// <returns>是否成功</returns>
        bool SelectSheet(string sheetName);

        /// <summary>
        /// 選擇工作表 (依索引)
        /// </summary>
        /// <param name="sheetIndex">工作表索引 (0-based)</param>
        /// <returns>是否成功</returns>
        bool SelectSheet(int sheetIndex);

        /// <summary>
        /// 串流讀取所有行
        /// </summary>
        /// <returns>行的列舉器</returns>
        IEnumerable<IStreamingRow> ReadRows();

        /// <summary>
        /// 讀取 Header 行 (第一行)
        /// </summary>
        /// <returns>Header 名稱陣列</returns>
        string[] ReadHeader();
    }
}
