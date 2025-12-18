namespace FluentNPOI.Streaming.Abstractions
{
    /// <summary>
    /// 串流讀取的單行資料介面
    /// </summary>
    public interface IStreamingRow
    {
        /// <summary>
        /// 行號 (0-based)
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// 欄位數量
        /// </summary>
        int ColumnCount { get; }

        /// <summary>
        /// 取得指定欄位的值
        /// </summary>
        /// <param name="columnIndex">欄位索引 (0-based)</param>
        /// <returns>欄位值，可能為 null</returns>
        object GetValue(int columnIndex);

        /// <summary>
        /// 取得指定欄位的值並轉換為指定型別
        /// </summary>
        /// <typeparam name="T">目標型別</typeparam>
        /// <param name="columnIndex">欄位索引 (0-based)</param>
        /// <returns>轉換後的值</returns>
        T GetValue<T>(int columnIndex);

        /// <summary>
        /// 檢查指定欄位是否為空
        /// </summary>
        bool IsNull(int columnIndex);
    }
}
