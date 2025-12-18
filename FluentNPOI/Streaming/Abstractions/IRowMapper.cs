namespace FluentNPOI.Streaming.Abstractions
{
    /// <summary>
    /// 行資料對應到 DTO 的介面
    /// </summary>
    /// <typeparam name="T">目標 DTO 型別</typeparam>
    public interface IRowMapper<T>
    {
        /// <summary>
        /// 將串流行轉換為 DTO
        /// </summary>
        /// <param name="row">串流行資料</param>
        /// <returns>轉換後的 DTO</returns>
        T Map(IStreamingRow row);
    }
}
