using System.Collections.Generic;
using FluentNPOI.Streaming.Abstractions;
using FluentNPOI.Streaming.Mapping;
using FluentNPOI.Streaming.Pipeline;
using FluentNPOI.Streaming.Readers;

namespace FluentNPOI.Streaming
{
    /// <summary>
    /// FluentNPOI 串流讀取入口
    /// </summary>
    public static class FluentExcelReader
    {
        /// <summary>
        /// 使用 Header 自動對應讀取 Excel
        /// </summary>
        /// <typeparam name="T">目標型別</typeparam>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="sheetName">工作表名稱 (選填)</param>
        /// <returns>物件列舉</returns>
        public static IEnumerable<T> Read<T>(string filePath, string sheetName = null) where T : new()
        {
            using (var reader = new ExcelDataReaderAdapter(filePath))
            {
                if (!string.IsNullOrEmpty(sheetName))
                    reader.SelectSheet(sheetName);

                // 讀取 Header 建立自動 Mapping
                var headers = reader.ReadHeader();
                var mapper = CreateAutoMapper<T>(headers);

                foreach (var row in reader.ReadRows())
                {
                    yield return mapper.Map(row);
                }
            }
        }

        /// <summary>
        /// 使用 FluentMapping 讀取 Excel
        /// </summary>
        /// <typeparam name="T">目標型別</typeparam>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="mapping">Mapping 設定</param>
        /// <param name="sheetName">工作表名稱 (選填)</param>
        /// <returns>物件列舉</returns>
        public static IEnumerable<T> Read<T>(string filePath, FluentMapping<T> mapping, string sheetName = null) where T : new()
        {
            using (var reader = new ExcelDataReaderAdapter(filePath))
            {
                if (!string.IsNullOrEmpty(sheetName))
                    reader.SelectSheet(sheetName);

                var pipeline = StreamingPipelineBuilder.CreatePipeline(reader, mapping)
                    .SkipHeader();

                foreach (var item in pipeline.ToEnumerable())
                {
                    yield return item;
                }
            }
        }

        /// <summary>
        /// 建立管線以進行更細緻的控制
        /// </summary>
        public static StreamingPipeline<T> CreatePipeline<T>(string filePath, FluentMapping<T> mapping) where T : new()
        {
            var reader = new ExcelDataReaderAdapter(filePath);
            return StreamingPipelineBuilder.CreatePipeline(reader, mapping);
        }

        private static IRowMapper<T> CreateAutoMapper<T>(string[] headers) where T : new()
        {
            var mapping = new FluentMapping<T>();
            var properties = typeof(T).GetProperties();

            for (int i = 0; i < headers.Length; i++)
            {
                var headerName = headers[i]?.Trim();
                if (string.IsNullOrEmpty(headerName))
                    continue;

                // 找到名稱匹配的屬性
                foreach (var prop in properties)
                {
                    if (string.Equals(prop.Name, headerName, System.StringComparison.OrdinalIgnoreCase))
                    {
                        // 使用反射設定 mapping (簡化版)
                        var columnMappings = mapping.GetMappings();
                        // 這裡需要更完整的實作來動態新增 mapping
                        break;
                    }
                }
            }

            return mapping;
        }
    }
}
