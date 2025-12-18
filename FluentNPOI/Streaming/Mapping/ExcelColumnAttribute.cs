using System;

namespace FluentNPOI.Streaming.Mapping
{
    /// <summary>
    /// Excel 欄位對應屬性 (選用)
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// 欄位索引 (0-based)
        /// </summary>
        public int Index { get; set; } = -1;

        /// <summary>
        /// 欄位名稱 (Header 對應)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 標題 (寫入時使用)
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 格式
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 以索引建立
        /// </summary>
        public ExcelColumnAttribute(int index)
        {
            Index = index;
        }

        /// <summary>
        /// 以名稱建立
        /// </summary>
        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }

        /// <summary>
        /// 預設建構子
        /// </summary>
        public ExcelColumnAttribute()
        {
        }
    }
}
