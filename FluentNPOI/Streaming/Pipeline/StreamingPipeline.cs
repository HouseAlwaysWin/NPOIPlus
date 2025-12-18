using System;
using System.Collections.Generic;
using FluentNPOI.Streaming.Abstractions;

namespace FluentNPOI.Streaming.Pipeline
{
    /// <summary>
    /// 串流處理管線，提供 Fluent 鏈式 API
    /// </summary>
    /// <typeparam name="T">目標型別</typeparam>
    public class StreamingPipeline<T> where T : new()
    {
        private readonly IStreamingReader _reader;
        private readonly IRowMapper<T> _mapper;
        private int _skipRows;
        private Func<IStreamingRow, bool> _filter;

        internal StreamingPipeline(IStreamingReader reader, IRowMapper<T> mapper)
        {
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
            _mapper = mapper ?? throw new ArgumentNullException(nameof(mapper));
        }

        /// <summary>
        /// 跳過指定行數 (例如 Header)
        /// </summary>
        public StreamingPipeline<T> Skip(int rowCount)
        {
            _skipRows = rowCount;
            return this;
        }

        /// <summary>
        /// 跳過 Header (等同 Skip(1))
        /// </summary>
        public StreamingPipeline<T> SkipHeader()
        {
            return Skip(1);
        }

        /// <summary>
        /// 篩選行
        /// </summary>
        public StreamingPipeline<T> Where(Func<IStreamingRow, bool> predicate)
        {
            _filter = predicate;
            return this;
        }

        /// <summary>
        /// 執行管線並回傳結果 (延遲執行)
        /// </summary>
        public IEnumerable<T> ToEnumerable()
        {
            int skipped = 0;

            foreach (var row in _reader.ReadRows())
            {
                // 跳過指定行數
                if (skipped < _skipRows)
                {
                    skipped++;
                    continue;
                }

                // 套用篩選
                if (_filter != null && !_filter(row))
                    continue;

                // 對應並回傳
                yield return _mapper.Map(row);
            }
        }

        /// <summary>
        /// 執行管線並回傳 List
        /// </summary>
        public List<T> ToList()
        {
            var result = new List<T>();
            foreach (var item in ToEnumerable())
            {
                result.Add(item);
            }
            return result;
        }
    }

    /// <summary>
    /// 管線建構器
    /// </summary>
    public static class StreamingPipelineBuilder
    {
        /// <summary>
        /// 從 Reader 建立管線
        /// </summary>
        public static StreamingPipeline<T> CreatePipeline<T>(
            IStreamingReader reader,
            IRowMapper<T> mapper) where T : new()
        {
            return new StreamingPipeline<T>(reader, mapper);
        }
    }
}
