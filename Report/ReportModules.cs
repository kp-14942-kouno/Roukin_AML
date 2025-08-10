using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.Report
{
    public static class ReportModules
    {
        /// <summary>
        /// 指定のレコード数でリストをチャンクに分割する
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="chunkSize"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static List<List<T>> ChunkBy<T>(List<T> source, int chunkSize)
        {
            if (chunkSize <= 0)
                throw new ArgumentException("Chunk size must be greater than zero.", nameof(chunkSize));
            return source.Select((x, i) => new { Index = i, Value = x })
                         .GroupBy(x => x.Index / chunkSize)
                         .Select(g => g.Select(x => x.Value).ToList())
                         .ToList();
        }
    }
}
