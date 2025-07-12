using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.RoukinClass
{
    public static class RoukinModules
    {
        /// <summary>
        /// 半角カナ小文字wを半角カナ大文字に変換
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string ReplaceSmallKanaWithLargeKana(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            const string smallKana = "ｧｨｩｪｫｬｭｮｯ";
            const string largeKana = "ｱｲｳｴｵﾔﾕﾖﾂ";

            return string.Create(value.Length, value, (span, src) =>
            {
                var small = smallKana.AsSpan();
                var large = largeKana.AsSpan();

                for (int i = 0; i < src.Length; i++)
                {
                    var idx = small.IndexOf(src[i]);
                    span[i] = idx >= 0 ? large[idx] : src[i];
                }
            });
        }

        /// <summary>
        /// 囲い文字を付与（ダブルクォーテーション）
        /// </summary>
        /// <param name="value"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public static string SetDc(string value)
        {
            // valueをダブルクォーテーションで囲む
            return $"\"{value}\"";
        }
    }
}
