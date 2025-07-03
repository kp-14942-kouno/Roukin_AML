using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate
{
    public class RoukinModule
    {
        private static readonly Encoding JisEncoding = Encoding.GetEncoding("Shift_JIS");

        private static readonly byte[] Allowed1ByteCodes = { 0x00, 0x20, 0x28, 0x29, 0x5C, 0xA2, 0xA3 };

        /// <summary>
        /// 1バイト文字の許可リスト
        /// </summary>
        /// <param name="b"></param>
        /// <returns></returns>

        private static bool IsAllowed1Byte(byte b)
        {
            return
                (b >= 0x2C && b <= 0x39) ||  // , ～ 9
                (b >= 0x41 && b <= 0x5A) || // A ～ Z
                (b >= 0xA6 && b <= 0xDF) || // 半角カナ
                Array.Exists(Allowed1ByteCodes, x => x == b);
        }

        /// <summary>
        /// 2バイト文字の許可リスト
        /// </summary>
        /// <param name="b1"></param>
        /// <param name="b2"></param>
        /// <returns></returns>
        private static bool IsAllowed2Byte(byte b1, byte b2)
        {
            int jisCode = (b1 << 8) | b2;
            return
                (jisCode >= 0x8140 && jisCode <= 0x84BE) ||  // JIS 第1水準
                (jisCode >= 0x889F && jisCode <= 0xEAA2);    // JIS 第2水準
        }

        /// <summary>
        /// 1バイト文字の不正な文字を取得
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string? GetInvalid1ByteChars(string input)
        {
            List<char> invalids = new();

            foreach (char c in input)
            {
                byte[] encoded = JisEncoding.GetBytes(c.ToString());

                // エスケープシーケンス除去（先頭がESCでなければ実体）
                int i = 0;
                while (i < encoded.Length && encoded[i] == 0x1B)
                {
                    // ESC シーケンススキップ
                    while (i < encoded.Length && encoded[i] != 'B' && encoded[i] != 'J' && encoded[i] != 'I') i++;
                    i++;
                }

                // 残ったバイトが1つのみなら1バイトとみなす
                if (i == encoded.Length - 1 && IsAllowed1Byte(encoded[i]) == false)
                {
                    invalids.Add(c);
                }
                else if (encoded.Length > i + 1)
                {
                    // 2バイト文字は対象外（ここではNGとする）
                    invalids.Add(c);
                }
            }

            return invalids.Count > 0 ? string.Join(",", invalids) : null;
        }

        /// <summary>
        /// 2バイト文字の不正な文字を取得
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string? GetInvalid2ByteChars(string input)
        {
            List<char> invalids = new();

            foreach (char c in input)
            {
                byte[] encoded = JisEncoding.GetBytes(c.ToString());

                int i = 0;
                while (i < encoded.Length && encoded[i] == 0x1B)
                {
                    while (i < encoded.Length && encoded[i] != 'B' && encoded[i] != 'J' && encoded[i] != '$') i++;
                    i++;
                }

                if (encoded.Length - i != 2 || !IsAllowed2Byte(encoded[i], encoded[i + 1]))
                {
                    invalids.Add(c);
                }
            }

            return invalids.Count > 0 ? string.Join(",", invalids) : null;
        }

        /// <summary>
        /// 混在文字の不正な文字を取得
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string? GetInvalidMixedChars(string input)
        {
            List<char> invalids = new();

            foreach (char c in input)
            {
                byte[] encoded = JisEncoding.GetBytes(c.ToString());

                int i = 0;
                while (i < encoded.Length && encoded[i] == 0x1B)
                {
                    while (i < encoded.Length && encoded[i] != 'B' && encoded[i] != 'J' && encoded[i] != '$') i++;
                    i++;
                }

                int payloadLen = encoded.Length - i;

                if (payloadLen == 1)
                {
                    if (!IsAllowed1Byte(encoded[i]))
                        invalids.Add(c);
                }
                else if (payloadLen == 2)
                {
                    if (!IsAllowed2Byte(encoded[i], encoded[i + 1]))
                        invalids.Add(c);
                }
                else
                {
                    invalids.Add(c); // エスケープが不正 or 不明な形式
                }
            }

            return invalids.Count > 0 ? string.Join(",", invalids) : null;
        }

    }
}
