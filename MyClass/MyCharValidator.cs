using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.MyClass
{
    /// <summary>
    /// 文字の検証クラス
    /// </summary>
    public class MyCharValidator
    {
        private static readonly Encoding JisEncoding = Encoding.GetEncoding("Shift_JIS");

        private List<(int start, int end)> _allowed1ByteRanges = new();
        private List<(int start, int end)> _allowed2ByteRanges = new();
        public List<(int start, int end)> _allowedUniRanges = new();

        /// <summary>
        /// コンストラクタ
        /// SJIS
        /// </summary>
        /// <param name="encoding"></param>
        /// <param name="allowed1Byte"></param>
        /// <param name="allowed2Byte"></param>
        public MyCharValidator(string? allowed1Byte, string? allowed2Byte, string? allowdUnicode)
        {
            _allowed1ByteRanges = ParseRanges(allowed1Byte);
            _allowed2ByteRanges = ParseRanges(allowed2Byte);
            _allowedUniRanges = ParseRanges(allowdUnicode);
        }

        /// <summary>
        /// 範囲文字列を解析して、許可された範囲のリストを生成する
        /// </summary>
        /// <param name="ranges"></param>
        /// <returns></returns>
        private List<(int start, int end)> ParseRanges(string? ranges)
        {
            var result = new List<(int start, int end)>();

            var matches = System.Text.RegularExpressions.Regex.Matches(ranges ?? string.Empty, @"\[\s*(0x)?[0-9A-Fa-f]{1,6}\s*,\s*(0x)?[0-9A-Fa-f]{1,6}\s*\]");

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                string content = match.Value.Trim('[', ']');
                var parts = content.Split(',', StringSplitOptions.TrimEntries);

                if (parts.Length == 2)
                {
                    if (TryParserHex(parts[0], out int start) &&
                        TryParserHex(parts[1], out int end))
                    {
                        result.Add((start, end));
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 16進数の文字列を整数に変換する
        /// </summary>
        /// <param name="hex"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private static bool TryParserHex(string hex, out int value)
        {
            value = 0;
            hex = hex.Trim();
            if (hex.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                hex = hex.Substring(2);
            }
            return int.TryParse(hex, NumberStyles.HexNumber, null, out value);
        }

        /// <summary>
        /// 1バイト文字の許可範囲をチェック
        /// </summary>
        /// <param name="b"></param>
        /// <returns></returns>
        private bool IsAllowed1Byte(byte b)
        {
            if (_allowed1ByteRanges.Count == 0) return true; // 1バイトの許可範囲が指定されていない場合は全て許可 
            foreach (var (start, end) in _allowed1ByteRanges)
            {
                if (b >= start && b <= end)
                {
                    return true;
                }
            }
            return false; // 範囲外の場合は許可しない
        }

        /// <summary>
        /// 2バイト文字の許可範囲をチェック
        /// </summary>
        /// <param name="b1"></param>
        /// <param name="b2"></param>
        /// <returns></returns>
        private bool IsAllowed2Byte(byte b1, byte b2)
        {
            if (_allowed2ByteRanges.Count == 0) return true; // 2バイトの許可範囲が指定されていない場合は全て許可
            int codePoint = (b1 << 8) | b2; // 2バイトを1つのコードポイントに変換
            foreach (var (start, end) in _allowed2ByteRanges)
            {
                if (codePoint >= start && codePoint <= end)
                {
                    return true;
                }
            }
            return false; // 範囲外の場合は許可しない
        }

        /// <summary>
        /// Unicodeの許可範囲をチェック
        /// </summary>
        /// <param name="codePoint"></param>
        /// <returns></returns>
        private bool IsAllowedUnicode(int codePoint)
        {
            if (_allowedUniRanges.Count == 0) return true; // Unicodeの許可範囲が指定されていない場合は全て許可
            foreach (var (start, end) in _allowedUniRanges)
            {
                if (codePoint >= start && codePoint <= end)
                {
                    return true;
                }
            }
            return false; // 範囲外の場合は許可しない
        }

        /// <summary>
        /// 1バイト文字の不正な文字を取得
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public string? GetInvalid1ByteChars(string input)
        {
            List<char> invalids = new();

            foreach (char c in input)
            {
                byte[] encoded = JisEncoding.GetBytes(new[] { c });

                int i = 0;
                while (i < encoded.Length && encoded[i] == 0x1B)
                {
                    while (i < encoded.Length && encoded[i] != 'B' && encoded[i] != 'J' && encoded[i] != 'I') i++;
                    i++;
                }

                int len = encoded.Length - i;

                if (len == 1)
                {
                    if (!IsAllowed1Byte(encoded[i]))
                        invalids.Add(c);
                }
                else
                {
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
        public string? GetInvalid2ByteChars(string input)
        {
            List<char> invalids = new();

            foreach (char c in input)
            {
                byte[] encoded = JisEncoding.GetBytes(new[] { c });

                int i = 0;
                while (i < encoded.Length && encoded[i] == 0x1B)
                {
                    while (i < encoded.Length && encoded[i] != 'B' && encoded[i] != 'J' && encoded[i] != '$') i++;
                    i++;
                }

                int len = encoded.Length - i;

                if (len == 2)
                {
                    if (!IsAllowed2Byte(encoded[i], encoded[i + 1]))
                        invalids.Add(c);
                }
                else
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
        public string? GetInvalidMixedChars(string input)
        {
            List<char> invalids = new();

            foreach (char c in input)
            {
                byte[] encoded = JisEncoding.GetBytes(new[] { c });

                int i = 0;
                while (i < encoded.Length && encoded[i] == 0x1B)
                {
                    while (i < encoded.Length && encoded[i] != 'B' && encoded[i] != 'J' && encoded[i] != '$') i++;
                    i++;
                }

                int len = encoded.Length - i;

                if (len == 1)
                {
                    if (!IsAllowed1Byte(encoded[i]))
                        invalids.Add(c);
                }
                else if (len == 2)
                {
                    if (!IsAllowed2Byte(encoded[i], encoded[i + 1]))
                        invalids.Add(c);
                }
                else
                {
                    invalids.Add(c);
                }
            }
            return invalids.Count > 0 ? string.Join(",", invalids) : null;
        }

        /// <summary>
        /// Unicodeの不正な文字を取得
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public string? GetInvalidUnicodeChars(string input)
        {
            List<char> invalids = new();
            foreach (var rune in input.EnumerateRunes())
            {
                if (!IsAllowedUnicode(rune.Value))
                {
                    invalids.Add(rune.ToString()[0]);
                }
            }
            return invalids.Count > 0 ? string.Join(",", invalids) : null;
        }
    }
}
