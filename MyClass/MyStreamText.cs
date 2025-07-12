using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyLibrary.MyModules;
using MyLibrary;
using DocumentFormat.OpenXml.Office2013.Drawing.Chart;

namespace MyTemplate.Class
{
    /// <summary>
    /// MystreamTextクラス
    /// </summary>
    /// <param name="textSetting"></param>
    internal class MyStreamText : IDisposable
    {
        string _filePath;
        Encoding _encoding;
        MyEnum.Delimiter _delimiter;
        MyEnum.Separator _separator;
        private bool _byteOrderMarks = true;    //  detectEncodingFromByteOrderMarks

        private System.IO.StreamReader? _stream;
        public int RecordNum { get; private set; }

        public MyStreamText(string filePath, Encoding encoding, MyEnum.Delimiter delimiter = MyEnum.Delimiter.Comma, MyEnum.Separator separator = MyEnum.Separator.DoubleQuotation)
        {
            _filePath = filePath;
            _encoding = encoding;
            _delimiter = delimiter;
            _separator = separator;

            // 文字コードによって「detectEncodingFromByteOrderMarks」を変更
            // 932: Shift_jis 65001: utf8 1200:utf16le 1201: utf16be
            var codePage = encoding.CodePage;
            _byteOrderMarks = codePage switch
            {
                932 => false,
                65001 or 1200 or 1201 => true,
                _ => true
            };
        }

        /// <summary>
        /// リソースの解放
        /// </summary>
        public void Dispose()
        {
            _stream?.Dispose();
        }

        /// <summary>
        /// ファイルを開く
        /// </summary>
        public void Open()
        {
            // ファイルを開く
            _stream = new System.IO.StreamReader(_filePath, _encoding, _byteOrderMarks);
        }

        /// <summary>
        /// レコードをスキップ
        /// </summary>
        /// <param name="SkipCount"></param>
        public void SkipRecord(int SkipCount)
        {
            for (int i = 1; i < SkipCount; i++)
            {
                RecordNum++;
                if (_stream?.ReadLine() == null)
                {
                    break;
                }
            }
        }

        /// <summary>
        /// 終端を取得
        /// </summary>
        /// <returns></returns>
        public bool EndOfStream()
        {
            return _stream.EndOfStream;
        }

        /// <summary>
        /// 行を読込み配列化
        /// </summary>
        /// <returns></returns>
        public string[] ReadLine()
        {
            var line = new StringBuilder();
            bool inSeparator = false;
            char quoteChar = '\0';
            char serpartorChar = GetSeparatorChar(_separator);
            char delimitorChar = GetDelimitorChar(_delimiter);

            while (!_stream.EndOfStream)
            {
                RecordNum++;

                var currentLine = _stream.ReadLine();

                if (currentLine == null)
                {
                    break;
                }

                for (int i = 0; i < currentLine.Length; i++)
                {
                    char currentChar = currentLine[i];
                    if (inSeparator)
                    {
                        // セパレータ内にいる場合
                        if (currentChar == serpartorChar)
                        {
                            inSeparator = false;
                        }
                    }
                    else
                    {
                        // セパレータ外にいる場合
                        if (currentChar == serpartorChar)
                        {
                            inSeparator = true;
                            quoteChar = currentChar;
                        }
                    }
                    line.Append(currentChar);
                }

                if (!inSeparator)
                {
                    // セパレータ外にいる場合
                    break;
                }
                else
                {
                    // セパレータ内にいる場合、改行を追加
                    line.Append(Environment.NewLine);
                }
            }
            // 行を分割
            return SplitLine(line.ToString(), delimitorChar, serpartorChar);
        }

        /// <summary>
        /// 行を分割
        /// </summary>
        /// <param name="line"></param>
        /// <param name="delimitorChar"></param>
        /// <param name="serpartorChar"></param>
        /// <returns></returns>
        private string[] SplitLine(string line, char delimitorChar, char serpartorChar)
        {
            var result = new List<string>();
            var field = new StringBuilder();
            bool inSeparator = false;
            char quoteChar = '\0';

            for (int i = 0; i < line.Length; i++)
            {
                char currentChar = line[i];
                if (inSeparator)
                {
                    // セパレータ内にいる場合
                    if (currentChar == serpartorChar)
                    {
                        inSeparator = false;
                    }
                    else
                    {
                        field.Append(currentChar);
                    }
                }
                else
                {
                    // セパレータ外にいる場合
                    if (currentChar == serpartorChar)
                    {
                        inSeparator = true;
                        quoteChar = currentChar;
                    }
                    else if (currentChar == delimitorChar)
                    {
                        result.Add(field.ToString());
                        field.Clear();
                    }
                    else
                    {
                        field.Append(currentChar);
                    }
                }
            }
            result.Add(field.ToString());
            return result.ToArray();
        }

        /// <summary>
        /// レコード数を取得
        /// </summary>
        /// <returns></returns>
        public int RecordCount()
        {
            int recordCount = 0;
            char serpartorChar = GetSeparatorChar(_separator);

            using (var reader = new System.IO.StreamReader(_filePath, _encoding, _byteOrderMarks))
            {
                while (!reader.EndOfStream)
                {
                    char[] buffer = new char[8192];
                    int charsRead;
                    bool inSeparator = false;

                    while ((charsRead = reader.ReadBlock(buffer, 0, buffer.Length)) > 0)
                    {
                        for (int i = 0; i < charsRead; i++)
                        {
                            char currentChar = buffer[i];
                            if (inSeparator)
                            {
                                // セパレータ内にいる場合
                                if (currentChar == serpartorChar)
                                {
                                    inSeparator = false;
                                }
                            }
                            else
                            {
                                // セパレータ外にいる場合
                                if (currentChar == serpartorChar)
                                {
                                    inSeparator = true;
                                }
                                else if (currentChar == '\n')
                                {
                                    recordCount++;
                                }
                            }
                        }
                    }
                }
            }
            return recordCount;
        }

        /// <summary>
        /// 指定行の指定長さのデータを配列で取得
        /// </summary>
        /// <param name="lengths"></param>
        /// <param name="isByte"></param>
        /// <returns></returns>
        public string[] ReadLine(List<int> lengths, int totalLength, bool isByte = false)
        {
            var line = new List<string>();
            int lineLength;

            if (!_stream.EndOfStream)
            {
                string currentLine = _stream.ReadLine();

                if (isByte)
                {
                    // バイト単位で長さを取得
                    lineLength = currentLine.LenBSjis();
                }
                else
                {
                    // 文字単位で長さを取得
                    lineLength = currentLine.LengthInTextElements();
                }

                if (lineLength != totalLength)
                {
                    return new string[] { currentLine };
                }

                if (isByte)
                {
                    // バイト単位で分割
                    line = SplitByte(lengths, currentLine);
                }
                else
                {
                    // 文字単位で分割
                    line = SplitLen(lengths, currentLine);
                }
            }
            return line.ToArray();
        }

        /// <summary>
        /// 文字列を指定長さで分割
        /// </summary>
        /// <param name="lengths"></param>
        /// <param name="currentLine"></param>
        /// <returns></returns>
        private List<string> SplitLen(List<int> lengths, string currentLine)
        {
            var result = new List<string>();
            int num = 0;
            // 各項目の長さに基づいて分割
            foreach (int length in lengths)
            {
                result.Add(currentLine.Mid(num, length));
                num += length;
            }
            return result;
        }

        /// <summary>
        ///　バイト単位で文字列を分割
        /// </summary>
        /// <param name="lengths"></param>
        /// <param name="currentLine"></param>
        /// <returns></returns>
        private List<string> SplitByte(List<int> lengths, string currentLine)
        {
            var result = new List<string>();
            int lenb = 0;
            int len = 0;
            // 文字列をSJISで置換
            string value = MyStringModules.SjisReplace(currentLine, '■');
            // 各項目の長さに基づいて分割
            foreach (int length in lengths)
            {
                // SJISに置き換えた文字列からバイト単位で分割
                string tmp = value.MidB(Encoding.GetEncoding("Shift_JIS"), lenb, length);
                lenb += length;
                // 元の文字列からバイト単位で分割した文字数で取得
                result.Add(currentLine.Mid(len, tmp.LengthInTextElements()));
                len += tmp.LengthInTextElements();
            }
            return result;
        }

        /// <summary>
        /// セパレータ文字を取得
        /// </summary>
        /// <param name="separator"></param>
        /// <returns></returns>
        private char GetSeparatorChar(MyEnum.Separator separator)
        {
            return separator == MyEnum.Separator.DoubleQuotation ? '"' :
                separator == MyEnum.Separator.SingleQuotation ? '\'' : '\0';
        }

        /// <summary>
        /// デリミタ文字を取得
        /// </summary>
        /// <param name="delimiter"></param>
        /// <returns></returns>
        private char GetDelimitorChar(MyEnum.Delimiter delimiter)
        {
            return delimiter == MyEnum.Delimiter.Comma ? ',' :
                delimiter == MyEnum.Delimiter.Tab ? '\t' : '\0';
        }
    }
}
