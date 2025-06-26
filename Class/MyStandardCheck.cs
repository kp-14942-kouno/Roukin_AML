using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MyLibrary.MyModules;
using MyLibrary;

namespace MyTemplate.Class
{
    /// <summary>
    /// MyStandardCheckクラス
    /// </summary>
    internal class MyStandardCheck
    {
        // 結果
        private string[] _resultAry = new string[] { "", "空白不可", "全角のみ", "全角数字のみ", "全角英数字のみ", "全角カナのみ", "半角のみ", "半角数字のみ", "半角英数字のみ",
                                                    "半角カナのみ", "日付(YYYYMMDD)形式のみ", "日付(YYYY/MM/DD)形式のみ", "存在しない日付", "文字長固定(@@@)", "文字長超過(@@@)", "文字長範囲以外(###～@@@)", "文字規則違反"};


        // 正規表現
        private static readonly Regex FullWidthRegex = new Regex(@"^([^ -~｡-ﾟ])+$", RegexOptions.Compiled);
        private static readonly Regex FullWidthDigitRegex = new Regex(@"^[０-９]+$", RegexOptions.Compiled);
        private static readonly Regex FullWidthAlnumRegex = new Regex(@"^[！-～]+$", RegexOptions.Compiled);
        private static readonly Regex FullWidthKanaRegex = new Regex(@"^([ァ-ー]|[　])+$", RegexOptions.Compiled);
        private static readonly Regex HalfWidthRegex = new Regex(@"^([ -~｡-ﾟ]|[ ])+$", RegexOptions.Compiled);
        private static readonly Regex HalfWidthDigitRegex = new Regex(@"^[0-9]+$", RegexOptions.Compiled);
        private static readonly Regex HalfWidthAlnumRegex = new Regex(@"^[!-~ ]+$", RegexOptions.Compiled);
        private static readonly Regex HalfWidthKanaRegex = new Regex(@"^[ｦ-ﾟ ]+$", RegexOptions.Compiled);
        private static readonly Regex DateYyyyMmDdRegex = new Regex(@"\d{8}$", RegexOptions.Compiled);
        private static readonly Regex DateYyyyMmSlashDdRegex = new Regex(@"\d{4}/\d{2}/\d{2}$", RegexOptions.Compiled);

        /// <summary>
        /// 結果番号・メッセージ取得
        /// </summary>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        /// <param name="dataByte"></param>
        /// <param name="nullFlg"></param>
        /// <param name="fixFlg"></param>
        /// <param name="regPattern"></param>
        /// <returns></returns>
        public (int, string) GetResult(string value, byte dataType, int dataByte, byte nullFlg, byte fixFlg, string regPattern)
        {
            return GetResult(value, dataType, 0, dataByte, nullFlg, fixFlg, regPattern);
        }

        public (int, string) GetResult(string value, byte dataType, int dataByteMin, int dataByteMax, byte nullFlg, byte fixFlg, string regPattern)
        {
            int result = ResultRun(value, dataType, dataByteMin, dataByteMax, nullFlg, fixFlg, regPattern);
            return (result, _resultAry[result].Replace("###", dataByteMin.ToString()).Replace("@@@", dataByteMax.ToString()));
        }

        /// <summary>
        /// チェック処理
        /// </summary>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        /// <param name="dataByteMin"></param>
        /// <param name="dataByteMax"></param>
        /// <param name="nullFlg"></param>
        /// <param name="fixFlg"></param>
        /// <param name="regPattern"></param>
        /// <returns></returns>
        private int ResultRun(string value, byte dataType, int dataByteMin, int dataByteMax, byte nullFlg, byte fixFlg, string regPattern)
        {
            value = MyStringModules.Nz(value);

            // 空白可否
            if (value == "" && nullFlg != 0) return 0;
            if (value == "" && nullFlg == 0) return 1;

            return TextCheck(value, dataType,dataByteMin, dataByteMax, fixFlg, regPattern);
        }

        /// <summary>
        /// 文字列チェック
        /// </summary>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        /// <param name="dataByte"></param>
        /// <param name="fixFlg"></param>
        /// <param name="regPattern"></param>
        /// <returns></returns>
        private int TextCheck(string value, byte dataType, int dataByteMin, int dataByteMax, byte fixFlg, string regPattern)
        {
            int result = 0;

            // データ型チェック
            result = DataTypeCheck(value, dataType);
            if (result != 0) return result;

            switch (dataType)
            {
                case 0:
                case 1:
                case 2:
                case 3:
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                    // 文字長チェック
                    if (dataByteMax > 0)
                    {
                        // 文字長固定
                        if (fixFlg != 0)
                        {
                            result = MyCompareModules.IIf(value.LengthInTextElements() != dataByteMax, 13, 0);
                        }
                        else
                        {
                            result = 0;
                            int len = value.LengthInTextElements();
                            if (len < dataByteMin || len > dataByteMax)
                                if (dataByteMin == 0) result = 14;
                                else result = 15;
                            
                            //result = MyCompareModules.IIf(value.LengthInTextElements() > dataByte, 14, 0);
                        }
                    }
                    break;

                case 11:
                    // バイト長チェック
                    if (dataByteMax > 0)
                    {
                        // 文字長固定
                        if (fixFlg != 0)
                        {
                            result = MyCompareModules.IIf(value.LenBSjis() != dataByteMax, 13, 0);
                        }
                        else
                        {
                            result = 0;
                            int len = value.LenBSjis();
                            if (len < dataByteMin || len > dataByteMax)
                                if (dataByteMin == 0) result = 14;
                                else result = 15;

                            //result = MyCompareModules.IIf(value.LenBSjis() > dataByte, 14, 0);
                        }
                    }
                    break;
            }

            if (result == 0)
            {
                // 正規表現
                if (MyStringModules.Nz(regPattern) != "")
                    if (!Regex.IsMatch(value, regPattern)) result = 16;
            }

            return result;
        }

        /// <summary>
        /// 文字型チェック
        /// </summary>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        private int DataTypeCheck(string value, byte dataType)
        {
            switch (dataType)
            {
                // 全角
                case 1:
                    if (!FullWidthRegex.IsMatch(value)) return 2;
                    break;

                // 全角数字
                case 2:
                    if (!FullWidthDigitRegex.IsMatch(value)) return 3;
                    break;

                // 全角英数字
                case 3:
                    if (!FullWidthAlnumRegex.IsMatch(value)) return 4;
                    break;

                // 全角カナ
                case 4:
                    if (!FullWidthKanaRegex.IsMatch(value)) return 5;
                    break;

                // 半角
                case 5:
                    if (!HalfWidthRegex.IsMatch(value)) return 6;
                    break;

                // 半角数字
                case 6:
                    if (!HalfWidthDigitRegex.IsMatch(value)) return 7;
                    break;

                // 半角英数字
                case 7:
                    if (!HalfWidthAlnumRegex.IsMatch(value)) return 8;
                    break;

                // 半角カナ
                case 8:
                    if (!HalfWidthKanaRegex.IsMatch(value)) return 9;
                    break;

                // 日付(yyyymmdd)
                case 9:
                    // 数値8桁チェック
                    if (!DateYyyyMmDdRegex.IsMatch(value)) return 10;
                    // yyyymmdd ⇒ yyyy/mm/dd に変換して日付チェック
                    string d = $"{value.Left(4)}/{value.Mid(4, 2)}/{value.Right(2)}";
                    if (!MyUtilityModules.IsDate(d)) return 12;
                    break;

                // 日付(yyyy/mm/dd)
                case 10:
                    if (!DateYyyyMmSlashDdRegex.IsMatch(value)) return 11;
                    if (!MyUtilityModules.IsDate(value)) return 12;
                    break;
            }
            return 0;
        }
    }
}
