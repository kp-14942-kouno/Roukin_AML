using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace MyTemplate.RoukinClass
{
    /// <summary>
    /// 団体　不備審査クラス
    /// </summary>
    public class FubiCheck : MyLibrary.MyLoading.Thread
    {
        private DataTable _dataTable;
        private DataTable _contoryCode;
        private DataTable _bussinessCode;
        private DataTable _bussinessCodePerson;

        public DataTable FubiData { get; set; } = new DataTable();
        public DataTable FixData { get; set; } = new DataTable();
        
        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FubiCheck()
        {
            // リソースの解放
            _dataTable?.Dispose();
            _contoryCode?.Dispose();
            _bussinessCode?.Dispose();
            _bussinessCodePerson?.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="dataTable"></param>
        public FubiCheck(DataTable dataTable)
        {
            _dataTable = dataTable;
            // 不備登録用テーブル
            FubiData = CreateFubiTable();
            // 正常登録用テーブル
            FixData = CreateFixTable();
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            try
            {
                // 各コードをテーブルで取得
                GetCodeData();
                // 実行
                Run();
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                Result = MyEnum.MyResult.Error;
            }
            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        private void Run()
        {
            ProcessName = "不備審査処理中…";
            ProgressMax = _dataTable.Rows.Count;
            ProgressValue = 0;

            // 人格コード（実質的支配者の不備審査が必要）
            HashSet<string> personCodes = new HashSet<string> { "21", "31", "81" };

            foreach (DataRow row in _dataTable.Rows)
            {
                ProgressValue++;

                // 不備コード用
                StringBuilder fubiCode = new StringBuilder();

                // 実質的支配者のチェック要否
                // 人格コード
                string personCode = row["bpo_person_cd"].ToString().Trim();
                // 実質的支配者1～3
                string[] person1 = GetPersonData(row, "ubo1");
                string[] person2 = GetPersonData(row, "ubo2");
                string[] person3 = GetPersonData(row, "ubo3");

                FubiCheck_01(row, fubiCode);
                FubiCheck_02(row, fubiCode);
                FubiCheck_03(row, fubiCode);
                FubiCheck_04(row, fubiCode);
                FubiCheck_05(row, fubiCode);
                FubiCheck_06(row, fubiCode);
                FubiCheck_07(row, fubiCode);
                FubiCheck_08(row, fubiCode);
                FubiCheck_09(row, fubiCode);
                FubiCheck_10(row, fubiCode);

                FubiCheck_12(row, fubiCode);
                FubiCheck_13(row, fubiCode);
                FubiCheck_14(row, fubiCode);

                // 200万円超現金取引の頻度、金額、原資
                string[] freq = {
                    row["cash_freq1"].ToString().Trim(),
                    row["cash_freq2"].ToString().Trim()
                };
                // 空欄は除外
                freq = freq.Where(x => !string.IsNullOrEmpty(x)).ToArray();

                string[] amt = {
                    row["cash_amt1"].ToString().Trim(),
                    row["cash_amt2"].ToString().Trim()
                };
                // 空欄は除外
                amt = amt.Where(x => !string.IsNullOrEmpty(x)).ToArray();

                string[] src = {
                    row["cash_src1"].ToString().Trim(),
                    row["cash_src2"].ToString().Trim(),
                    row["cash_src3"].ToString().Trim(),
                    row["cash_src4"].ToString().Trim()
                };
                // 空欄は除外
                src = src.Where(x => !string.IsNullOrEmpty(x)).ToArray();

                // 200万円超現金取引の頻度、金額、原資のどれかに値が入っている場合は不備チェック
                if (!(freq.Count() == 0 && amt.Count() == 0 && src.Count() == 0))
                {
                    // 200万円超現金取引の頻度、金額、原資のチェック
                    FubiCheck_15(row, freq, fubiCode);
                    FubiCheck_16(row, amt, fubiCode);
                    FubiCheck_17(row, src, fubiCode);
                }

                FubiCheck_18(row, fubiCode);
                FubiCheck_19(row, fubiCode);
                FubiCheck_20(row, fubiCode);
                FubiCheck_21(row, fubiCode);
                FubiCheck_22(row, fubiCode);

                // 実質的支配者1のチェック
                if (personCodes.Contains(personCode))
                {
                    // 人格コードが対象の場合は必ずチェック
                    FubiCheck_23(row, fubiCode, "ubo1", "23");
                    FubiCheck_24(row, fubiCode, "ubo1", "24");
                    FubiCheck_25(row, fubiCode, "ubo1", "25");
                    FubiCheck_26(row, fubiCode, "ubo1", "26");
                    FubiCheck_27(row, fubiCode, "ubo1", "27");
                    FubiCheck_28(row, fubiCode, "ubo1", "28");
                }

                // 実質的支配者2のチェック
                if (personCodes.Contains(personCode) && person2.Count() > 0)
                {
                    // 人格コードが対象で実質的支配者2のデータが存在する場合
                    FubiCheck_23(row, fubiCode, "ubo2", "29");
                    FubiCheck_24(row, fubiCode, "ubo2", "30");
                    FubiCheck_25(row, fubiCode, "ubo2", "31");
                    FubiCheck_26(row, fubiCode, "ubo2", "32");
                    FubiCheck_27(row, fubiCode, "ubo2", "33");
                    FubiCheck_28(row, fubiCode, "ubo2", "34");
                }

                // 実質的支配者3のチェック
                if (personCodes.Contains(personCode) && person3.Count() > 0)
                {
                    // 人格コードが対象で実質的支配者3のデータが存在する場合
                    FubiCheck_23(row, fubiCode, "ubo3", "35");
                    FubiCheck_24(row, fubiCode, "ubo3", "36");
                    FubiCheck_25(row, fubiCode, "ubo3", "37");
                    FubiCheck_26(row, fubiCode, "ubo3", "38");
                    FubiCheck_27(row, fubiCode, "ubo3", "39");
                    FubiCheck_28(row, fubiCode, "ubo3", "40");
                }

                FubiCheck_41(row, fubiCode);
                FubiCheck_42(row, fubiCode);

                // 不備コードが存在する場合は不備データに追加
                if (fubiCode.Length == 0)
                {
                    var fixRow = FixData.NewRow();
                    fixRow["bpo_num"]  = row["bpo_num"];
                    FixData.Rows.Add(fixRow);
                }
                else
                {
                    var fubiRow = FubiData.NewRow();
                    fubiRow["bpo_num"] = row["bpo_num"];
                    fubiRow["fubi_code"] = fubiCode.ToString().TrimEnd(';');
                    FubiData.Rows.Add(fubiRow);
                }
                fubiCode.Clear(); // 不備コードをクリア
            }
        }

        // 実質的支配者の項目を配列化　空欄を除外
        private string[] GetPersonData(DataRow row, string prefix)
        {
            // 指定されたプレフィックスに基づいて、各項目を取得
            string[] personData =
            {
                row[$"{prefix}_name"].ToString().Trim(),
                row[$"{prefix}_kana"].ToString().Trim(),
                row[$"{prefix}_bday"].ToString().Trim(),
                row[$"{prefix}_rel1"].ToString().Trim(),
                row[$"{prefix}_rel2"].ToString().Trim(),
                row[$"{prefix}_addr"].ToString().Trim(),
                row[$"{prefix}_job1"].ToString().Trim(),
                row[$"{prefix}_job2"].ToString().Trim(),
                row[$"{prefix}_peps"].ToString().Trim(),
                row[$"{prefix}_nat"].ToString().Trim(),
                row[$"{prefix}_natname"].ToString().Trim(),
                row[$"{prefix}_alpha"].ToString().Trim(),
                row[$"{prefix}_visa"].ToString().Trim(),
                row[$"{prefix}_exp"].ToString().Trim()
            };

            // 空欄を除外
            personData = personData.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            
            return personData;
        }

        /// <summary>
        /// 不備チェック　01　団体名漢字
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_01(DataRow row, StringBuilder fubiCode)
        {
            string flg = row["org_name_chg"].ToString().Trim();
            string name = row["org_name_new"].ToString().Trim();

            // 変更なしで団体名漢字が空はOK
            if (flg == "0" && string.IsNullOrEmpty(name)) return;
            // 変更有で団体名漢字に値が入っている場合はOK
            if (flg == "1" && !string.IsNullOrEmpty(name)) return;
            // 変更有無未選択で団体名漢字に値が入っている場合はOK
            if (string.IsNullOrEmpty(flg) && !string.IsNullOrEmpty(name)) return;

            // それ以外は不備コードをセット
            fubiCode.Append("01;");
        }

        /// <summary>
        /// 不備チェック 02　団体名カナ
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_02(DataRow row, StringBuilder fubiCode)
        {
            string flg = row["org_name_chg"].ToString().Trim();
            string name = row["org_kana_new"].ToString().Trim();

            // 変更なしで団体名カナが空はOK
            if (flg == "0" && string.IsNullOrEmpty(name)) return;
            // 変更有で団体名カナに値が入っている場合はOK
            if (flg == "1" && !string.IsNullOrEmpty(name)) return;
            // 変更有無未選択で団体名カナに値が入っている場合はOK
            if (string.IsNullOrEmpty(flg) && !string.IsNullOrEmpty(name)) return;

            // それ以外は不備コードをセット
            fubiCode.Append("02;");
        }

        /// <summary>
        /// 不備チェック　03　団体種類
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_03(DataRow row, StringBuilder fubiCode)
        {
            string[] grpType =
            {
                row["org_type1"].ToString().Trim(),
                row["org_type2"].ToString().Trim()
            };

            // 空欄は除外
            grpType = grpType.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 団体種類の選択が1つならOK
            if( grpType.Count() == 1) return;

            // それ以外は不備コードをセット
            fubiCode.Append("03;");
        }

        /// <summary>
        /// 不備チェック　04　事業内容 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_04(DataRow row, StringBuilder fubiCode)
        {
            // 事業内容を配列化
            string[] biz = {
                row["biz_type1"].ToString().Trim(),
                row["biz_type2"].ToString().Trim()
            };

            // 空欄は除外
            biz = biz.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 事業内容のコードで_bussinessCodeのコード以外が存在するか
            //bool hasIvalid = biz.Any(x => !_bussinessCode.AsEnumerable()
            //    .Any(y => y.Field<string>("code") == x));

            // 事業内容コードの選択が１つならOK
            if(biz.Count() == 1) return;

            // それ以外は不備コードをセット
            fubiCode.Append("04;");
        }

        /// <summary>
        /// 不備チェック　05　設立年月日
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_05(DataRow row, StringBuilder fubiCode)
        {
            string estdateFlg = row["est_date_chg"].ToString().Trim();
            string estdate = row["est_date"].ToString().Trim();

            // 変更なしで設立年月日が空はOK
            if (estdateFlg == "0" && string.IsNullOrEmpty(estdate)) return;

            // 設立年月日（yyyymmdd）をyyyy/MM/dd形式に変換
            DateTime? formattedDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString(estdate, "yyyyMMdd");

            // 設立年月日がyyyy/MM/dd形式に変換できた場合はOK
            if (!string.IsNullOrEmpty(formattedDate.ToString()))
            {
                // 設立年月日が未来日でなく
                if (formattedDate <= DateTime.Now)
                {
                    // 変更ありはOK
                    if (estdateFlg == "1") return;
                    // 変更有無未選択はOK
                    if (string.IsNullOrEmpty(estdateFlg)) return;
                }
            }

            // それ以外は不備コードをセット
            fubiCode.Append("05;");
        }

        /// <summary>
        /// 不備チェック　06　郵便番号
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_06(DataRow row, StringBuilder fubiCode)
        {
            string addrChgFlg = row["addr_chg"].ToString().Trim();
            string zipCode = row["zip_new"].ToString().Trim();

            // 住所変更なしで郵便番号が空はOK
            if (addrChgFlg == "0" && string.IsNullOrEmpty(zipCode)) return;

            // zipCodeが7桁の数字で
            if (System.Text.RegularExpressions.Regex.IsMatch(zipCode, @"^\d{7}$"))
            {
                // 住所変更ありならOK
                if (addrChgFlg == "1") return;
                // 住所変更有無未選択ならOK
                if(string.IsNullOrEmpty(addrChgFlg)) return;
            }

            // それ以外は不備コードをセット
            fubiCode.Append("06;");
        }

        /// <summary>
        /// 不備チェック　07　所在地
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_07(DataRow row, StringBuilder fubiCode)
        {
            string addrChgFlg = row["addr_chg"].ToString().Trim();
            string addr = row["pref_new"].ToString().Trim() +
                        row["city_new"].ToString().Trim() +
                        row["addr1_new"].ToString().Trim() +
                        row["addr2_new"].ToString().Trim();

            // 住所変更なしで住所が空はOK
            if (addrChgFlg == "0" && string.IsNullOrEmpty(addr)) return;

            // 住所があれば
            if (!string.IsNullOrEmpty(addr)) 
            {
                // 住所変更ありならOK
                if (addrChgFlg == "1") return;
                // 住所変更有無未選択ならOK
                if (string.IsNullOrEmpty(addrChgFlg)) return;
            }

            // それ以外は不備コードをセット
            fubiCode.Append("07;");
        }

        /// <summary>
        /// 不備チェック　08　本店所在国
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_08(DataRow row, StringBuilder fubiCode)
        {
            string cngFlg = row["addr_chg"].ToString().Trim();
            string hqFlg = row["hq_ctry_chg"].ToString().Trim();
            string hqCountry = row["hq_country"].ToString().Trim();
            // 日本コード
            string jpn = MyUtilityModules.AppSetting("roukin_setting", "jpn_code");

            // 変更なしで
            if (cngFlg == "0")
            {
                // 日本・日本以外未選択・国名未記入はOK
                if (string.IsNullOrEmpty(hqFlg) && string.IsNullOrEmpty(hqCountry)) return;
            }

            // 変更ありで
            if (cngFlg == "1")
            {
                // 日本選択・国名未記入はOK
                if (hqFlg == "01" && string.IsNullOrEmpty(hqCountry)) return;

                // 国籍テーブルから国コードを取得
                string? code = _contoryCode.AsEnumerable()
                    .Where(x => x.Field<string>("country_name") == hqCountry)
                    .Select(x => x.Field<string>("code"))
                    .FirstOrDefault();

                // 日本以外選択で国名が日本以外はOK
                if (hqFlg == "02" && !string.IsNullOrEmpty(code) && code != jpn) return;
            }

            // それ以外は不備コードをセット
            fubiCode.Append("08;");
        }

        /// <summary>
        /// 第一電話番号
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_09(DataRow row, StringBuilder fubiCode)
        {
            string tel = row["tel"].ToString().Trim().Replace("-", "");

            // 電話番号が数値10桁か11桁ならOK
            if (Regex.IsMatch(tel, @"^\d{10,11}$")) return;

            fubiCode.Append("09;");
        }

        /// <summary>
        /// 取引目的
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_10(DataRow row, StringBuilder fubiCode)
        {
            string flg = row["purp_chg"].ToString().Trim();

            // 取引目的を配列化
            string[] trans =
            {
                row["purp_cd1"].ToString().Trim(),
                row["purp_cd2"].ToString().Trim(),
                row["purp_cd3"].ToString().Trim(),
                row["purp_cd4"].ToString().Trim(),
                row["purp_cd5"].ToString().Trim(),
                row["purp_cd6"].ToString().Trim()
            };

            string purpCd7 = row["purp_cd7"].ToString().Trim();

            // 空欄は除外
            trans = trans.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 変更なしで取引目的なしはOK
            if (flg == "0" && trans.Count() == 0) return;

            // 取引目的1～6に値があり取引目的7に値が無ければOK
            if (trans.Count() > 0 && string.IsNullOrEmpty(purpCd7)) return;

            // 取引目的7に値があっても取引目的1～6の入力数が1以上6未満ならOK
            if(trans.Count() >=1 && trans.Count() < 6 && !string.IsNullOrEmpty(purpCd7)) return;

            // 正常取引目的コードハッシュリスト
            //var validCOdes = new HashSet<string> { "001", "002", "003", "004", "005", "006" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = trans.Any(x => !validCOdes.Contains(x));

            // 取引目的コードが存在し内容が001～006で入力数が6以下ならOK
            //if (!hasIvalid && trans.Count() > 0 && trans.Count() <= 6) return;

            // 取引目的コードの選択数が6を超えていた場合は不備コード11それ以外は10
            fubiCode.Append(trans.Count() == 0 ? "10;" : "11;");
        }

        /// <summary>
        /// 取引形態
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_12(DataRow row, StringBuilder fubiCode)
        {
            string[] trans =
            {
                row["deal_type1"].ToString().Trim(),
                row["deal_type2"].ToString().Trim()
            };

            // 空欄は除外
            trans = trans.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 正常取引形態コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3", "4", "5" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = trans.Any(x => !validCOdes.Contains(x));

            // 選択が1つならOK
            if (trans.Count() == 1) return;

            fubiCode.Append("12;");
        }

        /// <summary>
        /// 取引頻度
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_13(DataRow row, StringBuilder fubiCode)
        {
            string[] trans =
            {
                row["deal_freq1"].ToString().Trim(),
                row["deal_freq2"].ToString().Trim()
            };

            // 空欄は除外
            trans = trans.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 正常取引形態コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3", "4", "5", "6" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = trans.Any(x => !validCOdes.Contains(x));

            // 選択が1つはOK
            if (trans.Count() == 1) return;

            fubiCode.Append("13;");
        }

        /// <summary>
        /// 1回あたりの取引金額
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_14(DataRow row, StringBuilder fubiCode)
        {
            string[] trans =
            {
                row["deal_amt1"].ToString().Trim(),
                row["deal_amt2"].ToString().Trim()
            };

            // 空欄は除外
            trans = trans.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 正常取引形態コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3", "4", "5", "6" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = trans.Any(x => !validCOdes.Contains(x));

            // 選択が1つはOK
            if (trans.Count() == 1) return;

            fubiCode.Append("14;");
        }

        /// <summary>
        /// 200万円超現金取引の頻度
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_15(DataRow row, string[] freq, StringBuilder fubiCode)
        {
            // 正常取引形態コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3", "4", "5", "6" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = freq.Any(x => !validCOdes.Contains(x));

            // 選択が1つで正しいコードはOK
            if (freq.Count() == 1) return;

            fubiCode.Append("15;");
        }

        /// <summary>
        /// 200万円超現金取引の金額
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_16(DataRow row, string[] amt, StringBuilder fubiCode)
        {
            // 正常取引形態コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3", "4" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = amt.Any(x => !validCOdes.Contains(x));

            // 選択が1つはOK
            if (amt.Count() == 1) return;

            fubiCode.Append("16;");
        }

        /// <summary>
        /// 200万円超現金取引の原資
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_17(DataRow row, string[] src, StringBuilder fubiCode)
        {
            // 正常取引形態コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3", "4", "6", "7", "8", "9", "10" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = src.Any(x => !validCOdes.Contains(x));

            // 選択が1つから3つでOK
            if (src.Count() >= 1 && src.Count() <=3) return;

            fubiCode.Append("17;");
        }

        /// <summary>
        /// 代表者の漢字氏名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_18(DataRow row, StringBuilder fubiCode) 
        {
            string flg = row["rep_chg"].ToString().Trim();
            string name = row["rep_name"].ToString().Trim();

            // 変更なしで代表者の漢字氏名が空はOK
            if (flg == "0" && string.IsNullOrEmpty(name)) return;

            // 氏名が空でなく文字列内に全角スペースがあれば
            if (!string.IsNullOrEmpty(name) && name.Contains("　"))
            {
                // 変更ありならOK
                if (flg == "1") return;
                // 変更有無未選択ならOK
                if (string.IsNullOrEmpty(flg)) return;
            }

            fubiCode.Append("18;");
        }

        /// <summary>
        /// 代表者のカナ氏名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_19(DataRow row, StringBuilder fubiCode)
        {
            string flg = row["rep_chg"].ToString().Trim();
            string name = row["rep_kana"].ToString().Trim();

            // 変更なしで代表者のカナ氏名が空はOK
            if (flg == "0" && string.IsNullOrEmpty(name)) return;

            // 氏名が空でなく文字列内に半角スペースがあれば
            if (!string.IsNullOrEmpty(name) && name.Contains(" "))
            {   
                // 変更ありならOK
                if (flg == "1") return;
                // 変更有無未選択ならOK
                if (string.IsNullOrEmpty(flg)) return;
            }

            fubiCode.Append("19;");
        }

        /// <summary>
        /// 代表者の生年月日
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_20(DataRow row, StringBuilder fubiCode)
        {
            string flg = row["rep_chg"].ToString().Trim();
            string birth = row["rep_bday"].ToString().Trim();

            // 変更なしで代表者の生年月日が空はOK
            if (flg=="0" && string.IsNullOrEmpty(birth)) return;

            // 生年月日（yyyymmdd）をyyyy/MM/dd形式に変換
            DateTime? formattedDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString(birth, "yyyyMMdd");

            // 生年月日がyyyy/MM/dd形式に変換できた場合はOK
            if (!string.IsNullOrEmpty(formattedDate.ToString()))
            {
                // 生年月日が未来かどうか
                bool isFuture = formattedDate > DateTime.Now;
                // 変更ありで未来日でなければOK
                if (!isFuture && flg == "1") return;
                // 変更有無未選択で未来日でなければOK
                if (!isFuture && string.IsNullOrEmpty(flg)) return;
            }
            // それ以外は不備コードをセット
            fubiCode.Append("20;");
        }

        /// <summary>
        /// 代表者の役職
        /// </summary>
        /// <param name="row"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_21(DataRow row, StringBuilder fubiCode)
        {
            string flg = row["rep_chg"].ToString().Trim();
            string title = row["rep_title"].ToString().Trim();

            // 変更なしで代表者の役職が空はOK
            if (flg == "0" && string.IsNullOrEmpty(title)) return;
            // 変更ありで代表者の役職に値が入っている場合はOK
            if (flg == "1" && !string.IsNullOrEmpty(title)) return;
            // 変更有無未選択で代表者の役職に値が入っている場合はOK
            if (string.IsNullOrEmpty(flg) && !string.IsNullOrEmpty(title)) return;

            // それ以外は不備コードをセット
            fubiCode.Append("21;");
        }

        /// <summary>
        /// 代表者の国籍
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_22(DataRow row, StringBuilder fubiCode)
        {
            // 国籍
            string nat = row["rep_nat"].ToString().Trim();
            // 国名
            string natname = row["rep_natname"].ToString().Trim();
            // 日本コード
            string jpn = MyUtilityModules.AppSetting("roukin_setting", "jpn_code");

            // 日本選択で国名が未記入はOK
            if (nat == "01" && string.IsNullOrEmpty(natname)) return;
            // 日本・日本以外未選択で国名未記入はOK
            if(nat == "" && string.IsNullOrEmpty(natname)) return;

            // 国籍テーブルから国コードを取得
            string? code = _contoryCode.AsEnumerable()
                .Where(x => x.Field<string>("country_name") == natname)
                .Select(x => x.Field<string>("code"))
                .FirstOrDefault();

            // 日本以外選択で日本以外の国名が存在する場合はOK
            if (nat == "02" && !string.IsNullOrEmpty(code) && code != jpn) return;
            // 日本・日本以外とも選択で国名が存在する場合はOK
            if(nat == "0102" && !string.IsNullOrEmpty(code)) return;
            // 日本・日本以外とも未選択で国名が存在する場合はOK
            if(string.IsNullOrEmpty(nat) && !string.IsNullOrEmpty(code)) return;
            // 日本選択・国名記入はOK
            if (nat == "01" && !string.IsNullOrEmpty(code)) return;
   
            // それ以外は不備コードをセット
            fubiCode.Append("22;");
        }

        /// <summary>
        /// 実質的支配者の氏名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_23(DataRow row, StringBuilder fubiCode, string preFix, string fubi)
        {
            // 実質的支配者の氏名
            string name = row[$"{preFix}_name"].ToString().Trim();

            // 氏名が空でなく文字列内に全角スペースがあればOK
            if (!string.IsNullOrEmpty(name) && name.Contains("　")) return;

            // それ以外は不備コードをセット
            fubiCode.Append(fubi + ";");
        }

        /// <summary>
        /// 実質的支配者カナ氏名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        /// <param name="preFix"></param>
        /// <param name="fubi"></param>
        private void FubiCheck_24(DataRow row, StringBuilder fubiCode, string preFix, string fubi)
        {
            // 実質的支配者の氏名
            string name = row[$"{preFix}_kana"].ToString().Trim();

            // 氏名が空でなく文字列内に半角スペースがあればOK
            if (!string.IsNullOrEmpty(name) && name.Contains(" ")) return;

            // それ以外は不備コードをセット
            fubiCode.Append(fubi + ";");
        }

        /// <summary>
        /// 実質的支配者の生年月日
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        /// <param name="preFix"></param>
        /// <param name="fubi"></param>
        private void FubiCheck_25(DataRow row, StringBuilder fubiCode, string preFix, string fubi)
        {
            // 実質的支配者の生年月日
            string birth = row[$"{preFix}_bday"].ToString().Trim();

            // 生年月日（yyyymmdd）をyyyy/MM/dd形式に変換
            DateTime? formattedDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString(birth, "yyyyMMdd");

            // 生年月日がyyyy/MM/dd形式に変換できた場合はOK
            if (!string.IsNullOrEmpty(formattedDate.ToString()))
            {
                // 生年月日が未来日でなければOK
                if (formattedDate <= DateTime.Now) return;
            }
            // それ以外は不備コードをセット
            fubiCode.Append(fubi + ";");
        }

        /// <summary>
        /// 実質的支配者の団体との関係
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        /// <param name="preFix"></param>
        /// <param name="fubi"></param>
        private void FubiCheck_26(DataRow row, StringBuilder fubiCode, string preFix, string fubi)
        {
            // 実質的支配者の団体との関係
            string[] rel = {
                row[$"{preFix}_rel1"].ToString().Trim(),
                row[$"{preFix}_rel2"].ToString().Trim()
            };

            // 空欄は除外
            rel = rel.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 正常コードハッシュリスト
            //var validCOdes = new HashSet<string> { "1", "2", "3" };
            // ハッシュリストのコード以外が存在するかチェック
            //bool hasIvalid = rel.Any(x => !validCOdes.Contains(x));

            // 選択が1つはOK
            if (rel.Count() == 1) return;

            // それ以外は不備コードをセット
            fubiCode.Append(fubi + ";");
        }

        /// <summary>
        /// 実質的支配者の職業・事業内容
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        /// <param name="preFix"></param>
        /// <param name="fubi"></param>
        private void FubiCheck_27(DataRow row, StringBuilder fubiCode, string preFix, string fubi)
        {
            // 実質的支配者の職業・事業内容
            string[] job = {
                row[$"{preFix}_job1"].ToString().Trim(),
                row[$"{preFix}_job2"].ToString().Trim()
            };

            // 空欄は除外
            job = job.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // jobのコードで_bussinessCodePersonのコード以外が存在するか
            //bool hasIvalid = job.Any(x => !_bussinessCodePerson.AsEnumerable()
            //    .Any(y => y.Field<string>("code") == x));

            // 選択が1つはOK
            if (job.Count() == 1) return;

            // それ以外は不備コードをセット
            fubiCode.Append(fubi + ";");
        }

        /// <summary>
        /// 実質的支配者の国籍・国名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        /// <param name="preFix"></param>
        /// <param name="fubi"></param>
        private void FubiCheck_28(DataRow row, StringBuilder fubiCode, string preFix, string fubi)
        {
            // 国籍
            string nat = row[$"{preFix}_nat"].ToString().Trim();
            // 国名
            string natname = row[$"{preFix}_natname"].ToString().Trim();
            // 日本コード
            string jpn = MyUtilityModules.AppSetting("roukin_setting", "jpn_code");

            // 日本選択・国名未記入はOK
            if (nat == "01" && string.IsNullOrEmpty(natname)) return;

            // 国籍テーブルから国コードを取得
            string? code = _contoryCode.AsEnumerable()
                .Where(x => x.Field<string>("country_name") == natname)
                .Select(x => x.Field<string>("code"))
                .FirstOrDefault();

            // 日本以外選択で日本外の国名記入はOK
            if (nat == "02" && !string.IsNullOrEmpty(code) && code != jpn) return;
            // 日本・日本以外未選択で国名記入はOK
            if (nat == "" && !string.IsNullOrEmpty(code)) return;
            // 日本・日本以外とも選択で国名記入はOK
            if (nat == "0102" && !string.IsNullOrEmpty(code)) return;
            // 日本選択で国名記入はOK
            if (nat == "01" && !string.IsNullOrEmpty(code)) return;

            // それ以外は不備コードをセット
            fubiCode.Append(fubi + ";");
        }

        /// <summary>
        /// 取引担当者の氏名
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_41(DataRow row, StringBuilder fubiCode)
        {
            string name = row["staff_name"].ToString().Trim();

            // 氏名が空でなく文字列内に全角スペースがあればOK
            if (!string.IsNullOrEmpty(name) && name.Contains("　")) return;

            fubiCode.Append("41;");
        }

        /// <summary>
        /// 取引担当者の電話番号
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beatle"></param>
        /// <param name="fubiCode"></param>
        private void FubiCheck_42(DataRow row, StringBuilder fubiCode)
        {
            string tel = row["staff_tel"].ToString().Trim().Replace("-", "");

            // 電話番号が数値10桁か11桁ならOK
            if (Regex.IsMatch(tel, @"^\d{10,11}$")) return;

            fubiCode.Append("42;");
        }

        /// <summary>
        /// 各コード変換用テーブルを取得
        /// </summary>
        private void GetCodeData()
        {
            using (var db = new MyDbData("code"))
            {
                // 国コード
                _contoryCode = db.ExecuteQuery("select * from t_country_code order by code;");
                // 業種コード（団体）
                _bussinessCode = db.ExecuteQuery("select * from t_business_code_organization order by code;");
                // 業種コード（個人）
                _bussinessCodePerson = db.ExecuteQuery("select * from t_business_code_person order by code;");
            }
        }

        /// <summary>
        /// 不備登録用テーブルを作成
        /// </summary>
        /// <returns></returns>
        private DataTable CreateFubiTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("bpo_num", typeof(string));
            table.Columns.Add("fubi_code", typeof(string));
            return table;
        }

        /// <summary>
        /// 正常登録用テーブルを作成
        /// </summary>
        /// <returns></returns>
        private DataTable CreateFixTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("bpo_num", typeof(string));
            return table;
        }


    }
}
