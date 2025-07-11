using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System.Data;
using System.IO;
using System.Text;

namespace MyTemplate.RoukinClass
{
    /// <summary>
    /// ろうきん　納品データ作成クラス
    /// </summary>
    public class NouhinClass : MyLibrary.MyLoading.Thread
    {
        private DataTable _dantai = new();
        private DataTable _kojin = new();
        private string _expPath;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="dantai"></param>
        /// <param name="kojin"></param>
        /// <param name="expPath"></param>
        public NouhinClass(DataTable dantai, DataTable kojin, string expPath)
        {
            _dantai = dantai;
            _kojin = kojin;
            _expPath = expPath;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            //try
            {
                // 開始ログ出力
                MyLogger.SetLogger("納品データ作成を開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {
                    // 実行
                    Run(codeDb);

                    // 結果メッセージ
                    ResultMessage = $"納品データ作成完了／各：{_kojin.Rows.Count + _dantai.Rows.Count}件"
                        + "・顧客情報変更データ<br/>・回答結果顧客管理データ<br/>・回答結果イメージ管理データ<br/>・金庫事務用データ";

                    // 終了ログ出力
                    MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);

                    Result = MyLibrary.MyEnum.MyResult.Ok;
                }

            }
            //catch (Exception ex)
            //{
            //    MyLogger.SetLogger(ex, MyLibrary.MyEnum.LoggerType.Error);
            //    Result = MyLibrary.MyEnum.MyResult.Error;
            //}
            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        /// <param name="codeDb"></param>
        /// <exception cref="Exception"></exception>
        private void Run(MyDbData codeDb)
        {
            // 顧客情報変更データ名
            string customerFile = MyUtilityModules.AppSetting("roukin_setting", "customer_info_name");
            // 回答顧客情報ファイル名
            string resuponseFile = MyUtilityModules.AppSetting("roukin_setting", "resuponse_res_name");
            // 回答結果イメージ管理データファイル名
            string answerImageFile = MyUtilityModules.AppSetting("roukin_setting", "answer_res_image_name");
            // 納品ファイル作成結果データ（個人）
            string resPerile = MyUtilityModules.AppSetting("roukin_setting", "exp_per_res_file_name", true, _kojin.Rows.Count);
            // 納品ファイル作成結果データ（団体）
            string resGrpFile = MyUtilityModules.AppSetting("roukin_setting", "exp_grp_res_file_name", true, _dantai.Rows.Count);
            // DP連携ファイル名
            string dpFile = MyUtilityModules.AppSetting("roukin_setting", "dp_file_name", true);
            // 勘定系（顧客情報変更データ）ディレクトリ
            string customerDir = MyUtilityModules.AppSetting("roukin_setting", "customer_dir", true);
            // 本陣確認システム（回答結果データ）ディレクトリ
            string ansDir = MyUtilityModules.AppSetting("roukin_setting", "answer_dir", true);

            string customerPath = Path.Combine(_expPath, customerDir);
            string ansPath = Path.Combine(_expPath, ansDir);

            // 各ディレクトリを作成
            Directory.CreateDirectory(customerPath);
            Directory.CreateDirectory(ansPath);

            // 金融機関コードをリスト化
            List<string> bankCodes = GetBankCodes();

            // 納品結果データ
            using (var resPer = new StreamWriter(Path.Combine(_expPath, resPerile), false, Encoding.GetEncoding("Shift_Jis")))
            using (var resGrp = new StreamWriter(Path.Combine(_expPath, resGrpFile), false, Encoding.GetEncoding("Shift_Jis")))
            // 顧客情報変更データ
            using (var customer = new StreamWriter(Path.Combine(customerPath, customerFile), false, Encoding.GetEncoding("Shift_Jis")))
            // 回答顧客情報データ
            using (var resuponse = new StreamWriter(Path.Combine(ansPath, resuponseFile), false, Encoding.GetEncoding("Shift_Jis")))
            // 回答結果イメージ管理データ
            using (var answerImage = new StreamWriter(Path.Combine(ansPath, answerImageFile), false, Encoding.GetEncoding("Shift_Jis")))
            // DP連携ファイル
            using (var dp = new StreamWriter(Path.Combine(_expPath, dpFile), false, Encoding.GetEncoding("Shift_Jis")))
            {
                ProgressMax = _dantai.Rows.Count + _kojin.Rows.Count;
                ProgressValue = 0;

                foreach (string bankCode in bankCodes)
                {
                    // 金融機関コードから金融機関名を取得
                    var bankData = codeDb.ExecuteQuery($"select * from t_financial_code where code = '{bankCode}'");

                    // 金融機関名が見つからない場合は例外を投げる
                    if (bankData.Rows.Count == 0)
                    {
                        throw new Exception($"金融機関コードが見つかりません: {bankCode}");
                    }

                    // 金融機関名を取得
                    string bankName = bankData.Rows[0]["financial_name"].ToString().Trim();
                    // 支店番号を取得
                    string branchNo = bankData.Rows[0]["branch_number"].ToString().Trim();

                    // 開放してDataTableを破棄
                    bankData.Dispose();

                    // 団体と個人のDataTableから金融機関コードに一致する行を抽出
                    var dantai = _dantai.AsEnumerable()
                        .Where(row => row.Field<string>("bpo_bank_code") == bankCode);

                    var kojin = _kojin.AsEnumerable()
                        .Where(row => row.Field<string>("bpo_bank_code") == bankCode);

                    // 個人データ作成
                    SetKojin(codeDb, kojin, customer, resuponse, answerImage, resPer, branchNo);
                    // 団体データ作成
                    SetDantai(codeDb, dantai, customer, resuponse, answerImage, resGrp, branchNo);
                }
            }
        }

        /// <summary>
        /// 団体処理
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="dantaiRows"></param>
        /// <param name="customer"></param>
        /// <param name="resuponse"></param>
        /// <param name="answerImage"></param>
        /// <param name="safeBox"></param>
        /// <param name="resGrp"></param>
        /// <param name="branchNo"></param>
        /// <exception cref="Exception"></exception>
        private void SetDantai(MyDbData codeDb, EnumerableRowCollection<DataRow> dantaiRows, StreamWriter customer, StreamWriter resuponse, StreamWriter answerImage, StreamWriter resGrp, string branchNo)
        {
            // 人格コード（実質的支配者のデータが必要なコード）
            HashSet<string> personCodes = new HashSet<string> { "21", "31", "81" };

            foreach (DataRow row in dantaiRows)
            {
                ProgressValue++;

                // 業種
                string industry = string.Empty;   // 業務コード
                string bussiness = string.Empty;  // 職業コード

                string[] blz =
                {
                    row["biz_type1"].ToString().Trim(),
                    row["biz_type2"].ToString().Trim()
                };
                // 空欄を除外
                blz = blz.AsEnumerable().Where(blz => !string.IsNullOrEmpty(blz)).ToArray();

                // その他以外
                if (blz[0].ToString() != "13")
                {
                    // 業種コード取得
                    var blzRow = codeDb.ExecuteQuery($"select * from t_business_code_organization where code='{blz[0].ToString()}';");
                    // データが無ければエラー
                    if (blzRow.Rows.Count == 0) throw new Exception($"（団体）業種・職業コードが取得できません。：{blz[0].ToString()}");

                    industry = blzRow.Rows[0]["industry_code"].ToString().Trim();    // 業務コード
                    bussiness = blzRow.Rows[0]["bussiness_code"].ToString().Trim();   // 職業コード

                    // 開放
                    blz = null;
                    blzRow.Dispose();
                }
                // その他の場合
                else
                {
                    industry = "000000"; // 業務コード
                    bussiness = "000";  // 職業コード
                }

                // 本店所在国
                string hqFlg = row["hq_ctry_chg"].ToString().Trim();
                string hqCountry = row["hq_country"].ToString().Trim();
                string? hqCode = null;

                // 本店所在国に記入ありの場合
                if (!string.IsNullOrEmpty(hqCountry))
                {
                    // 国名から国コードを取得
                    hqCode = codeDb.ExecuteScalar($"select code from t_country_code where country_name = '{hqCountry}'") as string;
                    // 国コードが見つからない場合は例外を投げる
                    if (string.IsNullOrEmpty(hqCode)) throw new Exception($"（団体）本店所在国コードが見つかりません: {hqCountry}");
                }
                // 国コードの取得有無で変更あり・なしフラグを設定
                hqFlg = string.IsNullOrEmpty(hqCode) ? "0" : "1";
                // 変更有無が0の場合は空文字にする
                hqCountry = hqFlg == "0" ? string.Empty : hqCountry;

                // 取引目的コード
                string[] purpose =
                {
                    row["purp_cd1"].ToString().Trim(),
                    row["purp_cd2"].ToString().Trim(),
                    row["purp_cd3"].ToString().Trim(),
                    row["purp_cd4"].ToString().Trim(),
                    row["purp_cd5"].ToString().Trim(),
                    row["purp_cd6"].ToString().Trim()
                };
                // 取引目的コードが空でないものを抽出し、ソートして配列に変換
                purpose = purpose
                    .Where(x => !string.IsNullOrEmpty(x))
                    .OrderBy(x => x)
                    .ToArray();

                // 取引目的コードの有無フラグ　取引目的コードに値がない場合は0、ある場合は1
                string purposeFlg = purpose.Count() == 0 ? "0" : "1";
                // 取引目的コードその他のテキスト
                string purposeTxt = purpose.Contains("006") ? row["purp_other"].ToString().Trim() : string.Empty;

                // 記入日
                string answerDate = "20" + row["entry_date"].ToString().Trim();

                // 案件毎番号
                string caseNo = $"{row["bpo_bank_code"].ToString()}-{branchNo}-{ProgressValue.ToString("0000")}";

                // 結果データ
                StringBuilder record = new StringBuilder();
                record.Append(row["bpo_num"].ToString() + "\t"); // 管理番号
                record.Append(caseNo);
                resGrp.WriteLine(record.ToString());
                record.Clear(); // 開放

                // 個人の顧客情報変更データを取得
                string customerInfo = GetCustomerInfoDantai(codeDb, row, answerDate, bussiness, industry, hqFlg, hqCode, purposeFlg, purpose, purposeTxt);
                customer.WriteLine(customerInfo);
                // 個人の回答結果顧客管理データを取得
                string responseDantai = GetResuponseDantai(codeDb, row, branchNo, caseNo, answerDate, ProgressValue);
                resuponse.Write(responseDantai);
                // 個人の回答結果イメージ管理データを取得
                string answerImageDantai = GetAnswerImageDantai(codeDb, row, branchNo, caseNo, answerDate, ProgressValue);
                answerImage.Write(answerImageDantai);
            }
        }

        /// <summary>
        /// 回答結果イメージ管理データ（団体）作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="brancNo"></param>
        /// <param name="caseNo"></param>
        /// <param name="answerDate"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        private string GetAnswerImageDantai(MyDbData codeDb, DataRow row, string brancNo, string caseNo, string answerDate, int num)
        {
            string delimiter = ",";

            string record = string.Empty;
            record += row["bpo_bank_code"].ToString() + delimiter;  // 金融機関コード
            record += brancNo + delimiter;  // 支店番号
            record += caseNo + delimiter; // 案件毎番号
            record += "1" + delimiter; // 代表種類区分
            record += num.ToString("0000") + delimiter; // イメージ毎番号
            record += Convert("", 13, ' '); // 予備
            record += "\r\n"; // 改行

            // レコードデータが40Byte以外はエラー
            if (record.LenBSjis() != 40)
            {
                //throw new Exception($"（個人）個人回答結果イメージ管理データのレコード長が40Byteではありません: {record.LenBSjis()}");
            }
            return record;
        }

        /// <summary>
        /// 回答結果顧客管理データ（団体）作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="brancNo"></param>
        /// <param name="caseNo"></param>
        /// <param name="answerDate"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        private string GetResuponseDantai(MyDbData codeDb, DataRow row, string brancNo, string caseNo, string answerDate, int num)
        {
            string delimiter = ",";

            // カナ団体名
            string kanaName = string.IsNullOrEmpty(row["org_kana_new"].ToString().Trim()) ? row["bpo_kana_name"].ToString().Trim() : row["org_kana_new"].ToString().Trim();
            // 漢字団体名
            string kanjiName = string.IsNullOrEmpty(row["org_name_new"].ToString().Trim()) ? row["bpo_org_kanji"].ToString().Trim() : row["org_name_new"].ToString().Trim();

            string record = string.Empty;
            record += row["bpo_bank_code"].ToString() + delimiter;  // 金融機関コード
            record += brancNo + delimiter;  // 支店番号
            record += row["bpo_cust_no"].ToString() + delimiter;    // 顧客番号
            record += Convert(kanaName, 30, ' ') + delimiter;       // カナ団体名
            record += Convert(kanjiName, 30, '　') + delimiter;     // 漢字団体名
            record += Convert("", 8, '0') + delimiter;  // 本人確認日
            record += row["bpo_branch_no"].ToString() + delimiter;  // 顧客管理番号
            record += caseNo + delimiter; // 案件毎番号
            record += answerDate + delimiter; // 取引日
            // 設立年月日　入力値が無い場合は BPOデータの生年月日を使用
            record += (string.IsNullOrEmpty(row["est_date"].ToString().Trim()) ? row["bpo_birth_date"].ToString() : row["est_date"].ToString().Trim()) + delimiter;
            record += Convert("", 11, ' '); // 予備
            record += "\r\n"; // 改行

            // レコードデータが160Byte以外はエラー
            if (record.LenBSjis() != 160)
            {
                //throw new Exception($"（個人）回答結果顧客管理データのレコード長が160Byteではありません: {record.LenBSjis()}");
            }

            return record;
        }

        /// <summary>
        /// 顧客情報変更データ（団体）作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="answerDate"></param>
        /// <param name="bussiness"></param>
        /// <param name="industry"></param>
        /// <param name="hqFlg"></param>
        /// <param name="hqCode"></param>
        /// <param name="purposeFlg"></param>
        /// <param name="purpose"></param>
        /// <returns></returns>
        private string GetCustomerInfoDantai(MyDbData codeDb, DataRow row, string answerDate, string bussiness, string industry, string hqFlg, string? hqCode, string purposeFlg, string[] purpose, string purposeTxt)
        {
            var record = string.Empty;
            string delimiter = ",";

            // 取引目的有無フラグが0の場合はBPOデータから取得する
            if (purposeFlg == "0")
            {
                purpose = new string[]
                {
                    row["bpo_purp_1"].ToString().Trim(),
                    row["bpo_purp_2"].ToString().Trim(),
                    row["bpo_purp_3"].ToString().Trim(),
                    row["bpo_purp_4"].ToString().Trim(),
                    row["bpo_purp_5"].ToString().Trim(),
                    row["bpo_purp_6"].ToString().Trim()
                };
            }

            record += Convert(row["bpo_bank_code"].ToString(), 4, ' ') + delimiter; // 金融機関コード
            record += Convert(row["bpo_cust_no"].ToString(), 10, '0') + delimiter;  // 顧客番号
            record += Convert("", 4, '0') + delimiter; // 店番
            record += Convert("", 2, '0') + delimiter; // 科目
            record += Convert("", 7, '0') + delimiter; // 口座番号

            record += industry + delimiter; // 業種コード

            record += Convert("", 6, '0') + delimiter;  // 業種コード2
            record += Convert("", 1, '0') + delimiter;  // 性別コード
            record += Convert("", 1, '0') + delimiter;  // 世帯主区分
            record += Convert("", 2, '0') + delimiter;  // 続柄コード
            record += Convert("", 1, '0') + delimiter;  // 非居住者区分
            record += Convert("", 8, '0') + delimiter;  // 生年月日・設立年月日
            record += Convert("", 7, '0') + delimiter;  // 郵便番号
            record += Convert("", 11, '0') + delimiter; // 住所コード

            record += Convert("", 84, ' ') + delimiter; // カナ補助住所
            record += Convert("", 50, '　') + delimiter; // 漢字補助住所

            record += Convert(row["tel"].ToString().Trim(), 13, ' ') + delimiter; // 第一電話番号
            record += Convert("", 1, '0') + delimiter;  // 電話番号区分
            record += Convert("", 13, ' ') + delimiter; // 第二電話番号
            record += Convert("", 13, ' ') + delimiter; // 第三電話番号
            record += Convert("", 13, ' ') + delimiter; // FAX番号
            record += Convert("", 2, '0') + delimiter;  // 担当者コード1
            record += Convert("", 2, '0') + delimiter;  // 担当者コード2
            record += Convert("", 5, '0') + delimiter;  // 勤務先コード
            record += Convert("", 4, '0') + delimiter;  // 事業所コード
            record += Convert("", 4, '0') + delimiter;  // 登録店番
            record += Convert("", 48, ' ') + delimiter;  // カナ勤務先名
            record += Convert("", 30, '　') + delimiter; // 勤務先名
            record += Convert("", 6, '0') + delimiter;  // 年収
            record += Convert("", 2, '0') + delimiter;  // 本人確認コード

            record += Convert(purposeTxt, 20, '　') + delimiter; // 補足事項

            record += Convert("", 2, '0') + delimiter;  // 家族人数
            record += Convert("", 6, '0') + delimiter;  // 居住開始年
            record += Convert("", 8, '0') + delimiter;  // 退職日
            record += Convert("", 8, '0') + delimiter;  // 出国日
            record += Convert("", 8, '0') + delimiter;  // 国内勤務開始日
            record += Convert("", 7, '0') + delimiter;  // 国税率
            record += Convert("", 4, '0') + delimiter;  // 外為与信稟議店番
            record += Convert("", 4, '0') + delimiter;  // 外為与信元帳店番
            record += Convert("", 8, '0') + delimiter;  // 取引開始日
            record += Convert("", 1, '0') + delimiter;  // 住所コード設定有無
            record += Convert("", 8, '0') + delimiter;  // 本人確認日
            record += Convert("", 8, '0') + delimiter;  // 勧誘有効期限
            record += Convert("", 2, '0') + delimiter;  // 勧誘上限回数
            record += Convert("", 9, ' ') + delimiter;  // 融資極度額
            record += Convert("", 4, '0') + delimiter;  // 対象店番号
            record += Convert("", 2, '0') + delimiter;  // 対象科目
            record += Convert("", 7, '0') + delimiter;  // 対象口座番号
            record += Convert("", 4, '0') + delimiter;  // 居住国

            // 本店所在国　変更なしはBPOデータを取得
            record += (hqFlg == "0" ? row["bpo_nation"].ToString() : hqCode) + delimiter;

            record += "90" + delimiter; // 顧客管理事項コード
            record += answerDate + delimiter; // 顧客管理事項確認日

            record += Convert(purpose.Count() > 0 ? purpose[0] : "", 3, '0') + delimiter; // 取引目的コード1

            // 職業事業内容コード 000（その他）の場合はBPOデータから取得
            record += Convert(bussiness == "000" ? row["bpo_job_type_cd"].ToString().Trim() : bussiness, 3, '0') + delimiter;

            record += Convert("", 1, '0') + delimiter;  // 実質的支配者確認
            record += Convert("", 2, '0') + delimiter;  // 年初都道府県コード
            record += Convert("", 2, '0') + delimiter;  // 現都道府県コード
            record += Convert("", 2, '0') + delimiter;  // PEPS確認コード
            record += Convert("", 1, '0') + delimiter;  // PEPS確認方法
            record += Convert("", 8, '0') + delimiter;  // PEPS確認日
            record += Convert("", 8, '0') + delimiter;  // 高齢者ATM振込確認済日
            record += Convert("", 8, '0') + delimiter;  // 高齢者基準額超支払確認済日

            record += Convert(purpose.Count() > 1 ? purpose[1] : "", 3, '0') + delimiter; // 取引目的コード2
            record += Convert(purpose.Count() > 2 ? purpose[2] : "", 3, '0') + delimiter; // 取引目的コード3
            record += Convert(purpose.Count() > 3 ? purpose[3] : "", 3, '0') + delimiter; // 取引目的コード4
            record += Convert(purpose.Count() > 4 ? purpose[4] : "", 3, '0') + delimiter; // 取引目的コード5
            record += Convert(purpose.Count() > 5 ? purpose[5] : "", 3, '0') + delimiter; // 取引目的コード6

            record += Convert("", 1, '0') + delimiter;  // 質問対象者
            record += Convert("", 8, '0') + delimiter;  // 質問送付日
            record += Convert("", 1, '0') + delimiter;  // 質問送付媒体区分

            record += answerDate + delimiter; // 質問回答日

            record += "1" + delimiter; // 顧客属性データ内訳区分
            record += Convert("", 10, '　') + delimiter; // その他業種名称
            record += Convert("", 14, ' ');  // FILLER

            // レコードデータが700Byte以外はエラー
            if (record.LenBSjis() != 700)
            {
                //throw new Exception($"（個人）顧客情報変更データのレコード長が700Byteではありません: {record.LenBSjis()}");
            }

            return record;
        }

        /// <summary>
        /// 個人データ作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="kojinRows"></param>
        /// <param name="customer"></param>
        /// <param name="resuponse"></param>
        /// <param name="answerImage"></param>
        /// <param name="safeBox"></param>
        /// <param name="branchNo"></param>
        /// <exception cref="Exception"></exception>
        private void SetKojin(MyDbData codeDb, EnumerableRowCollection<DataRow> kojinRows, StreamWriter customer, StreamWriter resuponse, StreamWriter answerImage, StreamWriter resPer, string branchNo)
        {
            // 業種その他コード
            string industryEtc = MyUtilityModules.AppSetting("roukin_setting", "industry_etc_code");

            foreach (DataRow row in kojinRows)
            {
                ProgressValue++;

                // 国籍
                string nationCode = string.Empty;
                // 変更あり
                if (row["nationchg_flg"].ToString() == "2")
                {
                    string region = string.Empty;
                    string country = string.Empty;

                    if (row["nation"].ToString() == "1")
                    {
                        region = "1";   // アジア
                        country = "0";  // 日本
                    }
                    else
                    {
                        // 国地域
                        region = row["region"].ToString();
                        // 国籍
                        country = row["region"].ToString() switch
                        {
                            "1" => row["nation_asia"].ToString(),   // アジア
                            "2" => row["nation_mide"].ToString(),   // 中近東
                            "3" => row["nation_weur"].ToString(),   // 西欧
                            "4" => row["nation_eeur"].ToString(),   // 東欧
                            "5" => row["nation_nam"].ToString(),    // 北米
                            "6" => row["nation_cari"].ToString(),   // カリブ 
                            "7" => row["nation_latm"].ToString(),   // 南米    
                            "8" => row["nation_afrc"].ToString(),   // アフリカ
                            "9" => row["nation_oce"].ToString(),    // 大洋州
                            "10" => row["nation_othr"].ToString(),  // その他
                            _ => throw new Exception("不明な地域コードです。")
                        };
                    }
                    // 国籍コードを取得
                    nationCode = codeDb.ExecuteScalar($"select code from t_country_code_person where region_code = '{region}' and country_code = '{country}'") as string;

                    // 国籍コードが見つからない場合は例外を投げる
                    if (string.IsNullOrEmpty(nationCode)) throw new Exception($"（個人）国籍コードが見つかりません: region={region}, country={country}");
                }

                // 取引目的コードを配列で取得
                string[] purpose =
                {
                row["tx_purpose1"].ToString(),
                row["tx_purpose2"].ToString(),
                row["tx_purpose3"].ToString(),
                row["tx_purpose4"].ToString(),
                row["tx_purpose5"].ToString(),
                row["tx_purpose6"].ToString() == "-999" ? "006" : ""
                };

                // 取引目的コードが空でないものを抽出し、ソートして配列に変換
                purpose = purpose
                    .Where(x => !string.IsNullOrEmpty(x))
                    .OrderBy(x => x)
                    .ToArray();

                // 取引目的コードが設定されていない場合は例外を投げる
                if (purpose.Count() == 0) throw new Exception("（個人）取引目的コードが設定されていません。");

                // 取引目的コードに"006"が含まれている場合は、その他の取引目的を取得
                string purposeTxt = purpose.Contains("006") ? row["tx_othr_txt"].ToString() : string.Empty;

                // 固定電話番号
                string telFiexed = row["tel_fixed"].ToString();
                // 携帯電話番号
                string telMobile = row["tel_mobile"].ToString();
                // 第一電話番号と第三電話番号の設定
                string tel1st = string.IsNullOrEmpty(telFiexed) ? telMobile : telFiexed;
                string tel3rd = string.IsNullOrEmpty(telFiexed) ? string.Empty : telMobile;

                // 業種コード
                string industry = row["industry"].ToString();
                // その他業種名称
                string industryTxt = string.Empty;
                // 業種コードが-999の場合はその他業種コードを設定
                industry = industry == "-999" ? industryEtc : industry;
                // 業種コードがその他の場合は、業種名称を取得
                industryTxt = industry == industryEtc ? row["indus_othr"].ToString() : string.Empty; // その他業種名称

                // 職業コード
                string job = row["job"].ToString() == "-999" ? "010" : row["job"].ToString(); // 職業コード
                // 職業コードがその他の場合は、職業名称を取得
                string jobTxt = job == "010" ? row["job_other"].ToString() : string.Empty; // その他職業名称

                // 回答日が日時に変換できない場合はエラー
                if (DateTime.TryParse(row["answer_date"].ToString(), out DateTime parsedDate) == false)
                {
                    throw new Exception($"（個人）回答日が不正です: {row["answer_date"].ToString()}");
                }

                // 案件毎番号
                string caseNo = $"{row["bpo_bank_code"].ToString()}-{branchNo}-{ProgressValue.ToString("0000")}";

                // 結果データ
                StringBuilder record = new StringBuilder();
                record.Append(row["bpo_num"].ToString() + "\t"); // 管理番号
                record.Append(caseNo);
                resPer.WriteLine(record.ToString());
                record.Clear(); // 開放

                // 個人の顧客情報変更データを取得
                string customerInfo = GetCustomerInfoKojin(codeDb, row, parsedDate, industryEtc, industry, industryTxt, nationCode, purpose, purposeTxt, tel1st, tel3rd, job, jobTxt);
                customer.WriteLine(customerInfo);
                // 個人の回答結果顧客管理データを取得
                string responseKojin = GetResuponseKojin(codeDb, row, branchNo, caseNo, parsedDate, ProgressValue);
                resuponse.Write(responseKojin);
                // 個人の回答結果イメージ管理データを取得
                string answerImageKojin = GetAnswerImageKojin(codeDb, row, branchNo, caseNo, parsedDate, ProgressValue);
                answerImage.Write(answerImageKojin);
            }
        }

        /// <summary>
        /// 回答結果イメージ管理データ（個人）作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="brancNo"></param>
        /// <param name="caseNo"></param>
        /// <param name="parsedDate"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        private string GetAnswerImageKojin(MyDbData codeDb, DataRow row, string brancNo, string caseNo, DateTime parsedDate, int num)
        {
            string delimiter = "";

            string record = string.Empty;
            record += row["bpo_bank_code"].ToString() + delimiter;  // 金融機関コード
            record += brancNo + delimiter;  // 支店番号
            record += caseNo + delimiter; // 案件毎番号
            record += "1" + delimiter; // 代表種類区分
            record += num.ToString("0000") + delimiter; // イメージ毎番号
            record += Convert("", 13, ' '); // 予備
            record += "\r\n"; // 改行

            // レコードデータが40Byte以外はエラー
            if (record.LenBSjis() != 40)
            {
                throw new Exception($"（個人）個人回答結果イメージ管理データのレコード長が40Byteではありません: {record.LenBSjis()}");
            }
            return record;
        }

        /// <summary>
        /// 回答結果顧客管理データ（個人）作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="brancNo"></param>
        /// <param name="caseNo"></param>
        /// <param name="parsedDate"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        private string GetResuponseKojin(MyDbData codeDb, DataRow row, string brancNo, string caseNo, DateTime parsedDate, int num)
        {
            string delimiter = ",";

            // カナ氏名
            string kanaName = string.Empty;
            // 漢字氏名
            string kanjiName = string.Empty;

            // 氏名変更有無
            if (row["namechg_flg"].ToString() == "2")
            {
                // 氏名変更ありの場合
                kanaName = row["lname_kana"].ToString() + " " + row["fname_kana"].ToString();
                kanjiName = row["lname_kanji"].ToString() + "　" + row["fname_kanji"].ToString();
            }
            else
            {
                // 氏名変更なしの場合
                kanaName = row["bpo_kana_name"].ToString();
                kanjiName = row["bpo_org_kanji"].ToString();
            }

            string record = string.Empty;
            record += row["bpo_bank_code"].ToString() + delimiter;  // 金融機関コード
            record += brancNo + delimiter;  // 支店番号
            record += row["bpo_cust_no"].ToString() + delimiter;    // 顧客番号
            record += Convert(RokinModules.ReplaceSmallKanaWithLargeKana(kanaName), 30, ' ') + delimiter;   // カナ氏名
            record += Convert(kanjiName, 30, '　') + delimiter;     // 漢字氏名
            record += Convert("", 8, '0') + delimiter;  // 本人確認日
            record += row["bpo_branch_no"].ToString() + delimiter;  // 顧客管理番号
            record += caseNo + delimiter; // 案件毎番号
            record += parsedDate.ToString("yyyyMMdd") + delimiter; // 取引日
            record += row["bpo_birth_date"].ToString() + delimiter; // 生年月日
            record += Convert("", 11, ' '); // 予備
            record += "\r\n"; // 改行

            // レコードデータが160Byte以外はエラー
            if (record.LenBSjis() != 160)
            {
                //throw new Exception($"（個人）回答結果顧客管理データのレコード長が160Byteではありません: {record.LenBSjis()}");
            }

            return record;
        }

        /// <summary>
        /// 顧客情報変更データ（個人）作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="industryEtc"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private string GetCustomerInfoKojin(MyDbData codeDb, DataRow row, DateTime parsedDate, string industryEtc, string industry, string industryTxt,
                        string nationCode, string[] sortedPurpose, string purposeTxt, string tel1st, string tel3rd, string job, string jobTxt)
        {
            var record = string.Empty;
            string delimiter = "";

            // BPOデータの業種コードもその他の場合はzeroで埋める
            if (industry == industryEtc && industry == row["bpo_biz_type_cd"].ToString())
            {
                industry = Convert("", 6, '0'); // 業種コード
            }

            // 補足事項
            string remarks = string.Empty;
            // 取引目的コードと職業コードが両方とも空でない場合は、／区切りで結合
            if (!string.IsNullOrEmpty(purposeTxt) && !string.IsNullOrEmpty(jobTxt))
            {
                remarks = $"{purposeTxt.Left(10)}／{jobTxt.Left(9)}";
            }
            else
            {
                // 片方が空の場合は、もう片方のテキストをそのまま使用
                remarks = purposeTxt + jobTxt;
            }

            record += Convert(row["bpo_bank_code"].ToString(), 4, ' ') + delimiter; // 金融機関コード
            record += Convert(row["bpo_cust_no"].ToString(), 10, ' ') + delimiter;      // 顧客番号
            record += Convert("", 4, '0') + delimiter; // 店番
            record += Convert("", 2, '0') + delimiter; // 科目
            record += Convert("", 7, '0') + delimiter; // 口座番号

            record += industry + delimiter; // 業種コード

            record += Convert("", 6, '0') + delimiter;  // 業種コード2
            record += Convert("", 1, '0') + delimiter;  // 性別コード
            record += Convert("", 1, '0') + delimiter;  // 世帯主区分
            record += Convert("", 2, '0') + delimiter;  // 続柄コード
            record += Convert("", 1, '0') + delimiter;  // 非居住者区分
            record += Convert("", 8, '0') + delimiter;  // 生年月日・設立年月日
            record += Convert("", 7, '0') + delimiter;  // 郵便番号
            record += Convert("", 11, '0') + delimiter; // 住所コード

            record += Convert("", 84, ' ') + delimiter; // カナ補助住所
            record += Convert("", 50, '　') + delimiter; // 漢字補助住所

            record += Convert(tel1st, 13, ' ') + delimiter; // 第一電話番号
            record += Convert("", 1, '0') + delimiter;      // 電話番号区分
            record += Convert(row["work_tel"].ToString(), 13, ' ') + delimiter; // 第二電話番号
            record += Convert(tel3rd, 13, ' ') + delimiter; // 第三電話番号

            record += Convert("", 13, ' ') + delimiter; // FAX番号
            record += Convert("", 2, '0') + delimiter;  // 担当者コード1
            record += Convert("", 2, '0') + delimiter;  // 担当者コード2
            record += Convert("", 5, '0') + delimiter;  // 勤務先コード
            record += Convert("", 4, '0') + delimiter;  // 事業所コード
            record += Convert("", 4, '0') + delimiter;  // 登録店番

            record += Convert(RokinModules.ReplaceSmallKanaWithLargeKana(row["worksch_kan"].ToString()), 48, ' ') + delimiter; // カナ勤務先名
            record += Convert(row["work_or_sch"].ToString(), 30, '　') + delimiter; // 勤務先名

            record += Convert("", 6, '0') + delimiter;  // 年収
            record += Convert("", 2, '0') + delimiter;  // 本人確認コード

            record += Convert(remarks, 20, '　') + delimiter; // 補足事項

            record += Convert("", 2, '0') + delimiter;  // 家族人数
            record += Convert("", 6, '0') + delimiter;  // 居住開始年
            record += Convert("", 8, '0') + delimiter;  // 退職日
            record += Convert("", 8, '0') + delimiter;  // 出国日
            record += Convert("", 8, '0') + delimiter;  // 国内勤務開始日
            record += Convert("", 7, '0') + delimiter;  // 国税率
            record += Convert("", 4, '0') + delimiter;  // 外為与信稟議店番
            record += Convert("", 4, '0') + delimiter;  // 外為与信元帳店番
            record += Convert("", 8, '0') + delimiter;  // 取引開始日
            record += Convert("", 1, '0') + delimiter;  // 住所コード設定有無
            record += Convert("", 8, '0') + delimiter;  // 本人確認日
            record += Convert("", 8, '0') + delimiter;  // 勧誘有効期限
            record += Convert("", 2, '0') + delimiter;  // 勧誘上限回数
            record += Convert("", 9, ' ') + delimiter;  // 融資極度額
            record += Convert("", 4, '0') + delimiter;  // 対象店番号
            record += Convert("", 2, '0') + delimiter;  // 対象科目
            record += Convert("", 7, '0') + delimiter;  // 対象口座番号
            record += Convert("", 4, '0') + delimiter;  // 居住国

            // 本店所在国　変更なしはBPOデータを取得
            record += Convert((row["nationchg_flg"].ToString() == "1" ? row["bpo_nation"].ToString().Trim() : nationCode), 4, '0') + delimiter;

            record += "90" + delimiter; // 顧客管理事項コード
            record += parsedDate.ToString("yyyyMMdd") + delimiter; // 顧客管理事項確認日

            record += (sortedPurpose.Count() > 0 ? sortedPurpose[0] : string.Empty) + delimiter; // 取引目的コード

            record += job + delimiter; // 職業事業内容コード

            record += Convert("", 1, '0') + delimiter;  // 実質的支配者確認
            record += Convert("", 2, '0') + delimiter;  // 年初都道府県コード
            record += Convert("", 2, '0') + delimiter;  // 現都道府県コード

            record += row["pep_flag"].ToString() + delimiter; // PEPS確認コード
            record += "1" + delimiter; // PEPS確認方法
            record += parsedDate.ToString("yyyyMMdd") + delimiter; // PEPS確認日

            record += Convert("", 8, '0') + delimiter;  // 高齢者ATM振込確認済日
            record += Convert("", 8, '0') + delimiter;  // 高齢者基準額超支払確認済日

            record += (sortedPurpose.Count() > 1 ? sortedPurpose[1] : Convert("", 3, '0')) + delimiter; // 取引目的コード2
            record += (sortedPurpose.Count() > 2 ? sortedPurpose[2] : Convert("", 3, '0')) + delimiter; // 取引目的コード3
            record += (sortedPurpose.Count() > 3 ? sortedPurpose[3] : Convert("", 3, '0')) + delimiter; // 取引目的コード4
            record += (sortedPurpose.Count() > 4 ? sortedPurpose[4] : Convert("", 3, '0')) + delimiter; // 取引目的コード5
            record += (sortedPurpose.Count() > 5 ? sortedPurpose[5] : Convert("", 3, '0')) + delimiter; // 取引目的コード6

            record += Convert("", 1, '0') + delimiter;  // 質問対象者
            record += Convert("", 8, '0') + delimiter;  // 質問送付日
            record += Convert("", 1, '0') + delimiter;  // 質問送付媒体区分

            record += parsedDate.ToString("yyyyMMdd") + delimiter; // 質問回答日
            record += "1" + delimiter; // 顧客属性データ内訳区分
            record += Convert(industryTxt, 10, '　') + delimiter; // その他業種名称
            record += Convert("", 14, ' ');  // FILLER

            // レコードデータが700Byte以外はエラー
            if (record.LenBSjis() != 700)
            {
                //throw new Exception($"（個人）顧客情報変更データのレコード長が700Byteではありません: {record.LenBSjis()}");
            }

            return record;
        }

        /// <summary>
        /// 文字列が指定の長さに満たない場合、指定の文字で埋める
        /// 指定の長さを超えている場合は、指定の文字で切り詰める
        /// </summary>
        /// <param name="value"></param>
        /// <param name="mojiByte"></param>
        /// <param name="lenght"></param>
        /// <returns></returns>
        public string Convert(string value, int lenght, char replaceChr)
        {
            if (value.Length > lenght)
            {
                return value.Substring(0, lenght);
            }
            else
            {
                return value.PadRight(lenght, replaceChr);
            }
        }

        /// <summary>
        /// 団体と個人のDataTableから緒副なしで金融機関コードを取得
        /// </summary>
        /// <returns></returns>
        private List<string> GetBankCodes()
        {
            var dantaiCodes = _dantai.AsEnumerable()
                .Select(row => row.Field<string>("bpo_bank_code"))
                .Where(code => !string.IsNullOrEmpty(code));

            var kojinCodes = _kojin.AsEnumerable()
                .Select(row => row.Field<string>("bpo_bank_code"))
                .Where(code => !string.IsNullOrEmpty(code));

            var result = dantaiCodes
                .Concat(kojinCodes)
                .Distinct()
                .ToList();

            return result;
        }
    }
}

