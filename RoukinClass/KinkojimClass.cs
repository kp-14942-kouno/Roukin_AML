using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.RoukinClass
{
    public class KinkojimClass : MyLibrary.MyLoading.Thread
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
        public KinkojimClass(DataTable dantai, DataTable kojin, string expPath)
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
                MyLogger.SetLogger("金庫事務用データ作成を開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {   
                    // 実行呼出し
                    Run(codeDb);

                    // 結果メッセージ
                    ResultMessage = $"金庫事務用データ：{_kojin.Rows.Count + _dantai.Rows.Count}件\r\n作成完了";
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
            // 金融機関コードをリスト化
            List<string> bankCodes = GetBankCodes();

            // 金庫事務用データファイル名
            string safeBoxFile = MyUtilityModules.AppSetting("roukin_setting", "safe_box_admin_name", true);
            // 金庫事務用ディレクトリ
            string safeBoxDir = MyUtilityModules.AppSetting("roukin_setting", "safe_box_admin_dir", true);

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

                // 金融機関名のディレクトリを作成
                string bankDir = Path.Combine(_expPath, safeBoxDir, bankCode + bankName);
                Directory.CreateDirectory(bankDir);

                // 団体と個人のDataTableから金融機関コードに一致する行を抽出
                var dantai = _dantai.AsEnumerable()
                    .Where(row => row.Field<string>("bpo_bank_code") == bankCode);

                var kojin = _kojin.AsEnumerable()
                    .Where(row => row.Field<string>("bpo_bank_code") == bankCode);

                ProgressMax = dantai.Count() + kojin.Count();
                ProgressValue = 0;

                // 金庫事務用データ
                using (var safeBox = new StreamWriter(Path.Combine(bankDir, safeBoxFile), false, Encoding.GetEncoding("Shift_Jis")))
                {
                    SetKojin(codeDb, kojin, safeBox, branchNo);
                    SetDantai(codeDb, dantai, safeBox, branchNo);
                }
            }
        }
        
        /// <summary>
        /// 団体処理
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="dantaiRows"></param>
        /// <param name="safeBox"></param>
        /// <param name="branchNo"></param>
        /// <exception cref="Exception"></exception>
        private void SetDantai(MyDbData codeDb, EnumerableRowCollection<DataRow> dantaiRows, StreamWriter safeBox, string branchNo)
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

                // 団体の金庫事務用データを取得
                string safeBoxDantai = GetSafeBoxDantai(codeDb, row, branchNo, caseNo, answerDate, hqFlg, hqCountry, hqCode, bussiness, industry, purposeFlg, purpose, purposeTxt, personCodes); ;
                safeBox.WriteLine(safeBoxDantai);
            }
        }

        /// <summary>
        /// 金庫事務用データ（団体）
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="brancNo"></param>
        /// <param name="caseNo"></param>
        /// <param name="answerDate"></param>
        /// <param name="hqFlg"></param>
        /// <param name="hqCounty"></param>
        /// <param name="hqCode"></param>
        /// <param name="busisness"></param>
        /// <param name="industry"></param>
        /// <param name="purposeFlg"></param>
        /// <param name="purpose"></param>
        /// <param name="purposeTxt"></param>
        /// <param name="personCodes"></param>
        /// <returns></returns>
        private string GetSafeBoxDantai(MyDbData codeDb, DataRow row, string brancNo, string caseNo, string answerDate, string hqFlg, string hqCounty, string hqCode,
                            string busisness, string industry, string purposeFlg, string[] purpose, string purposeTxt, HashSet<string> personCodes)
        {
            string delimiter = ",";
            string record = string.Empty;

            record += row["bpo_bank_code"].ToString() + delimiter;  // 金融機関コード
            record += row["bpo_cust_no"].ToString() + delimiter;  // 顧客番号
            record += row["bpo_member_no"].ToString() + delimiter;  // 会員番号
            record += row["bpo_persona_cd"].ToString() + delimiter;  // 人格コード
            record += row["bpo_kana_name"].ToString() + delimiter;  // カナ氏名
            record += row["bpo_org_kanji"].ToString() + delimiter;  // 漢字氏名

            record += answerDate + delimiter; // 回答日
            record += "" + delimiter;  // メールアドレス

            // 氏名変更有無（団体名変更有無）　漢字団体目に記入がなければ0、あれば1
            record += (string.IsNullOrEmpty(row["org_name_new"].ToString().Trim()) ? "0" : "1") + delimiter;
            record += row["org_kana_new"].ToString().Trim() + delimiter; // カナ氏名（団体名）
            record += row["org_name_new"].ToString().Trim() + delimiter; // 漢字氏名（団体名）


            // 住所変更有無判定用
            StringBuilder addr = new StringBuilder();
            addr.Append(row["pref_new"].ToString().Trim()); // 都道府県
            addr.Append(row["city_new"].ToString().Trim()); // 市区町村
            addr.Append(row["addr1_new"].ToString().Trim()); // 丁目・番地
            addr.Append(row["addr2_new"].ToString().Trim()); // マンション名・部屋番号

            // 住所変更有無  漢字住所に記入がなければ0、あれば1
            record += (addr.Length == 0 ? "0" : "1") + delimiter;
            addr.Clear(); // 開放

            record += row["zip_new"].ToString().Trim() + delimiter;     // 郵便番号
            record += row["pref_new"].ToString().Trim() + delimiter;    // 都道府県
            record += row["city_new"].ToString().Trim() + delimiter;    // 市区町村
            record += row["addr1_new"].ToString().Trim() + delimiter;   // 丁目・番地
            record += row["addr2_new"].ToString().Trim() + delimiter;   // マンション名・部屋番号
            record += row["kana_addr"].ToString().Trim() + delimiter;   // カナ住所

            record += row["tel"].ToString().Trim() + delimiter;　// 第一電話番号
            record += "" + delimiter;   // 第二電話番号
            record += "" + delimiter;   // 第三電話番号

            // 国籍変更有無
            record += hqFlg + delimiter;
            // 国籍・国名
            record += hqCounty + delimiter; // 国籍・国名

            record += "" + delimiter; // 国籍・アルファベット氏名
            record += "" + delimiter; // 在留資格
            record += "" + delimiter; // 在留期限
            record += "" + delimiter; // PEPS該否

            // 職業事業内容コード　その他（000）の場合はBPOデータの職業事業内容コードを使用
            record += (busisness.ToString() == "000" ? row["bpo_job_type_cd"].ToString() : busisness) + delimiter;

            record += "" + delimiter; // 職業その他
            record += "" + delimiter;   // 勤務先変更有無
            record += "" + delimiter; // 勤務先名
            record += "" + delimiter; // 勤務先名カナ

            // 職業　業種コードが000000（その他）の場合はnull
            record += (industry == "000000" ? "" : industry) + delimiter;
            // 業種その他　業種コードが000000（その他）の場合は入力値を使用
            record += (industry == "000000" ? row["biz_other"].ToString().Trim() : "") + delimiter;
            // 主な製品・サービス　業種コードが000000（その他）の場合は入力値を使用
            record += (industry == "000000" ? row["product_srv"].ToString().Trim() : "") + delimiter;

            record += "" + delimiter; // 副業有無
            record += "" + delimiter; // 副業の業種
            record += "" + delimiter; // 副業の業種その他

            record += purposeFlg + delimiter; // 取引目的コード変更有無
            record += (purpose.Count() > 0 ? purpose[0] : "") + delimiter; // 取引目的コード1
            record += (purpose.Count() > 1 ? purpose[1] : "") + delimiter; // 取引目的コード2
            record += (purpose.Count() > 2 ? purpose[2] : "") + delimiter; // 取引目的コード3
            record += (purpose.Count() > 3 ? purpose[3] : "") + delimiter; // 取引目的コード4
            record += (purpose.Count() > 4 ? purpose[4] : "") + delimiter; // 取引目的コード5
            record += (purpose.Count() > 5 ? purpose[5] : "") + delimiter; // 取引目的コード6
            record += purposeTxt + delimiter; // 取引目的コードその他

            // 取引形態
            string[] dealType =
            {
                row["deal_type1"].ToString().Trim(),
                row["deal_type2"].ToString().Trim()
            };
            // 空欄を除外
            dealType = dealType.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();
            record += dealType[0].ToString().Trim() + delimiter; // 取引形態
            dealType = null; // 開放

            // 取引頻度
            string[] dealFreq =
            {
                row["deal_freq1"].ToString().Trim(),
                row["deal_freq2"].ToString().Trim()
            };
            // 空欄を除外
            dealFreq = dealFreq.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();
            record += dealFreq[0].ToString().Trim() + delimiter; // 取引頻度
            dealFreq = null; // 開放

            // 取引金額
            string[] dealAmt =
            {
                row["deal_amt1"].ToString().Trim(),
                row["deal_amt2"].ToString().Trim()
            };
            // 空欄を除外
            dealAmt = dealAmt.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();
            record += dealAmt[0].ToString().Trim() + delimiter; // 1回あたり取引金額
            dealAmt = null; // 開放

            // 200万円超取引の頻度
            string[] cashFreq =
            {
                row["cash_freq1"].ToString().Trim(),
                row["cash_freq2"].ToString().Trim()
            };
            // 空欄を除外
            cashFreq = cashFreq.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 200万円超取引の金額
            string[] cashAmt =
            {
                row["cash_amt1"].ToString().Trim(),
                row["cash_amt2"].ToString().Trim()
            };
            // 空欄を除外
            cashAmt = cashAmt.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 200万円超現金取引の原資
            string[] cashSrc =
            {
                row["cash_src1"].ToString().Trim(),
                row["cash_src2"].ToString().Trim(),
                row["cash_src3"].ToString().Trim(),
                row["cash_src4"].ToString().Trim()
            };
            // 空欄を除外
            cashSrc = cashSrc.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();

            // 200万円超現金取引の原資その他
            string cashSrcOther = cashSrc.Contains("9") ? row["cash_src_oth"].ToString().Trim() : string.Empty;
            // 選択数
            int cash = cashFreq.Count() + cashAmt.Count() + cashSrc.Count();
            // 200万円超現金取引の選択フラグ
            string cashFlg = cash == 0 ? "0" : "1";

            record += cashFlg + delimiter; // 200万円超取引の有無
            record += (cashFreq.Count() > 0 ? cashFreq[0].ToString().Trim() : "") + delimiter;  // 200万円超取引の頻度
            record += (cashAmt.Count() > 0 ? cashAmt[0].ToString().Trim() : "") + delimiter;    // 200万円超取引の金額
            record += (cashSrc.Count() > 0 ? cashSrc[0].ToString() : "") + delimiter; // 200万円超現金取引の原資
            record += (cashSrc.Count() > 1 ? cashSrc[1].ToString() : "") + delimiter; // 200万円超現金取引の原資2
            record += (cashSrc.Count() > 2 ? cashSrc[2].ToString() : "") + delimiter; // 200万円超現金取引の原資3
            record += cashSrcOther + delimiter; // 200万円超現金取引の原資その他

            // 団体種類
            string[] orgType =
            {
                row["org_type1"].ToString().Trim(),
                row["org_type2"].ToString().Trim()
            };
            // 空欄を除外
            orgType = orgType.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();
            record += (orgType.Count() > 0 ? orgType[0].ToString().Trim() : "") + delimiter; // 団体種類
            orgType = null; // 開放

            // 設立年月日変更有無　設立年月日に記入がなければ0、あれば1
            record += (string.IsNullOrEmpty(row["est_date"].ToString().Trim()) ? "0" : "1") + delimiter;
            record += row["est_date"].ToString().Trim() + delimiter; // 設立年月日

            record += hqCounty + delimiter; // 本社所在国名

            // 代表者の漢字氏名変更有無　代表者の漢字締名に記入がなければ0、あれば1
            record += (string.IsNullOrEmpty(row["rep_name"].ToString().Trim()) ? "0" : "1") + delimiter;
            record += row["rep_name"].ToString().Trim() + delimiter; // 代表者の漢字氏名
            record += row["rep_kana"].ToString().Trim() + delimiter; // 代表者のカナ氏名
            record += row["rep_bday"].ToString().Trim() + delimiter; // 代表者の生年月日
            record += row["rep_title"].ToString().Trim() + delimiter; // 代表者の役職

            // 代表者の国籍・国名　国名が空欄の場合は「日本国」を設定
            record += (string.IsNullOrEmpty(row["rep_natname"].ToString().Trim()) ? "日本国" : row["rep_natname"].ToString().Trim()) + delimiter;
            record += row["rep_alpha"].ToString().Trim() + delimiter; // 代表者の国籍・アルファベット氏名
            record += row["rep_visa"].ToString().Trim() + delimiter; // 代表者の国籍・在留資格
            record += row["rep_exp"].ToString().Trim() + delimiter; // 代表者の国籍・在留期限

            // 実質的支配者1～3
            for (int ubo = 1; ubo <= 3; ubo++)
            {
                if (personCodes.Contains(row["bpo_persona_cd"].ToString().Trim()))
                {
                    record += row[$"ubo{ubo}_name"].ToString().Trim() + delimiter; // 実質的支配者の氏名
                    record += row[$"ubo{ubo}_kana"].ToString().Trim() + delimiter; // 実質的支配者のカナ氏名
                    record += row[$"ubo{ubo}_bday"].ToString().Trim() + delimiter; // 実質的支配者の生年月日

                    // 団体との関係
                    string[] rel =
                    {
                    row[$"ubo{ubo}_rel1"].ToString().Trim(),
                    row[$"ubo{ubo}_rel2"].ToString().Trim()
                    };
                    // 空欄を除外
                    rel = rel.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    record += (rel.Count() > 0 ? rel[0].ToString().Trim() : "") + delimiter; // 実質的支配者の団体との関係
                    rel = null; // 開放

                    record += row[$"ubo{ubo}_addr"].ToString().Trim() + delimiter; // 実質的支配者の住所

                    // 事業内容
                    string[] job =
                    {
                    row[$"ubo{ubo}_job1"].ToString().Trim(),
                    row[$"ubo{ubo}_job2"].ToString().Trim()
                    };
                    // 空欄を除外
                    job = job.AsEnumerable().Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    record += (job.Count() > 0 ? job[0].ToString().Trim() : "") + delimiter; // 実質的支配者の職業・事業内容
                    job = null; // 開放

                    // 実質的支配者のPEPS該否 空欄は0、それ以外は右から1桁目を取得
                    // 0 => 0, 1 => 1, 01 => 1
                    string peps = row[$"ubo{ubo}_peps"].ToString().Trim();
                    record += (string.IsNullOrEmpty(peps) ? "0" : peps.Right(1)) + delimiter;

                    // 実質的支配者の国籍・国名　国名が空欄の場合は「日本国」を設定
                    record += (string.IsNullOrEmpty(row[$"ubo{ubo}_natname"].ToString().Trim()) ? "日本国" : row[$"ubo{ubo}_natname"].ToString().Trim()) + delimiter;

                    record += row[$"ubo{ubo}_alpha"].ToString().Trim() + delimiter; // 実質的支配者のアルファベット氏名
                    record += row[$"ubo{ubo}_visa"].ToString().Trim() + delimiter; // 実質的支配者の在留資格
                    record += row[$"ubo{ubo}_exp"].ToString().Trim() + delimiter; // 実質的支配者の在留期限
                }
                else
                {
                    // 人格コードが21,31,81以外の場合は空欄を設定
                    record += "" + delimiter; // 実質的支配者の氏名
                    record += "" + delimiter; // 実質的支配者のカナ氏名
                    record += "" + delimiter; // 実質的支配者の生年月日
                    record += "" + delimiter; // 実質的支配者の団体との関係
                    record += "" + delimiter; // 実質的支配者の住所
                    record += "" + delimiter; // 実質的支配者の職業・事業内容
                    record += "" + delimiter; // 実質的支配者のPEPS該否
                    record += "" + delimiter; // 実質的支配者の国籍・国名
                    record += "" + delimiter; // 実質的支配者のアルファベット氏名
                    record += "" + delimiter; // 実質的支配者の在留資格
                    record += "" + delimiter; // 実質的支配者の在留期限
                }
            }

            record += row["staff_name"].ToString().Trim() + delimiter; // 取引担当者の氏名
            record += row["staff_kana"].ToString().Trim() + delimiter; // 取引担当者のカナ氏名
            record += row["staff_dept"].ToString().Trim() + delimiter; // 取引担当者の部署
            record += row["staff_title"].ToString().Trim() + delimiter; // 取引担当者の役職
            record += row["staff_tel"].ToString().Trim() + delimiter; // 取引担当者の電話番号
            record += row["staff_mail"].ToString().Trim() + delimiter; // 取引担当者のメールアドレス
            record += row["staff_zip"].ToString().Trim() + delimiter; // 取引担当者の郵便番号
            record += row["staff_addr"].ToString().Trim() + delimiter; // 取引担当者の住所
            record += row["staff_kanaad"].ToString().Trim() + delimiter; // 取引担当者のカナ住所

            record += row["bpo_ship_round"].ToString() + delimiter; // 発送希望回次
            record += row["bpo_zip_code"].ToString() + delimiter; // 変更前郵便番号
            record += row["bpo_address"].ToString() + delimiter; // 変更前住所（所在地）
            record += row["bpo_home_tel"].ToString() + delimiter; // 変更前第一電話番号
            record += row["bpo_work_tel"].ToString() + delimiter; // 変更前第二電話番号
            record += row["bpo_mobile_tel"].ToString() + delimiter; // 変更間第三電話番号
            record += row["bpo_nation"].ToString() + delimiter; // 変更前国籍
            record += row["bpo_visa_type"].ToString() + delimiter; // 変更前在留資格
            record += row["bpo_visa_exp"].ToString() + delimiter; // 変更前在留期限
            record += row["bpo_job_type_cd"].ToString() + delimiter; // 変更前職業事業内容コード
            record += row["bpo_biz_type_cd"].ToString() + delimiter; // 変更前業種コード
            record += row["bpo_biz_name_etc"].ToString() + delimiter; // 変更前その他業種名称

            record += row["bpo_purp_1"].ToString() + delimiter; // 変更前取引目的1
            record += row["bpo_purp_2"].ToString() + delimiter; // 変更前取引目的2
            record += row["bpo_purp_3"].ToString() + delimiter; // 変更前取引目的3
            record += row["bpo_purp_4"].ToString() + delimiter; // 変更前取引目的4
            record += row["bpo_purp_5"].ToString() + delimiter; // 変更前取引目的5
            record += row["bpo_purp_6"].ToString() + delimiter; // 変更前取引目的6
            record += row["bpo_birth_date"].ToString() + delimiter; // 変更間生年月日（設立年月日）
            record += row["bpo_kanji_2"].ToString() + delimiter; // 変更前漢字氏名2
            record += row["bpo_role_kanji"].ToString() + delimiter; // 変更前漢字役職名
            record += row["bpo_rep_kanji"].ToString(); // 変更前代表者の漢字氏名

            return record;
        }

        /// <summary>
        /// 個人データ作成
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="kojinRows"></param>
        /// <param name="safeBox"></param>
        /// <param name="branchNo"></param>
        /// <exception cref="Exception"></exception>
        private void SetKojin(MyDbData codeDb, EnumerableRowCollection<DataRow> kojinRows, StreamWriter safeBox, string branchNo)
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

                // 個人の金庫事務用データを取得
                string safeBoxKojin = GetSafeBoxKojin(codeDb, row, branchNo, caseNo, parsedDate, ProgressValue, tel1st, tel3rd, industryEtc, industry, industryTxt, job, jobTxt, nationCode, purpose, purposeTxt);
                safeBox.WriteLine(safeBoxKojin);
            }
        }

        /// <summary>
        /// 金庫事務用データ（個人）
        /// </summary>
        /// <param name="codeDb"></param>
        /// <param name="row"></param>
        /// <param name="brancNo"></param>
        /// <param name="caseNo"></param>
        /// <param name="parsedDate"></param>
        /// <param name="num"></param>
        /// <param name="tel1st"></param>
        /// <param name="tel3rd"></param>
        /// <param name="industryEtc"></param>
        /// <param name="industry"></param>
        /// <param name="industryTxt"></param>
        /// <param name="job"></param>
        /// <param name="jobTxt"></param>
        /// <param name="nationCode"></param>
        /// <param name="purpose"></param>
        /// <param name="purposeTxt"></param>
        /// <returns></returns>
        private string GetSafeBoxKojin(MyDbData codeDb, DataRow row, string brancNo, string caseNo, DateTime parsedDate, int num, string tel1st, string tel3rd,
                            string industryEtc, string industry, string industryTxt, string job, string jobTxt, string nationCode, string[] purpose, string purposeTxt)
        {
            string delimiter = ",";

            // カナ住所結合
            string kanaAddr = row["pref_kana"].ToString() + " " + row["city_kana"].ToString() + " "
                                + row["addrnum_kan"].ToString() + " " + row["bldgkana_nm"].ToString();

            string record = string.Empty;
            record += row["bpo_bank_code"].ToString() + delimiter;  // 金融機関コード
            record += row["bpo_cust_no"].ToString() + delimiter;  // 顧客番号
            record += row["bpo_member_no"].ToString() + delimiter;  // 会員番号
            record += row["bpo_persona_cd"].ToString() + delimiter;  // 人格コード
            record += row["bpo_kana_name"].ToString() + delimiter;  // カナ氏名
            record += row["bpo_org_kanji"].ToString() + delimiter;  // 漢字氏名
            record += parsedDate.ToString("yyyyMMdd") + delimiter; // 回答日
            record += row["email"].ToString() + delimiter;  // メールアドレス

            // 氏名変更有無　WEBCASデータでは1・2 ⇒ 0・1 に置換え
            record += (row["namechg_flg"].ToString() == "1" ? "0" : "1") + delimiter;
            record += (row["lname_kana"].ToString() == "2" ? row["fname_kana"].ToString() + " " + row["fname_kanji"].ToString() : "") + delimiter; // カナ氏名
            record += (row["namechg_flg"].ToString() == "2" ? row["lname_kanji"].ToString() + "　" + row["fname_kanji"].ToString() : "") + delimiter; // 漢字氏名

            // 住所変更有無　WEBCASデータでは1・2 ⇒ 0・1 に置換え
            record += (row["addrchg_flg"].ToString() == "1" ? "0" : "1") + delimiter;
            record += (row["addrchg_flg"].ToString() == "2" ? row["zipcode"].ToString() : "") + delimiter;  // 郵便番号

            string pref = string.Empty;
            if (row["addrchg_flg"].ToString() == "2")
            {
                // 都道府県コードから都道府県名を取得
                pref = codeDb.ExecuteScalar($"SELECT zip_name FROM t_zip_code WHERE code = '{row["pref"].ToString()}'") as string;
                // 都道府県名が見つからない場合は例外を投げる
                if (string.IsNullOrEmpty(pref)) throw new Exception($"（個人）都道府県コードが見つかりません: {row["pref"].ToString()}");
            }

            record += pref + delimiter;  // 都道府県
            record += (row["addrchg_flg"].ToString() == "2" ? row["city"].ToString() : "") + delimiter;  // 市区町村
            record += (row["addrchg_flg"].ToString() == "2" ? row["addr_num"].ToString() : "") + delimiter;  // 丁目・番地
            record += (row["addrchg_flg"].ToString() == "2" ? row["bldg_name"].ToString() : "") + delimiter;  // マンション名・部屋番号
            record += (row["addrchg_flg"].ToString() == "2" ? kanaAddr.Left(176) : "") + delimiter; // カナ住所
            record += tel1st + delimiter;　// 第一電話番号
            record += row["work_tel"].ToString() + delimiter;　// 第二電話番号
            record += tel3rd + delimiter; // 第三電話番号

            // 国籍変更有無　WEBCASデータでは1・2 ⇒　0・1 に置換え
            record += (row["nationchg_flg"].ToString() == "1" ? "0" : "1") + delimiter;
            // 国籍・国名　変更なしは空欄
            record += (row["nationchg_flg"].ToString() == "1" ? "" : nationCode) + delimiter; // 国籍・国名

            record += row["name_alpha"].ToString() + delimiter; // 国籍・アルファベット氏名
            record += row["visa"].ToString() + delimiter; // 在留資格
            record += row["visa_limit"].ToString() + delimiter; // 在留期限
            record += row["pep_flag"].ToString() + delimiter; // PEPS該否
            record += job + delimiter; // 職業
            record += jobTxt + delimiter; // 職業その他
            record += "" + delimiter; // 勤務先変更有無

            record += row["work_or_sch"].ToString() + delimiter; // 勤務先名
            record += row["worksch_kan"].ToString() + delimiter; // 勤務先名カナ
            record += job + delimiter; // 職業
            record += jobTxt + delimiter; // 職業その他
            record += "" + delimiter; // 主な製品・サービス
            record += row["sidejob_flg"].ToString() + delimiter; // 副業有無
            record += row["sidejob_typ"].ToString() + delimiter; // 副業の業種
            record += row["sideothr_tx"].ToString() + delimiter; // 副業の業種その他
            record += "" + delimiter; // 取引目的コード変更有無
            record += (purpose.Count() > 0 ? purpose[0] : "") + delimiter; // 取引目的コード1
            record += (purpose.Count() > 1 ? purpose[1] : "") + delimiter; // 取引目的コード2
            record += (purpose.Count() > 2 ? purpose[2] : "") + delimiter; // 取引目的コード3
            record += (purpose.Count() > 3 ? purpose[3] : "") + delimiter; // 取引目的コード4
            record += (purpose.Count() > 4 ? purpose[4] : "") + delimiter; // 取引目的コード5
            record += (purpose.Count() > 5 ? purpose[5] : "") + delimiter; // 取引目的コード6
            record += purposeTxt + delimiter; // 取引目的コードその他
            record += row["tx_type"].ToString() + delimiter; // 取引形態
            record += row["tx_freq"].ToString() + delimiter; // 取引頻度
            record += row["tx_amt_once"].ToString() + delimiter; // 1回あたり取引金額

            record += row["tx_over2m_f"].ToString() + delimiter; // 200万円超取引の有無
            record += row["tx_over2mfrq"].ToString() + delimiter; // 200万円超取引の頻度
            record += row["tx_over2mam"].ToString() + delimiter; // 200万円超取引の金額

            string[] src =
            {
                row["src_salary"].ToString(),   // 200万円超現金取引の原資_給与_退職金
                row["src_pension"].ToString(),  // 200万円超現金取引の原資_年金_社会保険_公的扶助
                row["src_insur"].ToString(),    // 200万円超現金取引の原資_保険_民間
                row["src_exec"].ToString(),     // 200万円超現金取引の原資_役員報酬
                row["src_bizin"].ToString(),    // 200万円超現金取引の原資_事業収入
                row["src_inheri"].ToString(),   // 200万円超現金取引の原資_相続_贈与
                row["src_invincm"].ToString(),  // 200万円超現金取引の原資_配当_利子_資産運用益_不動産収入
                row["src_saving"].ToString(),   // 200万円超現金取引の原資_貯蓄
                row["src_otherbk"].ToString(),  // 200万円超現金取引の原資_他の金融機関口座から引き出した現金
                row["src_home"].ToString(),     // 200万円超現金取引の原資_自宅保管現金
                row["src_loan"].ToString(),     // 200万円超現金取引の原資_借入金
                row["src_other"].ToString() == "-999" ? "12" : ""   // 200万円超現金取引の原資_それ以外
            };

            // 200万円超現金取引の原資の配列から空欄を除外
            var srcFiltered = src
                .Where(x => !string.IsNullOrEmpty(x))
                .ToArray();

            record += (srcFiltered.Count() > 0 ? srcFiltered[0].ToString() : "") + delimiter; // 200万円超現金取引の原資
            record += (srcFiltered.Count() > 1 ? srcFiltered[1].ToString() : "") + delimiter; // 200万円超現金取引の原資2
            record += (srcFiltered.Count() > 2 ? srcFiltered[2].ToString() : "") + delimiter; // 200万円超現金取引の原資3
            record += (srcFiltered.Contains("12") ? row["src_oth_txt"].ToString() : "") + delimiter; // 200万円超現金取引の原資その他

            record += "" + delimiter; // 団体種類
            record += "" + delimiter; // 設立年月日変更有無
            record += "" + delimiter; // 設立年月日
            record += "" + delimiter; // 本社所在国名
            record += "" + delimiter; // 代表者の漢字氏名変更有無
            record += "" + delimiter; // 代表者の漢字氏名
            record += "" + delimiter; // 代表者のカナ氏名
            record += "" + delimiter; // 代表者の生年月日
            record += "" + delimiter; // 代表者の役職
            record += "" + delimiter; // 代表者の国籍・国名
            record += "" + delimiter; // 代表者の国籍・アルファベット氏名
            record += "" + delimiter; // 代表者の国籍・在留資格
            record += "" + delimiter; // 代表者の国籍・在留期限
            record += "" + delimiter; // 実質的支配者の氏名
            record += "" + delimiter; // 実質的支配者のカナ氏名
            record += "" + delimiter; // 実質的支配者の生年月日
            record += "" + delimiter; // 実質的支配者の団体との関係
            record += "" + delimiter; // 実質的支配者の住所
            record += "" + delimiter; // 実質的支配者の職業・事業内容
            record += "" + delimiter; // 実質的支配者のPEPS該否
            record += "" + delimiter; // 実質的支配者の国籍・国名
            record += "" + delimiter; // 実質的支配者のアルファベット氏名
            record += "" + delimiter; // 実質的支配者の在留資格
            record += "" + delimiter; // 実質的支配者の在留期限
            record += "" + delimiter; // 実質的支配者の氏名2
            record += "" + delimiter; // 実質的支配者のカナ氏名2
            record += "" + delimiter; // 実質的支配者の生年月日2
            record += "" + delimiter; // 実質的支配者の団体との関係2
            record += "" + delimiter; // 実質的支配者の住所2
            record += "" + delimiter; // 実質的支配者の職業・事業内容2
            record += "" + delimiter; // 実質的支配者のPEPS該否2
            record += "" + delimiter; // 実質的支配者の国籍・国名2
            record += "" + delimiter; // 実質的支配者のアルファベット氏名2
            record += "" + delimiter; // 実質的支配者の在留資格2
            record += "" + delimiter; // 実質的支配者の在留期限2
            record += "" + delimiter; // 実質的支配者の氏名3
            record += "" + delimiter; // 実質的支配者のカナ氏名3
            record += "" + delimiter; // 実質的支配者の生年月日3
            record += "" + delimiter; // 実質的支配者の団体との関係3
            record += "" + delimiter; // 実質的支配者の住所3
            record += "" + delimiter; // 実質的支配者の職業・事業内容3
            record += "" + delimiter; // 実質的支配者のPEPS該否3
            record += "" + delimiter; // 実質的支配者の国籍・国名3
            record += "" + delimiter; // 実質的支配者のアルファベット氏名3
            record += "" + delimiter; // 実質的支配者の在留資格3
            record += "" + delimiter; // 実質的支配者の在留期限3
            record += "" + delimiter; // 取引担当者の氏名
            record += "" + delimiter; // 取引担当者のカナ氏名
            record += "" + delimiter; // 取引担当者の部署
            record += "" + delimiter; // 取引担当者の役職
            record += "" + delimiter; // 取引担当者の電話番号
            record += "" + delimiter; // 取引担当者のメールアドレス
            record += "" + delimiter; // 取引担当者の郵便番号
            record += "" + delimiter; // 取引担当者の住所
            record += "" + delimiter; // 取引担当者のカナ住所

            record += row["bpo_ship_round"].ToString() + delimiter; // 発送希望回次
            record += row["bpo_zip_code"].ToString() + delimiter; // 変更前郵便番号
            record += row["bpo_address"].ToString() + delimiter; // 変更前住所（所在地）
            record += row["bpo_home_tel"].ToString() + delimiter; // 変更前第一電話番号
            record += row["bpo_work_tel"].ToString() + delimiter; // 変更前第二電話番号
            record += row["bpo_mobile_tel"].ToString() + delimiter; // 変更間第三電話番号
            record += row["bpo_nation"].ToString() + delimiter; // 変更前国籍
            record += row["bpo_visa_type"].ToString() + delimiter; // 変更前在留資格
            record += row["bpo_visa_exp"].ToString() + delimiter; // 変更前在留期限
            record += row["bpo_job_type_cd"].ToString() + delimiter; // 変更前職業事業内容コード
            record += row["bpo_biz_type_cd"].ToString() + delimiter; // 変更前業種コード
            record += row["bpo_biz_name_etc"].ToString() + delimiter; // 変更前その他業種名称

            record += row["bpo_purp_1"].ToString() + delimiter; // 変更前取引目的1
            record += row["bpo_purp_2"].ToString() + delimiter; // 変更前取引目的2
            record += row["bpo_purp_3"].ToString() + delimiter; // 変更前取引目的3
            record += row["bpo_purp_4"].ToString() + delimiter; // 変更前取引目的4
            record += row["bpo_purp_5"].ToString() + delimiter; // 変更前取引目的5
            record += row["bpo_purp_6"].ToString() + delimiter; // 変更前取引目的6
            record += row["bpo_birth_date"].ToString() + delimiter; // 変更間生年月日（設立年月日）
            record += row["bpo_kanji_2"].ToString() + delimiter; // 変更前漢字氏名2
            record += row["bpo_role_kanji"].ToString() + delimiter; // 変更前漢字役職名
            record += row["bpo_rep_kanji"].ToString() + delimiter; // 変更前代表者の漢字氏名

            return record;
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
