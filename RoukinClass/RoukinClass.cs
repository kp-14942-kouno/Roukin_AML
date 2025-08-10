using DocumentFormat.OpenXml.Presentation;
using ICSharpCode.SharpZipLib.Zip;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.MyClass;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Interop;
using System.Windows.Markup;

namespace MyTemplate
{

    /// <summary>
    /// WEBCASデータ変換出力クラス
    /// </summary>
    public class EpxWebcasData : MyLibrary.MyLoading.Thread
    {
        DataTable _table;
        string _fileDir = string.Empty; // 出力先ディレクトリ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="fileDir"></param>
        public EpxWebcasData(DataTable table, string fileDir)
        {
            _fileDir = fileDir;
            _table = table;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public override int MultiThreadMethod()
        {
            try
            {
                Run();
                Result = MyEnum.MyResult.Ok;
            }
            catch(Exception ex)
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
            string fileName = MyUtilityModules.AppSetting("roukin_setting", "exp_webcas_file_name", true, _table.Rows.Count);
            string delimiter = MyUtilityModules.AppSetting("roukin_setting", "exp_webcas_file_delimiter");
            string mojiCode = MyUtilityModules.AppSetting("roukin_setting", "exp_webcas_file_mojicode");

            ProgressValue = 0;
            ProgressMax = _table.Rows.Count;
            ProcessName = "WEBCASデータ変換出力";

            using(var fsWriter = new System.IO.FileStream(System.IO.Path.Combine(_fileDir, fileName), System.IO.FileMode.CreateNew, System.IO.FileAccess.Write))
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(fsWriter, MyUtilityModules.GetEncoding(mojiCode)))
            {
                StringBuilder builder = new StringBuilder();

                foreach (DataRow row in _table.Rows)
                {
                    ProgressValue++;

                    builder.Clear();

                    builder.Append(row["ctrl_num"].ToString().Trim() + delimiter);
                    builder.Append(row["answer_date"].ToString().Trim() + delimiter);
                    builder.Append(row["namechg_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["lname_kanji"].ToString().Trim() + delimiter);
                    builder.Append(row["fname_kanji"].ToString().Trim() + delimiter);
                    builder.Append(row["lname_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["fname_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["addrchg_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["zipcode"].ToString().Trim() + delimiter);
                    builder.Append(row["pref"].ToString().Trim() + delimiter);
                    builder.Append(row["city"].ToString().Trim() + delimiter);
                    builder.Append(row["addr_num"].ToString().Trim() + delimiter);
                    builder.Append(row["bldg_name"].ToString().Trim() + delimiter);
                    builder.Append(row["pref_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["city_kana"].ToString().Trim() + delimiter);
                    builder.Append(row["addrnum_kan"].ToString().Trim() + delimiter);
                    builder.Append(row["bldgkana_nm"].ToString().Trim() + delimiter);
                    builder.Append(row["tel_fixed"].ToString().Trim() + delimiter);
                    builder.Append(row["tel_mobile"].ToString().Trim() + delimiter);
                    builder.Append(row["email"].ToString().Trim() + delimiter);
                    builder.Append(row["nationchg_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["nation"].ToString().Trim() + delimiter);
                    builder.Append(row["region"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_asia"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_mide"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_weur"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_eeur"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_nam"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_cari"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_latm"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_afrc"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_oce"].ToString().Trim() + delimiter);
                    builder.Append(row["nation_othr"].ToString().Trim() + delimiter);
                    builder.Append(row["name_alpha"].ToString().Trim() + delimiter);
                    builder.Append(row["visa"].ToString().Trim() + delimiter);
                    builder.Append(row["visa_limit"].ToString().Trim() + delimiter);
                    builder.Append(row["pep_flag"].ToString().Trim() + delimiter);
                    builder.Append(row["job"].ToString().Trim() + delimiter);
                    builder.Append(row["job_other"].ToString().Trim() + delimiter);
                    builder.Append(row["work_or_sch"].ToString().Trim() + delimiter);
                    builder.Append(row["worksch_kan"].ToString().Trim() + delimiter);
                    builder.Append(row["work_tel"].ToString().Trim() + delimiter);
                    builder.Append(row["industry"].ToString().Trim() + delimiter);
                    builder.Append(row["indus_othr"].ToString().Trim() + delimiter);
                    builder.Append(row["sidejob_flg"].ToString().Trim() + delimiter);
                    builder.Append(row["sidejob_typ"].ToString().Trim() + delimiter);
                    builder.Append(row["sideothr_tx"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose1"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose2"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose3"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose4"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose5"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_purpose6"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_othr_txt"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_type"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_freq"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_amt_once"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_over2m_f"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_over2mfrq"].ToString().Trim() + delimiter);
                    builder.Append(row["tx_over2mam"].ToString().Trim() + delimiter);
                    builder.Append(row["src_salary"].ToString().Trim() + delimiter);
                    builder.Append(row["src_pension"].ToString().Trim() + delimiter);
                    builder.Append(row["src_insur"].ToString().Trim() + delimiter);
                    builder.Append(row["src_exec"].ToString().Trim() + delimiter);
                    builder.Append(row["src_bizin"].ToString().Trim() + delimiter);
                    builder.Append(row["src_inheri"].ToString().Trim() + delimiter);
                    builder.Append(row["src_invincm"].ToString().Trim() + delimiter);
                    builder.Append(row["src_saving"].ToString().Trim() + delimiter);
                    builder.Append(row["src_otherbk"].ToString().Trim() + delimiter);
                    builder.Append(row["src_home"].ToString().Trim() + delimiter);
                    builder.Append(row["src_loan"].ToString().Trim() + delimiter);
                    builder.Append(row["src_other"].ToString().Trim() + delimiter);
                    builder.Append(row["src_oth_txt"].ToString().Trim());

                    // 書き込み
                    writer.WriteLine(builder);
                }

                ResultMessage = $"WEBCASデータ：{ProgressValue}件の変換出力が完了しました。";
            }
        }
    }

    /// <summary>
    /// 審査結果出力クラス
    /// </summary>
    public class ExpReviewResult : MyLibrary.MyLoading.Thread
    {
        DataTable _fubiData;
        DataTable _FixData;
        string _filePath = string.Empty; // 出力先ファイルパス
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="fubiData"></param>
        /// <param name="fixData"></param>
        public ExpReviewResult(DataTable fubiData, DataTable fixData, string filePath)
        {
            _fubiData = fubiData;
            _FixData = fixData;
            _filePath = filePath;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public override int MultiThreadMethod()
        {
            try
            {
                // 審査結果出力処理の開始ログ
                MyLogger.SetLogger("審査結果出力処理を開始", MyEnum.LoggerType.Info, false);

                ExpFubiData();
                ExpFixData();

                ResultMessage = $"審査不備：{_fubiData.Rows.Count}件 ／ 審査正常：{_FixData.Rows.Count}件"
                    + "\r\n\r\n審査結果の出力が完了しました。";

                // 審査結果出力処理の完了ログ
                MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);

                Result = MyEnum.MyResult.Ok;
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
        /// 審査不備データ作成
        /// </summary>
        private void ExpFubiData()
        {
            // 不備データが無い場合は処理を終了
            if (_fubiData.Rows.Count == 0) return;

            ProcessName = "審査不備データ出力";
            ProgressMax = _fubiData.Rows.Count;
            ProgressValue = 0;

            string fileName = MyUtilityModules.AppSetting("roukin_setting", "exp_rew_ng_file_name", true, _fubiData.Rows.Count);
            string expPath = System.IO.Path.Combine(_filePath, fileName);

            using(var fsWriter = new System.IO.FileStream(expPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write))
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(fsWriter, MyUtilityModules.GetEncoding(MyEnum.MojiCode.Utf8Bom)))
            {
                foreach (DataRow row in _fubiData.Rows)
                {
                    ProgressValue++;

                    string value = string.Empty;
                    value += row["bpo_num"].ToString().Trim() + ",";
                    value += row["fubi_code"].ToString().Trim();
                    // 書き込み
                    writer.WriteLine(value);
                }
            }
        }

        /// <summary>
        /// 審査正常データ作成
        /// </summary>
        private void ExpFixData()
        {
            // 正常データが無い場合は処理を終了
            if (_FixData.Rows.Count == 0) return;

            ProcessName = "審査正常データ出力";
            ProgressMax = _FixData.Rows.Count;
            ProgressValue = 0;

            string fileName = MyUtilityModules.AppSetting("roukin_setting", "exp_rew_ok_file_name", true, _FixData.Rows.Count);
            string expPath = System.IO.Path.Combine(_filePath, fileName);

            using (var fsWriter = new System.IO.FileStream(expPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write))
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(fsWriter, MyUtilityModules.GetEncoding(MyEnum.MojiCode.Utf8Bom)))
            {
                foreach (DataRow row in _FixData.Rows)
                {
                    ProgressValue++;

                    string value = string.Empty;
                    value += row["bpo_num"].ToString().Trim();
                    // 書き込み
                    writer.WriteLine(value);
                }
            }
        }
    }

    /// <summary>
    /// パンチ連携画像作成クラス
    /// </summary>
    public class ExpPunchImage : MyLibrary.MyLoading.Thread
    {
        DataTable _table = new DataTable();
        string _expPath = string.Empty; // 出力先パス

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        public ExpPunchImage(DataTable table, string expPath)
        {
            _expPath = expPath;
            _table = table;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public override int MultiThreadMethod()
        {
#if DEBUG
            {
                // 開始ログ
                MyLogger.SetLogger("パンチ連携画像作成処理を開始", MyEnum.LoggerType.Info, false);
                // 実行
                Run();

                // 完了メッセージ
                ResultMessage = $"パンチ連携画像の作成完了";
                // 完了ログ
                MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);
                Result = MyEnum.MyResult.Ok;
            }
#else
            try
            {
                // 開始ログ
                MyLogger.SetLogger("パンチ連携画像作成処理を開始", MyEnum.LoggerType.Info, false);
                // 実行
                Run();

                // 完了メッセージ
                ResultMessage = $"パンチ連携画像の作成完了";
                // 完了ログ
                MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);
                Result = MyEnum.MyResult.Ok;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                Result = MyEnum.MyResult.Error;
            }
#endif
            Completed = true;
            return 0;
        }

        /// <summary>
        /// パンチ画像作成
        /// </summary>
        private void Run()
        {
            ProcessName = "パンチ連携画像作成中...";
            ProgressBarType = MyEnum.MyProgressBarType.Percent;
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            // 画像ファイルパス
            string imgPath = MyUtilityModules.AppSetting("roukin_setting", "img_root_path");
            // パンチ画像ディレクトリ
            string punchDir = MyUtilityModules.AppSetting("roukin_setting", "punch_img_dir");
            // パンチ画像ZIPファイル名
            string zipName = MyUtilityModules.AppSetting("roukin_setting", "punch_zip_name", true, _table.Rows.Count);
            // パスワード
            string zipPassword = MyUtilityModules.AppSetting("roukin_setting", "punch_zip_password");

            // ZIPの作成
            using (var zipStream = new MyArchiveWriter(Path.Combine(_expPath, zipName), true, true))
            {
                foreach (DataRow row in _table.Rows)
                {
                    ProgressValue++;

                    // 対象パンチ画像のパスを取得
                    string sourcePath = Path.Combine(imgPath, punchDir, row["taba_num"].ToString(), row["bpo_num"].ToString() + ".PDF");
                    // ZIPに書出し
                    zipStream.WriteFile(row["bpo_num"].ToString() + ".PDF", sourcePath, zipPassword);
                }
            }
        }   
    }

    /// <summary>
    /// 不備納品データのZIP作成クラス
    /// </summary>
    public class FubiZipCreate : MyLibrary.MyLoading.Thread
    {
        private string _expPath = string.Empty; // 出力先パス

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="expPath"></param>
        public FubiZipCreate(string expPath)
        {
            _expPath = expPath;
        }

        public override int MultiThreadMethod()
        {
#if DEBUG
            {
                string msg = "不備納品ファイルのZIP作成";
                // 開始ログ出力
                MyLogger.SetLogger($"{msg}を開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {
                    // 実行
                    Run(codeDb);
                    // 結果メッセージ
                    ResultMessage = $"{msg}完了";
                    // 終了ログ出力
                    MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);

                    Result = MyLibrary.MyEnum.MyResult.Ok;
                }
            }
#else
            try
            {
                string msg = "不備納品ファイルのZIP作成";
                // 開始ログ出力
                MyLogger.SetLogger($"{msg}を開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {
                    // 実行
                    Run(codeDb);
                    // 結果メッセージ
                    ResultMessage = $"{msg}納品データ作成完了";
                    // 終了ログ出力
                    MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);

                    Result = MyLibrary.MyEnum.MyResult.Ok;
                }
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyLibrary.MyEnum.LoggerType.Error);
                Result = MyLibrary.MyEnum.MyResult.Error;
            }
#endif
            Completed = true;
            return 0;
        }

        private void Run(MyDbData codeDb)
        {
            ProgressBarType = MyEnum.MyProgressBarType.None;
            ProcessName = "不備納品ファイルのZIP作成中...";

            // 金庫事務用ディレクトリ
            string safeBoxDir = MyUtilityModules.AppSetting("roukin_setting", "safe_box_admin_dir", true);

            using (DbDataReader reader = codeDb.ExecuteReader("select * from t_financial_code order by code;"))
            {
                while (reader.Read())
                {
                    // 金庫名を取得
                    string bankName = reader["financial_name"].ToString();
                    bankName = bankName.Replace("労金", ""); // 労金を削除
                    // ZIP用パスワード
                    string password = reader["password"].ToString();

                    // 対象フォルダが存在するか確認
                    if (Directory.Exists(Path.Combine(_expPath, safeBoxDir, bankName)))
                    {
                        // ZIP作成
                        using(var archive = new MyArchiveWriter(Path.Combine(_expPath, safeBoxDir, bankName + ".zip"), true, true))
                        {
                            // ZIPに書き込み
                            archive.AddFolder(Path.Combine(_expPath, safeBoxDir, bankName), bankName, true, password);
                        }
                    }
                }
            }
        }
    }
}
