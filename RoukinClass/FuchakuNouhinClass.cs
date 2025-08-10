using DocumentFormat.OpenXml.Vml;
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
using System.Windows.Interop;

namespace MyTemplate.RoukinClass
{
    public class FuchakuNouhinClass :  MyLibrary.MyLoading.Thread
    {
        private DataTable _table = new DataTable(); // データテーブル
        string _msg = string.Empty; // メッセージ
        string _expPath = string.Empty; // 出力先パス

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        public FuchakuNouhinClass(DataTable table, string msg, string expPath)
        {
            _table = table;
            _msg = msg;
            _expPath = expPath;
        }

        public override int MultiThreadMethod()
        {
#if DEBUG
            {
                // 開始ログ出力
                MyLogger.SetLogger($"{_msg}作成開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {
                    // 実行呼出し
                    Run(codeDb);

                    // 結果メッセージ
                    ResultMessage = $"{_msg}全：{_table.Rows.Count} 件の作成完了";
                    // 終了ログ出力
                    MyLogger.SetLogger(ResultMessage, MyEnum.LoggerType.Info, false);

                    Result = MyLibrary.MyEnum.MyResult.Ok;
                }
            }
#else

            try
            {
                // 開始ログ出力
                MyLogger.SetLogger($"{_msg}作成開始", MyEnum.LoggerType.Info, false);

                using (var codeDb = new MyDbData("code"))
                {
                    // 実行呼出し
                    Run(codeDb);

                    // 結果メッセージ
                    ResultMessage = $"{_msg}全：{_table.Rows.Count} 件の作成完了";
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

        /// <summary>
        /// 実行
        /// </summary>
        /// <param name="codeDb"></param>
        /// <exception cref="Exception"></exception>
        private void Run(MyDbData codeDb)
        {
            ProgressBarType = MyLibrary.MyEnum.MyProgressBarType.Percent;
            ProcessName = $"不着納品作成中..."; // 処理名設定
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            string delimiter = ","; // レコードの区切り文字

            // 金庫事務用ディレクトリ
            string safeBoxDir = MyUtilityModules.AppSetting("roukin_setting", "safe_box_admin_dir", true);
            // 不備対象者データファイル名
            string fuchakuName = MyUtilityModules.AppSetting("roukin_setting", "fuchaku_name", true);

            // 金融機関コードを重複除外して取得
            var banks = _table.AsEnumerable().Select(x => x["bpo_bank_code"].ToString()).Distinct().ToList();
            // 作成日
            var date = DateTime.Now.ToString("yyyyMMdd");

            // 金融機関コードごとに処理
            foreach (string bank in banks)
            {
                // 金融機関コードから金融機関名を取得
                var bankData = codeDb.ExecuteQuery($"select * from t_financial_code where code = '{bank}'");

                // 金融機関名が見つからない場合は例外を投げる
                if (bankData.Rows.Count == 0)
                {
                    throw new Exception($"金融機関コードが見つかりません: {bank}");
                }

                // 金融機関名を取得
                string bankName = bankData.Rows[0]["financial_name"].ToString().Trim();
                bankName = bankName.Replace("労金", "");

                bankData.Dispose(); // データテーブルを破棄

                // 出力先パス作成
                string expDir = System.IO.Path.Combine(_expPath, safeBoxDir, bankName);

                // 出力先作成
                System.IO.Directory.CreateDirectory(expDir);

                var table = _table.AsEnumerable()
                    .Where(x => x["bpo_bank_code"].ToString().Trim() == bank)
                    .CopyToDataTable();

                // ファイル名を金庫名で置換え
                var fileName = fuchakuName.Replace("xxx", bankName);

                using (var fs = new FileStream(System.IO.Path.Combine(expDir, fileName), FileMode.CreateNew, FileAccess.Write))
                using (var writer = new StreamWriter(fs, MyLibrary.MyModules.MyUtilityModules.GetEncoding(MyLibrary.MyEnum.MojiCode.Sjis)))
                {
                    // ヘッダ行の作成
                    var header = string.Empty;
                    header += "金融機関コード" + delimiter;    // 金融機関コード
                    header += "顧客管理店番号" + delimiter;    // 顧客管理店番号
                    header += "顧客番号";                      // 顧客番号
                    writer.WriteLine(header);

                    // データ行の作成
                    foreach (DataRow row in table.Rows)
                    {
                        var record = string.Empty;
                        record += row["bpo_bank_code"].ToString() + delimiter;          // 金融機関コード
                        record += row["bpo_branch_no"].ToString().Trim() + delimiter;   // 顧客管理店番号
                        record += row["bpo_cust_no"].ToString().Trim();                 // 顧客番号
                        writer.WriteLine(record);
                    }
                }
            }
        }
    }
}
