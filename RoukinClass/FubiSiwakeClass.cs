using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.Report;
using MyTemplate.Report.Helpers;
using MyTemplate.Report.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;

namespace MyTemplate.RoukinClass
{
    public class FubiSiwakeClass : MyLibrary.MyLoading.Thread
    {
        public const int OP_PRINT = 1 << 0;         // 印刷
        public const int OP_MACHING = 1 << 2;       // マッチング

        private int _operation; // 実行する操作
        private DataTable _table; // 仕分けリストデータ
        private PrintQueue _printer; // 印刷するプリンター
        private string _msg; // メッセージ
        private List<Report.Models.BankModel> _bankModels; // 金融機関情報

        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="printer"></param>
        /// <param name="bankModels"></param>
        /// <param name="operation"></param>
        public FubiSiwakeClass(DataTable table, PrintQueue printer, List<BankModel> bankModels, int operation, string msg)
        {
            _table = table;
            _printer = printer;
            _bankModels = bankModels;
            _operation = operation;
            _msg = msg;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            ProgressBarType = MyEnum.MyProgressBarType.None;
            ProcessName = $"不備状不着仕分けリスト処理中...";
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            // マッチングデータ用
            StringBuilder maching = new();

            try
            {
                // 開始ログ
                MyLogger.SetLogger($"{_msg}作成開始", MyEnum.LoggerType.Info, false);

                // 実行
                Run(_bankModels, maching);

                // マッチングデータの保存
                if (maching.Length > 0 && (_operation & OP_MACHING) != 0)
                {
                    var path = MyUtilityModules.AppSetting("roukin_setting", "fubifuchaku_root_path");
                    var fileName = MyUtilityModules.AppSetting("roukin_setting", "fubifuchaku_name", true);
                    System.IO.Directory.CreateDirectory(path); // フォルダが存在しない場合は作成
                    System.IO.File.WriteAllText(System.IO.Path.Combine(path, fileName), maching.ToString());
                }

                // 完了ログ
                MyLogger.SetLogger($"{_msg}作成完了", MyEnum.LoggerType.Info, false);

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
        /// 実行
        /// </summary>
        /// <param name="sb"></param>
        private void Run(List<Report.Models.BankModel> bankModels, StringBuilder maching)
        {
            // 仕分け対象のデータの金融機関コードを取得
            var codes = _table.AsEnumerable().Select(x => x["bpo_bank_code"].ToString()).Distinct().OrderBy(x => x).ToList();

            foreach (var code in codes)
            {
                // 金融機関コードでフィルタリングし、束番号と束内連番でソート
                var rows = _table.AsEnumerable().Where(x => x.Field<string>("bpo_bank_code") == code)
                    .OrderBy(x => x["taba_num"].ToString())
                    .CopyToDataTable();

                var financial = bankModels.FirstOrDefault(x => x.code == code);

                FixedDocument document = null;

                // Dispatcherを使用してUIスレッドでFixedDocumentを作成
                Application.Current.Dispatcher.Invoke(() =>
                {
                    // FixedDocumentを作成
                    document = FubiFuchakuHelper.CreateFixedDocument(rows, code, financial.financial_name);
                });

                if ((_operation & OP_PRINT) != 0)
                {
                    // 印刷処理
                    // A4、縦向き、片面印刷、用紙トレイは自動選択
                    // Dispatcherを使用してUIスレッドで印刷処理を実行
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Modules.FixedDocumentPrint(document, _printer, ParperSize.A4, PageOrientation.Portrait, Duplexing.OneSided, InputBin.AutoSelect);
                    });
                }

                if ((_operation & OP_MACHING) != 0)
                {
                    // マッチングデータ作成処理
                    foreach (DataRow row in rows.Rows)
                    {
                        var value = "";
                        value += code + ","; // 金融機関コード
                        value += row["bpo_num"].ToString(); // BPO管理番号
                        maching.AppendLine(value);
                    }
                }

                document = null; // FixedDocumentの参照を解放

                System.Threading.Thread.Sleep(50); // documentオブジェクトが解放されガベージコレクションが正しく処理されるように少し待機
                GC.Collect(); // ガベージコレクションを実行
                GC.WaitForPendingFinalizers(); // ガベージコレクションの完了を待機
            }
        }

    }
}
