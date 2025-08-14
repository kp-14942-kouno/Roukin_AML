using Azure;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.Report;
using MyTemplate.Report.Helpers;
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
    public class SinseishoMeisaiPrintClass : MyLibrary.MyLoading.Thread
    {
        private DataTable _table; // 仕分けリストデータ
        private PrintQueue _printer; // 印刷するプリンター
        private List<Report.Models.BankModel> _bankModels; // 金融機関情報

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="printer"></param>
        /// <param name="bankModels"></param>
        public SinseishoMeisaiPrintClass(DataTable table, PrintQueue printer, List<Report.Models.BankModel> bankModels)
        {
            _table = table;
            _printer = printer;
            _bankModels = bankModels;
        }

        public override int MultiThreadMethod()
        {
            var msg = "申請書明細リスト印刷";

            ProgressBarType = MyEnum.MyProgressBarType.None;
            ProcessName = $"{msg}中...";

            try
            {
                // 開始ログ
                MyLogger.SetLogger($"{msg}開始", MyEnum.LoggerType.Info, false);

                // 実行
                Run(_bankModels);

                // 完了ログ
                MyLogger.SetLogger($"{msg}完了", MyEnum.LoggerType.Info, false);

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
        private void Run(List<Report.Models.BankModel> bankModels)
        {
            // 印刷対象の金融機関コードを取得
            var codes = _table.AsEnumerable().Select(x => x["bpo_bank_code"].ToString()).Distinct().OrderBy(x => x).ToList();

            foreach (var code in codes)
            {
                // 金融機関コードでフィルタリングし、人格コードでソート
                var rows = _table.AsEnumerable().Where(x => x.Field<string>("bpo_bank_code") == code)
                    .OrderBy(x => x["bpo_persona_cd"].ToString())
                    .CopyToDataTable();

                var financial = bankModels.FirstOrDefault(x => x.code == code);

                FixedDocument document = null;

                // Dispatcherを使用してUIスレッドでFixedDocumentを作成
                Application.Current.Dispatcher.Invoke(() =>
                {
                    // FixedDocumentを作成
                    document = SinseishoMeisaiHelper.CreateFixedDocument(rows, code, financial.financial_name, "");
                });

                // 印刷処理
                // A4、縦向き、片面印刷、用紙トレイは自動選択
                // Dispatcherを使用してUIスレッドで印刷処理を実行
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Modules.FixedDocumentPrint(document, _printer, ParperSize.A4, PageOrientation.Portrait, Duplexing.OneSided, InputBin.AutoSelect);
                });

                document = null; // FixedDocumentの参照を解放

                System.Threading.Thread.Sleep(50); // documentオブジェクトが解放されガベージコレクションが正しく処理されるように少し待機
                GC.Collect(); // ガベージコレクションを実行
                GC.WaitForPendingFinalizers(); // ガベージコレクションの完了を待機
            }
        }
    }
}
