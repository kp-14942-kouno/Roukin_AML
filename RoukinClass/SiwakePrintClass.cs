using ICSharpCode.SharpZipLib;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MySqlX.XDevAPI.Common;
using MyTemplate.Report;
using MyTemplate.Report.Helpers;
using MyTemplate.Report.Models;
using MyTemplate.Report.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using YamlDotNet.Core.Events;
using static MyLibrary.MyEnum;

namespace MyTemplate.RoukinClass
{
    public class SiwakePrintClass : MyLibrary.MyLoading.Thread
    {
        public const int OP_PRINT = 1 << 0;         // 印刷
        public const int OP_MACHING = 1 << 2;       // マッチング
        public const int OP_MEISAI = 1 << 3;        // 明細

        private int _operation; // 実行する操作
        private DataTable _table; // 仕分けリストデータ
        private PrintQueue _printer; // 印刷するプリンター
        private string _msg; // メッセージ
        private List<Report.Models.BankModel> _bankModels; // 金融機関情報

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="printer"></param>
        /// <param name="bankModels"></param>
        /// <param name="operation"></param>
        public SiwakePrintClass(DataTable table, PrintQueue printer, List<BankModel> bankModels, int operation, string msg)
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
            ProcessName = $"申請書仕分けリスト処理中...";
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            try
            {
                // 開始ログ
                MyLogger.SetLogger($"{_msg}作成開始", MyEnum.LoggerType.Info, false);

                // 実行
                Run(_bankModels);

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
        private void Run(List<Report.Models.BankModel> bankModels)
        {
            // 法人格の有無
            var ari = new string[] { "21", "31", "81", "83" };
            var nasi = new string[] { "12", "22" };

            // 日付
            var date = DateTime.Now.ToString("yyyyMMdd");

            // マッチング用ファイルの設定
            var matchingDir = MyUtilityModules.AppSetting("roukin_setting", "siwake_root_path");
            var matchingFile = MyUtilityModules.AppSetting("roukin_setting", "siwake_name");

            // 仕分け対象のデータの金融機関コードを取得
            var codes = _table.AsEnumerable().Select(x => x["bpo_bank_code"].ToString()).Distinct().OrderBy(x => x).ToList();

            foreach (var code in codes)
            {
                for (int i = 0; i <= 1; i++)
                {
                    DataTable? rows = null;

                    // マッチングデータ用
                    StringBuilder maching = new();

                    // 種別名
                    string typeName = string.Empty;

                    // 法人格の有無で分岐
                    if (i == 0)
                    {
                        typeName = "法人格あり";
                        // 対象の金融機関コードで法人格ありでフィルタリング
                        var tmp = _table.AsEnumerable().Where(x => x.Field<string>("bpo_bank_code") == code && ari.Contains(x.Field<string>("bpo_persona_cd")));

                        if (tmp.Any())
                            rows = tmp.CopyToDataTable();
                    }
                    else
                    {
                        typeName = "法人格なし";
                        // 対象の金融機関コードで法人格なしでフィルタリング
                        var tmp = _table.AsEnumerable().Where(x => x.Field<string>("bpo_bank_code") == code && nasi.Contains(x.Field<string>("bpo_persona_cd")));

                        if (tmp.Any())
                            rows = tmp.CopyToDataTable();
                    }

                    if (rows == null || rows.Rows.Count == 0)
                    {
                        // 該当するデータがない場合はスキップ
                        continue;
                    }

                    // 金融機関情報取得
                    var financial = bankModels.FirstOrDefault(x => x.code == code);

                    // QRコード値
                    var qrCode = $"{code}_{date}_{i}"; // 金融機関コード_日付_法人格区分

                    // 明細印刷
                    if((_operation & OP_MEISAI) != 0)
                    {
                        FixedDocument document = null;

                        // Dispatcherを使用してUIスレッドでFixedDocumentを作成
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            // FixedDocumentを作成
                            document = SinseishoMeisaiHelper.CreateFixedDocument(rows, code, financial.financial_name, typeName);
                        });

                        // 印刷処理
                        // A4、縦向き、片面印刷、用紙トレイは自動選択
                        // Dispatcherを使用してUIスレッドで印刷処理を実行
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            Modules.FixedDocumentPrint(document, _printer, ParperSize.A4, PageOrientation.Portrait, Duplexing.OneSided, InputBin.AutoSelect);
                        });

                        document = null; // documentオブジェクトを解放
                    }

                    // 仕分けリスト印刷処理
                    if ((_operation & OP_PRINT) != 0)
                    {
                        FixedDocument document = null;

                        // Dispatcherを使用してUIスレッドでFixedDocumentを作成
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            // FixedDocumentを作成
                            document = SiwakeHelper.CreateFixedDocument(rows, code, financial.financial_name, typeName, qrCode);
                        });

                        // 印刷処理
                        // A4、縦向き、片面印刷、用紙トレイは自動選択
                        // Dispatcherを使用してUIスレッドで印刷処理を実行
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            Modules.FixedDocumentPrint(document, _printer, ParperSize.A4, PageOrientation.Portrait, Duplexing.OneSided, InputBin.AutoSelect);
                        });

                        document = null; // documentオブジェクトを解放
                    }

                    // マッチングデータ作成処理
                    if ((_operation & OP_MACHING) != 0)
                    {
                        var record = string.Empty;

                        record += qrCode + ",";

                        // マッチングデータ作成処理
                        foreach (DataRow row in rows.Rows)
                        {
                            record += row["bpo_num"].ToString() + ","; // BPO管理番号
                        }

                        record = record.TrimEnd(','); // 最後のカンマを削除
                        maching.AppendLine(record); // マッチングデータに追加
                    }

                    if(maching.Length > 0)
                    {
                        // マッチングデータをファイルに保存
                        var filePath = Path.Combine(matchingDir, $"{matchingFile}_{financial.financial_name}_{typeName}_{DateTime.Now.ToString("yyyyMMdd_hhmmss")}.csv");
                        Directory.CreateDirectory(Path.GetDirectoryName(filePath)); // ディレクトリが存在しない場合は作成
                        File.WriteAllText(filePath, maching.ToString(), MyUtilityModules.GetEncoding(MyEnum.MojiCode.Sjis));
                        maching.Clear(); // マッチングデータをクリア
                    }

                    System.Threading.Thread.Sleep(50); // documentオブジェクトが解放されガベージコレクションが正しく処理されるように少し待機
                    GC.Collect(); // ガベージコレクションを実行
                    GC.WaitForPendingFinalizers(); // ガベージコレクションの完了を待機
                }
            }
        }
    }
}
