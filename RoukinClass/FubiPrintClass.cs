using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
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
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ZstdSharp.Unsafe;

namespace MyTemplate.RoukinClass
{
    /// <summary>
    /// 不備状印刷・画像作成・ビートルデータ作成を行うクラス
    /// </summary>
    public class FubiPrintDocument : MyLibrary.MyLoading.Thread
    {
        public const int OP_PRINT = 1 << 0;         // 印刷
        public const int OP_IMAGE = 1 << 1;         // 画像作成
        public const int OP_BEETLE = 1 << 2;        // ビートルデータ作成
        public const int OP_HIKINUKI = 1 << 3;      // 引抜き
        public const int OP_MACHING = 1 << 4;       // マッチング

        private int _operation; // 実行する操作
        private DataTable _table; // 不備状データ用テーブル
        private PrintQueue _printer; // 印刷するプリンター
        private string _msg; // メッセージ
        private List<Report.Models.DefectModel> _defectModels; // 不備文言のモデルリスト
        private List<BankModel> _bankModel; // 銀行情報のモデルリスト

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="defectDic"></param>
        /// <param name="printer"></param>
        /// <param name="operation"></param>
        public FubiPrintDocument(DataTable table, PrintQueue printer, List<DefectModel> defectModels, List<BankModel> bankModel, int operation, string msg)
        {
            _table = table;
            _printer = printer;
            _defectModels = defectModels;
            _bankModel = bankModel;
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
            ProcessName = $"不備状処理中...";
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            // ビートルデータ用
            StringBuilder btl = new();
            // マッチングデータ用
            StringBuilder maching = new();

            try
            {
                // 開始ログ
                MyLogger.SetLogger($"{_msg}作成開始", MyEnum.LoggerType.Info, false);

                // 実行
                Run(_defectModels, _bankModel, btl, maching);

                // ビートルデータの保存
                if (btl.Length > 0 && (_operation & OP_BEETLE) != 0)
                {
                    var path = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
                    var fileName = MyUtilityModules.AppSetting("roukin_setting", "fubi_print_name", true);
                    System.IO.Directory.CreateDirectory(path); // フォルダが存在しない場合は作成
                    System.IO.File.WriteAllText(System.IO.Path.Combine(path, fileName), btl.ToString());
                }

                // マッチングデータの保存
                if(maching.Length > 0 && (_operation & OP_MACHING) != 0)
                {
                    var path = MyUtilityModules.AppSetting("roukin_setting", "match_root_path");
                    var fileName = MyUtilityModules.AppSetting("roukin_setting", "match_name", true);
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
        private void Run(List<Report.Models.DefectModel> defectModels, List<Report.Models.BankModel> bankModel, StringBuilder blt, StringBuilder maching)
        {
            // 不備状データをtaba_numでグループ化
            var tabas = _table.AsEnumerable().Select(x => x["taba_num"].ToString()).Distinct().OrderBy(x => x).ToList();

            foreach (var taba in tabas)
            {
                // 各taba_numごとに行をフィルタリング
                var rows = _table.AsEnumerable().Where(x => x.Field<string>("taba_num") == taba).CopyToDataTable();

                FixedDocument document = null;

                if ((_operation & OP_HIKINUKI) != 0)
                {
                    // Dispatcher.Invokeを使用してUIスレッドでFixedDocumentを作成
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        document = HikinukiHelper.CreateFixedDocument(rows, taba);
                    });

                    // 印刷処理
                    // A4、縦向き、片面印刷、用紙トレイは自動選択
                    // Dispatcher.Invokeを使用してUIスレッドで印刷処理を実行
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Modules.FixedDocumentPrint(document, _printer, ParperSize.A4, PageOrientation.Portrait, Duplexing.OneSided, InputBin.AutoSelect);
                    });
                }

                // _operationに引抜以外も含まれている場合
                if ((_operation & (OP_PRINT | OP_IMAGE | OP_BEETLE | OP_MACHING)) != 0)
                {
                    foreach (DataRow row in rows.Rows)
                    {
                        // Dispatcher.Invokeを使用してUIスレッドでFixedDocumentを作成
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            document = FubiHelper.CreateFixedDocument(row, defectModels, bankModel);
                        });

                        if ((_operation & OP_PRINT) != 0)
                        {
                            // 印刷処理
                            // A4、縦向き、片面印刷、用紙トレイは自動選択
                            // Dispatcher.Invokeを使用してUIスレッドで印刷処理を実行
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                Modules.FixedDocumentPrint(document, _printer, ParperSize.A4, PageOrientation.Portrait, Duplexing.OneSided, InputBin.AutoSelect);
                            });
                        }
                        if ((_operation & OP_IMAGE) != 0)
                        {
                            var path = MyUtilityModules.AppSetting("roukin_setting", "img_root_path");
                            var dir = MyUtilityModules.AppSetting("roukin_setting", "fubi_img_dir");
                            var fullPath = System.IO.Path.Combine(path, dir, row["taba_num"].ToString());

                            // 画像作成処理
                            // Dispatcher.Invokeを使用してUIスレッドで画像作成処理を実行
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                FixedDocumentAsJpeg(document, fullPath);
                                FixedDocumentAsTiff(document, fullPath);
                            });
                        }
                        if ((_operation & OP_BEETLE) != 0)
                        {
                            // ビートルデータ作成処理
                            var value = "";
                            value += row["bpo_num"].ToString() + "\t";
                            value += document.Pages.Count.ToString();
                            blt.AppendLine(value);
                        }
                        if((_operation & OP_MACHING) != 0)
                        {
                            // マッチングデータ作成処理
                            var value = "";
                            value += row["bpo_num"].ToString() + ",";
                            value += row["bpo_num"].ToString() + "1,";
                            value += row["bpo_num"].ToString() + $"{document.Pages.Count.ToString()}";
                            maching.AppendLine(value);
                        }
                    }
                }
                document = null; // FixedDocumentの参照を解放

                System.Threading.Thread.Sleep(50); // documentオブジェクトが解放されガベージコレクションが正しく処理されるように少し待機
                GC.Collect(); // ガベージコレクションを実行
                GC.WaitForPendingFinalizers(); // ガベージコレクションの完了を待機
            }
        }

        /// <summary>
        /// FixedDocumentをTIFF画像として保存するメソッド
        /// </summary>
        /// <param name="document"></param>
        /// <param name="row"></param>
        private void FixedDocumentAsTiff(FixedDocument document, string fullPath)
        {
            // フォルダが存在しない場合は作成
            System.IO.Directory.CreateDirectory(fullPath);

            // FixedDocumentの各ページをJPEG画像として保存
            foreach (var pageContent in document.Pages)
            {
                // pageContentがPageContentであることを確認
                if (pageContent is PageContent pc)
                {
                    // PageContentからFixedPageを取得
                    var fixedPage = pc.GetPageRoot(false);
                    if (fixedPage == null) continue;

                    // FixedPageの子要素を取得
                    var child = fixedPage.Children[0];
                    string qrCode = string.Empty;

                    // FubiPageまたはFubiPageNのDataContextからqr_codeを取得
                    if (child is FubiPage view)
                    {
                        qrCode = view.QrCode;
                    }
                    // FubiPageNの場合も同様に処理
                    else if (child is FubiPageN viewN)
                    {
                        qrCode = viewN.QrCode;
                    }

                    // 画像のサイズを計算
                    const double dpi = 200;
                    double width = fixedPage.Width;
                    double height = fixedPage.Height;
                    // ピクセル単位に変換
                    var pixelWidth = (int)(width * dpi / 96);
                    var pixelHeight = (int)(height * dpi / 96);
                    // RenderTargetBitmapを作成
                    var renderBitmap = new RenderTargetBitmap(pixelWidth, pixelHeight, dpi, dpi, PixelFormats.Pbgra32);
                    fixedPage.Measure(new Size(width, height));
                    fixedPage.Arrange(new Rect(new Size(width, height)));
                    renderBitmap.Render(fixedPage);

                    // 2値化（白黒）に変換
                    var bwBitmap = new FormatConvertedBitmap();
                    bwBitmap.BeginInit();
                    bwBitmap.Source = renderBitmap;
                    bwBitmap.DestinationFormat = PixelFormats.BlackWhite;
                    bwBitmap.EndInit();

                    // RenderTargetBitmapをTIFF形式で保存
                    var encoder = new TiffBitmapEncoder();
                    encoder.Compression = TiffCompressOption.Ccitt4; // CCITT4圧縮を使用
                    encoder.Frames.Add(BitmapFrame.Create(bwBitmap));

                    string outputFile = System.IO.Path.Combine(fullPath, $"{qrCode}.tiff");
                    using (var stream = new FileStream(outputFile, FileMode.Create))
                    {
                        encoder.Save(stream);
                    }

                    encoder = null; // TiffBitmapEncoderの参照を解放
                    pc = null; // PageContentの参照を解放
                    child = null; // 子要素の参照を解放
                    renderBitmap = null; // RenderTargetBitmapの参照を解放
                    bwBitmap = null; // 2値化されたBitmapの参照を解放
                    fixedPage = null; // FixedPageの参照を解放
                }
            }
        }

        /// <summary>
        /// FixedDocumentをJPEG画像として保存するメソッド
        /// </summary>
        /// <param name="document"></param>
        private void FixedDocumentAsJpeg(FixedDocument document, string fullPath)
        {
            // フォルダが存在しない場合は作成
            System.IO.Directory.CreateDirectory(fullPath);

            // FixedDocumentの各ページをJPEG画像として保存
            foreach (var pageContent in document.Pages)
            {
                // pageContentがPageContentであることを確認
                if (pageContent is PageContent pc)
                {
                    // PageContentからFixedPageを取得
                    var fixedPage = pc.GetPageRoot(false);
                    if (fixedPage == null) continue;

                    // FixedPageの子要素を取得
                    var child = fixedPage.Children[0];
                    string qrCode = string.Empty;

                    // FubiPageまたはFubiPageNのDataContextからqr_codeを取得
                    if (child is FubiPage view)
                    {
                        qrCode = view.QrCode;
                    }
                    // FubiPageNの場合も同様に処理
                    else if (child is FubiPageN viewN)
                    {
                        qrCode = viewN.QrCode;
                    }

                    // 画像のサイズを計算
                    const double dpi = 200;
                    double width = fixedPage.Width;
                    double height = fixedPage.Height;
                    // ピクセル単位に変換
                    var pixelWidth = (int)(width * dpi / 96);
                    var pixelHeight = (int)(height * dpi / 96);
                    // RenderTargetBitmapを作成
                    var renderBitmap = new RenderTargetBitmap(pixelWidth, pixelHeight, dpi, dpi, PixelFormats.Pbgra32);
                    fixedPage.Measure(new Size(width, height));
                    fixedPage.Arrange(new Rect(new Size(width, height)));
                    renderBitmap.Render(fixedPage);

                    // RenderTargetBitmapをJPEG形式で保存
                    var encoder = new JpegBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

                    string outputFile = System.IO.Path.Combine(fullPath, $"{qrCode}.jpg");
                    using (var stream = new FileStream(outputFile, FileMode.Create))
                    {
                        encoder.Save(stream);
                    }

                    encoder = null; // JpegBitmapEncoderの参照を解放
                    pc = null; // PageContentの参照を解放
                    child = null; // 子要素の参照を解放
                    renderBitmap = null; // RenderTargetBitmapの参照を解放
                    fixedPage = null; // FixedPageの参照を解放
                }
            }
        }
    }
}
