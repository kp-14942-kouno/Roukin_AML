using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.Report;
using MyTemplate.Report.Helpers;
using MyTemplate.Report.ViewModels;
using MyTemplate.Report.Views;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using YamlDotNet.Serialization.BufferedDeserialization;

namespace MyTemplate.RoukinForm
{
    /// <summary>
    /// FubiPrint.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiPrint : Window, IDisposable
    {
        // 不備状データ用テーブル
        private DataTable _table = new();
        // 不備文言用辞書
        Dictionary<string, string> _defectDic = new();

        /// <summary>
        /// リソースの開放
        /// </summary>
        public void Dispose()
        {
            // リソースの解放
            _table?.Dispose();
            _defectDic?.Clear();
        }

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FubiPrint()
        {
            // Disposeメソッドを呼び出す
            Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FubiPrint()
        {
            InitializeComponent();
            Owner = Application.Current.MainWindow;
            Modules.SetPrinterList(cmb_PrinterList);
            SetCount();
        }

        /// <summary>
        /// 件数表示
        /// </summary>
        private void SetCount()
        {
            tb_fubiCount.Text = $"{_table.Rows.Count.ToString()}件";
        }

        /// <summary>
        /// 不備文言取得
        /// </summary>
        /// <returns></returns>
        private bool GetDefectDic()
        {
            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog())
            {
                var dic = new GetDefectDic();
                dlg.ThreadClass(dic);
                dlg.ShowDialog();

                if(dic.Result != MyLibrary.MyEnum.MyResult.Ok)
                {
                    return false;
                }
                _defectDic = dic.DefectDic;
            }
            return true;
        }

        /// <summary>
        /// 不備状データ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InpFubiData_Click(object sender, RoutedEventArgs e)
        {
            // ファイル読込みプロパティ
            using (FileLoadProperties file = new FileLoadProperties())
            {
                // ファイル読込み設定
                if (!FileLoadClass.GetFileLoadSetting(50, file)) return;
                // ファイル読込み
                if (FileLoadClass.FileLoad(this, file) != MyLibrary.MyEnum.MyResult.Ok) return;

                _table = new DataTable();
                _table = file.LoadData;

                // テーブルに選択項目を追加
                var isSelectedColumn = new DataColumn("IsSelected", typeof(bool))
                {
                    DefaultValue = false
                };
                _table.Columns.Add(isSelectedColumn);

                dg_FubiList.ItemsSource = _table.DefaultView;

                SetCount();
            }
        }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 不備文言を取得
            if (!GetDefectDic()) Close();
        }

        /// <summary>
        /// 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Closedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closed(object sender, EventArgs e)
        {
            Dispose();
        }

        /// <summary>
        /// 全選択ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_AllSelect_Click(object sender, RoutedEventArgs e)
        {
            // 全行のIsSelected列をtrueに設定
            _table.AsEnumerable().ToList().ForEach(row => row["IsSelected"] = true);

            // DataGridを更新
            dg_FubiList.Items.Refresh();
        }

        /// <summary>
        /// 全解除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_AllUnSelect_Click(object sender, RoutedEventArgs e)
        {
            // 全行のIsSelected列をtrueに設定
            _table.AsEnumerable().ToList().ForEach(row => row["IsSelected"] = false);

            // DataGridを更新
            dg_FubiList.Items.Refresh();
        }

        /// <summary>
        /// 不備状プレビューボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Preview_Click(object sender, RoutedEventArgs e)
        {
            // 選択された行を取得
            var selectedRows = SelectedItem();
            // 選択された行がない場合は処理を終了
            if (selectedRows.Rows.Count == 0) return;

            // プレビューで選択できるのは1つのみ
            if (selectedRows.Rows.Count > 1)
            {
                MyMessageBox.Show("プレビューで選択できるのは１つのみです。");
                return;
            }
            // 選択された行の不備状のFixedDocumentを作成
            var document = FubiHelper.CreateFixedDocument(selectedRows.Rows[0], _defectDic);
            // FixedDocumentをReportPreivewに渡して表示
            var form = new ReportPreivew(document);
            form.ShowDialog();

            // 開放
            document = null;
        }

        /// <summary>
        /// DataGridの選択項目を取得
        /// </summary>
        /// <returns></returns>
        private DataTable SelectedItem()
        {
            var selectedRows = _table.AsEnumerable()
                .Where(row => row.Field<bool>("IsSelected"));

            if(!selectedRows.Any())
            {
                return new DataTable();
            }
            return selectedRows.CopyToDataTable();
        }

        /// <summary>
        /// 不備状印刷・画像作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Print_Click(object sender, RoutedEventArgs e)
        {
            int oparation = 0;

            oparation |= chk_FubiPrint.IsChecked == true ? FubiPrintDocument.OP_PRINT : 0;
            oparation |= chk_FubiImage.IsChecked == true ? FubiPrintDocument.OP_IMAGE : 0;
            oparation |= chk_ExpData.IsChecked == true ? FubiPrintDocument.OP_BEETLE : 0;
            oparation |= chk_HikinukiPrit.IsChecked == true ? FubiPrintDocument.OP_HIKINUKI : 0;
            oparation |= chk_ExpMaching.IsChecked == true ? FubiPrintDocument.OP_MACHING : 0;

            // チェックが無ければ中断
            if (oparation == 0) return;

            ExpPrintAndData(oparation);
        }


        /// <summary>
        /// 印刷・画像作成・ビートルデータ作成を一括で行うメソッド
        /// </summary>
        /// <param name="operation"></param>
        private void ExpPrintAndData(int operation)
        {
            var print = string.Empty;
            var image = string.Empty;
            var beetle = string.Empty;

            if((operation & FubiPrintDocument.OP_PRINT) != 0)
            {
                print = "・印刷";
            }
            if ((operation & FubiPrintDocument.OP_IMAGE) != 0)
            {
                image = "・画像作成";
            }
            if ((operation & FubiPrintDocument.OP_BEETLE) != 0)
            {
                beetle = "・ビートルデータ作成";
            }
            if((operation & FubiPrintDocument.OP_HIKINUKI) != 0)
            {
                print += "・引抜きリスト";
            }
            if((operation & FubiPrintDocument.OP_MACHING) != 0)
            {
                image += "・マッチングリスト";
            }

            var message = $"{print}{image}{beetle}";

            // 選択された行を取得
            var selectedRows = SelectedItem();
            // 選択された行がない場合は処理を終了
            if (selectedRows.Rows.Count == 0) return;
            // 確認メッセージ
            if (MyMessageBox.Show($"選択された不備状の{message}を開始します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.Info) != MyEnum.MessageBoxResult.Yes) return;

            // 選択されたプリンターを取得
            if (cmb_PrinterList.SelectedItem is not PrintQueue printer)
            {
                MyMessageBox.Show("プリンターが選択されていません。");
                return;
            }
            // プログレスダイアログ準備
            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 不備状印刷・画像作成スレッドを実行
                var thread = new FubiPrintDocument(selectedRows, _defectDic, printer, operation, message);
                dlg.ThreadClass(thread);
                dlg.ShowDialog();
                if (thread.Result != MyLibrary.MyEnum.MyResult.Ok)
                {
                    MyMessageBox.Show($"{message}に失敗しました。");
                }
                else
                {
                    MyMessageBox.Show($"{message}が完了しました。");
                }
            }
        }
    }

    /// <summary>
    /// 不備文言データを取得するクラス
    /// </summary>
    public class GetDefectDic : MyLibrary.MyLoading.Thread
    {
        // 不備文言用辞書
        public Dictionary<string, string> DefectDic { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            ProgressBarType = MyEnum.MyProgressBarType.None;
            ProcessName = "不備文言データ取得中";

            try
            {
                // 不備文言データを取得
                using (MyDbData db = new MyDbData("code"))
                {
                    using (DbDataReader reader = db.ExecuteReader("select * from t_fubi_code order by fubi_code"))
                    {
                        while (reader.Read())
                        {
                            DefectDic.Add(reader["fubi_code"].ToString(), reader["fubi_caption"].ToString());
                        }
                    }
                }
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
    }

    /// <summary>
    /// 不備状印刷・画像作成・ビートルデータ作成を行うクラス
    /// </summary>
    public class FubiPrintDocument : MyLibrary.MyLoading.Thread
    {
        public const int OP_PRINT = 1 << 0;  // 印刷
        public const int OP_IMAGE = 1 << 1;  // 画像作成
        public const int OP_BEETLE = 1 << 2; // ビートルデータ作成
        public const int OP_HIKINUKI = 1 << 3; // 引抜き
        public const int OP_MACHING = 1 << 4; // マッチング

        private int _operation; // 実行する操作
        private DataTable _table; // 不備状データ用テーブル
        private Dictionary<string, string> _defectDic; // 不備文言用辞書
        private PrintQueue _printer; // 印刷するプリンター
        private string _msg; // メッセージ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="defectDic"></param>
        /// <param name="printer"></param>
        /// <param name="operation"></param>
        public FubiPrintDocument(DataTable table, Dictionary<string, string> defectDic, PrintQueue printer, int operation, string msg)
        {
            _table = table;
            _defectDic = defectDic;
            _printer = printer;
            _operation = operation;
            _msg = msg;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            ProgressBarType = MyEnum.MyProgressBarType.Percent;
            ProcessName = $"不備状{_msg}処理中...";
            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;

            // ビートルデータ用
            StringBuilder sb = new();

            try
            {
                foreach (DataRow row in _table.Rows)
                {
                    ProgressValue++;

                    // 不備状のFixedDocumentを作成
                    // Dispatcher.Invokeを使用してUIスレッドでFixedDocumentを作成
                    FixedDocument document = null;
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        document = FubiHelper.CreateFixedDocument(row, _defectDic);
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
                        // 画像作成処理
                        // Dispatcher.Invokeを使用してUIスレッドで画像作成処理を実行
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            FixedDocumentAsJpeg(document, row);
                            FixedDocumentAsTiff(document, row);
                        });
                    }
                    if ((_operation & OP_BEETLE) != 0)
                    {
                        // ビートルデータ作成処理
                        var value = "";
                        value += row["qr_code"].ToString() + "\t";
                        value += document.Pages.Count.ToString();
                        sb.AppendLine(value);
                    }

                    document = null; // FixedDocumentの参照を解放

                    System.Threading.Thread.Sleep(50); // documentオブジェクトが解放されガベージコレクションが正しく処理されるように少し待機
                    GC.Collect(); // ガベージコレクションを実行
                    GC.WaitForPendingFinalizers(); // ガベージコレクションの完了を待機
                }

                // ビートルデータの保存
                if (sb.Length > 0 && (_operation & OP_BEETLE) != 0)
                {
                    var path = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
                    var fileName = MyUtilityModules.AppSetting("roukin_setting", "fubi_print_name", true);
                    System.IO.File.WriteAllText(System.IO.Path.Combine(path, fileName), sb.ToString());
                }

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
        /// FixedDocumentをTIFF画像として保存するメソッド
        /// </summary>
        /// <param name="document"></param>
        /// <param name="row"></param>
        private void FixedDocumentAsTiff(FixedDocument document, DataRow row)
        {
            var path = MyUtilityModules.AppSetting("roukin_setting", "img_root_path");
            var dir = "不備状";
            var fullPath = System.IO.Path.Combine(path, dir, row["taba_num"].ToString());

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
                        if (view.DataContext is PersonViewModel person)
                        {
                            qrCode = person.Item.qr_code;
                        }

                    }
                    // FubiPageNの場合も同様に処理
                    else if (child is FubiPageN viewN)
                    {
                        if (viewN.DataContext is PersonViewModel person)
                        {
                            qrCode = person.Item.qr_code;
                        }
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
        private void FixedDocumentAsJpeg(FixedDocument document, DataRow row)
        {
            var path = MyUtilityModules.AppSetting("roukin_setting", "img_root_path");
            var dir = "不備状";
            var fullPath = System.IO.Path.Combine(path, dir, row["taba_num"].ToString());

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
                        if (view.DataContext is PersonViewModel person)
                        {
                            qrCode = person.Item.qr_code;
                        }

                    }
                    // FubiPageNの場合も同様に処理
                    else if (child is FubiPageN viewN)
                    {
                        if (viewN.DataContext is PersonViewModel person)
                        {
                            qrCode = person.Item.qr_code;
                        }
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
