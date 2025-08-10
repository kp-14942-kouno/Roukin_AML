using DocumentFormat.OpenXml.Presentation;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.Report;
using MyTemplate.Report.Helpers;
using MyTemplate.Report.Models;
using MyTemplate.Report.ViewModels;
using MyTemplate.Report.Views;
using MyTemplate.RoukinClass;
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
using System.Windows.Interop;
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
        // 不備文言
        private List<DefectModel> _defectModels = new();
        // 金融機関情報
        private List<BankModel> _bankModel = new();

        /// <summary>
        /// リソースの開放
        /// </summary>
        public void Dispose()
        {
            // リソースの解放
            _table?.Dispose();
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
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            // プリンターリストを設定
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
                if (!FileLoadClass.GetFileLoadSetting(4, file)) return;
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

                _table = _table.AsEnumerable().OrderBy(x => x["taba_num"]).ThenBy(x => int.Parse(x["taba_count"].ToString())).CopyToDataTable();
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
            try
            {
                // マウスカーソルをwaitにする
                Mouse.OverrideCursor = Cursors.Wait;
                // 不備文言の取得
                _defectModels = MyTemplate.Class.DataTableExtensions.ToList<DefectModel>("code", "t_fubi_code");
                // 不備文言が取得できていない場合は例外を投げる
                if (_defectModels.Count == 0) throw new Exception("不備文言の取得に失敗しました。");
                // 金融機関情報を取得
                _bankModel = MyTemplate.Class.DataTableExtensions.ToList<BankModel>("code", "t_financial_code");
                // 不備文言が取得できていない場合は例外を投げる
                if (_bankModel.Count == 0) throw new Exception("金融機関情報の取得に失敗しました。");
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error, true);
                this.Close();
            }
            finally
            {
                // マウスカーソルを元に戻す
                Mouse.OverrideCursor = null;
            }
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
            var document = FubiHelper.CreateFixedDocument(selectedRows.Rows[0], _defectModels, _bankModel);
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
            var msg = string.Empty;

            if ((operation & FubiPrintDocument.OP_PRINT) != 0)
            {
                msg += "・印刷\r\n";
            }
            if ((operation & FubiPrintDocument.OP_IMAGE) != 0)
            {
                msg += "・画像作成\r\n";
            }
            if ((operation & FubiPrintDocument.OP_BEETLE) != 0)
            {
                msg += "・ビートルデータ作成\r\n";
            }
            if((operation & FubiPrintDocument.OP_HIKINUKI) != 0)
            {
                msg += "・引抜きリストの印刷\r\n";
            }
            if((operation & FubiPrintDocument.OP_MACHING) != 0)
            {
                msg += "・マッチングリストの作成\r\n";
            }

            msg = $"不備状\r\n{msg}";
           
            // 選択された行を取得
            var selectedRows = SelectedItem();
            // 選択された行がない場合は処理を終了
            if (selectedRows.Rows.Count == 0) return;
            // 確認メッセージ
            if (MyMessageBox.Show($"{msg}を開始します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.Info) != MyEnum.MessageBoxResult.Yes) return;

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
                var thread = new FubiPrintDocument(selectedRows, printer, _defectModels, _bankModel, operation, msg);
                dlg.ThreadClass(thread);
                dlg.ShowDialog();
                if (thread.Result != MyLibrary.MyEnum.MyResult.Ok)
                {
                    MyMessageBox.Show($"{msg}に失敗しました。");
                }
                else
                {
                    MyMessageBox.Show($"{msg}が完了しました。");
                }
            }
        }
    }
}
