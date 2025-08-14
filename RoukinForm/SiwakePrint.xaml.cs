using MyLibrary;
using MyLibrary.MyClass;
using MyTemplate.ImportClass;
using MyTemplate.Report.Models;
using MyTemplate.RoukinClass;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace MyTemplate.RoukinForm
{
    /// <summary>
    /// ShiwakePrint.xaml の相互作用ロジック
    /// </summary>
    public partial class SiwakePrint : Window
    {
        DataTable _table = new DataTable();
        List<BankModel> _bankModel = new List<BankModel>();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SiwakePrint()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            // プリンターリストを設定
            Modules.SetPrinterList(cmb_PrinterList);

            // 件数表示
            SetCount();
        }

        /// <summary>
        /// 件数表示
        /// </summary>
        private void SetCount()
        {
            tb_ShiwakeCount.Text = _table.Rows.Count.ToString();
        }

        /// <summary>
        /// 仕分け対象読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InpShiwake_Click(object sender, RoutedEventArgs e)
        {
            // ファイル読込みプロパティ
            using (FileLoadProperties file = new FileLoadProperties())
            {
                // ファイル読込み設定
                if (!FileLoadClass.GetFileLoadSetting(10, file)) return;
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

                // 金融機関コード・束番号・束内連番でソート
                _table = _table.AsEnumerable().OrderBy(x => x["bpo_bank_code"])
                    .ThenBy(x => x["taba_num"].ToString())
                    .ThenBy(x => int.Parse(x["taba_count"].ToString())).CopyToDataTable();
                
                dg_List.ItemsSource = _table.DefaultView;

                SetCount();
            }
        }

        /// <summary>
        /// 印刷・データ作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Print_Click(object sender, RoutedEventArgs e)
        {
            int oparation = 0;

            oparation |= chk_ShiwakePrint.IsChecked == true ? SiwakePrintClass.OP_PRINT : 0;
            oparation |= chk_ShiwakeData.IsChecked == true ? SiwakePrintClass.OP_MACHING : 0;
            oparation |= chk_MeisaiPrint.IsChecked == true ? SiwakePrintClass.OP_MEISAI : 0;

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

            if((operation & SiwakePrintClass.OP_MEISAI) != 0)
            {
                msg += "・明細\r\n";
            }
            if ((operation & SiwakePrintClass.OP_PRINT) != 0)
            {
                msg += "・印刷\r\n";
            }
            if ((operation & SiwakePrintClass.OP_MACHING) != 0)
            {
                msg += "・データ作成\r\n";
            }

            msg = $"仕分けリスト\r\n{msg}";

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
                var thread = new SiwakePrintClass(selectedRows, printer, _bankModel, operation, msg);
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

        /// <summary>
        /// DataGridの選択項目を取得
        /// </summary>
        /// <returns></returns>
        private DataTable SelectedItem()
        {
            var selectedRows = _table.AsEnumerable()
                .Where(row => row.Field<bool>("IsSelected"));

            if (!selectedRows.Any())
            {
                return new DataTable();
            }
            return selectedRows.CopyToDataTable();
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
            dg_List.Items.Refresh();
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
            dg_List.Items.Refresh();
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
        /// Loadedイベントハンドラ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // マウスカーソルをwaitにする
                Mouse.OverrideCursor = Cursors.Wait;
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
    }
}
