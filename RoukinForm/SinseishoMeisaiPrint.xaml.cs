using Azure;
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
    /// SinseishoMeisaiPrint.xaml の相互作用ロジック
    /// </summary>
    public partial class SinseishoMeisaiPrint : Window
    {
        private DataTable _table = new DataTable(); // 明細リストデータ
        List<BankModel> _bankModel = new List<BankModel>();

        public SinseishoMeisaiPrint()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            // プリンターリストを設定
            Modules.SetPrinterList(cmb_PrinterList);

            SetCount();
        }


        private void bt_InpMeisaiData_Click(object sender, RoutedEventArgs e)
        {
            // ファイル読込みプロパティ
            using (FileLoadProperties file = new FileLoadProperties())
            {
                // ファイル読込み設定
                if (!FileLoadClass.GetFileLoadSetting(13, file)) return;
                // ファイル読込み
                if (FileLoadClass.FileLoad(this, file) != MyLibrary.MyEnum.MyResult.Ok) return;

                _table = new DataTable();
                _table = file.LoadData;

                SetCount();
            }

        }

        private void SetCount()
        {
            tb_MeisaiCount.Text = _table.Rows.Count.ToString();
        }

        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {

        }

        private void bt_Print_Click(object sender, RoutedEventArgs e)
        {
            // 確認メッセージ
            if (MyMessageBox.Show($"明細リストの印刷を開始します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.Info) != MyEnum.MessageBoxResult.Yes) return;

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
                var thread = new SinseishoMeisaiPrintClass(_table, printer, _bankModel);
                dlg.ThreadClass(thread);
                dlg.ShowDialog();
                if (thread.Result != MyLibrary.MyEnum.MyResult.Ok)
                {
                    MyMessageBox.Show($"明細の印刷に失敗しました。");
                }
                else
                {
                    MyMessageBox.Show($"明細の印刷が完了しました。");
                }
            }
        }

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
