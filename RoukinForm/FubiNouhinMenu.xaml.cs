using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.RoukinClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
    /// FubiNouhinMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiNouhinMenu : Window
    {
        DataTable _table = new DataTable();


        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FubiNouhinMenu()
        {
            _table?.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FubiNouhinMenu()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            SetCount();
        }

        /// <summary>
        /// 件数
        /// </summary>
        private void SetCount()
        {
            tb_FubiCount.Text = _table.Rows.Count.ToString();
        }

        /// <summary>
        /// 不備納品対象データ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsFubiData_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(7, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                _table = load.LoadData;
                SetCount();
            }
        }

        /// <summary>
        /// 納品データ作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_ExpNouhin_Click(object sender, RoutedEventArgs e)
        {
            // 不備納品対象データがない場合は処理を中止
            if (_table.Rows.Count == 0)
            {
                MyMessageBox.Show("不備納品対象データがありません。");
                return;
            }

            // 出力対象が選択されてない場合は中止
            int operation = 0;
            operation |= chk_ExpFubiData.IsChecked == true ? FubiNouhinClass.OP_DATA : 0;
            operation |= chk_ExpFubiCall.IsChecked == true ? FubiNouhinClass.OP_CALL : 0;

            if (operation == 0)
            {
                MyMessageBox.Show("出力対象が選択されていません。");
                return;
            }

            string msg = string.Empty;
            if ((operation & FubiNouhinClass.OP_DATA) != 0)
            {
                msg += "・不備対象データ\r\n";
            }
            if ((operation & FubiNouhinClass.OP_CALL) != 0)
            {
                msg += "・コール連携画像\r\n";
            }

            // 出力先設定を取得
            string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");

            // 出力先指定
            expPath = MyTemplate.Modules.MyFolderDialog(expPath);
            // 出力先が指定されていない場合は処理を中止
            if (string.IsNullOrEmpty(expPath)) return;
            // 確認
            if (MyMessageBox.Show($"{msg}作成を開始します。よろしいですか？", "確認",
                MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

            // ローディングダイアログを表示
            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 処理実行
                var exp = new FubiNouhinClass(_table, operation, msg);
                dlg.ThreadClass(exp);
                dlg.ShowDialog();
                // 結果を確認
                if (exp.Result != MyEnum.MyResult.Ok) return;
                // 結果メッセージを表示
                MyMessageBox.Show(exp.ResultMessage);
            }
        }

        /// <summary>
        /// / 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
