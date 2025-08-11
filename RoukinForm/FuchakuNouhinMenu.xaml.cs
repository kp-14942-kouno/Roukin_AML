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
    /// FuchakuNouhinMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class FuchakuNouhinMenu : Window
    {
        DataTable _table = new DataTable();

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FuchakuNouhinMenu()
        {
            _table?.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FuchakuNouhinMenu()
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
            tb_FuchakuCount.Text = _table.Rows.Count.ToString();
        }

        /// <summary>
        /// 不着納品対象読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsFuchakuData_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(12, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                _table = load.LoadData;
                SetCount();
            }
        }

        private void bt_ExpNouhin_Click(object sender, RoutedEventArgs e)
        {
            // 不着納品対象データがない場合は処理を中止
            if (_table.Rows.Count == 0)
            {
                MyMessageBox.Show("不着納品対象データがありません。");
                return;
            }

            string msg = "不着納品データ";

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
                var exp = new FuchakuNouhinClass(_table, msg, expPath);
                dlg.ThreadClass(exp);
                dlg.ShowDialog();
                // 結果を確認
                if (exp.Result != MyEnum.MyResult.Ok) return;
                // 結果メッセージを表示
                MyMessageBox.Show(exp.ResultMessage);
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
    }
}
