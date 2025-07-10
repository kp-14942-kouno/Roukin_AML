using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.RoukinClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
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
    /// Nouhin.xaml の相互作用ロジック
    /// </summary>
    public partial class NouhinMenu : Window
    {
        private DataTable _kojin = new();
        private DataTable _dantai = new();

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~NouhinMenu()
        {
            // データテーブルの破棄
            _kojin?.Dispose();
            _dantai?.Dispose(); 
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public NouhinMenu()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            SetCount();
        }

        /// <summary>
        /// 件数表示
        /// </summary>
        private void SetCount()
        {
            tb_DantaiCount.Text = _dantai.Rows.Count.ToString();
            tb_KojinCount.Text = _kojin.Rows.Count.ToString();
        }

        /// <summary>
        /// 納品データ作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_ExpNouhin_Click(object sender, RoutedEventArgs e)
        {
            List<string> zero = new();

            if(_dantai.Rows.Count == 0) zero.Add("団体");
            if(_kojin.Rows.Count == 0) zero.Add("個人");

            // 納品対象データが存在しない場合はメッセージを表示
            if (zero.Count == 2)
            {
                if (MyMessageBox.Show("納品対象データが存在しません。0件データを作成しますか？", "確認",
                                            MyEnum.MessageBoxButtons.YesNo, window:this) != MyEnum.MessageBoxResult.Yes) return;
            }
            else if (zero.Count == 1)
            {
                if (MyMessageBox.Show($"{zero[0].ToString()}のデータが存在しません。作成しますか？", "確認",
                                            MyEnum.MessageBoxButtons.YesNo, window:this) != MyEnum.MessageBoxResult.Yes) return;
            }

            // 出力先設定を取得
            string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");

            // 出力先指定
            expPath = MyTemplate.Modules.MyFolderDialog(expPath);
            // 出力先が指定されていない場合は処理を中止
            if (string.IsNullOrEmpty(expPath)) return;
            // 確認
            if (MyMessageBox.Show("納品データを出力します。よろしいですか？", "確認", 
                MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

            // ローディングダイアログを表示
            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 処理実行
                var exp = new NouhinClass(_dantai, _kojin, expPath);
                dlg.ThreadClass(exp);
                dlg.ShowDialog();
                // 結果を確認
                if (exp.Result != MyEnum.MyResult.Ok) return;
                // 結果メッセージを表示
                MyMessageBox.Show(exp.ResultMessage);
            }
        }

        /// <summary>
        /// （個人）納品対象データ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsKojinData_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(6, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                _kojin = load.LoadData;
                SetCount();
            }
        }

        /// <summary>
        /// （団体）納品対象データ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsDantaiData_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(5, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                _dantai = load.LoadData;
                SetCount();
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
