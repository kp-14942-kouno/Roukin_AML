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
        DataTable _dantai = new DataTable();
        DataTable _kojin = new DataTable();

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FuchakuNouhinMenu()
        {
            _dantai?.Dispose();
            _kojin?.Dispose();
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
            List<string> lst = new List<string>();

            if(_dantai.Rows.Count > 0)
            {
                lst.Add("団体不着");
            }

            if(_kojin.Rows.Count > 0)
            {
                lst.Add("個人不着");
            }

            // 不着納品対象データがない場合は処理を中止
            if (lst.Count == 0)
            {
                MyMessageBox.Show("不着納品対象データがありません。");
                return;
            }

            // 出力するメッセージを作成
            var msg = string.Join("\r\n", lst) + "\r\n";

            // 出力先設定を取得
            string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");

            // 出力先指定
            expPath = MyTemplate.Modules.MyFolderDialog(expPath);
            // 出力先が指定されていない場合は処理を中止
            if (string.IsNullOrEmpty(expPath)) return;
            // 確認
            if (MyMessageBox.Show($"{msg.ToString()}納品データ作成を開始します。よろしいですか？", "確認",
                MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

            // ローディングダイアログを表示
            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 処理実行
                var exp = new FuchakuNouhinClass(_dantai, _kojin, msg.ToString(), expPath);
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

        /// <summary>
        /// 個人・不着納品対象読込ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsKojin_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(13, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                _kojin = load.LoadData;
                SetCount();
            }
        }

        /// <summary>
        /// 団体・不着納品対象読込ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsDantai_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(12, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                _dantai = load.LoadData;
                SetCount();
            }
        }
    }
}
