using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
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
    /// WebcasCngMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class WebcasCngMenu : Window
    {
        // WEBCASデータ読込みテーブル
        private DataTable _table = new DataTable();

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~WebcasCngMenu()
        {
            _table = null;
            _table?.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public WebcasCngMenu()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            SetCount();
        }

        /// <summary>
        /// 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// WEBCASデータ変換作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_ExpWebcas_Click(object sender, RoutedEventArgs e)
        {
            // 出力先ディレクトリを取得
            string fileDir = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
            // ディレクトリ選択ダイアログを表示
            fileDir = MyTemplate.Modules.MyFolderDialog(fileDir);
            // 選択されたディレクトリが空の場合は処理を中止
            if (string.IsNullOrEmpty(fileDir)) return;

            // 確認ダイアログを表示
            if (MyMessageBox.Show("WEBCASの変換ファイルを作成します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

            // ログ作成
            MyLogger.SetLogger("WEBCASデータ変換処理を開始", MyEnum.LoggerType.Info, false);

            // プログレスダイアログを準備
            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // WEBCASデータ変換クラスを生成
                var exp = new EpxWebcasData(_table, fileDir);
                dlg.ThreadClass(exp);
                // ダイアログを表示（実行）
                dlg.ShowDialog();
                // 結果の確認
                if (exp.Result != MyEnum.MyResult.Ok) return;

                // ログに結果を出力
                MyLogger.SetLogger(exp.ResultMessage, MyEnum.LoggerType.Info, true);
            }
        }

        /// <summary>
        /// WEBCASデータ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsWebcas_Click(object sender, RoutedEventArgs e)
        {
            using (var load = new ImportClass.FileLoadProperties())
            {
                // ファイル読込み設定取得
                if (!FileLoadClass.GetFileLoadSetting(1, load)) return;
                // ファイル読込み処理
                if (FileLoadClass.FileLoad(this, load) != MyEnum.MyResult.Ok) return;

                _table = load.LoadData;
            }
            SetCount();
        }

        /// <summary>
        /// 件数表示
        /// </summary>
        private void SetCount()
        {
            tb_WebcasCount.Text = $"{_table.Rows.Count}件";
        }
    }
}
