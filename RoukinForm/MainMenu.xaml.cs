using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.Report;
using MyTemplate.Report.Views;
using System;
using System.Collections.Generic;
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
    /// RoukinMainMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class MainMenu : Window
    {
        /// <summary>
        ///  コンストラクタ
        /// </summary>
        public MainMenu()
        {
            InitializeComponent();

            //var title = MyUtilityModules.AppSetting("projectSettings", "projectName");
            //var version = MyUtilityModules.AppSetting("projectSettings", "projectVersion");

            var title = Assembly
                            .GetExecutingAssembly()
                            .GetCustomAttribute<AssemblyTitleAttribute>()?
                            .Title;

            var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
            Console.WriteLine(version);

            Title = $"{title} 　Ver.{version}";
        }

        /// <summary>
        /// WEBCASデータ変換ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_WebcasCngMenu_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new WebcasCngMenu();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 不備状印刷MENUボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FubiPrint_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.FubiPrint();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 不備審査MENUボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FubiCheck_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.FubiMenu();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 勘定系・本人確認納品MENUボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_RoukinNouhinMenu_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.NouhinMenu();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 金庫事務納品MENUボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_KinkoNouhinMenu_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.KinkojimMenu();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 不備納品MENUボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FubiNouhinMenu_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.FubiNouhinMenu();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 申請書仕分リストMENUボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_ShiwakeList_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.SiwakePrint();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// パンチ画像作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Punc_Click(object sender, RoutedEventArgs e)
        {
            using(ImportClass.FileLoadProperties load = new ImportClass.FileLoadProperties())
            {
                // ファイル読込み設定取得
                if(!ImportClass.FileLoadClass.GetFileLoadSetting(2, load)) return;
                // ファイル読込み実行
                if (ImportClass.FileLoadClass.FileLoad(this, load) != MyEnum.MyResult.Ok) return;

                // 確認
                if (MyMessageBox.Show("パンチ画像をZIP化します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

                // 出力先パス
                string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
                // パスが存在しない場合はダイアログで取得
                expPath = MyTemplate.Modules.MyFolderDialog(expPath);
                if (string.IsNullOrEmpty(expPath)) return;

                using(var dlg = new MyLibrary.MyLoading.Dialog(this))
                {
                    // パンチ画像作成クラスを生成
                    var exp = new ExpPunchImage(load.LoadData, expPath);
                    // スレッド実行
                    dlg.ThreadClass(exp);
                    // ダイアログを表示
                    dlg.ShowDialog();
                    // 結果確認
                    if (exp.Result != MyEnum.MyResult.Ok) return;
                    // 結果メッセージを表示
                    MyMessageBox.Show(exp.ResultMessage);
                }
            }
        }

        private void bt_FubiFuchakuList_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.FubiFuchakuPrint();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        private void bt_FuchakuMenu_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.FuchakuNouhinMenu();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }

        private void bt_CreateZip_Click(object sender, RoutedEventArgs e)
        {
            // 確認
            if (MyMessageBox.Show("不備納品ファイルをZIP化します。", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.None) != MyEnum.MessageBoxResult.Yes) return;

            // 出力先パス
            string expPath = MyUtilityModules.AppSetting("roukin_setting", "exp_root_path");
            // パスが存在しない場合はダイアログで取得
            expPath = MyTemplate.Modules.MyFolderDialog(expPath);
            if (string.IsNullOrEmpty(expPath)) return;

            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // パンチ画像作成クラスを生成
                var exp = new FubiZipCreate(expPath);
                // スレッド実行
                dlg.ThreadClass(exp);
                // ダイアログを表示
                dlg.ShowDialog();
                // 結果確認
                if (exp.Result != MyEnum.MyResult.Ok) return;
                // 結果メッセージを表示
                MyMessageBox.Show(exp.ResultMessage);
            }
        }

        private void bt_SinseishoMeisaiList_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
            var form = new RoukinForm.SinseishoMeisaiPrint();
            form.ShowDialog();
            this.Visibility = Visibility.Visible;
        }
    }
}
