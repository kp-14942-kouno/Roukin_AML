using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.Report;
using MyTemplate.Report.Views;
using System;
using System.Collections.Generic;
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

            var title = MyUtilityModules.AppSetting("projectSettings", "projectName");
            var version = MyUtilityModules.AppSetting("projectSettings", "projectVersion");

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

        private void bt_Kinou_Click(object sender, RoutedEventArgs e)
        {
            string reg = @"^([6][2][0][0]([0-2][0-9]|[3][0-3])|[\-][9][9][9])$";
            bool isMatch = System.Text.RegularExpressions.Regex.IsMatch("620034", reg);

            MyMessageBox.Show($"正規表現マッチ: {isMatch}", "確認", MyEnum.MessageBoxButtons.Ok, MyEnum.MessageBoxIcon.Info);
        }
    }
}
