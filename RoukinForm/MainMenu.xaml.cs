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

        private void bt_FubiPrint_Click(object sender, RoutedEventArgs e)
        {
            var form = new RoukinForm.FubiPrint();
            form.ShowDialog();
        }

        private void bt_NouhinMenu_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
 
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

        private void bt_RoukinNouhinMenu_Click(object sender, RoutedEventArgs e)
        {
            var form = new RoukinForm.NouhinMenu();
            form.ShowDialog();
        }

        private void bt_KinkoNouhinMenu_Click(object sender, RoutedEventArgs e)
        {
            var form = new RoukinForm.KinkojimMenu();
            form.ShowDialog();
        }

        private void bt_FubiNouhinMenu_Click(object sender, RoutedEventArgs e)
        {
            var form = new RoukinForm.FubiNouhinMenu();
            form.ShowDialog();
        }
    }
}
