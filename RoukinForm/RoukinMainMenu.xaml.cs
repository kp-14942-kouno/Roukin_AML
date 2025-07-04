using MyLibrary.MyModules;
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
    public partial class RoukinMainMenu : Window
    {
        public RoukinMainMenu()
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
    }
}
