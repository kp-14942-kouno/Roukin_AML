using MyLibrary;
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
            var form = new RoukinForm.Nouhin();
            form.ShowDialog();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DateTime? formattedDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString("19870102", "yyyyMMdd", MyEnum.CalenderType.Western);

            // 設立年月日がyyyy/MM/dd形式に変換できた場合はOK
            if (!string.IsNullOrEmpty(formattedDate.ToString()))
            {
                // 設立年月日が未来日でなければOK
                if (formattedDate <= DateTime.Now) return;
            }
        }
    }
}
