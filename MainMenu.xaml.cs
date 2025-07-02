using DocumentFormat.OpenXml.Drawing.Diagrams;
using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.Report;
using MyTemplate.Report.Models;
using MyTemplate.Report.ViewModels;
using MyTemplate.Report.Views;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
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

namespace MyTemplate
{
    /// <summary>
    /// MainMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class MainMenu : Window
    {
        public MainMenu()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var form = new FubiPrint();
            form.ShowDialog();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (var load = new ImportClass.FileLoadProperties())
            {
                // ファイル読込み設定取得
                if (!FileLoadClass.GetFileLoadSetting(5, load)) return;
                // ファイル読込み処理
                if (FileLoadClass.FileLoad(this, load) != MyEnum.MyResult.Ok) return;

                dg_View.ItemsSource = load.LoadData.DefaultView;
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string date = "20231031";

            DateTime? formatterDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString(date, "yyyyMMdd", MyEnum.CalenderType.Western);
            
            MessageBox.Show(formatterDate?.ToString("yyyy/MM/dd") ?? "日付の変換に失敗しました。", "日付変換結果", MessageBoxButton.OK, MessageBoxImage.Information);

        }
    }
}
