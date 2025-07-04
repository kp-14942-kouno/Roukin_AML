using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2016.Excel;
using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.Report;
using MyTemplate.Report.Models;
using MyTemplate.Report.ViewModels;
using MyTemplate.Report.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Permissions;
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
            var form = new RoukinForm.FubiPrint();
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
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MyMessageBox.Show("今日は<font color='Blue' size='24'><b><u>雨</u></b></font><br/>明日は<font color='Red' size='24'><b><u>晴れ</u></b><br/></font>");

            /*
            string allowd1Byte = "[0x00,0x00][0x20,0x20][0x28,0x28][0x29,0x29][0x2C,0x39][0x41,0x5A][0x5C,0x5C][0xA2,0xA2][0xA3,0xA3][0xA6,0xDF]";
            string allowd2Byte = "[8140,84BE][889F,EAA2]";
            //string allowd2Byte = string.Empty;
            string allowdUni = "[0x3040,0x309F][0x4E00,0x9FFF][0x1F300,0x1F5FF][0x41,0x5A]";

            string value = "ABCDEFG";
            //string value = "ｱｲｳｴｵﾔﾕﾖｬｭｮ㎝";
            //string value = "ＥＦＧ７８９０あいうえおアイウエオ亜居鵜江尾№";
            //string value = "1234567890";

            //var validator = new  CharValidator(Encoding.GetEncoding("Shift_JIS"), allowd1Byte, allowd2Byte);

            //var validator = new CharValidator(allowdUni);

            //string? result = validator.GetInvalidUnicodeChars(value);

            //MessageBox.Show(result ?? "全て範囲内", "文字チェック結果", MessageBoxButton.OK, MessageBoxImage.Information);

            //string date = "20231031";

            //DateTime? formatterDate = MyLibrary.MyModules.MyUtilityModules.ParseDateString(date, "yyyyMMdd", MyEnum.CalenderType.Western);

            //MessageBox.Show(formatterDate?.ToString("yyyy/MM/dd") ?? "日付の変換に失敗しました。", "日付変換結果", MessageBoxButton.OK, MessageBoxImage.Information);
            */
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// WEBCASデータ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsWebcas_Click(object sender, RoutedEventArgs e)
        {

        }

        private void bt_ExpWebcas_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
