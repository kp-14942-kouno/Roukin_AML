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
using MyTemplate.Class;

namespace MyTemplate.Forms
{
    /// <summary>
    /// ExcelSheetList.xaml の相互作用ロジック
    /// </summary>
    public partial class ExcelSheetList : Window
    {
        private string _sheetName = string.Empty;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="excel"></param>
        public ExcelSheetList(MyStreamExcel excel)
        {
            InitializeComponent();

            this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            this.Owner = Application.Current.MainWindow;

            lst_Sheet.ItemsSource = excel.GetSheetLists();

        }

        /// <summary>
        /// 選択されたシート名を取得
        /// </summary>
        /// <returns></returns>
        public string GetSheetName()
        {
            return _sheetName;
        }

        /// <summary>
        /// OKボタン押下時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_OK_Click(object sender, RoutedEventArgs e)
        {
            SheetSelect();
        }

        /// <summary>
        /// リストボックスのダブルクリック時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lst_Sheeet_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            SheetSelect();
        }

        /// <summary>
        /// リストボックスの選択時の処理
        /// </summary>
        private void SheetSelect()
        {
            if (lst_Sheet.SelectedIndex >= 0)
            {
                _sheetName = lst_Sheet.SelectedValue.ToString();
                this.Close();
            }
        }

        /// <summary>
        /// キャンセルボタン押下時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Cancel_Click(object sender, RoutedEventArgs e)
        {
            _sheetName = string.Empty;
            this.Close();
        }
    }
}
