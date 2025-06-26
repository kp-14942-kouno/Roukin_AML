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

namespace MyTemplate
{
    /// <summary>
    /// TemplateWndow.xaml の相互作用ロジック
    /// </summary>
    public partial class TemplateWindow : Window
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public TemplateWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// テーブル作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_TableCreate_Click(object sender, RoutedEventArgs e)
        {
            var form = new Forms.TableList();
            form.ShowDialog();
        }

        /// <summary>
        /// 選択ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_TableSelect_Click(object sender, RoutedEventArgs e)
        {
            if (cmb_TableList.SelectedItem == null) return;
            // 選択されたテーブルのIDを取得
            var tableId = int.Parse(cmb_TableList.SelectedValue.ToString());

            var form = new Forms.DataSearch(tableId);
            form.ShowDialog();
        }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Modules.SetTableList(cmb_TableList);
        }
    }
}
