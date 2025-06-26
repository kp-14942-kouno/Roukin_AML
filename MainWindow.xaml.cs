using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using Jint;
using Jint.Native;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace MyTemplate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public event EventHandler LogoutRequested;

        public MainWindow()
        {
            InitializeComponent();

            // ログインユーザー名を表示
            this.Title = $"{UserInfo.UserName} : {UserInfo.AuthorityName}";

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

        /// <summary>
        /// テーブル作成ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_CreataTable_Click(object sender, RoutedEventArgs e)
        {
            var form = new Forms.TableList();
            form.ShowDialog();
        }

        /// <summary>
        /// ファイル読込（Text）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FileLoadText_Click(object sender, RoutedEventArgs e)
        {
            using FileLoadProperties load = new FileLoadProperties();

            if(!FileLoadClass.GetFileLoadSetting(1, load)) return;
            if (FileLoadClass.FileLoad(this, load) != MyEnum.MyResult.Ok) return;

            dg_Viewer.ItemsSource = load.LoadData.DefaultView;
        }

        /// <summary>
        /// ファイル読込（Excel）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FileLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            using FileLoadProperties load = new FileLoadProperties();

            if (!FileLoadClass.GetFileLoadSetting(2, load)) return;
            if (FileLoadClass.FileLoad(this, load) != MyEnum.MyResult.Ok) return;

            dg_Viewer.ItemsSource = load.LoadData.DefaultView;
        }

        /// <summary>
        /// フィアル取込（Text）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FileImportText_Click(object sender, RoutedEventArgs e)
        {
            using FileImportProperties import = new FileImportProperties();

            if (!FileImportClass.GetFileImportSetting(1, import)) return;
            if (FileImportClass.FileImport(this, import) != MyEnum.MyResult.Ok) return;
        }

        private void bt_FileLoadTextFromImport_Click(object sender, RoutedEventArgs e)
        {
            using FileLoadProperties load = new FileLoadProperties();

            if (!FileLoadClass.GetFileLoadSetting(1, load)) return;
            if (FileLoadClass.FileLoad(this, load) != MyEnum.MyResult.Ok) return;

            using FileImportProperties import = new FileImportProperties();

            if(!FileImportClass.GetFileImportSetting(2, import)) return;
            if (FileImportClass.FileImport(this, import, load.LoadData) != MyEnum.MyResult.Ok) return;
        }

        private void bt_InputBoxTest_Click(object sender, RoutedEventArgs e)
        {
            var value = MyInputBox.Show(this, "テスト", "入力してください", isPassword: true);

            Debug.Print($"入力値: {value}");
        }

        private void bt_UserRegist_Click(object sender, RoutedEventArgs e)
        {
            var form = new Forms.UserRegistration(1, this);
            form.ShowDialog();
        }

        private void bt_UserView_Click(object sender, RoutedEventArgs e)
        {
            var form = new Forms.UserEdit(1, this);
            form.ShowDialog();
        }

        private void bt_Logout_Click(object sender, RoutedEventArgs e)
        {
            LogoutRequested?.Invoke(this, EventArgs.Empty);
        }

        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}