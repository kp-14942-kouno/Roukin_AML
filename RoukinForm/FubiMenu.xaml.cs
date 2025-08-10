using DocumentFormat.OpenXml.Drawing.Diagrams;
using MyLibrary;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using MyTemplate.RoukinClass;
using Org.BouncyCastle.Asn1.X509;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
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
    /// FubiMenu.xaml の相互作用ロジック
    /// </summary>
    public partial class FubiMenu : Window
    {
        DataTable _table = new();
        DataTable _fubiData = new();
        DataTable _fixData = new();

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FubiMenu() 
        {
            // リソースの解放
            _table?.Dispose();
            _fubiData?.Dispose();
            _fixData?.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FubiMenu()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            SetCount();
        }

        /// <summary>
        /// 件数表示
        /// </summary>
        private void SetCount()
        {
            tb_PnchCount.Text = $"{_table.Rows.Count.ToString()}件";
            tb_FubiCount.Text = $"{_fubiData.Rows.Count.ToString()}件";
            tb_FixCount.Text = $"{_fixData.Rows.Count.ToString()}件";
        }

        /// <summary>
        /// 審査対象データ読込みボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_InsPnch_Click(object sender, RoutedEventArgs e)
        {
            using(var load = new FileLoadProperties())
            {
                if (!FileLoadClass.GetFileLoadSetting(3, load)) return;
                if (FileLoadClass.FileLoad(this, load) != MyLibrary.MyEnum.MyResult.Ok) return;

                // 読込み完了で_tableにデータをセット
                _table = load.LoadData;

                // 各テーブルを初期化
                _fubiData = new DataTable();
                _fixData = new DataTable();

                SetCount();
            }
        }

        /// <summary>
        /// 不備審査処理ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_FubiCheck_Click(object sender, RoutedEventArgs e)
        {
            // 審査対象データが無い場合は処理を中止
            if (_table.Rows.Count == 0)
            {
                MyMessageBox.Show("審査対象データの読込みをしてから実行してください。", "確認", MyEnum.MessageBoxButtons.Ok, window: this);
                return;
            }
            // 審査実行確認
            else if (MyMessageBox.Show("不備審査を実行します。よろしいですか？", "確認", MyEnum.MessageBoxButtons.YesNo, window: this) != MyEnum.MessageBoxResult.Yes)
            {
                return;
            }

            // プログレスダイアログ準備
            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 審査実行
                var chk = new FubiCheck(_table);
                dlg.ThreadClass(chk);
                dlg.ShowDialog();

                _fubiData = chk.FubiData;
                _fixData = chk.FixData;

                if(chk.Result == MyEnum.MyResult.Ok) 
                    MyMessageBox.Show(chk.ResultMessage, buttons: MyEnum.MessageBoxButtons.Ok, window: this);

                SetCount();
            }
        }

        /// <summary>
        /// 審査結果出力ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_ExpData_Click(object sender, RoutedEventArgs e)
        {
            if(_fubiData.Rows.Count == 0 && _fixData.Rows.Count == 0)
            {
                // 審査結果が無い場合は処理を中止
                MyMessageBox.Show("不備審査を実行してから出力してください。", "確認", MyEnum.MessageBoxButtons.Ok, window: this);
                return;
            }

            // 出力先取得
            string expPath = MyTemplate.Modules.MyFolderDialog(MyUtilityModules.AppSetting("roukin_setting", "exp_root_path", false));
            // 出力先が指定されていない場合は処理を中止
            if (string.IsNullOrEmpty(expPath)) return;
            // 確認
            if (MyMessageBox.Show("審査結果を出力します。よろしいですか？", "確認", MyEnum.MessageBoxButtons.YesNo, window: this) != MyEnum.MessageBoxResult.Yes) return;

            // プログレスダイアログ準備
            using (var dlg = new MyLibrary.MyLoading.Dialog(this))
            {
                // 審査結果出力
                var exp = new ExpReviewResult(_fubiData, _fixData, expPath);
                dlg.ThreadClass(exp);
                dlg.ShowDialog();

                if (exp.Result != MyLibrary.MyEnum.MyResult.Ok) return;

                // 出力完了メッセージ
                MyMessageBox.Show(exp.ResultMessage, buttons: MyEnum.MessageBoxButtons.Ok, window: this);
            }
        }

        /// <summary>
        /// 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 不備対象データ一覧表示イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tb_FubiCount_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            // 項目設定
            List<string> columns = new List<string> { "bpo_num", "fubi_code" };
            List<string> headers = new List<string> { "管理番号", "不備コード" };
            // データ一覧表示
            var viewer = new MyLibrary.MyDataViewer(this, _fubiData.DefaultView, "不備データ一覧", columnNames:columns, columnHeaders:headers);
            viewer.ShowDialog();
        }
    }
}
