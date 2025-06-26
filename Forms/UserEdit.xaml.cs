using DocumentFormat.OpenXml.Office.LongProperties;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.Class;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Windows;
using System.Windows.Data;

namespace MyTemplate.Forms
{
    /// <summary>
    /// UserEdit.xaml の相互作用ロジック
    /// </summary>
    public partial class UserEdit : Window
    {
        private LoginParam _loginParam = new LoginParam();
        private int _id;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="id"></param>
        /// <param name="windw"></param>
        public UserEdit(int id, Window? windw = null)
        {
            InitializeComponent();

            this.DataContext = this;

            _id = id;

            // ウィンドウの初期設定
            this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            if (windw != null)
                this.Owner = windw;
            else
                this.Owner = Application.Current.MainWindow;
        }

        /// <summary>
        /// 権限用辞書（xamlのBindingで使用するためpublic）
        /// </summary>
        public　static Dictionary<string, object> AuthorityDict { get; set; }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 初期設定呼出し
            if (!Initialize()) this.Close();
        }

        /// <summary>
        /// 初期設定
        /// </summary>
        /// <returns></returns>
        private bool Initialize()
        {
            try
            {
                // 権限情報の取得
                AuthorityDict = MyUtilityModules.AppSettings("authority");

                // DB接続
                using MyDbData db = new MyDbData("setting");
                // ログイン設定取得
                MyPropertyModules.GetCreateProperties(db, typeof(LoginParam), _loginParam, "t_login_setting", "id", _id);
                if (_loginParam == null) throw new Exception("ログイン設定が取得できませんでした。");

                // ユーザー登録DB接続
                using MyDbData loginDb = new MyDbData(_loginParam.schema);
                // ユーザー登録情報取得
                var table = loginDb.ExecuteQuery($"select * from {_loginParam.table_name};");
                // DataGridに表示
                dg_Users.ItemsSource = table.DefaultView;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初期化に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// 削除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Delete_Click(object sender, RoutedEventArgs e)
        {
            // 選択行の取得
            var selectRow = dg_Users.SelectedItem as System.Data.DataRowView;
            if (selectRow == null) return;

            // ユーザーIDとユーザー名の取得
            var userId = selectRow["user_id"].ToString();
            var userName = selectRow["user_name"].ToString();
            if (string.IsNullOrEmpty(userId)) return;
            // 確認
            if (MyMessageBox.Show($@"ユーザー: {userName} / {userId}　を削除しますか？", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.Warning) != MyEnum.MessageBoxResult.Yes) return;

            try
            {
                // DB接続
                using MyDbData db = new MyDbData(_loginParam.schema);
                // ユーザー情報の削除
                var prms = new List<DbParameter>
                {
                    db.CreateParameter("@user_id", userId)
                };
                db.ExecuteNonQuery($"delete from {_loginParam.table_name} where user_id = @user_id", prms);
                // DataGridの更新
                dg_Users.ItemsSource = db.ExecuteQuery($"select * from {_loginParam.table_name};").DefaultView;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
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
    }
}
