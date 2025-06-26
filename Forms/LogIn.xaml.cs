using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System.Data.Common;
using System.Windows;

namespace MyTemplate.Forms
{
    /// <summary>
    /// LogIn.xaml の相互作用ロジック
    /// </summary>
    public partial class LogIn : Window, IDisposable
    {
        private MyDbData _userDb;
        private LoginParam _loginParam = new();
        private int _id;

        public event EventHandler LoginSucceded;

        /// <summary>
        /// 開放
        /// </summary>
        public void Dispose()
        {
            _userDb?.Dispose();
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="id"></param>
        public LogIn(int id)
        {
            InitializeComponent();
            _id = id;

            // ユーザー情報の初期化
            UserInfo.UserName = string.Empty;
            UserInfo.Authority = 0;

            UserId.Focus();

            try
            {
                // LogIn設定を取得
                using MyDbData db = new MyDbData("setting");
                MyPropertyModules.GetCreateProperties(db, typeof(LoginParam), _loginParam, "t_login_setting", "id", _id);
                // ログイン情報のDBを取得
                _userDb = new MyDbData(_loginParam.schema);
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                this.Close();
            }
        }

        /// <summary>
        /// Loginボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_LogIn_Click(object sender, RoutedEventArgs e)
        {
            // ユーザーIDとパスワードのどちらかが未入力は中断
            if (string.IsNullOrEmpty(UserId.Value) || string.IsNullOrEmpty(Password.Value))
            {
                MyMessageBox.Show(@"<\ #FF0000 ユーザーIDとパスワードを入力してください \>", title:"ログイン失敗", icon:MyEnum.MessageBoxIcon.Warning);
                return;
            }

            // ユーザーIDとパスワードを取得
            string userId = UserId.Value;
            string pass = Password.Value;

            // パスワードをハッシュ化
            pass = MyEncryptModules.Sha256Hash(pass);

            // パラメータ作成
            List<DbParameter> prms = new List<DbParameter>();
            prms.Add(_userDb.CreateParameter("@user_id", userId));
            prms.Add(_userDb.CreateParameter("@password", pass));
            // SQLインジェクション対策のため、パラメータを使用してクエリを実行
            var result = _userDb.ExecuteQuery($"select * from {_loginParam.table_name} where user_id=@user_id and password=@password;", prms);
            // クエリ実行結果を確認
            if (result != null && result.Rows.Count != 0)
            {
                // ログイン情報を取得
                UserInfo.UserName = result.Rows[0]["user_name"].ToString();
                UserInfo.Authority = int.Parse(result.Rows[0]["authority"].ToString());

                // ログイン成功イベント
                LoginSucceded?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                MyMessageBox.Show(@"<\ #FF0000 ユーザーIDかパスワードが違います \>", title: "ログイン失敗", icon: MyEnum.MessageBoxIcon.Warning, window: this);
            }
        }

        /// <summary>
        /// Cancelボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}
