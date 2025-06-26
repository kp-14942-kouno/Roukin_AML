using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using System.Data.Common;
using System.Windows;
using System.Windows.Controls;

namespace MyTemplate.Forms
{
    /// <summary>
    /// UserEdit.xaml の相互作用ロジック
    /// </summary>
    public partial class UserRegistration : Window
    {
        private Dictionary<string, object> _authoritys = [];
        private LoginParam _loginParam = new();
        private int _id;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="id"></param>
        /// <param name="windw"></param>
        public UserRegistration(int id, Window? windw = null)
        {
            InitializeComponent();

            _id = id;

            this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            if (windw != null)
                this.Owner = windw;
            else
                this.Owner = Application.Current.MainWindow;
        }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 初期設定呼出し
            if (!Initialize()) this.Close();

            // 初期フォーカス
            UserId.Focus();
        }

        /// <summary>
        /// 初期設定
        /// </summary>
        /// <returns></returns>
        private bool Initialize()
        {
            try
            {
                // DB接続
                using MyDbData db = new MyDbData("setting");
                // ログイン設定取得
                MyPropertyModules.GetCreateProperties(db, typeof(LoginParam), _loginParam, "t_login_setting", "id", _id);

                if(_loginParam == null) throw new Exception("ログイン設定が取得できませんでした。");

                // UserName設定
                this.UserName.Parameter = new UserControls.MyEditTextBox.EditParameter
                {
                    DataType = 0,
                    LengthMin = _loginParam.user_name_length_min,
                    LengthMax = _loginParam.user_name_length_max,
                    IsPassword = false
                };

                // UserId設定
                this.UserId.Parameter = new UserControls.MyEditTextBox.EditParameter
                {
                    DataType = 5,
                    LengthMin = _loginParam.user_length_min,
                    LengthMax = _loginParam.user_length_max,
                    IsPassword = false
                };

                // PassWord設定
                this.PassWord.Parameter = new UserControls.MyEditTextBox.EditParameter
                {
                    DataType = 7,
                    LengthMin = _loginParam.pass_length_min,
                    LengthMax = _loginParam.pass_length_max,
                    IsPassword = true
                };

                // 権限設定呼出し
                SetAuthority();

                return true;
            }
            catch(Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return false;
            }
        }
        
        /// <summary>
        /// 権限を設定
        /// </summary>
        private void SetAuthority()
        {
            _authoritys = MyUtilityModules.AppSettings("authority");

            foreach (var item in _authoritys)
            {
                cmb_Authority.Items.Add(new ComboBoxItem
                {
                    Content = item.Key,
                    Tag = item.Value
                });
            }
            cmb_Authority.SelectedIndex = 0; // 初期選択
        }

        /// <summary>
        /// 登録ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Save_Click(object sender, RoutedEventArgs e)
        {
            // 入力チェック
            if (UserId.IsError || UserName.IsError || PassWord.IsError)
            {
                MyMessageBox.Show("入力内容に誤りがあります。", "エラー", icon: MyEnum.MessageBoxIcon.Error, window: this);
                return;
            }

            // 確認
            if (MyMessageBox.Show("登録しますか？", "確認", MyEnum.MessageBoxButtons.YesNo, MyEnum.MessageBoxIcon.Info, window: this) != MyEnum.MessageBoxResult.Yes) return;
          
            try
            {
                // DB接続
                using MyDbData db = new MyDbData(_loginParam.schema);

                // ユーザIDの重複チェック
                var userId = db.ExecuteScalar($"select count(*) from {_loginParam.table_name} where user_id = @user_id", new List<DbParameter> { db.CreateParameter("@user_id", UserId.Value) });
                if (userId != null && Convert.ToInt32(userId) > 0)
                {
                    MyMessageBox.Show("このユーザIDはすでに登録されています。", "エラー", icon: MyEnum.MessageBoxIcon.Error, window: this);
                    return;
                }

                try
                {
                    // トランザクション開始
                    db.BindTransaction();

                    // ユーザ情報の登録
                    var prms = new List<DbParameter>();
                    prms.Add(db.CreateParameter("@user_id", UserId.Value));
                    prms.Add(db.CreateParameter("@user_name", UserName.Value));
                    prms.Add(db.CreateParameter("@password", MyEncryptModules.Sha256Hash(PassWord.Value)));

                    var authority = cmb_Authority.SelectedItem as ComboBoxItem;
                    prms.Add(db.CreateParameter("@authority", authority.Tag));

                    db.ExecuteNonQuery($"insert into {_loginParam.table_name} (user_id, user_name, password, authority) values (@user_id, @user_name, @password, @authority)", prms);
                    
                    // コミット
                    db.CommitTransaction();
                  
                    MyMessageBox.Show("登録が完了しました。", "完了", icon: MyEnum.MessageBoxIcon.Info, window: this);
                    this.Close();
                }
                catch (Exception ex)
                {
                    db.RollbackTransaction();
                    throw;
                }
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                this.Close();
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
