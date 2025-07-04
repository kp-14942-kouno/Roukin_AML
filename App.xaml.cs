using MyTemplate.Forms;
using System.Configuration;
using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Documents;

namespace MyTemplate
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        // Mutex発行
        private System.Threading.Mutex _mutex = new System.Threading.Mutex(false, MyLibrary.MyModules.MyUtilityModules.AppSetting("projectSettings", "projectName"));

        /// <summary>
        /// アプリケーション開始
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Application_Start(object sender, StartupEventArgs e)
        {
            // 文字エンコーディングプロバイダーの登録
            // .NET5以降ではデフォルトで利用できる文字エンコーディングがUTF-8やUTF-16に限定されている
            // SJISやEUC-JPなどの日本語を含むレガシーエンコーディングを利用するためRegisterProviderに登録が必要
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Mutexの所有権を要求
            if (!_mutex.WaitOne(0, false))
            {
                // すでに起動していると判断して終了
                MessageBox.Show("すでに起動しています。");
                _mutex.Dispose();
                _mutex.Close();
                this.Shutdown();
            }

            // 例外処理イベントを取得
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            var windwos = new RoukinForm.RoukinMainMenu();
            Application.Current.MainWindow = windwos;
            windwos.Show();


            //var mainWindow = new MainWindow();
            //Application.Current.MainWindow = mainWindow;
            //mainWindow.Show();

            // ログイン画面表示
            //ShowLogin();
        }

        /// <summary>
        /// ログイン画面表示
        /// </summary>
        private void ShowLogin()
        {
            // ログイン設定ID取得
            var id = int.Parse(MyLibrary.MyModules.MyUtilityModules.AppSetting("projectSettings", "login_id"));
            var login = new Forms.LogIn(id);
            var isLogIn = false;
            // ログインイベント取得
            login.LoginSucceded += (s, args) =>
            {
                // ログイン成功でMainWindow表示
                isLogIn = true;
                ShowMainWindow();
                login.Close();
            };

            // Close時は終了
            login.Closed += (s, args) =>
            {
                if (!isLogIn) Shutdown();
            };
            // ログイン画面表示
            Application.Current.MainWindow = login;
            login.Show();
        }

        /// <summary>
        /// MainWindow表示
        /// </summary>
        private void ShowMainWindow()
        {
            var mainWindow = new MainWindow();
            var isLogOut = false;

            // ログアウトの時はLogin画面を表示
            mainWindow.LogoutRequested += (s, args) =>
            {
                isLogOut = true;
                ShowLogin();
                mainWindow.Close();
            };

            // Closeの時は終了
            mainWindow.Closed += (s, args) =>
            {
                if(!isLogOut) Shutdown();
            };

            // MainWindow表示
            Application.Current.MainWindow = mainWindow;
            mainWindow.Show();
        }

        /// <summary>
        /// Exit処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Application_Exit(object sender, ExitEventArgs e)
        {
            if (_mutex != null)
            {
                // Mutexの解放
                _mutex.ReleaseMutex();
                _mutex?.Close();
            }
        }

        /// <summary>
        /// 例外処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MyLibrary.MyClass.MyLogger.SetLogger(((Exception)e.ExceptionObject).ToString(), MyLibrary.MyEnum.LoggerType.Error);
        }
    }
}
