using MyTemplate.Class;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MyTemplate.UserControls
{
    /// <summary>
    /// MyEditTextBox.xaml の相互作用ロジック
    /// </summary>
    public partial class MyEditTextBox : UserControl
    {
        private static readonly MyStandardCheck _check = new MyStandardCheck();
        private EditParameter _paraeter = new EditParameter();

        /// <summary>
        /// 編集用入力値パラメータ
        /// </summary>
        public class EditParameter
        {
            public byte DataType { get; set; } = 0;
            public int LengthMin { get; set; } = 0;
            public int LengthMax { get; set; } = 20;
            public bool IsPassword { get; set; } = false;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MyEditTextBox()
        {
            InitializeComponent();

            this.my_TextBox.BorderThickness = new Thickness(0, 0, 0, 1);
            this.my_TextBox.GotFocus += MyEditTextBox_Focused;
            this.my_TextBox.LostFocus += MyEditTextBox_LostFocus;
        }

        /// <summary>
        /// パラメータのgetter/setter
        /// </summary>
        public EditParameter Parameter
        {
            get => _paraeter;
            set
            {
                _paraeter = value;
                my_TextBox.IsPassword = _paraeter.IsPassword;
            }
        }

        /// <summary>
        /// 入力値の取得
        /// </summary>
        public string Value => my_TextBox.Value;

        /// <summary>
        /// エラー状態の取得
        /// </summary>
        public bool IsError { get => my_TextBox.ToolTip != null; }

        /// <summary>
        /// my_TextBoxのフォーカスイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyEditTextBox_Focused(object sender, RoutedEventArgs e)
        {
            my_TextBox.Focus();
            my_TextBox.Background = Brushes.LightYellow;

            // IMEモード設定
            MyLibrary.MyClass.MyImeHelper.SetInputMode(Window.GetWindow(this), my_TextBox, Parameter.DataType);
        }

        /// <summary>
        /// my_TextBoxのフォーカスが外れたときのイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyEditTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // 入力値チェック
            var result = _check.GetResult(my_TextBox.Value, Parameter.DataType, Parameter.LengthMin, Parameter.LengthMax, 0, 0, string.Empty);

            // チェック結果に応じて背景色とツールチップを設定
            if (result.Item1 != 0)
            {
                my_TextBox.Background = Brushes.LightPink;
                my_TextBox.ToolTip = result.Item2;
            }
            else
            {
                my_TextBox.Background = Brushes.Transparent;
                my_TextBox.ToolTip = null;
            }
        }
    }
}
