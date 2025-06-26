using System.Windows;
using System.Windows.Controls;

namespace MyTemplate.Class.MyCustomControlClass
{

    #region MyButton
    /// <summary>
    /// カスタムボタンコントロールのクラス
    /// 角にRadiusを持たせるためのプロパティを追加
    /// </summary>
    public class MyButton : Button
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        static MyButton()
        {
            // デフォルトのスタイルを設定
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MyButton), new FrameworkPropertyMetadata(typeof(MyButton)));
        }

        /// <summary>
        /// Radiusを持たせるためのプロパティ
        /// </summary>
        public CornerRadius CornerRadius
        {
            get => (CornerRadius)GetValue(CornerRadiusProperty);
            set => SetValue(CornerRadiusProperty, value);
        }

        /// <summary>
        /// Radiusプロパティの依存関係プロパティ
        /// </summary>
        public static readonly DependencyProperty CornerRadiusProperty =
            DependencyProperty.Register(nameof(CornerRadius), typeof(CornerRadius), typeof(MyButton), new PropertyMetadata(new CornerRadius(1)));
    }
    #endregion
}
