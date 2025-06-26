using System;
using System.Collections;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MyTemplate.UserControls
{
    /// <summary>
    /// UserInfoListView コントロールは、2つのListView（固定列と内容列）を持ち
    /// 右側のListViewの垂直スクロールに合わせて左側のListViewも同期スクロール
    /// データバインディングやヘッダーのカスタマイズも可能
    /// </summary>
    public partial class UserInfoListView : UserControl
    {
        // 右側ListViewのScrollViewer
        private ScrollViewer rightScroll;
        // 左側ListViewのScrollViewer
        private ScrollViewer leftScroll;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public UserInfoListView()
        {
            InitializeComponent();
            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
        }

        /// <summary>
        /// コントロールのLoaded時にScrollViewerを取得し
        /// 右側ListViewのスクロールイベントを登録
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            rightScroll = GetScrollViewer(RightListView);
            leftScroll = GetScrollViewer(LeftListView);

            if (rightScroll != null && leftScroll != null)
            {
                // 右側ListViewのスクロールに合わせて左側もスクロール
                rightScroll.ScrollChanged += RightScroll_ScrollChanged;
            }
        }

        /// <summary>
        /// コントロールのUnloaded時にイベントハンドラを解除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            if (rightScroll != null)
                rightScroll.ScrollChanged -= RightScroll_ScrollChanged;
        }

        /// <summary>
        /// 右側ListViewの垂直スクロールに合わせて左側ListViewも同期してスクロール
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ev"></param>
        private void RightScroll_ScrollChanged(object sender, ScrollChangedEventArgs ev)
        {
            if (ev.VerticalChange != 0)
                leftScroll?.ScrollToVerticalOffset(ev.VerticalOffset);
        }

        /// <summary>
        /// 指定したDependencyObjectからScrollViewerを再帰的に取得
        /// </summary>
        /// <param name="obj">検索対象のオブジェクト</param>
        /// <returns>ScrollViewer（見つからない場合はnull）</returns>
        /// <returns></returns>
        private ScrollViewer GetScrollViewer(DependencyObject obj)
        {
            if (obj is ScrollViewer sv) return sv;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                var child = VisualTreeHelper.GetChild(obj, i);
                var result = GetScrollViewer(child);
                if (result != null) return result;
            }
            return null;
        }

        /// <summary>
        /// ListViewにバインドするデータソース
        /// </summary>
        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(nameof(ItemsSource), typeof(IEnumerable), typeof(UserInfoListView), new PropertyMetadata(null));

        /// <summary>
        /// ListViewのデータソースを取得または設定
        /// </summary>
        public IEnumerable ItemsSource
        {
            get => (IEnumerable)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }

        /// <summary>
        /// 左側（固定列）のヘッダー文字列
        /// </summary>
        public static readonly DependencyProperty FixedColumnHeaderProperty =
            DependencyProperty.Register(nameof(FixedColumnHeader), typeof(string), typeof(UserInfoListView), new PropertyMetadata(""));

        /// <summary>
        /// 左側（固定列）のヘッダーを取得または設定
        /// </summary>
        public string FixedColumnHeader
        {
            get => (string)GetValue(FixedColumnHeaderProperty);
            set => SetValue(FixedColumnHeaderProperty, value);
        }

        /// <summary>
        /// 右側（内容列）のヘッダー文字列
        /// </summary>
        public static readonly DependencyProperty ContentColumnHeaderProperty =
            DependencyProperty.Register(nameof(ContentColumnHeader), typeof(string), typeof(UserInfoListView), new PropertyMetadata(""));

        /// <summary>
        /// 右側（内容列）のヘッダーを取得または設定
        /// </summary>
        public string ContentColumnHeader
        {
            get => (string)GetValue(ContentColumnHeaderProperty);
            set => SetValue(ContentColumnHeaderProperty, value);
        }
    }
}
