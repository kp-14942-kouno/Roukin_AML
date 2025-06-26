using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace MyTemplate.Forms
{
    /// <summary>
    /// DataSearch.xaml の相互作用ロジック
    /// </summary>
    public partial class DataSearch : Window, IDisposable
    {
        private MyDbData _db;
        private TableSetting _tableSetting { get; set; } = new ImportClass.TableSetting();
        private List<TableFields> _tableFields { get; set; } = [];
        private TableFields _selectField = null;
        int _tableId;

        string[] accdb = { "accdb", "accdb2019" };
        string[] like = { "like", "not like" };

        int _maxRecord;
        int _maxPage = 1;
        int _pageNum = 1;
        int _limitRecord = 0;

        #region◆イベント
        /// <summary>
        /// 解放
        /// </summary>
        public void Dispose()
        {
            _db?.Dispose();
            _db = null;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public DataSearch(int tableId)
        {
            InitializeComponent();

            // WindowのClosedイベント時にDisposeを呼出す
            this.Closed += (s, e) => Dispose();

            // Windowの表示位置
            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            // ページ移動の初期は非表示
            gd_Page.Visibility = Visibility.Collapsed;

            // 検索上限値
            _limitRecord = int.Parse(MyUtilityModules.AppSetting("systeSetting", "searchLimit"));
            // テーブルID
            _tableId = tableId;
        }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 初期設定呼び出し
                Initialize();
                // DB接続
                _db = new MyDbData(_tableSetting.schema);
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                this.Close();
            }
        }

        /// <summary>
        /// 検索項目の変更時イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lb_Columns_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lb_Columns.SelectedItem != null)
            {
                _selectField = (TableFields)lb_Columns.SelectedItem;
            }
            // 比較演算子
            Comparison();
        }

        /// <summary>
        /// 検索項目追加ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Add_Click(object sender, RoutedEventArgs e)
        {
            // 未選択があれば中断
            if (_selectField == null) return;
            if (cmb_Comparison.SelectedItem == null) return;
            if (cmb_Conjunction.SelectedItem == null) return;

            var item = new WhereItem();
            var value = tx_Value.Text;

            if (like.Contains(cmb_Comparison.SelectedValue.ToString().ToLower()))
            {
                value = _db.Provider.ToLower() switch
                {
                    "accdb" => $"*{value}*",
                    "accdb2019" => $"*{value}*",
                    "sqlserver" => $"%{value}%",
                    "sqlite" => $"%{value}%",
                    "mysql" => $"%{value}%",
                    _ => $"*{value}*"
                };
            }

            // where用item作成
            item = new WhereItem();
            item.field_name = _selectField.field_name;
            item.field_caption = _selectField.field_caption;
            item.conjunctino_type = cmb_Conjunction.SelectedValue.ToString();

            // bool型はtrueとfalse
            if (_selectField.field_type == 5)
            {
                item.comparison_type = "=";
                item.value = cmb_Comparison.SelectedValue.ToString();
            }
            // bool型以外
            else
            {
                item.comparison_type = cmb_Comparison.SelectedValue.ToString();
                item.value = value;
            }
            // 条件を追加
            dg_Where.Items.Add(item);
        }

        /// <summary>
        /// 条件文全体削除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Delete_Click(object sender, RoutedEventArgs e)
        {
            dg_Where.Items.Clear();
        }

        /// <summary>
        /// 検索ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Search_Click(object sender, RoutedEventArgs e)
        {
            _pageNum = 1;
            Search();
        }

        /// <summary>
        /// 閉じるボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// 検索値のTetBoxのFocusイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tx_Value_GotFocus(object sender, RoutedEventArgs e)
        {
            // 検索項目が非選択なら中断
            if (_selectField == null) return;

            InputScope inputScope = new InputScope();
            InputScopeName inputScopeName = new InputScopeName();

            inputScopeName.NameValue = _selectField.data_type switch
            {
                0 or 1 or 11 => InputScopeNameValue.Hiragana,                           // 全角
                2 or 3 => InputScopeNameValue.AlphanumericFullWidth,                    // 全角英数
                4 => InputScopeNameValue.KatakanaFullWidth,                             // 全角カナ
                5 or 6 or 7 or 9 or 10 => InputScopeNameValue.AlphanumericHalfWidth,    // 半角
                8 => InputScopeNameValue.KatakanaHalfWidth,                             // 半角カナ
                _ => InputScopeNameValue.Default
            };
            inputScope.Names.Add(inputScopeName);
            tx_Value.InputScope = inputScope;
        }

        /// <summary>
        /// 条件文削除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_DeleteWhere_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            var selectItem = button.DataContext;
            // 選択されたWhere文を削除
            if (selectItem != null) dg_Where.Items.Remove(selectItem);
        }

        /// <summary>
        /// 削除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (MyMessageBox.Show("選択レコードを<\\ #FF0000 削除\\>します。よろしいですか？", "確認", MyEnum.MessageBoxButtons.OkCancel,
                                                                        MyEnum.MessageBoxIcon.Info, this) == MyEnum.MessageBoxResult.Ok)
            {
                SearchDataDelete(sender);
                Search();
            }
        }

        /// <summary>
        /// 編集ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            DataRowView row = button.DataContext as DataRowView;

            var edit = new DataEdit(_db, _tableSetting, _tableFields, row);
            edit.Owner = this;
            edit.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            edit.ShowDialog();

            Search();
        }

        /// <summary>
        /// ページ移動（前）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Previouse_Click(object sender, RoutedEventArgs e)
        {
            if(_pageNum == 1) return;
            _pageNum--;
            Search();
        }

        /// <summary>
        /// ページ移動（次）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Next_Click(object sender, RoutedEventArgs e)
        {
            if (_pageNum == _maxPage) return;
            _pageNum++;
            Search();
        }

        /// <summary>
        /// ページのTextBoxのKeyDownイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tx_Page_KeyDown(object sender, KeyEventArgs e)
        {
            // Enterキー以外は中断
            if (!(e.Key == Key.Enter)) return;

            // ページ数を取得
            var page = tx_Page.Text.Trim();
            // ページ数が空白なら中断
            if (!int.TryParse(page, out int pageNum)) return;
            // ページ数が最大ページ数を超えたら中断
            if (pageNum > _maxPage) return;
            // ページ数が1未満なら中断
            if (pageNum < 1) return;

            _pageNum = pageNum;
            Search();
        }

        #endregion

        #region◆メソッド
        /// <summary>
        /// 検索
        /// </summary>
        private void Search()
        {
            try
            {
                // 検索用sql文作成
                var sql = CreateSelectSql();

                // 検索条件があれば検索実行
                if (!string.IsNullOrEmpty(sql.Item1))
                {
                    // 検索実行
                    using var dlg = new MyLibrary.MyLoading.Dialog(this);
                    var search = new DataSearchThread(_db, sql.Item1, sql.Item2);
                    dlg.ThreadClass(search);
                    dlg.ShowDialog();

                    // エラーが発生時は閉じる
                    if (search.Result == MyEnum.MyResult.Error) Close();

                    // 検索結果をDataGridに表示
                    dg_Search.ItemsSource = search.SearchResult.DefaultView;
                    tb_Count.Text = search.SearchResult.Rows.Count.ToString();

                    // スクロールバーを一番上に移動
                    dg_Search.UpdateLayout();
                    if(VisualTreeHelper.GetChild(dg_Search,0) is Decorator border && border.Child is ScrollViewer scrollViewr)
                    {
                        scrollViewr.ScrollToVerticalOffset(0);
                    }
                }

                // Access以外はページ処理あり
                if (!accdb.Contains(_db.Provider.ToLower()))
                {
                    // 最大ページ数を取得
                    _maxPage = (_maxRecord + _limitRecord - 1) / _limitRecord;

                    if (_maxPage > 1)
                    {
                        // 複数ページ数ならページ移動を表示
                        gd_Page.Visibility = Visibility.Visible;
                        
                        tx_Page.Text = _pageNum.ToString();
                        tb_PageCount.Text = $"　／　{_maxPage}　｜　Record： {_maxRecord}";
                    }
                    else
                    {
                        gd_Page.Visibility = Visibility.Collapsed;
                    }
                }
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                Close();
            }
        }

        /// <summary>
        /// 検索用sql文作成（パラメータークエリ）
        /// </summary>
        /// <returns></returns>
        private (string, Dictionary<string, object>) CreateSelectSql()
        {
            var where = string.Empty;
            var param = new Dictionary<string, object>();
            var colNum = 0;

            // 条件文とパラメーターを作成
            foreach (WhereItem item in dg_Where.Items)
            {
                colNum++;
                where += item.conjunctino_type + " ";

                if (accdb.Contains(_db.Provider.ToLower()))
                {
                    where += $"{item.field_name} {item.comparison_type} ? ";
                }
                else
                {
                    where += $"{item.field_name} {item.comparison_type} @col_{colNum} ";
                }
                param.Add($"col_{colNum}", item.value);
            }
            // 先端文字を削除
            where = where.TrimStart("and".ToCharArray());
            where = where.TrimStart("or".ToCharArray());

            // 検索全件数を取得
            _maxRecord = int.Parse(_db.ExecuteScalar($"select count(*) from {_tableSetting.table_name} where {where};", param).ToString());

            string sql = string.Empty;

            // sql文作成
            if (accdb.Contains(_db.Provider.ToLower()))
            {
                // access... ページングなし
                sql = $@"select * from {_tableSetting.table_name} where {where}";
            }
            else if(_db.Provider.ToLower() == "sqlserver")
            {
                // sqlserver ... ページング「order by 1」は主キーが無い場合でも対応できる様にダミーで指定　※order byは必須
                sql = $@"select * from {_tableSetting.table_name} where {where} order by 1 
                        offset {_limitRecord * (_pageNum - 1)} rows fetch next {_limitRecord} rows only;";
            }
            else
            {
                // sqlite, mysql ... ページング
                sql = $@"select * from {_tableSetting.table_name} where {where} 
                        limit {_limitRecord} offset {_limitRecord * (_pageNum - 1)};";
            }
            return (sql, param);
        }

        /// <summary>
        /// Where用Itemパラメーター
        /// </summary>
        private class WhereItem
        {
            public string field_name { get; set; }
            public string field_caption { get; set; }
            public string conjunctino_type { get; set; }
            public string comparison_type { get; set; }
            public string value { get; set; }
        }

        /// <summary>
        /// 初期設定
        /// </summary>
        /// <param name="tableId"></param>
        private void Initialize()
        {
            // 接続詞
            Conjunction();
            // 比較演算子
            Comparison();

            lb_Columns.DisplayMemberPath = "field_caption";
            lb_Columns.SelectedValuePath = "field_name";

            // テーブル設定
            using (MyDbData db = new MyDbData("setting"))
            {
                MyPropertyModules.GetCreateProperties(db, typeof(ImportClass.TableSetting), _tableSetting, "t_table_setting", "table_id", _tableId);
                MyPropertyModules.GetCreateProperties(db, typeof(ImportClass.TableFields), _tableFields, $"select * from t_table_fields where table_id={_tableId} order by num;");
            }

            // 主キーが存在すれば「削除ボタン」を追加
            if (_tableFields.Count(x => x.primary_flg == 1) != 0)
            {
                DataGridTemplateColumn buttonColumn = new DataGridTemplateColumn();
                buttonColumn.Header = "";

                FrameworkElementFactory buttonFactory = new FrameworkElementFactory(typeof(Button));
                buttonFactory.SetValue(Button.ContentProperty, "削除");
                buttonFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler(Delete_Click));

                DataTemplate cellTemplate = new DataTemplate();
                cellTemplate.VisualTree = buttonFactory;
                buttonColumn.CellTemplate = cellTemplate;
                buttonColumn.Width = 64;

                dg_Search.Columns.Add(buttonColumn);
            }

            // 編集ボタンを追加
            {
                DataGridTemplateColumn buttonColumn = new DataGridTemplateColumn();
                buttonColumn.Header = "";

                FrameworkElementFactory buttonFactory = new FrameworkElementFactory(typeof(Button));
                buttonFactory.SetValue(Button.ContentProperty, "編集");
                buttonFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler(Edit_Click));

                DataTemplate cellTemplate = new DataTemplate();
                cellTemplate.VisualTree = buttonFactory;
                buttonColumn.CellTemplate = cellTemplate;
                buttonColumn.Width = 64;

                dg_Search.Columns.Add(buttonColumn);
            }

            // テーブル項目を追加
            foreach (var field in _tableFields)
            {
                DataGridTextColumn column = new DataGridTextColumn();
                column.Header = field.field_caption;
                column.Width = DataGridLength.Auto;

                Binding binding = new Binding(field.field_name);

                // 日付項目のフォーマット変更
                if (field.field_type == 6)
                {
                    binding.StringFormat = "yyyy/MM/dd HH:mm:ss";
                }

                column.Binding = binding;
                dg_Search.Columns.Add(column);

                lb_Columns.Items.Add(field);
            }
        }

        /// <summary>
        /// 削除処理
        /// </summary>
        /// <param name="sender"></param>
        private void SearchDataDelete(object sender)
        {
            Button button = sender as Button;
            DataRowView row = button.DataContext as DataRowView;

            try
            {
                string sql = string.Empty;
                string where = string.Empty;
                var param = new Dictionary<string, object>();

                // 主キーでWhere分を作成
                foreach (var tableField in _tableFields.Where(x => x.primary_flg == 1))
                {
                    if (accdb.Contains(_db.Provider.ToLower()))
                    {
                        where += $"{tableField.field_name}= ? and ";
                    }
                    else
                    {
                        where += $"{tableField.field_name}=@{tableField.field_name} and ";
                    }
                    param.Add(tableField.field_name, row[tableField.field_name].ToString());
                }
                // 終端文字（and）を削除
                where = where.TrimEnd(" and ".ToCharArray());

                // 削除実行
                sql = $"delete from {_tableSetting.table_name} where {where};";
                _db.ExecuteNonQuery(sql, param);
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
        }

        /// <summary>
        /// 接続詞
        /// </summary>
        private void Conjunction()
        {
            DataTable table = new DataTable();

            string[] captionAry = new string[] { "and", "or" };
            string[] conjunctionAry = new string[] { "and", "or" };

            cmb_Conjunction.DisplayMemberPath = "caption";
            cmb_Conjunction.SelectedValuePath = "conjunction";

            cmb_Conjunction.ItemsSource = ControlItems("conjunction", conjunctionAry, captionAry).DefaultView;
            cmb_Conjunction.SelectedValue = "and";
        }

        /// <summary>
        /// 比較演算子
        /// </summary>
        private void Comparison()
        {
            cmb_Comparison.DisplayMemberPath = "caption";
            cmb_Comparison.SelectedValuePath = "comparison";

            string[] captionAry;
            string[] comarisonAry;

            string initialVlue = string.Empty;

            if (_selectField == null) return;

            switch (_selectField.field_type.ToString())
            {
                case "5":
                    captionAry = new string[] { "true", "false" };
                    comarisonAry = new string[] { "true", "false" };
                    initialVlue = "true";
                    break;

                case "6":
                    captionAry = new string[] { "=", "<>", "<", "<=", ">", ">=" };
                    comarisonAry = new string[] { "=", "<>", "<", "<=", ">", ">=" };
                    initialVlue = "=";
                    break;

                default:
                    captionAry = new string[] { "=", "<>", "<", "<=", ">", ">=", "like", "not like" };
                    comarisonAry = new string[] { "=", "<>", "<", "<=", ">", ">=", "like", "not like" };
                    initialVlue = "=";
                    break;
            }

            cmb_Comparison.ItemsSource = ControlItems("comparison", comarisonAry, captionAry).DefaultView;
            cmb_Comparison.SelectedValue = initialVlue;
        }

        /// <summary>
        /// Control用のItemをDataTable化
        /// </summary>
        /// <param name="valueMenber"></param>
        /// <param name="valueAry"></param>
        /// <param name="captionAry"></param>
        /// <returns></returns>
        private DataTable ControlItems(string valueMenber, string[] valueAry, string[] captionAry)
        {
            DataTable table = new DataTable();

            table.Columns.Add(valueMenber, typeof(string));
            table.Columns.Add("caption", typeof(string));

            for (int index = 0; index < valueAry.Length; index++)
            {
                DataRow row = table.NewRow();
                row[valueMenber] = valueAry[index].ToString();
                row["caption"] = captionAry[index].ToString();
                table.Rows.Add(row);
            }
            return table;
        }

    }
    #endregion

    /// <summary>
    /// 検索処理スレッド
    /// </summary>
    internal partial class DataSearchThread : MyLibrary.MyLoading.Thread
    {
        MyDbData _db;
        string _sql;
        Dictionary<string, object> _param;

        public DataTable SearchResult { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="db"></param>
        /// <param name="sql"></param>
        /// <param name="parameters"></param>
        public DataSearchThread(MyDbData db, string sql, Dictionary<string, object> parameters)
        {
            _db = db;
            _sql = sql;
            _param = parameters;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            ProgressBarType = MyEnum.MyProgressBarType.None;
            ProcessName = "検索中...";

            try
            {
                SearchResult = new DataTable();
                SearchResult = _db.ExecuteQuery(_sql, _param);

                Result = MyEnum.MyResult.Ok;
            }
            catch (Exception ex)
            {
                Result = MyEnum.MyResult.Error;
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
            Completed = true;
            return 0;
        }
    }
}
