using DocumentFormat.OpenXml.Wordprocessing;
using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MyTemplate.ImportClass;
using Org.BouncyCastle.Bcpg.OpenPgp;
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
    /// DataEdit.xaml の相互作用ロジック
    /// </summary>
    public partial class DataEdit : Window
    {
        // チェッククラス
        Class.MyStandardCheck check = new Class.MyStandardCheck();

        // 各設定
        MyDbData _db;
        ImportClass.TableSetting _tableSetting;
        List<ImportClass.TableFields> _tableFieds;
        DataRowView _rowView;

        // accessのプロバイダ名
        string[] accdb = {"accdb","accdb2019"};

        #region◆イベント
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public DataEdit(MyDbData db, ImportClass.TableSetting tableSetting, List<ImportClass.TableFields> tableFieds, DataRowView rowView)
        {
            InitializeComponent();

            _db = db;
            _tableSetting = tableSetting;
            _tableFieds = tableFieds;
            _rowView = rowView;
        }

        /// <summary>
        /// // データ保存ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Save_Click(object sender, RoutedEventArgs e)
        {
            // 編集項目のTextBoxにエラーがあればメッセージを表示
            foreach (var field in _tableFieds.Where(x => x.primary_flg ==0))
            {
                // 編集項目のTextBoxを取得
                var textBox = EditTextBox(field.field_name);
                if (textBox == null) continue;
                // ToolTipがnullでなければエラー
                if (textBox.ToolTip != null)
                {
                    MyMessageBox.Show($"<\\ #ff0000 エラー \\>が存在します。修正後に再度実行してください。", "エラー", MyEnum.MessageBoxButtons.Ok, MyEnum.MessageBoxIcon.Error);
                    return;
                }
            }

            // 確認メッセージを表示
            if (MyMessageBox.Show("編集結果を保存しますか？", "確認", MyEnum.MessageBoxButtons.OkCancel, MyEnum.MessageBoxIcon.Info) != MyEnum.MessageBoxResult.Ok) return;

            // 更新処理のSQL分とパラメータを取得
            var query = UpdateQuery();

            using var dlg = new MyLibrary.MyLoading.Dialog(this);
            var editSave = new EditSave(_db, query.Item1, query.Item2);
            dlg.ThreadClass(editSave);
            dlg.ShowDialog();

            if(editSave.Result == MyEnum.MyResult.Ok)
            {
                // 更新成功メッセージを表示
                MyMessageBox.Show("更新が完了しました。", "完了", MyEnum.MessageBoxButtons.Ok, MyEnum.MessageBoxIcon.Info);
                // 編集画面を閉じる
                Close();
            }
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
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Initialize();
        }

        /// <summary>
        /// TextBoxフォーカス取得イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox.ToolTip == null)
            {
                // フォーカスが当たったとき背景色をWiteに変更
                textBox.Background = Brushes.White;
            }
        }

        /// <summary>
        /// TextBoxフォーカス喪失イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // 背景色をLightGrayに変更
            var textBox = sender as TextBox;
            textBox.Background = Brushes.LightGray;

            // 対象のTabFieldsを取得
            var field = _tableFieds.Where(x => x.field_name == textBox.Name).FirstOrDefault();
            // チェック
            var result = check.GetResult(textBox.Text, field.data_type, field.data_length, 1, field.fix_flg, field.regpattern);
            // エラーがあればメッセージをToolTipに表示
            if (result.Item1 != 0)
            {
                textBox.ToolTip = result.Item2;
                textBox.Background = Brushes.LightPink;
            }
            else
            {
                textBox.ToolTip = null;
            }
        }

        #endregion

        #region◆メソッド
        /// <summary>
        /// 編集画面初期設定
        /// </summary>
        private void Initialize()
        {
            // 主キーがあれば保存ボタンを表示
            bt_Save.Visibility = _tableFieds.Where(x => x.primary_flg == 1).Count() > 0 ? Visibility.Visible : Visibility.Collapsed;

            // 編集項目作成
            foreach (var field in _tableFieds)
            {
                // 編集項目の表示位置を設定
                var panel = new Grid
                {
                    Margin = new Thickness(0, 0, 0, 1)
                };
                panel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
                panel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                // 表題作成
                var label = new Label
                {
                    Content = field.field_caption,
                    FontFamily = new System.Windows.Media.FontFamily("MS UI Gothic"),
                    FontSize = 10,
                    Background = Brushes.DodgerBlue,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Width = 100,
                    Height = 32,
                    Margin = new Thickness(0, 0, 8, 0)
                };
                Grid.SetColumn(label, 0);
                panel.Children.Add(label);

                // 編集項目作成
                var textBox = new TextBox
                {
                    Name = field.field_name,
                    FontFamily = new System.Windows.Media.FontFamily("MS UI Gothic"),
                    FontSize = 12,
                    Background = Brushes.White,
                    Foreground = Brushes.Black,
                    VerticalAlignment = VerticalAlignment.Top,
                    Text = _rowView[field.field_name].ToString(),
                    Height = 32
                };

                textBox.GotFocus += TextBox_GotFocus;
                textBox.LostFocus += TextBox_LostFocus;
                TextBox_LostFocus(textBox, new RoutedEventArgs());

                // 主キー項目は背景色を変更し、編集不可にする
                if (field.primary_flg == 1)
                {
                    textBox.Background = Brushes.PeachPuff;
                    textBox.IsEnabled = false;
                    textBox.IsReadOnly = true;
                }

                Grid.SetColumn(textBox, 1);
                panel.Children.Add(textBox);
                sp_Edit.Children.Add(panel);
            }
        }

        /// <summary>
        /// 編集データ保存用SQL文とパラメータを作成
        /// </summary>
        /// <returns></returns>
        private (string ,Dictionary<string, object>) UpdateQuery()
        {
            string field = string.Empty;
            string where = string.Empty;
            var param = new Dictionary<string, object>();

            // accessのsqlとparameter作成
            if (accdb.Contains(_db.Provider))
            {
                foreach (var tableField in _tableFieds.FindAll(x => x.primary_flg == 0))
                {
                    field += $"{tableField.field_name} = ?,";
                    param.Add(tableField.field_name, EditTextBoxObject(tableField));
                }
                foreach (var tableField in _tableFieds.FindAll(x => x.primary_flg == 1))
                {
                    where += $"{tableField.field_name} = ?,";
                    param.Add(tableField.field_name, EditTextBoxObject(tableField));
                }
            }
            // access以外のsqlとparameter作成
            else
            {
                foreach (var tableField in _tableFieds.FindAll(x => x.primary_flg == 0))
                {
                    field += $"{tableField.field_name} = @{tableField.field_name},";
                    param.Add(tableField.field_name, EditTextBoxObject(tableField));
                }
                foreach (var tableField in _tableFieds.FindAll(x => x.primary_flg == 1))
                {
                    where += $"{tableField.field_name} = @{tableField.field_name},";
                    param.Add(tableField.field_name, EditTextBoxObject(tableField));
                }
            }
            field = field.TrimEnd(',');
            where = where.TrimEnd(',');
            return ($"UPDATE {_tableSetting.table_name} SET {field} WHERE {where}", param);
        }

        /// <summary>
        /// 編集TextBoxの値を取得
        /// </summary>
        /// <param name="tableFields"></param>
        /// <returns></returns>
        private object EditTextBoxObject(TableFields tableFields)
        {
            var textBox = EditTextBox(tableFields.field_name);

            if(textBox.Text == string.Empty && tableFields.field_type == 6)
            {
                return DBNull.Value;
            }

            return textBox.Text.ToString();
        }

        /// <summary>
        /// 編集TextBox取得
        /// </summary>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        private TextBox EditTextBox(string fieldName)
        {
            // 編集用StackPanelの子要素を取得
            foreach (UIElement element in sp_Edit.Children)
            {
                // Grid要素を取得
                if (!(element is Grid panel)) continue;
                // 子要素を取得
                foreach (UIElement child in panel.Children)
                {
                    // TextBox要素を取得
                    if (!(child is TextBox textBox)) continue;

                    if (textBox.Name == fieldName)
                    {
                        return textBox;
                    }
                }
            }
            return null;
        }
    }
    #endregion

    /// <summary>
    /// 編集データ保存クラス
    /// </summary>
    internal class EditSave : MyLibrary.MyLoading.Thread
    {
        MyDbData _db;
        string _sql;
        Dictionary<string, object> _param;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableSetting"></param>
        /// <param name="tableFields"></param>
        public EditSave(MyDbData db, string sql, Dictionary<string, object> param)
        {
            _db = db;
            _sql = sql;
            _param = param;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            try
            {
                // トランザクション開始
                _db.BindTransaction();
                // SQL実行
                _db.ExecuteNonQuery(_sql, _param);
                // コミット
                _db.CommitTransaction();

                Result = MyEnum.MyResult.Ok;
            }
            catch(Exception ex)
            {
                // ロールバック
                _db.RollbackTransaction();
                Result = MyEnum.MyResult.Error;
                // エラーログ出力
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
            Completed = true;
            return 0;
        }
    }
}
