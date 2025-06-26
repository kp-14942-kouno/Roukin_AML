using MyLibrary;
using MyLibrary.MyClass;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace MyTemplate.Forms
{
    /// <summary>
    /// TableList.xaml の相互作用ロジック
    /// </summary>
    public partial class TableList : Window
    {
        #region◆イベント
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public TableList()
        {
            InitializeComponent();

            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
        }

        /// <summary>
        /// Loadedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // テーブル一覧表示を呼出し（エラー時はWindowを閉じる）
            if (!LoadTableList())
            {
                this.Close();
                return;
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
        /// テーブル作成・削除ボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Button_Click(object sender, RoutedEventArgs e)
        {
            string buttonName = ((Button)sender).Content.ToString()!.Right(2);

            if (dg_TableList.ItemsSource is DataView rowView)
            {
                if (rowView.Table == null) return;

                    // 選択された行を取得
                    var selectedRows = rowView.Table.AsEnumerable().Where(x => x.Field<bool>("isSelected") == true).ToList();

                if (selectedRows.Any())
                {
                    // schemaでグループ化
                    var schemas = selectedRows.Select(x => x.Field<string>("schema")).Distinct();

                    foreach (var schema in schemas)
                    {
                        // schema毎にフィルタリング
                        var selectedRow = selectedRows.Where(x => x.Field<string>("schema") == schema);

                        if (buttonName == "作成")
                        {
                            if (!CreateTable(schema, selectedRow))
                            {
                                // テーブル作成に失敗した場合は処理を終了
                                MessageBox.Show($"テーブル作成に失敗しました。プロバイダー名：{schema}");
                                return;
                            }
                        }
                        else if (buttonName == "削除")
                        {
                            if (!DeleteTable(schema, selectedRow))
                            {
                                // テーブル削除に失敗した場合は処理を終了
                                MessageBox.Show($"テーブル削除に失敗しました。プロバイダー名：{schema}");
                                return;
                            }
                        }
                    }
                    // テーブル作成・削除完了メッセージを表示
                    MessageBox.Show($"テーブル{buttonName}完了しました。", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        /// <summary>
        /// チェックボックスを全選択にするボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Select_Click(object sender, RoutedEventArgs e)
        {
            // チェックボックスを全選択にする
            if (dg_TableList.ItemsSource is DataView rowView)
            {
                if(rowView.Table == null) return;

                rowView.Table.AsEnumerable().ToList().ForEach(x => x.SetField("isSelected", true));
            }
        }

        /// <summary>
        /// チェックボックスを全解除にするボタンクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_Clear_Click(object sender, RoutedEventArgs e)
        {
            // チェックボックスを全解除にする
            if (dg_TableList.ItemsSource is DataView rowView)
            {
                if (rowView.Table == null) return;

                rowView.Table.AsEnumerable().ToList().ForEach(x => x.SetField("isSelected", false));
            }
        }

        #endregion

        #region◆メソッド

        /// <summary>
        /// テーブル一覧を表示
        /// </summary>
        private bool LoadTableList()
        {
            try
            {
                // 設定DBを開く
                using (MyDbData db = new MyDbData("setting"))
                {
                    // テーブル設定一覧を取得
                    var dt = db.ExecuteQuery("select * from t_table_setting");

                    // isSelected列を追加
                    DataColumn column = new DataColumn("isSelected", typeof(bool))
                    {
                        DefaultValue = false
                    };
                    dt.Columns.Add(column);

                    // DataGridのItemsSourceにDataTableを設定
                    dg_TableList.ItemsSource = dt.DefaultView;
                }
                return true;
            }
            catch(Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return false;
            }
        }

        /// <summary>
        /// テーブル作成
        /// </summary>
        /// <param name="provider"></param>
        /// <param name="selected"></param>
        /// <returns></returns>
        private bool CreateTable(string schema, IEnumerable<DataRow> selected)
        {
            try
            {
                // 設定DBを開く
                using var settingDb = new MyDbData("setting");
                // テーブル作成対象DBを開く
                using var db = new MyDbData(schema);

                try
                {
                    // トランザクション開始
                    db.BindTransaction();

                    // テーブル設定分繰り返す
                    foreach (var row in selected)
                    {
                        // テーブル項目設定を取得
                        var fields = settingDb.ExecuteQuery(
                                    $"select * from t_table_fields where table_id={row["table_id"]} order by num");

                        // プロバイダー名に応じてテーブル作成メソッドを呼び出す
                        switch (db.Provider.ToLower())
                        {
                            // SqlServer
                            case "sqlserver":
                                CreateTableSqlserver(db, row, fields);
                                break;
                            // Access
                            case "accdb":
                            case "accdb2019":
                                CreateTableAccess(db, row, fields);
                                break;
                            // SQLite
                            case "sqlite":
                                CreateTableSQLite(db, row, fields);
                                break;
                            // MySQL
                            case "mysql":
                                CreateTableMySql(db, row, fields);
                                break;
                        }
                    }
                    // トランザクションコミット
                    db.CommitTransaction();
                }
                catch
                {
                    // トランザクションロールバック
                    db.RollbackTransaction();
                    throw;
                }
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return false;
            }
        }

        /// <summary>
        /// テーブル作成SQLServer
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableSetting"></param>
        /// <param name="tableFields"></param>
        private void CreateTableSqlserver(MyDbData db, DataRow tableSetting, DataTable tableFields)
        {
            // テーブルが存在する場合は削除
            if (db.IsTable(tableSetting["table_name"].ToString()))
            {
                db.ExecuteNonQuery($"drop table {tableSetting["table_name"]};");
            }

            string query = string.Empty;
            string primaryKey = string.Empty;
            List<string> indexes = new List<string>();

            // 項目設定
            foreach(DataRow row in tableFields.Rows)
            {
                query += row["field_type"].ToString() switch
                {
                    "0" => $"{row["field_name"]} nvarchar({row["data_length"]}),", // text
                    "1" => $"{row["field_name"]} tinyint,", // byte
                    "2" => $"{row["field_name"]} int,", // int
                    "3" => $"{row["field_name"]} bigint,", // long
                    "4" => $"{row["field_name"]} nvarchar(max),", // memo
                    "5" => $"{row["field_name"]} bit,", // bool
                    "6" => $"{row["field_name"]} datetime,", // date
                    "7" => $"{row["field_name"]} int identity(1,1),", // autonumber
                    _ => $"{row["field_name"]} nvarchar({row["data_length"]}),",
                };

                // PRIMARY KEY フラグが立っている場合、PRIMARY KEY を追加
                if (row["primary_flg"].ToString() == "1")
                {
                    primaryKey += $"{row["field_name"]},";
                }
                // インデックスフラグが立っている場合、インデックスを追加
                if (row["primary_flg"].ToString() == "0" && row["index_flg"].ToString() == "1")
                {
                    indexes.Add(row["field_name"].ToString());
                }
            }

            // 最後のカンマを削除
            query = query.TrimEnd(',');
            primaryKey = primaryKey.TrimEnd(',');

            if (!string.IsNullOrEmpty(primaryKey))
            {
                query += $", primary key({primaryKey})";
            }
            
            query = $"create table {tableSetting["table_name"]} ({query});";
            // テーブル作成SQLを実行
            db.ExecuteNonQuery(query);

            // インデックス作成
            foreach (var index in indexes)
            {
                db.ExecuteNonQuery($"create index {tableSetting["table_name"]}_{index} on {tableSetting["table_name"]}({index});");
            }
        }

        /// <summary>
        /// テーブル作成Access
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableSetting"></param>
        /// <param name="tableFields"></param>
        private void CreateTableAccess(MyDbData db, DataRow tableSetting, DataTable tableFields)
        {
            // テーブルが存在する場合は削除
            if (db.IsTable(tableSetting["table_name"].ToString()))
            {
                db.ExecuteNonQuery($"DROP TABLE [{tableSetting["table_name"]}];");
            }

            string query = string.Empty;
            string primaryKey = string.Empty;
            List<string> indexes = new List<string>();

            // 項目設定
            foreach (DataRow row in tableFields.Rows)
            {
                query += row["field_type"].ToString() switch
                {
                    "0" => $"[{row["field_name"]}] TEXT({row["data_length"]}),", // text
                    "1" => $"[{row["field_name"]}] BYTE,", // byte
                    "2" => $"[{row["field_name"]}] INTEGER,", // int
                    "3" => $"[{row["field_name"]}] LONG,", // long
                    "4" => $"[{row["field_name"]}] MEMO,", // memo
                    "5" => $"[{row["field_name"]}] YESNO,", // bool
                    "6" => $"[{row["field_name"]}] DATETIME,", // date
                    "7" => $"[{row["field_name"]}] COUNTER,", // autonumber
                    _ => $"[{row["field_name"]}] TEXT({row["data_length"]}),",
                };
                // PRIMARY KEY フラグが立っている場合、PRIMARY KEY を追加
                if (row["primary_flg"].ToString() == "1")
                {
                    primaryKey += $"[{row["field_name"]}],";
                }
                // インデックスフラグが立っている場合、インデックスを追加
                if (row["primary_flg"].ToString() == "0" && row["index_flg"].ToString() == "1")
                {
                    indexes.Add(row["field_name"].ToString());
                }
            }

            // 最後のカンマを削除
            query = query.TrimEnd(',');
            primaryKey = primaryKey.TrimEnd(',');

            if (!string.IsNullOrEmpty(primaryKey))
            {
                query += $", PRIMARY KEY({primaryKey})";
            }

            query = $"CREATE TABLE [{tableSetting["table_name"]}] ({query});";
            // テーブル作成SQLを実行
            db.ExecuteNonQuery(query);

            // インデックス作成
            foreach (var index in indexes)
            {
                db.ExecuteNonQuery($"CREATE INDEX [{tableSetting["table_name"]}_{index}] ON [{tableSetting["table_name"]}]([{index}]);");
            }
        }

        /// <summary>
        /// テーブル作成SQLite
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableSetting"></param>
        /// <param name="tableFields"></param>
        private void CreateTableSQLite(MyDbData db, DataRow tableSetting, DataTable tableFields)
        {
            // テーブルが存在する場合は削除
            if (db.IsTable(tableSetting["table_name"].ToString()))
            {
                db.ExecuteNonQuery($"DROP TABLE IF EXISTS {tableSetting["table_name"]};");
            }

            string query = string.Empty;
            string primaryKey = string.Empty;
            List<string> indexes = new List<string>();

            // 項目設定
            foreach (DataRow row in tableFields.Rows)
            {
                query += row["field_type"].ToString() switch
                {
                    "0" => $"{row["field_name"]} TEXT,", // text
                    "1" => $"{row["field_name"]} INTEGER,", // byte (SQLite では INTEGER 型を使用)
                    "2" => $"{row["field_name"]} INTEGER,", // int
                    "3" => $"{row["field_name"]} INTEGER,", // long
                    "4" => $"{row["field_name"]} TEXT,", // memo
                    "5" => $"{row["field_name"]} INTEGER,", // bool (SQLite では 0/1 を INTEGER で表現)
                    "6" => $"{row["field_name"]} TEXT,", // date (SQLite では ISO8601 形式の文字列を推奨)
                    "7" => $"{row["field_name"]} INTEGER PRIMARY KEY AUTOINCREMENT,", // autonumber
                    _ => $"{row["field_name"]} TEXT,", // デフォルトは TEXT
                };

                // PRIMARY KEY フラグが立っている場合、PRIMARY KEY を追加（AUTOINCREMENTの項目は除外）
                if (row["primary_flg"].ToString() == "1" && row["field_type"].ToString() !="7")
                {
                    primaryKey += $"{row["field_name"]},";
                }
                // インデックスフラグが立っている場合、インデックスを追加
                if (row["primary_flg"].ToString() == "0" && row["index_flg"].ToString() == "1")
                {
                    indexes.Add(row["field_name"].ToString());
                }
            }

            // 最後のカンマを削除
            query = query.TrimEnd(',');
            primaryKey = primaryKey.TrimEnd(',');

            // PRIMARY KEY が指定されている場合、PRIMARY KEY を追加
            if (!string.IsNullOrEmpty(primaryKey))
            {
                query += $", PRIMARY KEY({primaryKey})";
            }
            query = $"CREATE TABLE {tableSetting["table_name"]} ({query});";
            // テーブル作成SQLを実行
            db.ExecuteNonQuery(query);

            // インデックス作成
            foreach (var index in indexes)
            {
                db.ExecuteNonQuery($"CREATE INDEX {tableSetting["table_name"]}_{index} ON {tableSetting["table_name"]}({index});");
            }
        }

        /// <summary>
        /// テーブル作成MySQL
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableSetting"></param>
        /// <param name="tableFields"></param>
        private void CreateTableMySql(MyDbData db, DataRow tableSetting, DataTable tableFields)
        {
            // テーブルが存在する場合は削除
            if (db.IsTable(tableSetting["table_name"].ToString()))
            {
                db.ExecuteNonQuery($"DROP TABLE IF EXISTS `{tableSetting["table_name"]}`;");
            }

            string query = string.Empty;
            string primaryKey = string.Empty;
            List<string> indexes = new List<string>();

            // 項目設定
            foreach (DataRow row in tableFields.Rows)
            {
                query += row["field_type"].ToString() switch
                {
                    "0" => $"`{row["field_name"]}` VARCHAR({row["data_length"]}),", // text
                    "1" => $"`{row["field_name"]}` TINYINT,", // byte
                    "2" => $"`{row["field_name"]}` INT,", // int
                    "3" => $"`{row["field_name"]}` BIGINT,", // long
                    "4" => $"`{row["field_name"]}` TEXT,", // memo
                    "5" => $"`{row["field_name"]}` BOOLEAN,", // bool
                    "6" => $"`{row["field_name"]}` DATETIME,", // date
                    "7" => $"`{row["field_name"]}` INT AUTO_INCREMENT,", // autonumber
                    _ => $"`{row["field_name"]}` VARCHAR({row["data_length"]}),", // デフォルトは VARCHAR
                };
                // PRIMARY KEY フラグが立っている場合、PRIMARY KEY を追加
                if (row["primary_flg"].ToString() == "1")
                {
                    primaryKey += $"`{row["field_name"]}`,";
                }
                // インデックスフラグが立っている場合、インデックスを追加
                if (row["primary_flg"].ToString() == "0" && row["index_flg"].ToString() == "1")
                {
                    indexes.Add(row["field_name"].ToString());
                }
            }

            // 最後のカンマを削除
            query = query.TrimEnd(',');
            primaryKey = primaryKey.TrimEnd(',');

            if (!string.IsNullOrEmpty(primaryKey))
            {
                query += $", PRIMARY KEY({primaryKey})";
            }
            // テーブル作成実行
            query = $"CREATE TABLE `{tableSetting["table_name"]}` ({query}) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
            // テーブル作成SQLを実行
            db.ExecuteNonQuery(query);

            // インデックス作成
            foreach (var index in indexes)
            {
                db.ExecuteNonQuery($"CREATE INDEX `{tableSetting["table_name"]}_{index}` ON `{tableSetting["table_name"]}`(`{index}`);");
            }
        }

        /// <summary>
        /// テーブル削除
        /// </summary>
        /// <param name="provider"></param>
        /// <param name="selected"></param>
        /// <returns></returns>
        private bool DeleteTable(string? schema, IEnumerable<DataRow> selected)
        {
            try
            {
                // 設定DBを開く
                using var settingDb = new MyDbData("setting");
                // テーブル作成対象DBを開く
                using var db = new MyDbData(schema);
                try
                {
                    // トランザクション開始
                    db.BindTransaction();
                    // テーブル設定分繰り返す
                    foreach (var row in selected)
                    {
                        // テーブルが存在する場合は削除
                        if (db.IsTable(row["table_name"].ToString()))
                        {
                            db.ExecuteNonQuery($"DROP TABLE IF EXISTS {row["table_name"]};");
                        }
                    }
                    // トランザクションコミット
                    db.CommitTransaction();
                }
                catch
                {
                    // トランザクションロールバック
                    db.RollbackTransaction();
                    throw;
                }
                return true;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return false;
            }
        }

        #endregion
    }
}
