using MyLibrary;
using MyLibrary.MyClass;
using MyLibrary.MyModules;
using MySql.Data.MySqlClient;
using MySqlX.XDevAPI.Common;
using MyTemplate.Class;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Text;
using System.Windows;

namespace MyTemplate.ImportClass
{
    /// <summary>
    /// 設定プロパティ
    /// </summary>
    internal class FileImportProperties : FileLoadProperties
    {
        public TableSetting TableSetting { get; set; } = new TableSetting();    // テーブル設定
        public List<TableFields> TableFields { get; set; } = [];                // テーブルフィールド設定
        public ImportSetting ImportSetting { get; set; } = new ImportSetting(); // インポート設定
        public List<ImportFields> ImportFields { get; set; } = [];              // インポートフィールド設定

        /// <summary>
        /// 解放
        /// </summary>
        public new void Dispose()
        {
            base.Dispose();
        }
    }

    /// <summary>
    /// ファイル取込みクラス
    /// </summary>
    internal static class FileImportClass
    {
        /// <summary>
        /// ファイル取込設定取得
        /// </summary>
        /// <param name="id"></param>
        /// <param name="fileImport"></param>
        /// <returns></returns>
        public static bool GetFileImportSetting(int id, FileImportProperties import)
        {
            try
            {
                // 初期化
                import.TableSetting = new TableSetting();
                import.TableFields = [];
                import.ImportSetting = new ImportSetting();
                import.ImportFields = [];
                import.FilePaths = [];
                import.LoadData = new DataTable();
                import.ErrorLog = new ErrorLog();

                // ファイル取込設定取得
                using (var db = new MyDbData("setting"))
                {
                    MyPropertyModules.GetCreateProperties(db, typeof(ImportSetting), import.ImportSetting, "t_import_setting", "import_id", id);
                    MyPropertyModules.GetCreateProperties(db, typeof(TableSetting), import.TableSetting, "t_table_setting", "table_id", import.ImportSetting.table_id);
                    MyPropertyModules.GetCreateProperties(db, typeof(TableFields), import.TableFields, "t_table_fields", "table_id", "num", import.ImportSetting.table_id);

                    string sql = @$"select i.*, t.table_id, t.field_caption, t.field_type, t.data_type, t.data_length, t.fix_flg, t.null_flg, t.regpattern, t.primary_flg, t.index_flg 
                                    from (select * from t_import_fields where field_id = {import.ImportSetting.field_id}) as i 
                                    left join (select * from t_table_fields where table_id = {import.TableSetting.table_id}) as t 
                                    on i.field_name = t.field_name order by i.num";
                    MyPropertyModules.GetCreateProperties(db, typeof(ImportFields), import.ImportFields, sql);

                    // data_typeが0（file）
                    if (import.ImportSetting.data_type == 0)
                    {
                        // ファイル読込み設定取得
                        MyPropertyModules.GetCreateProperties(db, typeof(FileSetting), import.FileSetting, "t_file_setting", "file_id", import.ImportSetting.load_id);
                        MyPropertyModules.GetCreateProperties(db, typeof(FileColumns), import.FileColumns, "t_file_columns", "columns_id", "num", import.FileSetting.columns_id);
                        MyPropertyModules.GetCreateProperties(db, typeof(AddSetting), import.AddSettings, "t_file_add_setting", "add_id", "num", import.FileSetting.add_id);
                    }
                    // data_typeが1（table）か2（reader）
                    else
                    {
                        // 追加情報のみ取得
                        MyPropertyModules.GetCreateProperties(db, typeof(AddSetting), import.AddSettings, "t_file_add_setting", "add_id", "num", import.ImportSetting.add_id);
                    }
                }

                // 取込み先DBチェック
                using (var db = new MyDbData(import.TableSetting.schema))
                {
                    // テーブル存在チェック
                    if (!db.IsTable(import.TableSetting.table_name))
                    {
                        throw new Exception($"テーブルが存在しません。{import.TableSetting.table_name}");
                    }

                    // カラム存在チェック
                    if (import.ImportFields.Where(x => string.IsNullOrEmpty(x.table_id.ToString())).Count() > 0)
                    {
                        throw new Exception($"カラムが存在しません。");
                    }

                    // 主キーチェック
                    if (import.ImportSetting.import_type < 4)
                    {
                        var primaryKeys = import.TableFields.Where(x => x.primary_flg == 1).Select(x => x.field_name).ToArray();
                        if (primaryKeys.Length == 0 || !db.IsPrimaryKey(import.TableSetting.table_name, primaryKeys))
                        {
                            throw new Exception($"主キーが存在しません。{string.Join(",", primaryKeys)}");
                        }
                    }

                    // 取込み先テーブルのスキーマ取得
                    import.LoadData = db.GetSchema(import.TableSetting.table_name);
                }

                // JavaScript
                if (!string.IsNullOrEmpty(import.ImportSetting.js_node_key))
                {
                    import.js = new MyJScriptClass(import.ImportSetting.js_node_key);
                }
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return false;
            }
            return true;
        }

        /// <summary>
        /// ファイル取込み共通処理
        /// 各種データタイプ（File, DataTable, DbdataReader）に対応
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="dataType"></param>
        /// <param name="importAction"></param>
        /// <returns></returns>
        private static MyEnum.MyResult FileImportCore(Window window, FileImportProperties import, int dataType,
                                            Func<MyDbData, (MyEnum.MyResult, string)> importAction)
        {
            try
            {
                // 取込み初期処理（データ種別チェック・ファイル選択・追加情報処理）
                var res = FileImportInitialize(window, import, dataType);
                if (res != MyEnum.MyResult.Ok) return res;

                // 確認メッセージ
                if (MyMessageBox.Show(
                    $"<\\ {import.ImportSetting.process_name} \\>\r\n処理を実行しますか？",
                    "確認", MyEnum.MessageBoxButtons.OkCancel,
                    window: window) == MyEnum.MessageBoxResult.Cancel) return MyEnum.MyResult.Cancel;

                // DB接続
                using var db = new MyDbData(import.TableSetting.schema);

                try
                {
                    // トランザクション開始
                    db.BindTransaction();
                    // 取込み処理（importActionデリゲートで呼び出し）
                    var (result, message) = importAction(db);
                    // 取込み結果処理（コミット・ロールバック・エラーログ）
                    FileImportResult(window, db, import, result, message);
                    return result;
                }
                catch
                {
                    // ロールバック
                    db.RollbackTransaction();
                    throw;
                }
                finally
                {
                    // DB接続終了
                    db.Dispose();
                }
            }
            catch (Exception ex)
            {
                // 例外処理
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return MyEnum.MyResult.Error;
            }
        }

        /// <summary>
        /// ファイル取込み処理（File）
        /// 各種データタイプごとにFileImportCoreへ処理を移譲
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <returns></returns>
        public static MyEnum.MyResult FileImport(Window window, FileImportProperties import)
            => FileImportCore(window, import, 0, (db) => FileImportRun(window, import, db));

        /// <summary>
        /// ファイル取込み処理（DataTable）
        /// 各種データタイプごとにFileImportCoreへ処理を移譲
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static MyEnum.MyResult FileImport(Window window, FileImportProperties import, DataTable table)
            => FileImportCore(window, import, 1, (db) => FileImportRun(window, import, db, table));

        /// <summary>
        /// ファイル取込み処理（DbDataReader）
        /// 各種データタイプごとにFileImportCoreへ処理を移譲
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="reader"></param>
        /// <param name="recordCount"></param>
        /// <returns></returns>
        public static MyEnum.MyResult FileImport(Window window, FileImportProperties import, DbDataReader reader, int recordCount)
            => FileImportCore(window, import, 2, (db) => FileImportRun(window, import, db, reader, recordCount));

        /// <summary>
        /// ファイル取込み共通処理　外部DB
        /// 各種データタイプ（File, DataTable, DbdataReader）に対応
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="dataType"></param>
        /// <param name="db"></param>
        /// <param name="importAction"></param>
        /// <returns></returns>
        private static (MyEnum.MyResult, string) FileImportCore(Window window, FileImportProperties import, int dataType, MyDbData db,
                        Func<MyDbData, (MyEnum.MyResult, string)> importAction)
        {
            try
            {
                var result = FileImportInitialize(window, import, 0);
                if (result != MyEnum.MyResult.Ok) return (result, string.Empty);

                return importAction(db);
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return (MyEnum.MyResult.Error, string.Empty);
            }
        }

        /// <summary>
        /// ファイル取込み処理（File）　外部DB
        /// 各種データタイプごとにFileImportCoreへ処理を移譲
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static (MyEnum.MyResult, string) FileImport(Window window, FileImportProperties import, MyDbData db)
            => FileImportCore(window, import, 0, db, db2 => FileImportRun(window, import, db));

        /// <summary>
        /// ファイル取込み処理（DataTable） 　外部DB
        /// 各種データタイプごとにFileImportCoreへ処理を移譲
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static (MyEnum.MyResult, string) FileImport(Window window, FileImportProperties import, MyDbData db, DataTable table)
            => FileImportCore(window, import, 0, db, db2 => FileImportRun(window, import, db, table));

        /// <summary>
        /// ファイル取込み処理（DbDataReader）　外部DB
        /// 各種データタイプごとにFileImportCoreへ処理を移譲
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <param name="reader"></param>
        /// <param name="recordCount"></param>
        /// <returns></returns>
        public static (MyEnum.MyResult, string) FileImport(Window window, FileImportProperties import, MyDbData db, DbDataReader reader, int recordCount)
            => FileImportCore(window, import, 0, db, db2 => FileImportRun(window, import, db, reader, recordCount));

        /// <summary>
        /// 結果処理
        /// </summary>
        /// <param name="window"></param>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="result"></param>
        /// <param name="message"></param>
        private static void FileImportResult(Window window, MyDbData db, FileImportProperties import, MyEnum.MyResult result, string message)
        {
            // 結果で分岐
            switch (result)
            {
                // 完了
                case MyEnum.MyResult.Ok:
                    // コミット
                    db.CommitTransaction();
                    // ファイル移動処理
                    if (import.ImportSetting.data_type == 0)
                    {
                        FileLoadClass.FileMove(import);
                    }
                    // ログ出力
                    MyLogger.SetLogger($"{import.ImportSetting.process_name}\r\n{message}取込み処理完了", MyEnum.LoggerType.Info, false);
                    MyMessageBox.Show($"{message}\r\n <\\ #4169E1 取込み処理完了 \\>");
                    break;
                // 異常終了
                case MyEnum.MyResult.Error:
                    db.RollbackTransaction();
                    break;
                // 結果エラー
                case MyEnum.MyResult.Ng:
                    db.RollbackTransaction();
                    MyLogger.SetLogger($"{import.ImportSetting.process_name}：取込み処理結果エラー中止", MyEnum.LoggerType.Info, false);
                    var viewer = new MyDataViewer(window, import.ErrorLog.DefaultView, title: import.ImportSetting.process_name, columnNames: import.ErrorLog.ErrorLogField(), columnHeaders: import.ErrorLog.ErrorLogCaption());
                    viewer.ShowDialog();
                    break;
            }
        }

        /// <summary>
        /// ファイル取込み初期処理
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static MyEnum.MyResult FileImportInitialize(Window window, FileImportProperties import, int dataType)
        {
            // データ種別チェック（設定と呼び出しメソッドが一致しているか）
            if (import.ImportSetting.data_type != dataType) throw new Exception("取込データ種別が不正です。");

            // データ種別がファイルならファイル選択
            if (import.ImportSetting.data_type == 0)
            {
                // ファイル選択
                var selectResult = FileLoadClass.FileSelect(window, import);
                if (selectResult != MyEnum.MyResult.Ok) return selectResult;
            }

            // 追加情報処理
            var addResult = Add.Run(import, window);
            if (addResult != MyEnum.MyResult.Ok) return addResult;

            return MyEnum.MyResult.Ok;
        }

        /// <summary>
        /// ファイル取込み実行 File
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        private static (MyEnum.MyResult, string) FileImportRun(Window window, FileImportProperties import, MyDbData db)
        {
            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(window))
            {
                MyLibrary.MyLoading.Thread thread;

                // Sqlserverとそれ以外で呼出しを変える
                thread = db.Provider.ToLower() switch
                {
                    "sqlserver" => new ImportFileSqlserver(db, import),
                    "mysql" => new ImportFileMySql(db, import),
                    _ => new ImportFile(db, import)
                };

                dlg.ThreadClass(thread);
                dlg.ShowDialog();

                return (thread.Result, thread.ResultMessage);
            }
        }

        /// <summary>
        /// ファイル取込み実行 DataTable
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        private static (MyEnum.MyResult, string) FileImportRun(Window window, FileImportProperties import, MyDbData db, DataTable table)
        {
            // コマンド作成
            //CreateCommands(db, import);

            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(window))
            {
                MyLibrary.MyLoading.Thread thread;

                // Sqlserverとそれ以外で呼出しを変える
                thread = db.Provider.ToLower() switch
                {
                    "sqlserver" => new ImportFileSqlserver(db, import, table),
                    "mysql" => new ImportFileMySql(db, import),
                    _ => new ImportFile(db, import, table)
                };

                dlg.ThreadClass(thread);
                dlg.ShowDialog();

                return (thread.Result, thread.ResultMessage);
            }
        }

        /// <summary>
        /// ファイル取込み実行 DbDataReader
        /// </summary>
        /// <param name="window"></param>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <param name="reader"></param>
        /// <returns></returns>
        private static (MyEnum.MyResult, string) FileImportRun(Window window, FileImportProperties import, MyDbData db, DbDataReader reader, int recordCount)
        {
            // コマンド作成
            //CreateCommands(db, import);

            using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(window))
            {
                MyLibrary.MyLoading.Thread thread;

                // Sqlserverとそれ以外で呼出しを変える
                thread = db.Provider.ToLower() switch
                {
                    "sqlserver" => new ImportFileSqlserver(db, import, reader, recordCount),
                    "mysql" => new ImportFileMySql(db, import),
                    _ => new ImportFile(db, import, reader, recordCount)
                };

                dlg.ThreadClass(thread);
                dlg.ShowDialog();

                return (thread.Result, thread.ResultMessage);
            }
        }

        /// <summary>
        /// 取込先テーブルの指定項目の最大値取得
        /// </summary>
        /// <param name="import"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static Dictionary<string, int> GetSirials(FileImportProperties import, MyDbData db)
        {
            Dictionary<string, int> sirials = new Dictionary<string, int>();
            string[] sirialCol = import.ImportFields.Where(x => x.item_type == 5).Select(x => x.field_name).ToArray();
            foreach (var col in sirialCol)
            {
                string sql = $"SELECT MAX({col}) FROM {import.TableSetting.table_name}";
                int sirial = int.Parse(db.ExecuteScalar(sql).ToString());
                sirials.Add(col, sirial);
            }
            return sirials;
        }

        /// <summary>
        /// 取込先テーブルの初期化
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        public static void InitialDeleteTable(MyDbData db, FileImportProperties import)
        {
            // 新規処理時はテーブルデータ削除
            if (MyCompareModules.OrEx(import.ImportSetting.import_type, 0, 4))
            {
                db.ExecuteNonQuery($"DELETE FROM {import.TableSetting.table_name}");
            }
        }

        /// <summary>
        /// 通常項目・追加情報のチェック
        /// </summary>
        /// <param name="import"></param>
        /// <param name="check"></param>
        /// <param name="row"></param>
        /// <param name="fileName"></param>
        /// <param name="recordNum"></param>
        /// <returns></returns>
        public static bool ValidateItemData(FileImportProperties import, MyStandardCheck check, CharValidator validator, Dictionary<string, string> row, string fileName, int recordNum)
        {
            // チェック前のエラーログ件数
            int errCount = import.ErrorLog.Count;

            // 通常項目
            foreach (ImportFields field in import.ImportFields)
            {
                // エラーチェック
                if (field.check != 0)
                {
                    var chk = check.GetResult(row[field.column_name].ToString(), field.data_type, field.data_length, field.null_flg, field.fix_flg, field.regpattern);

                    if (chk.Item1 != 0)
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"項番:{field.num} [{field.field_caption}] {chk.Item2}", (byte)chk.Item1);
                    }
                }

                string? result = null;

                // 文字コード範囲チェック 1byte
                if (field.check == 1)
                {
                    result = validator.GetInvalid1ByteChars(row[field.column_name].ToString());

                    if (!string.IsNullOrEmpty(result))
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"項番:{field.num} [{field.field_caption}]指定文字コード範囲外[{result}]", 54);
                    }
                }

                // 文字コード範囲チェック 2byte
                if (field.check == 2)
                {
                    result = validator.GetInvalid2ByteChars(row[field.column_name].ToString());

                    if (!string.IsNullOrEmpty(result))
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"項番:{field.num} [{field.field_caption}]指定文字コード範囲外[{result}]", 54);
                    }
                }

                // 文字コード範囲チェック 混在
                if (field.check == 3)
                {
                    result = validator.GetInvalidMixedChars(row[field.column_name].ToString());

                    if (!string.IsNullOrEmpty(result))
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"項番:{field.num} [{field.field_caption}]指定文字コード範囲外[{result}]", 54);
                    }
                }

                // 文字コード範囲チェック Unicode
                if (field.check == 4)
                {
                    result = validator.GetInvalidUnicodeChars(row[field.column_name].ToString());

                    if (!string.IsNullOrEmpty(result))
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"項番:{field.num} [{field.field_caption}]指定文字コード範囲外[{result}]", 54);
                    }
                }
            }
            // エラーログに追加が無ければエラーなし
            return import.ErrorLog.Count == errCount;
        }

        /// <summary>
        /// 通常項目・追加情報をDictionaryにセットする共通メソッド
        /// IRecordAccessorを利用することで、DataRow/DbDataReader/配列などに対応
        /// </summary>
        /// <param name="import"></param>
        /// <param name="accessor"></param>
        /// <param name="data_name"></param>
        /// <param name="num"></param>
        /// <param name="totalRecord"></param>
        /// <param name="sirials"></param>
        /// <param name="excel"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public static Dictionary<string, string> SetRecordRow(FileImportProperties import,
                                                                IRecordAccessor accessor,
                                                                string data_name,
                                                                int num, int totalRecord,
                                                                Dictionary<string, int> sirials,
                                                                MyStreamExcel? excel = null)
        {
            var row = new Dictionary<string, string>();

            foreach (ImportFields field in import.ImportFields)
            {
                string? value = field.item_type switch
                {
                    0 => accessor.GetValue(field.column_name) ?? "",
                    1 => import.AddSettings.Where(x => x.column_name == field.column_name).Select(x => x.value).FirstOrDefault() ?? "",
                    2 => data_name,
                    3 => num.ToString(),
                    4 => totalRecord.ToString(),
                    5 => (sirials[field.field_name] + totalRecord).ToString(),
                    _ => throw new InvalidOperationException($"不正なitem_type値: {field.item_type} (field_name: {field.field_name}, column_name: {field.column_name})")
                };

                // Trim
                if (field.trim_flg == 1) value = value.Trim();

                // 文字列の切り抜き
                if (field.start_position > 0 && field.length > 0)
                {
                    value = value.Mid(field.start_position - 1, field.length);
                }

                // JavaScript
                if (!string.IsNullOrEmpty(field.method_name))
                {
                    if (excel == null)
                    {
                        value = import.js.Invoke(field.method_name, value, accessor).ToString();
                    }
                    else
                    {
                        value = import.js.Invoke(field.method_name, excel, value, accessor).ToString();
                    }
                }

                row.Add(field.column_name, value);
            }
            return row;
        }

        /// <summary>
        /// Dictionaryから主キー値を取得
        /// </summary>
        /// <param name="primaryFields"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static string PrimaryKey(List<ImportFields> primaryFields, Dictionary<string, string> row)
        {
            var primaryKey = string.Empty;

            foreach (var field in primaryFields)
            {
                primaryKey += row[field.column_name] + ":";
            }
            return primaryKey.TrimEnd(':');
        }

        /// <summary>
        /// DbDataReaderから主キー値を取得
        /// </summary>
        /// <param name="primaryFields"></param>
        /// <param name="reader"></param>
        /// <returns></returns>
        public static string Primarykey(List<ImportFields> primaryFields, DbDataReader reader)
        {
            // 主キー値を取得
            var primaryKey = string.Empty;
            foreach (var field in primaryFields)
            {
                primaryKey += reader[field.field_name] + ":";
            }
            return primaryKey.TrimEnd(':');
        }

        /// <summary>
        /// 主キー重複チェック
        /// </summary>
        /// <param name="import"></param>
        /// <param name="primaryFields"></param>
        /// <param name="dataRow"></param>
        /// <param name="fileName"></param>
        /// <param name="primaryKeyHash"></param>
        /// <param name="recordNum"></param>
        /// <returns></returns>
        public static bool IsPrimaryKeyDuplicate(FileImportProperties import, List<ImportFields> primaryFields, Dictionary<string, string> dataRow, string fileName, HashSet<string> primaryKeyHash, int recordNum)
        {
            // 取込種別が新規 or 主キーなしならtrueを返す
            if (import.ImportSetting.import_type is 0 or 4 or 5) return true;

            // 主キー値を取得
            var primaryKey = PrimaryKey(primaryFields, dataRow);

            // 主キーの重複チェック
            if (!primaryKeyHash.Add(primaryKey))
            {
                import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"主キー【{primaryKey}】は重複しています", 52);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 登録情報チェック BulkInsert用
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="primaryFields"></param>
        /// <param name="schema"></param>
        public static void ValidateExistingEntry(MyDbData db, FileImportProperties import, List<ImportFields> primaryFields, string schema)
        {
            string join = string.Empty;
            string sql = string.Empty;
            string primaryJoin = string.Empty;
            string primaryKey = string.Empty;

            // PrimaryKeyのJoin句を作成
            foreach (var field in primaryFields)
            {
                // 主キー項目
                join += $"main.{field.field_name} = temp.{field.field_name} and ";
                // 主キー値
                primaryKey += field.field_name + ",";
            }
            // 終端文字の削除
            join = join.TrimEnd(" and ".ToCharArray());
            primaryKey = primaryKey.TrimEnd(',');
            // テーブル名
            var tableName = import.TableSetting.table_name;

            // 主キーの有る設定なら重複チェック
            if (!new[] { 4, 5 }.Contains(import.ImportSetting.import_type))
            {
                sql = $@"select main.* from #{tableName} main
                    inner join (
                        select {primaryKey}
                        from #{tableName}
                        group by {primaryKey}
                        having count(*) > 1
                        ) temp
                        on {join};";

                ValidateExistingEntryLog(db, import, primaryFields, sql, $"主キー【@@@】は重複", 52);
            }

            switch (import.ImportSetting.import_type)
            {
                // 新規登録
                case 1:
                    // 対象が存在する場合はエラー
                    sql = $@"select * from {schema}{tableName} as temp 
                                where exists (
                                    select 1 
                                    from {tableName} as main
                                    where {join});";

                    ValidateExistingEntryLog(db, import, primaryFields, sql, $"主キー【@@@】は登録済み", 51);
                    break;
                // 更新登録
                case 2:
                    // 対象が存在しない場合はエラー
                    sql = @$"select * from {schema}{tableName} as temp 
                                where not exists (
                                    select 1 
                                    from {tableName} as main
                                    where {join});";

                    ValidateExistingEntryLog(db, import, primaryFields, sql, $"主キー【@@@】は登録未登録", 52);
                    break;
            }

            // 条件文があれば処理
            if (!string.IsNullOrEmpty(import.ImportSetting.where) && new[] { 2, 3 }.Contains(import.ImportSetting.import_type))
            {
                // 条件に不一致はエラー
                sql = @$"select temp.* from {schema}{tableName} as temp 
                            inner join (
                                select * from {tableName} where not ({import.ImportSetting.where})
                                ) as main on {join};";

                ValidateExistingEntryLog(db, import, primaryFields, sql, $"主キー【@@@】は条件不一致", 53);
            }
        }

        /// <summary>
        /// 登録情報のエラーログを作成 BulkInsert用
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="primaryFields"></param>
        /// <param name="sql"></param>
        /// <param name="message"></param>
        /// <param name="errorCode"></param>
        private static void ValidateExistingEntryLog(MyDbData db, FileImportProperties import, List<ImportFields> primaryFields, string sql, string message, byte errorCode)
        {
            using DbDataReader reader = db.ExecuteReader(sql);

            while (reader.Read())
            {
                var dataName = reader["@data_name"].ToString();
                var rec = reader["@record_num"].ToString();

                // 主キー値を取得
                var primaryKey = FileImportClass.Primarykey(primaryFields, reader);
                // エラーログを追加
                import.ErrorLog.AddErrorLog(import.FileSetting.process_name, reader["@data_name"].ToString(), int.Parse(reader["@record_num"].ToString()), message.Replace("@@@", primaryKey), errorCode);
            }
        }
    }

    /// <summary>
    /// レコードデータへの抽象的なアクセスを提供するインターフェース
    /// DataRow/DbDataReader/配列などのデータソースに対して統一的なアクセスを提供
    /// </summary>
    public interface IRecordAccessor
    {
        // 指定カラムの値を取得
        string? GetValue(string columnName);
    }

    /// <summary>
    /// DataRowをラップし、IRecordAccessorインターフェースを実装するクラス
    /// </summary>
    public class DataRowAccessor : IRecordAccessor
    {
        private readonly DataRow _row;
        public DataRowAccessor(DataRow row) => _row = row;
        // DataRowの指定カラムの値を取得
        public string? GetValue(string columnName) => _row[columnName]?.ToString();

    }

    /// <summary>
    /// DbDataReaderをラップし、IRecordAccessorインターフェースを実装するクラス
    /// </summary>
    public class DataReaderAccessor : IRecordAccessor
    {
        private readonly DbDataReader _reaader;
        public DataReaderAccessor(DbDataReader reader) => _reaader = reader;
        // DbDataReaderの指定カラムの値を取得
        public string? GetValue(string columnName) => _reaader[columnName]?.ToString();
    }

    /// <summary>
    /// 配列をラップし、IRecordAccessorインターフェースを実装するクラス
    /// </summary>
    public class ArrayAccessor : IRecordAccessor
    {
        private readonly string[] _record;
        public readonly List<FileColumns> _fileCoulumns;
        public ArrayAccessor(string[] record, List<FileColumns> fileColumns)
        {
            _record = record; ;
            _fileCoulumns = fileColumns;
        }
        // 配列の指定カラムの値を取得
        public string? GetValue(string columnName)
        {
            var col = _fileCoulumns.FirstOrDefault(x => x.column_name == columnName);
            return _record[col.num - 1];
        }
    }

    /// <summary>
    /// ファイル取込みクラス Access, SQLite
    /// </summary>
    internal class ImportFile : MyLibrary.MyLoading.Thread
    {
        private MyDbData _db;
        private FileImportProperties _import;
        private MyStandardCheck _check = new MyStandardCheck();
        private DataTable _table = new DataTable();
        private DbDataReader _reader;
        private int _recordCount = 0;
        private CharValidator _validator;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        public ImportFile(MyDbData db, FileImportProperties import)
        {
            _db = db;
            _import = import;
            ResultMessage = string.Empty;
        }
        public ImportFile(MyDbData db, FileImportProperties import, DataTable table)
        {
            _db = db;
            _import = import;
            _table = table;
            ResultMessage += string.Empty;
        }
        public ImportFile(MyDbData db, FileImportProperties import, DbDataReader reader, int recordCount)
        {
            _db = db;
            _import = import;
            _reader = reader;
            _recordCount = recordCount;
            ResultMessage += string.Empty;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            try
            {
                // 文字コード範囲チェッククラス
                _validator = new CharValidator(_import.ImportSetting.jis_range_1byte,
                                                _import.ImportSetting.jis_range_2byte,
                                                _import.ImportSetting.unicode_range);
                // コマンド作成
                CreateCommands(_db, _import);

                // 取込先テーブルの初期化
                FileImportClass.InitialDeleteTable(_db, _import);
                Run();
                Result = _import.ErrorLog.Count == 0 ? MyEnum.MyResult.Ok : MyEnum.MyResult.Ng;
            }
            catch (Exception ex)
            {
                Result = MyEnum.MyResult.Error;
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        private void Run()
        {
            // 対象項目の最大値を取得
            Dictionary<string, int> sirials = FileImportClass.GetSirials(_import, _db);
            // 主キー項目
            var primaryFields = _import.ImportFields.FindAll(x => x.primary_flg == 1).ToList();

            switch (_import.ImportSetting.data_type)
            {
                // file
                case 0:
                    switch (_import.FileSetting.file_type)
                    {
                        // TEXT
                        case 0:
                            ImportText(sirials, primaryFields);
                            break;

                        // EXCEL
                        case 1:
                            ImportExce(sirials, primaryFields);
                            break;
                    }
                    break;

                // DataTable
                case 1:
                    ImportDataTable(sirials, primaryFields);
                    break;
                // DbDataReader
                case 2:
                    ImportDataReader(sirials, primaryFields);
                    break;
            }
        }

        /// <summary>
        /// 取込処理（Text）
        /// </summary>
        /// <exception cref="Exception"></exception>
        private void ImportText(Dictionary<string, int> sirials, List<ImportFields> primaryFields)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 主キーの重複チェック用ハッシュテーブル
            var primaryKeyHash = new HashSet<string>();
            // 行数
            int recordNum = 0;
            // 全行数
            int totalRecord = 0;

            foreach (var filePath in _import.FilePaths)
            {
                // ファイル名を取得
                var fileName = Path.GetFileName(filePath);

                using MyStreamText text = new MyStreamText(filePath, ImportModules.GetEncoding(_import.FileSetting.moji_code)
                                                                                        , ImportModules.GetDelimiter(_import.FileSetting.delimiter)
                                                                                        , ImportModules.GetSeparator(_import.FileSetting.separator));

                recordNum = 0; // 処理件数
                ProgressMax = text.RecordCount();
                ProgressValue = _import.FileSetting.start_record - 1;   // 実行数
                ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
                ProcessName = $"{fileName} : 取込処理中...";

                // ファイルを開く
                text.Open();
                // 開始行を移動
                text.SkipRecord(_import.FileSetting.start_record);

                while (!text.EndOfStream())
                {
                    recordNum++;
                    totalRecord++;
                    ProgressValue++;

                    // 区切り種別で読込を分岐
                    string[] record = _import.FileSetting.delimiter switch
                    {
                        2 => text.ReadLine(lens, totalLength, false),   // 固定長（文字数）
                        3 => text.ReadLine(lens, totalLength, true),    // 固定長（バイト数）
                        _ => text.ReadLine(),                           // その他の通常読込
                    };

                    // 項目数チェック呼出し
                    if (!FileLoadClass.CheckColumnCount(_import, record, fileName, ProgressValue)) continue;
                    // 通常項目と追加情報のDictionaryを作成
                    var accessor = new ArrayAccessor(record, _import.FileColumns);
                    var recordRow = FileImportClass.SetRecordRow(_import, accessor, fileName, recordNum, totalRecord, sirials);

                    // 項目チェック
                    if (!FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, fileName, ProgressValue)) continue;
                    // 登録情報チェック

                    _import.LoadData.Clear();
                    if (!ValidateRegistrationInfo(_db, _import, primaryFields, recordRow, fileName, primaryKeyHash, ProgressValue)) continue;
                    // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                    AddRowData(_import, recordRow);

                    // DBへ登録
                    int ra = _db.Adapter.Update(_import.LoadData.Select(null, null, System.Data.DataViewRowState.CurrentRows));
                    // 登録が失敗した場合はエラーを発生させる
                    if (ra == 0)
                    {
                        throw new Exception($"DB登録に失敗しました。 ／ {fileName} : {ProgressValue}行目");
                    }
                }
                // メッセージ
                ResultMessage += $"{fileName} : {recordNum}件\r\n";
            }
        }

        /// <summary>
        /// 取込み処理（Excel）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportExce(Dictionary<string, int> sirials, List<ImportFields> primaryFields)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 主キーの重複チェック用ハッシュテーブル
            var primaryKeyHash = new HashSet<string>();
            // 全行数
            int totalRecord = 0;

            foreach (var filePath in _import.FilePaths)
            {
                // ファイル名を取得
                var fileName = Path.GetFileName(filePath);
                // 読み取り項目Rangeを配列化
                var ranges = _import.FileColumns.Select(x => x.column_range).ToArray();

                using MyStreamExcel excel = new MyStreamExcel(filePath);

                ProgressMax = excel.LastRowNumber(_import.SheetName);
                ProgressValue = 0;
                ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
                ProcessName = $"{fileName} : 取込処理中...";

                for (int rowCount = _import.FileSetting.start_record; rowCount <= ProgressMax; rowCount++)
                {
                    totalRecord++;
                    ProgressValue++;

                    // 読込み
                    string[] record = excel.ReadLine(_import.SheetName, rowCount, ranges);

                    // 通常項目と追加情報のDictionaryを作成
                    var accessor = new ArrayAccessor(record, _import.FileColumns);
                    var recordRow = FileImportClass.SetRecordRow(_import, accessor, fileName, ProgressValue, totalRecord, sirials, excel);

                    // 項目チェック
                    if (!FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, fileName, rowCount)) continue;
                    // 登録情報チェック
                    if (!ValidateRegistrationInfo(_db, _import, primaryFields, recordRow, fileName, primaryKeyHash, rowCount)) continue;
                    // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                    AddRowData(_import, recordRow);

                    // DBへ登録
                    int ra = _db.Adapter.Update(_import.LoadData.Select(null, null, System.Data.DataViewRowState.CurrentRows));
                    // 登録が失敗した場合はエラーを発生させる
                    if (ra == 0)
                    {
                        throw new Exception($"DB登録に失敗しました。 ／ {fileName} : {rowCount}行目");
                    }
                }
                // メッセージ
                ResultMessage += $"{fileName} : {ProgressValue}件\r\n";
            }
        }

        /// <summary>
        /// 取込み処理（DataTable）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportDataTable(Dictionary<string, int> sirials, List<ImportFields> primaryFields)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 主キーの重複チェック用ハッシュテーブル
            var primaryKeyHash = new HashSet<string>();
            // 行数
            int recordNum = 0;
            // 全行数
            int totalRecord = 0;
            // データ名
            var dataName = _import.ImportSetting.data_name;

            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;
            ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
            ProcessName = $"{_import.ImportSetting.data_name}：取込み処理中...";

            foreach (DataRow row in _table.Rows)
            {
                ProgressValue++;

                // 通常項目と追加情報のDictionaryを作成
                var accessor = new DataRowAccessor(row);
                var recordRow = FileImportClass.SetRecordRow(_import, accessor, dataName, ProgressValue, totalRecord, sirials);

                // 項目チェック
                if (!FileImportClass.ValidateItemData(_import, _check,_validator, recordRow, dataName, recordNum)) continue;
                // 登録情報チェック
                if (!ValidateRegistrationInfo(_db, _import, primaryFields, recordRow, dataName, primaryKeyHash, recordNum)) continue;
                // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                AddRowData(_import, recordRow);

                // DBへ登録
                int ra = _db.Adapter.Update(_import.LoadData.Select(null, null, System.Data.DataViewRowState.CurrentRows));
                // 登録が失敗した場合はエラーを発生させる
                if (ra == 0)
                {
                    throw new Exception($"DB登録に失敗しました。 ／ {dataName} : {recordNum}行目");
                }
            }
            // メッセージ
            ResultMessage += $"{dataName} : {ProgressValue}件\r\n";
        }

        /// <summary>
        /// 取込み処理（DataTable）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportDataReader(Dictionary<string, int> sirials, List<ImportFields> primaryFields)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 主キーの重複チェック用ハッシュテーブル
            var primaryKeyHash = new HashSet<string>();
            // 行数
            int recordNum = 0;
            // 全行数
            int totalRecord = 0;
            // データ名
            var dataName = _import.ImportSetting.data_name;

            ProgressMax = _recordCount;
            ProgressValue = 0;
            ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
            ProcessName = $"{_import.ImportSetting.data_name}：取込み処理中...";

            while (_reader.Read())
            {
                ProgressValue++;

                // 通常項目と追加情報のDictionaryを作成
                var accessor = new DataReaderAccessor(_reader);
                var recordRow = FileImportClass.SetRecordRow(_import, accessor, dataName, ProgressValue, totalRecord, sirials);

                // 項目チェック
                if (!FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, dataName, recordNum)) continue;
                // 登録情報チェック
                if (!ValidateRegistrationInfo(_db, _import, primaryFields, recordRow, dataName, primaryKeyHash, recordNum)) continue;
                // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                AddRowData(_import, recordRow);

                // DBへ登録
                int ra = _db.Adapter.Update(_import.LoadData.Select(null, null, System.Data.DataViewRowState.CurrentRows));
                // 登録が失敗した場合はエラーを発生させる
                if (ra == 0)
                {
                    throw new Exception($"DB登録に失敗しました。 ／ {dataName} : {recordNum}行目");
                }
            }
            // メッセージ
            ResultMessage += $"{dataName} : {ProgressValue}件\r\n";
        }

        /// <summary>
        /// Adapter.Update用のDataTableを作成
        /// </summary>
        /// <param name="import"></param>
        /// <param name="sourceRow"></param>
        private static void AddRowData(FileImportProperties import, Dictionary<string, string> sourceRow)
        {
            var isNew = false;
            DataRow row;

            // テーブルにDataRowが無ければ新規追加
            if (import.LoadData.Rows.Count == 0)
            {
                isNew = true;
                row = import.LoadData.NewRow();
            }
            // テーブルにDataRowがあれば更新
            else
            {
                row = import.LoadData.Rows[0];
            }

            // 取込項目の値をセット
            foreach (var item in import.ImportFields)
            {
                var value = sourceRow[item.column_name];
                // 値が空なら
                if (string.IsNullOrEmpty(value))
                {
                    // 追記
                    if (isNew)
                    {
                        // 追加
                        row[item.field_name] = sourceRow[item.column_name];
                    }
                    // 更新で空欄更新なら
                    else if (item.null_update == 1)
                    {
                        row[item.field_name] = sourceRow[item.column_name];
                    }
                }
                else
                {
                    row[item.field_name] = sourceRow[item.column_name];
                }
            }

            // 新規はDataTableに追加
            if (isNew)
            {
                import.LoadData.Rows.Add(row);
            }
        }

        /// <summary>
        /// 主キーから登録情報をチェック
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="primaryFields"></param>
        /// <param name="dataRow"></param>
        /// <param name="fileName"></param>
        /// <param name="primaryKeyHash"></param>
        /// <param name="recordNum"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static bool ValidateRegistrationInfo(MyDbData db, FileImportProperties import, List<ImportFields> primaryFields, Dictionary<string, string> dataRow, string fileName, HashSet<string> primaryKeyHash, int recordNum)
        {
            // データ処理用テーブルの初期化
            import.LoadData.Clear();

            // 取込種別が新規 or 主キーなしならtrueを返す
            if (import.ImportSetting.import_type is 0 or 4 or 5) return true;

            // 主キー値を取得
            var primaryKey = FileImportClass.PrimaryKey(primaryFields, dataRow);

            // 主キーの重複チェック
            if (!primaryKeyHash.Add(primaryKey))
            {
                import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"主キー【{primaryKey}】は重複しています", 52);
                return false;
            }

            // 対象取得用パラメータクエリ
            var (sql, param) = ImportModules.SelectParam(import.TableSetting.table_name, primaryFields, dataRow, new[] { "accdb", "accdb2019" }.Contains(db.Provider.ToLower()));

            //　対象取得
            import.LoadData = db.ExecuteQuery(sql, param);

            // 対象が複数ある場合は異常なので中断
            if (import.LoadData.Rows.Count > 1) throw new Exception($"主キー【{primaryKey}】が複数登録されています。");

            switch (import.ImportSetting.import_type)
            {
                // 新規登録
                case 1:
                    // 対象が存在する場合はエラー
                    if (import.LoadData.Rows.Count > 0)
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"主キー【{primaryKey}】は登録済み", 51);
                        return false;
                    }
                    break;
                // 更新登録
                case 2:
                    // 対象が存在しない場合はエラー
                    if (import.LoadData.Rows.Count == 0)
                    {
                        import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"主キー【{primaryKey}】は未登録", 50);
                        return false;
                    }
                    break;
            }

            // 条件文があれば処理
            if (!string.IsNullOrEmpty(import.ImportSetting.where) && import.LoadData.Rows.Count != 0)
            {
                var rows = import.LoadData.Select(import.ImportSetting.where);
                // 条件に不一致はエラー
                if (rows.Count() == 0)
                {
                    import.ErrorLog.AddErrorLog(import.FileSetting.process_name, fileName, recordNum, $"主キー【{primaryKey}】は条件不一致", 53);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// コマンド作成
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        private static void CreateCommands(MyDbData db, FileImportProperties import)
        {
            // selectコマンド作成
            var select = db.Connection.CreateCommand();
            select.CommandType = CommandType.Text;
            select.CommandTimeout = 0;
            select.CommandText = $"select * from {import.TableSetting.table_name};";
            db.Adapter.SelectCommand = select;
            db.Adapter.SelectCommand.Transaction = db.Transaction;

            // insertコマンド作成
            var insert = db.Connection.CreateCommand();
            insert.CommandType = CommandType.Text;
            insert.CommandTimeout = 0;
            var insertCommand = InsertCommand(db, import.TableSetting.table_name, import.ImportFields, new[] { "accdb", "accdb2019" }.Contains(db.Provider.ToLower()));
            insert.CommandText = insertCommand.Item1;
            insert.Parameters.AddRange(insertCommand.Item2.ToArray());
            db.Adapter.InsertCommand = insert;
            db.Adapter.InsertCommand.Transaction = db.Transaction;

            // updateコマンド作成
            var update = db.Connection.CreateCommand();
            update.CommandType = CommandType.Text;
            update.CommandTimeout = 0;
            var updateCommand = UpdateCommand(db, import.TableSetting.table_name, import.ImportFields, new[] { "accdb", "accdb2019" }.Contains(db.Provider.ToLower()));
            update.CommandText = updateCommand.Item1;
            update.Parameters.AddRange(updateCommand.Item2.ToArray());
            db.Adapter.UpdateCommand = update;
            db.Adapter.UpdateCommand.Transaction = db.Transaction;
        }

        /// <summary>
        /// Insertコマンド作成（パラメータ）
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableName"></param>
        /// <param name="importFields"></param>
        /// <param name="isAccess"></param>
        /// <returns></returns>
        private static (string, List<DbParameter>) InsertCommand(MyDbData db, string tableName, List<ImportFields> importFields, bool isAccess = false)
        {
            var fields = string.Empty;
            var values = string.Empty;

            List<DbParameter> param = new List<DbParameter>();

            foreach (var item in importFields)
            {
                fields += $"{item.field_name},";
                values += isAccess == false ? $"@{item.field_name}," : $"?,";

                var p = db.Factory.CreateParameter();
                p.DbType = ImportModules.GetDbType(item.field_type);
                p.SourceColumn = item.field_name;
                p.ParameterName = item.field_name;
                param.Add(p);
            }

            fields = fields.TrimEnd(',');
            values = values.TrimEnd(',');

            var query = $"insert into {tableName} ({fields}) values ({values});";

            return (query, param);
        }

        /// <summary>
        /// Updateコマンド作成（パラメータ）
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tableName"></param>
        /// <param name="importFields"></param>
        /// <param name="isAccess"></param>
        /// <returns></returns>
        private static (string, List<DbParameter>) UpdateCommand(MyDbData db, string tableName, List<ImportFields> importFields, bool isAccess = false)
        {
            var fields = string.Empty;
            var where = string.Empty;
            List<DbParameter> param = new List<DbParameter>();

            // 主キー項目以外
            foreach (var item in importFields.Where(x => x.primary_flg == 0))
            {
                fields += isAccess == false ? $"{item.field_name} = @{item.field_name}," : $"{item.field_name} = ?,";
                var p = db.Factory.CreateParameter();
                p.DbType = ImportModules.GetDbType(item.field_type);
                p.SourceColumn = item.field_name;
                p.ParameterName = item.field_name;
                param.Add(p);
            }

            // 主キー項目
            foreach (var item in importFields.Where(x => x.primary_flg == 1))
            {
                where += isAccess == false ? $"{item.field_name} = @{item.field_name: ?} and " : $"{item.field_name} = ? and ";
                var p = db.Factory.CreateParameter();
                p.DbType = ImportModules.GetDbType(item.field_type);
                p.SourceColumn = item.field_name;
                p.ParameterName = item.field_name;
                param.Add(p);
            }

            fields = fields.TrimEnd(',');
            where = where.TrimEnd(" and ".ToCharArray());
            var query = $"update {tableName} set {fields} where {where};";
            return (query, param);
        }
    }

    /// <summary>
    /// ファイル取込みクラス　SqlServer用
    /// </summary>
    internal class ImportFileSqlserver : MyLibrary.MyLoading.Thread
    {
        private MyDbData _db;
        private FileImportProperties _import;
        private MyStandardCheck _check = new MyStandardCheck();
        private DataTable _table = new DataTable();
        private DbDataReader _reader;
        private int _recordCount = 0;
        const int _bulkCopyBatchSize = 10000; // バッチサイズ
        private CharValidator _validator;

        /// <summary>
        /// コンストラクタ File
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        public ImportFileSqlserver(MyDbData db, FileImportProperties import)
        {
            _db = db;
            _import = import;
            ResultMessage = string.Empty;
        }
        /// <summary>
        /// コンストラクタ DataTable
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="table"></param>
        public ImportFileSqlserver(MyDbData db, FileImportProperties import, DataTable table)
        {
            _db = db;
            _import = import;
            _table = table;
            ResultMessage += string.Empty;
        }
        /// <summary>
        /// コンストラクタ DbDataReader
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="reader"></param>
        /// <param name="recordCount"></param>
        public ImportFileSqlserver(MyDbData db, FileImportProperties import, DbDataReader reader, int recordCount)
        {
            _db = db;
            _import = import;
            _reader = reader;
            _recordCount = recordCount;
            ResultMessage += string.Empty;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            try
            {
                // 文字コード範囲チェッククラス
                _validator = new CharValidator(_import.ImportSetting.jis_range_1byte,
                                                _import.ImportSetting.jis_range_2byte,
                                                _import.ImportSetting.unicode_range);

                // 取込先テーブルの初期化
                FileImportClass.InitialDeleteTable(_db, _import);
                Run();
                Result = _import.ErrorLog.Count == 0 ? MyEnum.MyResult.Ok : MyEnum.MyResult.Ng;
            }
            catch (Exception ex)
            {
                Result = MyEnum.MyResult.Error;
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
            finally
            {
                // 一時テーブルの削除
                _db.ExecuteNonQuery($"drop table if exists #{_import.TableSetting.table_name};");
            }
            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        private void Run()
        {
            // BulkInsert用のテンポラリテーブルを作成
            _db.ExecuteNonQuery($"select * into #{_import.TableSetting.table_name} from {_import.TableSetting.table_name} where 1=0;");
            _db.ExecuteNonQuery($"alter table #{_import.TableSetting.table_name} add [@data_name] nvarchar(max);");
            _db.ExecuteNonQuery($"alter table #{_import.TableSetting.table_name} add [@record_num] int;");
            // BulkInsert用のデータ保管テーブルにシステム項目を追加
            _import.LoadData.Columns.Add("@data_name", typeof(string));
            _import.LoadData.Columns.Add("@record_num", typeof(int));

            // 対象項目の最大値を取得
            Dictionary<string, int> sirials = FileImportClass.GetSirials(_import, _db);
            // 主キー項目
            var primaryFields = _import.ImportFields.FindAll(x => x.primary_flg == 1).ToList();
            // BulkCopyクラス
            BulkCopyClass bulkCopy = new BulkCopyClass(_db, _import.TableSetting.table_name, _import.ImportFields);

            switch (_import.ImportSetting.data_type)
            {
                // file
                case 0:
                    switch (_import.FileSetting.file_type)
                    {
                        // TEXT
                        case 0:
                            ImportText(sirials, primaryFields, bulkCopy);
                            break;

                        // EXCEL
                        case 1:
                            ImportExcel(sirials, primaryFields, bulkCopy);
                            break;
                    }
                    break;

                // DataTable
                case 1:
                    ImportDataTable(sirials, primaryFields, bulkCopy);
                    break;
                // DbDataReader
                case 2:
                    ImportDataReader(sirials, primaryFields, bulkCopy);
                    break;
            }

            // 登録情報チェック
            FileImportClass.ValidateExistingEntry(_db, _import, primaryFields, "#");

            // エラーが無ければテンポラリテーブルのデータを取込み先テーブルに登録
            if (_import.ErrorLog.Count == 0)
            {
                // テンポラリーテーブルのデータ数を取得
                var tempCount = int.Parse(_db.ExecuteScalar($"select count(*) from #{_import.TableSetting.table_name};").ToString());
                // テンポラリテーブルから取込み先テーブルへの登録
                int ra = _db.ExecuteNonQuery(bulkCopy.BulkSql);

                // 登録が失敗した場合はエラーを発生させる
                if (ra != tempCount)
                {
                    throw new Exception($"DB登録に失敗しました。 ／ {_import.TableSetting.table_name}");
                }
            }
        }

        /// <summary>
        /// 取込処理（Text）
        /// </summary>
        /// <exception cref="Exception"></exception>
        private void ImportText(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 行数
            int recordNum = 0;
            // 全行数
            int totalRecord = 0;
            // BulkCopy用のDataTable
            var bulkTable = _import.LoadData.Clone();
            bulkTable.PrimaryKey = null;    // 主キーを設定しない（BulkCopyでの主キー重複チェックを行わないため）

            foreach (var filePath in _import.FilePaths)
            {
                // ファイル名を取得
                var fileName = Path.GetFileName(filePath);

                using MyStreamText text = new MyStreamText(filePath, ImportModules.GetEncoding(_import.FileSetting.moji_code)
                                                                                        , ImportModules.GetDelimiter(_import.FileSetting.delimiter)
                                                                                        , ImportModules.GetSeparator(_import.FileSetting.separator));

                recordNum = 0;  // 処理件数
                ProgressMax = text.RecordCount();
                ProgressValue = _import.FileSetting.start_record - 1;   // 実行数
                ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
                ProcessName = $"{fileName} : 取込処理中...";

                // ファイルを開く
                text.Open();
                // 開始行を移動
                text.SkipRecord(_import.FileSetting.start_record);

                while (!text.EndOfStream())
                {
                    recordNum++;
                    totalRecord++;
                    ProgressValue++;

                    // 区切り種別で読込を分岐
                    string[] record = _import.FileSetting.delimiter switch
                    {
                        2 => text.ReadLine(lens, totalLength, false),   // 固定長（文字数）
                        3 => text.ReadLine(lens, totalLength, true),    // 固定長（バイト数）
                        _ => text.ReadLine(),                           // その他の通常読込
                    };

                    // 項目数チェック呼出し
                    if (FileLoadClass.CheckColumnCount(_import, record, fileName, ProgressValue))
                    {
                        // 通常項目と追加情報のDataRowを作成
                        var accessor = new ArrayAccessor(record, _import.FileColumns);
                        var recordRow = FileImportClass.SetRecordRow(_import, accessor, fileName, ProgressValue, totalRecord, sirials);

                        // 項目チェック
                        if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, fileName, ProgressValue))
                        {
                            // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                            AddRowData(_import, recordRow, bulkTable, fileName, ProgressValue);
                        }
                    }

                    // DBへ登録 BulikInsert
                    if (bulkTable.Rows.Count != 0 && (text.EndOfStream() || bulkTable.Rows.Count % _bulkCopyBatchSize == 0))
                    {
                        // テンポラリテーブルにデータを登録
                        bulkCopy.BulkInsert.WriteToServer(bulkTable);
                        bulkTable.Dispose();
                        bulkTable = _import.LoadData.Clone();
                    }
                }
                // メッセージ
                ResultMessage += $"{fileName} : {recordNum}件\r\n";
            }
        }

        /// <summary>
        /// 取込み処理（Excel）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportExcel(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 全行数
            int totalRecord = 0;

            // BulkCopy用のDataTable
            var bulkTable = _import.LoadData.Clone();

            foreach (var filePath in _import.FilePaths)
            {
                // ファイル名を取得
                var fileName = Path.GetFileName(filePath);
                // 読み取り項目Rangeを配列化
                var ranges = _import.FileColumns.Select(x => x.column_range).ToArray();

                using MyStreamExcel excel = new MyStreamExcel(filePath);

                ProgressMax = excel.LastRowNumber(_import.SheetName);
                ProgressValue = 0;
                ProgressBarType = MyEnum.MyProgressBarType.Percent;
                ProcessName = $"{fileName} : 取込処理中...";

                for (int rowCount = _import.FileSetting.start_record; rowCount <= ProgressMax; rowCount++)
                {
                    totalRecord++;
                    ProgressValue++;

                    // 読込み
                    string[] record = excel.ReadLine(_import.SheetName, rowCount, ranges);

                    // 通常項目と追加情報のDictionaryを作成
                    var accessor = new ArrayAccessor(record, _import.FileColumns);
                    var recordRow = FileImportClass.SetRecordRow(_import, accessor, fileName, ProgressValue, totalRecord, sirials, excel);

                    // 項目チェック
                    if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, fileName, rowCount))
                    {
                        // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                        AddRowData(_import, recordRow, bulkTable, fileName, ProgressValue);
                    }

                    // DBへ登録 BulikInsert
                    if (bulkTable.Rows.Count != 0 && (rowCount == ProgressMax || bulkTable.Rows.Count % _bulkCopyBatchSize == 0))
                    {
                        // テンポラリテーブルにデータを登録
                        bulkCopy.BulkInsert.WriteToServer(bulkTable);
                        bulkTable.Clear();
                    }
                }
                // メッセージ
                ResultMessage += $"{fileName} : {ProgressValue}件\r\n";
            }
        }

        /// <summary>
        /// 取込み処理（DataTable）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportDataTable(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 全行数
            int totalRecord = 0;
            // データ名
            var dataName = _import.ImportSetting.data_name;

            // BulkCopy用のDataTable
            var bulkTable = _import.LoadData.Clone();

            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;
            ProgressBarType = MyEnum.MyProgressBarType.Percent;
            ProcessName = $"{_import.ImportSetting.data_name}：取込み処理中...";

            foreach (DataRow row in _table.Rows)
            {
                ProgressValue++;

                // 通常項目と追加情報のDictionaryを作成
                var accessor = new DataRowAccessor(row);
                var recordRow = FileImportClass.SetRecordRow(_import, accessor, dataName, ProgressValue, totalRecord, sirials);

                // 項目チェック
                if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, dataName, ProgressValue))
                {
                    // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                    AddRowData(_import, recordRow, bulkTable, dataName, ProgressValue);
                }

                // DBへ登録 BulikInsert
                if (bulkTable.Rows.Count != 0 && (ProgressMax == ProgressValue || bulkTable.Rows.Count % _bulkCopyBatchSize == 0))
                {
                    // テンポラリテーブルにデータを登録
                    bulkCopy.BulkInsert.WriteToServer(bulkTable);
                    bulkTable.Clear();
                }
            }
            // メッセージ
            ResultMessage += $"{dataName} : {ProgressValue}件\r\n";
        }

        /// <summary>
        /// 取込み処理（DataReader）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportDataReader(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 行数
            int recordNum = 0;
            // 全行数
            int totalRecord = 0;
            // データ名
            var dataName = _import.ImportSetting.data_name;

            // BulkCopy用のDataTable
            var bulkTable = _import.LoadData.Clone();

            ProgressMax = _recordCount;
            ProgressValue = 0;
            ProgressBarType = MyEnum.MyProgressBarType.Percent;
            ProcessName = $"{_import.ImportSetting.data_name}：取込み処理中...";

            while (_reader.Read())
            {
                ProgressValue++;

                // 通常項目と追加情報のDictionaryを作成
                var accessor = new DataReaderAccessor(_reader);
                var recordRow = FileImportClass.SetRecordRow(_import, accessor, dataName, ProgressValue, totalRecord, sirials);

                // 項目チェック
                if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, dataName, ProgressValue))
                {
                    // 登録先テーブルのスキーマ情報を取得したDataTableに値をセット
                    AddRowData(_import, recordRow, bulkTable, dataName, ProgressValue);
                }

                // DBへ登録 BulikInsert
                if (bulkTable.Rows.Count != 0 && (ProgressMax == ProgressValue || bulkTable.Rows.Count % _bulkCopyBatchSize == 0))
                {
                    // テンポラリテーブルにデータを登録
                    bulkCopy.BulkInsert.WriteToServer(bulkTable);
                    bulkTable.Clear();
                }
            }
            // メッセージ
            ResultMessage += $"{dataName} : {ProgressValue}件\r\n";
        }

        /// <summary>
        /// BuilInsert用のテーブルに取込値の追加
        /// </summary>
        /// <param name="import"></param>
        /// <param name="sourceRow"></param>
        /// <param name="bulkTable"></param>
        /// <param name="dataName"></param>
        /// <param name="recordNum"></param>
        private static void AddRowData(FileImportProperties import, Dictionary<string, string> sourceRow, DataTable bulkTable, string dataName, int recordNum)
        {
            var row = bulkTable.NewRow();
            // 取込項目の値をセット
            foreach (var item in import.ImportFields)
            {
                row[item.field_name] = sourceRow[item.column_name];
            }

            row["@data_name"] = dataName;
            row["@record_num"] = recordNum;

            bulkTable.Rows.Add(row);
        }

        /// <summary>
        /// BulkCopyクラス
        /// </summary>
        private class BulkCopyClass
        {
            public Microsoft.Data.SqlClient.SqlBulkCopy BulkInsert { get; private set; } = null!;
            public string BulkSql { get; private set; } = string.Empty;

            public BulkCopyClass(MyDbData db, string tableName, List<ImportFields> impFields)
            {
                BulkInsert = new Microsoft.Data.SqlClient.SqlBulkCopy((Microsoft.Data.SqlClient.SqlConnection)db.Connection, Microsoft.Data.SqlClient.SqlBulkCopyOptions.KeepNulls, (Microsoft.Data.SqlClient.SqlTransaction)db.Transaction);

                string primary = string.Empty;
                string update = string.Empty;
                string insert = string.Empty;
                string values = string.Empty;

                // BulkCopy先のテンポラリテーブル名
                BulkInsert.DestinationTableName = $"#{tableName}";

                foreach (var field in impFields)
                {
                    // マッピング
                    BulkInsert.ColumnMappings.Add(field.field_name, field.field_name);

                    // 主キー項目
                    if (field.primary_flg == 1) primary += $"t.{field.field_name} = s.{field.field_name} and ";
                    // 登録項目
                    insert += $"{field.field_name},";

                    // 空欄時更新フラグ
                    if (field.null_update == 1)
                    {
                        // 更新項目
                        update += $"t.{field.field_name} = s.{field.field_name},";
                    }
                    else
                    {
                        // 空欄時は更新しない
                        update += $"t.{field.field_name} = case when s.{field.field_name} is not null and s.{field.field_name} <> '' then s.{field.field_name} else t.{field.field_name} end,";
                    }
                    // 値項目
                    values += $"s.{field.field_name},";
                }
                // 終端文字の削除
                primary = primary.TrimEnd(" and ".ToCharArray());
                insert = insert.TrimEnd(",".ToCharArray());
                update = update.TrimEnd(",".ToCharArray());
                values = values.TrimEnd(",".ToCharArray());
                // PrimaryKeyが無い場合はinsertのみ処理
                primary = string.IsNullOrEmpty(primary) ? "1=0" : primary;

                BulkSql = $@"
                MERGE INTO {tableName} AS t
                USING #{tableName} AS s
                ON {primary}
                WHEN MATCHED THEN
                    UPDATE SET {update}
                WHEN NOT MATCHED THEN
                    INSERT ({insert}) VALUES ({values});";

                BulkInsert.ColumnMappings.Add("@data_name", "@data_name");
                BulkInsert.ColumnMappings.Add("@record_num", "@record_num");
            }
        }
    }

    /// <summary>
    /// ファイル取込みクラス　MySQL用
    /// </summary>
    internal class ImportFileMySql : MyLibrary.MyLoading.Thread
    {
        private MyDbData _db;
        private FileImportProperties _import;
        private MyStandardCheck _check = new MyStandardCheck();
        private DataTable _table = new DataTable();
        private DbDataReader _reader;
        private int _recordCount = 0;
        const int _bulkCopyBatchSize = 500; // バッチサイズ
        private CharValidator _validator;

        /// <summary>
        /// コンストラクタ File
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        public ImportFileMySql(MyDbData db, FileImportProperties import)
        {
            _db = db;
            _import = import;
            ResultMessage = string.Empty;
        }
        /// <summary>
        /// コンストラクタ DataTable
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="table"></param>
        public ImportFileMySql(MyDbData db, FileImportProperties import, DataTable table)
        {
            _db = db;
            _import = import;
            _table = table;
            ResultMessage += string.Empty;
        }
        /// <summary>
        /// コンストラクタ DbDataReader
        /// </summary>
        /// <param name="db"></param>
        /// <param name="import"></param>
        /// <param name="reader"></param>
        /// <param name="recordCount"></param>
        public ImportFileMySql(MyDbData db, FileImportProperties import, DbDataReader reader, int recordCount)
        {
            _db = db;
            _import = import;
            _reader = reader;
            _recordCount = recordCount;
            ResultMessage += string.Empty;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            try
            {
                // 文字コード範囲チェッククラス
                _validator = new CharValidator(_import.ImportSetting.jis_range_1byte,
                                                _import.ImportSetting.jis_range_2byte,
                                                _import.ImportSetting.unicode_range);

                // 取込先テーブルの初期化
                FileImportClass.InitialDeleteTable(_db, _import);
                Run();
                Result = _import.ErrorLog.Count == 0 ? MyEnum.MyResult.Ok : MyEnum.MyResult.Ng;
            }
            catch (Exception ex)
            {
                Result = MyEnum.MyResult.Error;
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
            finally
            {
                // MySqlの一時テーブル削除
                _db.ExecuteNonQuery($"drop temporary table if exists tmp_{_import.TableSetting.table_name};");
            }

            Completed = true;
            return 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        private void Run()
        {
            // BulkInsert用のテンポラリテーブルを作成
            _db.ExecuteNonQuery($"create temporary table tmp_{_import.TableSetting.table_name} like {_import.TableSetting.table_name};");
            _db.ExecuteNonQuery($"alter table tmp_{_import.TableSetting.table_name} add `@data_name` text;");
            _db.ExecuteNonQuery($"alter table tmp_{_import.TableSetting.table_name} add `@record_num` int;");

            // BulkInsert用のデータ保管テーブルにシステム項目を追加
            _import.LoadData.Columns.Add("@data_name", typeof(string));
            _import.LoadData.Columns.Add("@record_num", typeof(int));

            // 対象項目の最大値を取得
            Dictionary<string, int> sirials = FileImportClass.GetSirials(_import, _db);
            // 主キー項目
            var primaryFields = _import.ImportFields.FindAll(x => x.primary_flg == 1).ToList();
            // BulkCopyクラス
            BulkCopyClass bulkCopy = new BulkCopyClass(_db, _import.TableSetting.table_name, _import.ImportFields);

            switch (_import.ImportSetting.data_type)
            {
                // file
                case 0:
                    switch (_import.FileSetting.file_type)
                    {
                        // TEXT
                        case 0:
                            ImportText(sirials, primaryFields, bulkCopy);
                            break;

                        // EXCEL
                        case 1:
                            ImportExcel(sirials, primaryFields, bulkCopy);
                            break;
                    }
                    break;

                // DataTable
                case 1:
                    ImportDataTable(sirials, primaryFields, bulkCopy);
                    break;
                // DbDataReader
                case 2:
                    ImportDataReader(sirials, primaryFields, bulkCopy);
                    break;
            }

            // 登録情報チェック
            FileImportClass.ValidateExistingEntry(_db, _import, primaryFields, "tmp_");

            // エラーが無ければテンポラリテーブルのデータを取込み先テーブルに登録
            if (_import.ErrorLog.Count == 0)
            {
                // テンポラリーテーブルのデータ数を取得
                var tempCount = int.Parse(_db.ExecuteScalar($"select count(*) from tmp_{_import.TableSetting.table_name};").ToString());
                // テンポラリテーブルから取込み先テーブルへの登録
                int ra = _db.ExecuteNonQuery(bulkCopy.BulkSql);
            }
        }

        /// <summary>
        /// 取込処理（Text）
        /// </summary>
        /// <exception cref="Exception"></exception>
        private void ImportText(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 行数
            int recordNum = 0;
            // 全行数
            int totalRecord = 0;
            int bulkNum = 0; // BulkCopyの件数カウント
            var bulkObject = new BulkCopyClass.BulkObject();

            foreach (var filePath in _import.FilePaths)
            {
                // ファイル名を取得
                var fileName = Path.GetFileName(filePath);

                using MyStreamText text = new MyStreamText(filePath, ImportModules.GetEncoding(_import.FileSetting.moji_code)
                                                                                        , ImportModules.GetDelimiter(_import.FileSetting.delimiter)
                                                                                        , ImportModules.GetSeparator(_import.FileSetting.separator));

                recordNum = 0;  // 処理件数
                ProgressMax = text.RecordCount();
                ProgressValue = _import.FileSetting.start_record - 1;   // 実行数
                ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
                ProcessName = $"{fileName} : 取込処理中...";

                // ファイルを開く
                text.Open();
                // 開始行を移動
                text.SkipRecord(_import.FileSetting.start_record);

                while (!text.EndOfStream())
                {
                    recordNum++;
                    totalRecord++;
                    ProgressValue++;

                    // 区切り種別で読込を分岐
                    string[] record = _import.FileSetting.delimiter switch
                    {
                        2 => text.ReadLine(lens, totalLength, false),   // 固定長（文字数）
                        3 => text.ReadLine(lens, totalLength, true),    // 固定長（バイト数）
                        _ => text.ReadLine(),                           // その他の通常読込
                    };

                    // 項目数チェック呼出し
                    if (FileLoadClass.CheckColumnCount(_import, record, fileName, ProgressValue))
                    {
                        // 通常項目と追加情報のDataRowを作成
                        var accessor = new ArrayAccessor(record, _import.FileColumns);
                        var recordRow = FileImportClass.SetRecordRow(_import, accessor, fileName, ProgressValue, totalRecord, sirials);

                        // 項目チェック
                        if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, fileName, ProgressValue))
                        {
                            bulkNum++;
                            // 複数insert用のValueを追加
                            bulkCopy.BulkValueAdd(bulkNum, _import.ImportFields, recordRow, fileName, ProgressValue, bulkObject);
                        }
                    }

                    // DBへ登録 BulikInsert
                    if (bulkNum != 0 && (text.EndOfStream() || bulkNum == _bulkCopyBatchSize))
                    {
                        int ra = bulkCopy.BulkInsert(_db, bulkObject);
                        bulkObject.Dispose();
                        bulkObject = new BulkCopyClass.BulkObject();

                        // 登録が失敗した場合はエラーを発生させる
                        if (ra != bulkNum)
                        {
                            throw new Exception($"DB登録に失敗しました。 ／ {_import.TableSetting.table_name}");
                        }
                        bulkNum = 0; // BulkCopyの件数カウントをリセット
                    }
                }
                // メッセージ
                ResultMessage += $"{fileName} : {recordNum}件\r\n";
            }
        }

        /// <summary>
        /// 取込み処理（Excel）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportExcel(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 全行数
            int totalRecord = 0;

            int bulkNum = 0; // BulkCopyの件数カウント
            var bulkObject = new BulkCopyClass.BulkObject();

            foreach (var filePath in _import.FilePaths)
            {
                // ファイル名を取得
                var fileName = Path.GetFileName(filePath);
                // 読み取り項目Rangeを配列化
                var ranges = _import.FileColumns.Select(x => x.column_range).ToArray();

                using MyStreamExcel excel = new MyStreamExcel(filePath);

                ProgressMax = excel.LastRowNumber(_import.SheetName);
                ProgressValue = 0;
                ProgressBarType = MyEnum.MyProgressBarType.Percent;
                ProcessName = $"{fileName} : 取込処理中...";

                for (int rowCount = _import.FileSetting.start_record; rowCount <= ProgressMax; rowCount++)
                {
                    totalRecord++;
                    ProgressValue++;

                    // 読込み
                    string[] record = excel.ReadLine(_import.SheetName, rowCount, ranges);

                    // 通常項目と追加情報のDictionaryを作成
                    var accessor = new ArrayAccessor(record, _import.FileColumns);
                    var recordRow = FileImportClass.SetRecordRow(_import, accessor, fileName, ProgressValue, totalRecord, sirials, excel);

                    // 項目チェック
                    if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, fileName, rowCount))
                    {
                        bulkNum++;
                        // 複数insert用のValueを追加
                        bulkCopy.BulkValueAdd(bulkNum, _import.ImportFields, recordRow, fileName, ProgressValue, bulkObject);
                    }

                    // DBへ登録 BulikInsert
                    if (bulkNum != 0 && (rowCount == ProgressMax || bulkNum == _bulkCopyBatchSize))
                    {
                        int ra = bulkCopy.BulkInsert(_db, bulkObject);
                        bulkObject.Dispose();
                        bulkObject = new BulkCopyClass.BulkObject();

                        // 登録が失敗した場合はエラーを発生させる
                        if (ra != bulkNum)
                        {
                            throw new Exception($"DB登録に失敗しました。 ／ {_import.TableSetting.table_name}");
                        }
                        bulkNum = 0; // BulkCopyの件数カウントをリセット
                    }
                }
                // メッセージ
                ResultMessage += $"{fileName} : {ProgressValue}件\r\n";
            }
        }

        /// <summary>
        /// 取込み処理（DataTable）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportDataTable(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 全行数
            int totalRecord = 0;
            // データ名
            var dataName = _import.ImportSetting.data_name;

            int bulkNum = 0; // BulkCopyの件数カウント
            var bulkObject = new BulkCopyClass.BulkObject();

            ProgressMax = _table.Rows.Count;
            ProgressValue = 0;
            ProgressBarType = MyEnum.MyProgressBarType.Percent;
            ProcessName = $"{_import.ImportSetting.data_name}：取込み処理中...";

            foreach (DataRow row in _table.Rows)
            {
                ProgressValue++;

                // 通常項目と追加情報のDictionaryを作成
                var accessor = new DataRowAccessor(row);
                var recordRow = FileImportClass.SetRecordRow(_import, accessor, dataName, ProgressValue, totalRecord, sirials);

                // 項目チェック
                if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, dataName, ProgressValue))
                {
                    bulkNum++;
                    // 複数insert用のValueを追加
                    bulkCopy.BulkValueAdd(bulkNum, _import.ImportFields, recordRow, dataName, ProgressValue, bulkObject);
                }

                // DBへ登録 BulikInsert
                if (bulkNum != 0 && (ProgressMax == ProgressValue || bulkNum == _bulkCopyBatchSize))
                {
                    int ra = bulkCopy.BulkInsert(_db, bulkObject);
                    bulkObject.Dispose();
                    bulkObject = new BulkCopyClass.BulkObject();

                    // 登録が失敗した場合はエラーを発生させる
                    if (ra != bulkNum)
                    {
                        throw new Exception($"DB登録に失敗しました。 ／ {_import.TableSetting.table_name}");
                    }
                    bulkNum = 0; // BulkCopyの件数カウントをリセット
                }
            }
            // メッセージ
            ResultMessage += $"{dataName} : {ProgressValue}件\r\n";
        }

        /// <summary>
        /// 取込み処理（DataTable）
        /// </summary>
        /// <param name="sirials"></param>
        /// <param name="primaryFields"></param>
        /// <exception cref="Exception"></exception>
        private void ImportDataReader(Dictionary<string, int> sirials, List<ImportFields> primaryFields, BulkCopyClass bulkCopy)
        {
            // 固定長用文字数をLISTで取得
            List<int> lens = _import.FileColumns.Select(x => (int)x.column_length).ToList<int>();
            // 固定長用文字数の合計値
            int totalLength = _import.FileColumns.Sum(x => (int)x.column_length);
            // 全行数
            int totalRecord = 0;
            // データ名
            var dataName = _import.ImportSetting.data_name;

            int bulkNum = 0; // BulkCopyの件数カウント
            var bulkObject = new BulkCopyClass.BulkObject();

            ProgressMax = _recordCount;
            ProgressValue = 0;
            ProgressBarType = MyEnum.MyProgressBarType.Percent;
            ProcessName = $"{_import.ImportSetting.data_name}：取込み処理中...";

            while (_reader.Read())
            {
                ProgressValue++;

                // 通常項目と追加情報のDictionaryを作成
                var accessor = new DataReaderAccessor(_reader);
                var recordRow = FileImportClass.SetRecordRow(_import, accessor, dataName, ProgressValue, totalRecord, sirials);

                // 項目チェック
                if (FileImportClass.ValidateItemData(_import, _check, _validator, recordRow, dataName, ProgressValue))
                {
                    bulkNum++;
                    // 複数insert用のValueを追加
                    bulkCopy.BulkValueAdd(bulkNum, _import.ImportFields, recordRow, dataName, ProgressValue, bulkObject);

                }

                // DBへ登録 BulikInsert
                if (bulkNum != 0 && (ProgressMax == ProgressValue || bulkNum == _bulkCopyBatchSize))
                {
                    int ra = bulkCopy.BulkInsert(_db, bulkObject);
                    bulkObject.Dispose();
                    bulkObject = new BulkCopyClass.BulkObject();

                    // 登録が失敗した場合はエラーを発生させる
                    if (ra != bulkNum)
                    {
                        throw new Exception($"DB登録に失敗しました。 ／ {_import.TableSetting.table_name}");
                    }
                    bulkNum = 0; // BulkCopyの件数カウントをリセット
                }
            }
            // メッセージ
            ResultMessage += $"{dataName} : {ProgressValue}件\r\n";
        }

        /// <summary>
        /// BulkCopyクラス
        /// </summary>
        private class BulkCopyClass
        {
            public string BulkSql { get; private set; } = string.Empty;
            private string _bulkInsert = string.Empty;
            private List<MySqlParameter> _bulkPramaters = new List<MySqlParameter>();
            private List<string> _values = new List<string>();
            public BulkCopyClass(MyDbData db, string tableName, List<ImportFields> impFields)
            {
                string fields = string.Empty;
                string update = string.Empty;
                string values = string.Empty;

                foreach (var field in impFields)
                {
                    // 更新項目は主キー以外
                    if (field.primary_flg == 0)
                    {
                        // 空の時に更新
                        if (field.null_update == 1)
                        {
                            // 更新項目
                            update += $"{field.field_name} = values({field.field_name}),";
                        }
                        else
                        {
                            // 空の時に更新しない
                            update += $"{field.field_name} = if(values({field.field_name}) is null or values({field.field_name}) = '', {tableName}.{field.field_name}, values({field.field_name})),";
                        }
                    }

                    // 項目
                    fields += $"{field.field_name},";
                    values += $"{field.field_name},";
                }
                // 終端文字の削除
                update = update.TrimEnd(",".ToCharArray());
                values = values.TrimEnd(",".ToCharArray());
                fields = fields.TrimEnd(",".ToCharArray());

                BulkSql = $@"
                    insert into {tableName} ({values})
                    select {values} from tmp_{tableName}
                    on duplicate key update {update};";

                // BuklInsert用SQL文作成
                fields += ",`@data_name`,`@record_num`";
                _bulkInsert = @$"insert into tmp_{tableName} ({fields}) values ";
            }

            /// <summary>
            /// BalukValueの追加
            /// </summary>
            /// <param name="num"></param>
            /// <param name="fields"></param>
            /// <param name="record"></param>
            /// <param name="dataName"></param>
            /// <param name="recrodNum"></param>
            public void BulkValueAdd(int num, List<ImportFields> fields, Dictionary<string, string> record, string dataName, int recrodNum, BulkObject bulkObject)
            {
                string value = string.Empty;

                foreach (var field in fields)
                {
                    // values作成
                    value += $"@{field.field_name}_{num},";
                    // パラメータを追加
                    bulkObject.BulkPramaters.Add(new MySqlParameter($"@{field.field_name}_{num}", record[field.column_name]));
                }

                // システム項目を追加
                value += $"@data_name_{num},@record_num_{num},";
                bulkObject.BulkPramaters.Add(new MySqlParameter($"@data_name_{num}", dataName));
                bulkObject.BulkPramaters.Add(new MySqlParameter($"@record_num_{num}", recrodNum));
                // values作成
                bulkObject.Values.Add($"({value.TrimEnd(',').ToString()})");
            }

            /// <summary>
            /// BulikInsert実行
            /// </summary>
            /// <param name="db"></param>
            /// <returns></returns>
            public int BulkInsert(MyDbData db, BulkObject bulk)
            {
                // バルクインサートのSQL文を生成
                var sb = new StringBuilder();
                sb.Append(_bulkInsert);
                sb.Append(string.Join(",", bulk.Values));
                // コマンド作成
                var cmb = (MySqlCommand)db.Connection.CreateCommand();
                cmb.CommandText = sb.ToString();
                cmb.Parameters.AddRange(bulk.BulkPramaters.ToArray());
                var ra = cmb.ExecuteNonQuery();
                cmb.Dispose();
                sb = null;
                return ra;
            }

            /// <summary>
            /// BulkInsert用のデータオブジェクト
            /// </summary>
            public class BulkObject : IDisposable
            {
                private List<MySqlParameter> _bulkPramaters = new List<MySqlParameter>();
                private List<string> _values = new List<string>();
                public List<MySqlParameter> BulkPramaters => _bulkPramaters;
                public List<string> Values => _values;

                /// <summary>
                /// リソース開放
                /// </summary>
                public void Dispose()
                {
                    _bulkPramaters.Clear();
                    _values.Clear();
                }
            }
        }
    }
}
