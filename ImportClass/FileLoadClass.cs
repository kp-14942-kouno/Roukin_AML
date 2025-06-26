using System.Data;
using MyLibrary.MyModules;
using MyLibrary;
using MyTemplate.Class;
using System.Windows;
using MyLibrary.MyClass;

namespace MyTemplate.ImportClass
{
    /// <summary>
    /// ファイル読込み設定
    /// </summary>
    internal class FileSettingProperties: IDisposable
    {
        public FileSetting FileSetting { get; set; } = new FileSetting();   // ファイル読込み設定
        public List<FileColumns> FileColumns { get; set; } = [];            // カラム設定
        public List<AddSetting> AddSettings { get; set; } = [];             // 追加情報
        public MyJScriptClass? js = null;   // JavaScriptクラス

        /// <summary>
        /// 解放
        /// </summary>
        public void Dispose()
        {
            js?.Dispose();
        }
    }

    /// <summary>
    /// ファイル読込みプロパティ
    /// </summary>
    internal class FileLoadProperties : FileSettingProperties
    {
        public DataTable LoadData { get; set; } = new DataTable();  // 読込データ
        public ErrorLog ErrorLog { get; set; } = new ErrorLog();    // エラーログ
        public MyEnum.MyResult Result { get; set; }                 // 結果
        public List<string> FilePaths { get; set; } = [];           // 読込ファイルパス
        public string SheetName { get; set; } = string.Empty;       // シート名

        /// <summary>
        /// 解放
        /// </summary>
        public new void Dispose()
        {
            base.Dispose();
            LoadData.Dispose();
            ErrorLog.Dispose();
        }
    }

    /// <summary>
    /// ファイル読込みクラス
    /// </summary>
    internal static class FileLoadClass
    {
        /// <summary>
        /// ファイル読込み設定取得
        /// </summary>
        /// <param name="id"></param>
        /// <param name="fileLoadProperties"></param>
        /// <returns></returns>
        public static bool GetFileLoadSetting(int id, FileLoadProperties load)
        {
            try
            {
                // 初期化
                load.FileSetting = new FileSetting();
                load.FileColumns = new List<FileColumns>();
                load.AddSettings = new List<AddSetting>();
                load.LoadData = new DataTable();
                load.ErrorLog = new ErrorLog();
                load.FilePaths = new List<string>();

                // ファイル読込み設定取得
                using (var db = new MyDbData("setting"))
                {
                    MyPropertyModules.GetCreateProperties(db, typeof(FileSetting), load.FileSetting, "t_file_setting", "file_id", id);
                    MyPropertyModules.GetCreateProperties(db, typeof(FileColumns), load.FileColumns, "t_file_columns", "columns_id", "num", load.FileSetting.columns_id);
                    MyPropertyModules.GetCreateProperties(db, typeof(AddSetting), load.AddSettings, "t_file_add_setting", "add_id", "num", load.FileSetting.add_id);
                }

                // 読込データのカラムを作成
                CreateFileLoadTable(load);

                // JavaScript
                if (!string.IsNullOrEmpty(load.FileSetting.js_node_key))
                {
                    load.js = new MyJScriptClass(load.FileSetting.js_node_key);
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
        /// ファイル選択
        /// </summary>
        /// <param name="window"></param>
        /// <param name="fileLoadProperties"></param>
        /// <returns></returns>
        public static MyEnum.MyResult FileSelect(Window window, FileLoadProperties fileLoadProperties)
        {
            try
            {
                // フォルダ選択ダイアログ
                MyFileDialog dialog = new MyFileDialog();

                string fileExtension = fileLoadProperties.FileSetting.file_type switch
                {
                    0 => fileLoadProperties.FileSetting.delimiter == 0 ? "csv" : "txt", // テキストファイル
                    1 => "xlsx", // Excelファイル
                    _ => throw new NotImplementedException("未対応なファイル種別です。")
                };

                // ファイル選択処理
                var result = fileLoadProperties.FileSetting.select_type switch
                {
                    0 => dialog.Single(fileLoadProperties.FileSetting.file_path, new string[] { fileExtension }, new string[] { fileLoadProperties.FileSetting.file_name_reg }),
                    1 => dialog.Mulit(window, fileLoadProperties.FileSetting.file_path, new string[] { fileLoadProperties.FileSetting.file_name_reg }),
                    2 => dialog.DragDrop(window, new string[] { fileLoadProperties.FileSetting.file_name_reg }),
                    _ => throw new NotImplementedException("未対応なファイル選択種別です。")
                };

                // ファイル選択キャンセル
                if (result != MyEnum.MyResult.Ok)
                {
                    return result;
                }

                // 選択されたファイルパスを取得
                fileLoadProperties.FilePaths = dialog.FilePaths;

                // 読込みファイルがExcel
                if (fileLoadProperties.FileSetting.file_type == 1)
                {
                    // ファイル選択がSingleの場合はSheetチェックを呼出し
                    if (fileLoadProperties.FileSetting.select_type == 0)
                    {
                        // Sheetチェック呼出し
                        ValidateExcelSheet(fileLoadProperties);
                    }
                    else
                    {
                        // それ以外はFileSettingのSheet名をセット
                        fileLoadProperties.SheetName = fileLoadProperties.FileSetting.sheet_name;
                    }
                    // Sheetが空は中断
                    return string.IsNullOrEmpty(fileLoadProperties.SheetName) ? MyEnum.MyResult.Cancel : MyEnum.MyResult.Ok;
                }

                return MyEnum.MyResult.Ok;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return MyEnum.MyResult.Error;
            }
        }

        /// <summary>
        /// Excelシート名の存在確認
        /// </summary>
        /// <param name="fileLoadProperties"></param>
        /// <exception cref="Exception"></exception>
        /// <exception cref="NotImplementedException"></exception>
        private static void ValidateExcelSheet(FileLoadProperties fileLoadProperties)
        {
            using var excel = new MyStreamExcel(fileLoadProperties.FilePaths[0]);
            string sheetName = fileLoadProperties.FileSetting.sheet_flg switch
            {
                0 => excel.IsSheet(fileLoadProperties.FileSetting.sheet_name, true) ?? throw new Exception("指定されたシートが存在しません。"),
                1 => excel.IsSheet(fileLoadProperties.FileSetting.sheet_name, true) ?? throw new Exception("指定されたシートが存在しません。"),
                _ => throw new NotImplementedException("未対応なシート指定フラグです。")
            };

            // シート名を取得
            fileLoadProperties.SheetName = sheetName;
        }

        /// <summary>
        /// ファイル読込みテーブル作成
        /// </summary>
        /// <param name="fileLoadProperties"></param>
        /// <returns></returns>
        public static void CreateFileLoadTable(FileLoadProperties fileLoadProperties)
        {
            // DataTable初期化
            fileLoadProperties.LoadData = new DataTable();

            // 読込データのカラムを作成
            var columns = fileLoadProperties.FileColumns
                .Select(col => new DataColumn(col.column_name)
                {
                    Caption = col.column_caption,
                    DataType = typeof(string)
                });

            // 追加情報のカラムを作成
            var addColumns = fileLoadProperties.AddSettings
                .Select(add => new DataColumn(add.column_name)
                {
                    Caption = add.column_caption,
                    DataType = typeof(string)
                });

            // システム項目
            var systemColumns = new[]
            {
                new DataColumn("@data_name", typeof(string)) { Caption = "データ名" },
                new DataColumn("@record_num", typeof(int)) { Caption = "行番号" }
            };

            // 全てのカラムを一括追加
            fileLoadProperties.LoadData.Columns.AddRange(columns.Concat(addColumns).Concat(systemColumns).ToArray());
        }

        /// <summary>
        /// ファイル読込処理
        /// </summary>
        /// <param name="window"></param>
        /// <param name="fileLoadProperties"></param>
        /// <param name="showErrorLog"></param>
        /// <returns></returns>
        internal static MyEnum.MyResult FileLoad(Window window, FileLoadProperties fileLoadProperties, bool showErrorLog = true)
        {
            try
            {
                // ファイル選択
                MyEnum.MyResult selectResult = FileSelect(window, fileLoadProperties);
                if(selectResult != MyEnum.MyResult.Ok) return selectResult;

                // 確認
                if (MyMessageBox.Show($"{fileLoadProperties.FileSetting.process_name}\r\n処理を実行しますか？", "確認", MyEnum.MessageBoxButtons.OkCancel, window: window) == MyEnum.MessageBoxResult.Cancel)
                    return MyEnum.MyResult.Cancel;

                // 追加情報処理
                MyEnum.MyResult addResult = Add.Run(fileLoadProperties, window);
                if (addResult != MyEnum.MyResult.Ok) return addResult;

                // ファイル読込処理
                using (MyLibrary.MyLoading.Dialog dlg = new MyLibrary.MyLoading.Dialog(window))
                {
                    FileLoadThread thread = new FileLoadThread(fileLoadProperties);
                    dlg.ThreadClass(thread);
                    dlg.ShowDialog();

                    // エラーログがあれば表示
                    if (thread.Result == MyEnum.MyResult.Ng && showErrorLog)
                    {
                        MyLibrary.MyDataViewer viewr = new MyDataViewer(window, fileLoadProperties.ErrorLog.DefaultView, 
                            columnNames: fileLoadProperties.ErrorLog.ErrorLogField(), columnHeaders: fileLoadProperties.ErrorLog.ErrorLogCaption());

                       viewr.ShowDialog();
                    }

                    // 結果を返す
                    return thread.Result;
                }
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                return MyEnum.MyResult.Error;
            }
        }

        /// <summary>
        /// 通常項目と追加情報をDataRowへセット
        /// </summary>
        /// <param name="record"></param>
        /// <param name="fileName"></param>
        /// <param name="recordNum"></param>
        /// <param name="excel"></param>
        /// <returns></returns>
        public static DataRow SetRecordRow(FileLoadProperties fileLoadProperties, string[] record, string fileName, int recordNum, MyStreamExcel? excel = null)
        {
            // DataRow
            var row = fileLoadProperties.LoadData.NewRow();

            row["@data_name"] = fileName;
            row["@record_num"] = recordNum;

            // 通常項目カラム設定分繰り返す
            foreach (FileColumns column in fileLoadProperties.FileColumns)
            {
                string? value = record[column.num - 1].ToString();

                // Trim
                if (column.column_trim == 1)
                {
                    value = value.Trim();
                }

                // 文字列の切抜き
                if (column.start_position > 0 && column.length > 0)
                {
                    value = value.Mid(column.start_position - 1, column.length);
                }

                // メソッド実行（JavaScript）
                if (!string.IsNullOrEmpty(column.method_name))
                {
                    if(excel == null)
                    {
                        value = fileLoadProperties.js.Invoke(column.method_name, value, record).ToString();
                    }
                    else
                    {
                        value = fileLoadProperties.js.Invoke(column.method_name, excel, value, record).ToString();
                    }
                }

                row[column.column_name] = value;
            }
            // 追加情報分繰り返す
            foreach (AddSetting add in fileLoadProperties.AddSettings)
            {
                row[add.column_name] = add.value?.ToString();
            }
            return row;
        }

        /// <summary>
        /// 項目数チェック
        /// </summary>
        /// <param name="record"></param>
        /// <param name="fileName"></param>
        /// <param name="recordNum"></param>
        /// <returns></returns>
        public static bool CheckColumnCount(FileLoadProperties fileLoadProperties, string[] record, string fileName, int recordNum)
        {
            // 項目数チェック
            if (record.Length != fileLoadProperties.FileColumns.Count)
            {
                // エラーログに追加
                fileLoadProperties.ErrorLog.AddErrorLog(fileLoadProperties.FileSetting.process_name, fileName, recordNum, $"項目数が異なります。[設定値:{fileLoadProperties.FileColumns.Count} ／ データ値:{record.Length}]", 99);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 通常項目・追加情報のチェック
        /// </summary>
        /// <param name="row"></param>
        /// <param name="fileName"></param>
        /// <param name="recordNum"></param>
        /// <returns></returns>
        public static bool CheckColumn(FileLoadProperties fileLoadProperties, MyStandardCheck check, DataRow row, string fileName, int recordNum)
        {
            // チェック前のエラーログ件数
            int errCount = fileLoadProperties.ErrorLog.Count;

            // 通常項目
            foreach (FileColumns column in fileLoadProperties.FileColumns)
            {
                // エラーチェック
                if (column.check_flg == 1)
                {
                    var (errCode, errMsg) = check.GetResult(row[column.column_name].ToString(), column.column_type, column.column_length, column.column_null, column.column_fix, column.column_reg);

                    if (errCode != 0)
                    {
                        fileLoadProperties.ErrorLog.AddErrorLog(fileLoadProperties.FileSetting.process_name, fileName, recordNum, $"項番:{column.num} [{column.column_caption}] {errMsg}", (byte)errCode);
                    }
                }
            }

            // エラーログに追加が無ければエラーなし
            return fileLoadProperties.ErrorLog.Count == errCount;
        }

        /// <summary>
        /// ファイル移動処理
        /// </summary>
        /// <param name="load"></param>
        /// <exception cref="NotImplementedException"></exception>
        public static void FileMove(FileLoadProperties load)
        {
            try
            {
                // ファイル移動フラグが0以外でファイルパスが存在する場合
                if (load.FileSetting.file_move_flg != 0 && load.FilePaths != null && load.FilePaths.Count > 0)
                {
                    foreach (var filePath in load.FilePaths)
                    {
                        var fielRoot = System.IO.Path.GetDirectoryName(filePath);
                        var dirDate = DateTime.Now.ToString("yyyyMMdd_hhmmss");

                        // ファイル移動先パスを取得
                        var destFilePath = load.FileSetting.file_move_flg switch
                        {
                            1 => System.IO.Path.Combine(load.FileSetting.file_move_path, dirDate), // 固定パス
                            2 => System.IO.Path.Combine(fielRoot, load.FileSetting.file_move_path, DateTime.Now.ToString("yyyyMMdd")), // 日付フォルダ
                            _ => throw new NotImplementedException("未対応なファイル移動フラグです。")
                        };

                        // 移動先ディレクトリ作成
                        System.IO.Directory.CreateDirectory(destFilePath);
                        // ファイル名取得
                        var fileName = System.IO.Path.GetFileName(filePath);
                        // ファイル移動
                        System.IO.File.Move(filePath, System.IO.Path.Combine(destFilePath, fileName));
                    }
                }
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
            }
        }
    }

    /// <summary>
    /// ファイル読込み実行クラス
    /// </summary>
    internal class FileLoadThread : MyLibrary.MyLoading.Thread
    {
        FileLoadProperties _fileLoadProperties = new FileLoadProperties();
        MyStandardCheck _check = new MyStandardCheck();

        // コンストラクタ
        public FileLoadThread(FileLoadProperties fileLoadProperties)
        {
            _fileLoadProperties = fileLoadProperties;
        }

        /// <summary>
        /// 排他処理
        /// </summary>
        /// <returns></returns>
        public override int MultiThreadMethod()
        {
            try
            {
                switch(_fileLoadProperties.FileSetting.file_type)
                {
                    case 0:
                        // テキストファイル読込み
                        TextFileLoad();
                        break;
                    case 1:
                        // Excelファイル読込み
                        ExcelFileLoad();
                        break;
                    default:
                        throw new NotImplementedException("未対応なファイル種別です。");
                }
                // 結果
                Result = _fileLoadProperties.ErrorLog.Count > 0 ? MyEnum.MyResult.Ng : MyEnum.MyResult.Ok;
            }
            catch (Exception ex)
            {
                MyLogger.SetLogger(ex, MyEnum.LoggerType.Error);
                Result = MyEnum.MyResult.Error;
            }
            finally
            {
                // 処理結果がOKならばファイル移動処理を実行
                if (Result == MyEnum.MyResult.Ok)
                    // ファイル移動処理
                    FileLoadClass.FileMove(_fileLoadProperties);
            }
            Completed = true;
            return 0;
        }

        /// <summary>
        /// テキストファイル読込
        /// </summary>
        private void TextFileLoad()
        {
            // 固定長用の文字数を配列で取得
            List<int> lens = _fileLoadProperties.FileColumns.AsEnumerable().Select(x => (int)x.column_length).ToList<int>();
            // 固定長用の文字数の合計値を取得
            int totalLength = lens.Sum();

            // ファイル数分繰り返す
            foreach (string filePath in _fileLoadProperties.FilePaths)
            {
                // ファイル名取得
                string fileName = System.IO.Path.GetFileName(filePath);

                // MyStreamText
                using (MyStreamText text = new MyStreamText(filePath, ImportModules.GetEncoding(_fileLoadProperties.FileSetting.moji_code)
                                                                                        , ImportModules.GetDelimiter(_fileLoadProperties.FileSetting.delimiter)
                                                                                        , ImportModules.GetSeparator(_fileLoadProperties.FileSetting.separator)))
                {
                    // プログレスバー設定
                    ProgressMax = text.RecordCount();
                    ProgressValue = _fileLoadProperties.FileSetting.start_record - 1;
                    ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
                    ProcessName = $"{fileName} \r\n 読込中...";

                    text.Open();
                    text.SkipRecord(_fileLoadProperties.FileSetting.start_record); // 開始レコード分スキップ

                    while (!text.EndOfStream())
                    {
                        ProgressValue++;

                        // 区切り種別で読込を分岐
                        string[] record = _fileLoadProperties.FileSetting.delimiter switch
                        {
                            2 => text.ReadLine(lens, totalLength, false),    // 固定長（文字数）
                            3 => text.ReadLine(lens, totalLength, true),     // 固定長（バイト数）
                            _ => text.ReadLine()                // その他通常読込
                        };

                        // 項目数チェック呼出し
                        if (!FileLoadClass.CheckColumnCount(_fileLoadProperties, record, fileName, ProgressValue)) continue;
                        // 通常項目と追加情報のDataRowを作成
                        var recordRow = FileLoadClass.SetRecordRow(_fileLoadProperties, record, fileName, ProgressValue);
                        // 項目チェック
                        if(!FileLoadClass.CheckColumn(_fileLoadProperties, _check, recordRow, fileName, ProgressValue)) continue;
                        // エラーが無ければDataTableに追加
                        _fileLoadProperties.LoadData.Rows.Add(recordRow);
                    }
                }
            }
        }

        /// <summary>
        /// Excelファイル読込
        /// </summary>
        private void ExcelFileLoad()
        {
            // ファイル数分繰り返す
            foreach (string filePath in _fileLoadProperties.FilePaths)
            {
                // ファイル名取得
                string fileName = System.IO.Path.GetFileName(filePath);
                // 読み取りRangeを配列化
                var ranges = _fileLoadProperties.FileColumns.Select(x => x.column_range).ToArray();

                // MyStreamText
                using (MyStreamExcel excel = new MyStreamExcel(filePath))
                {
                    // プログレスバー設定
                    ProgressMax = excel.LastRowNumber(_fileLoadProperties.SheetName);
                    ProgressValue = _fileLoadProperties.FileSetting.start_record - 1;
                    ProgressBarType = MyEnum.MyProgressBarType.PercentFraction;
                    ProcessName = $"{fileName} \r\n 読込中...";

                    for (int row = _fileLoadProperties.FileSetting.start_record; row <= ProgressMax; row++)
                    {
                        ProgressValue = row;

                        string[] record = excel.ReadLine(_fileLoadProperties.SheetName, row, ranges);

                        // 通常項目と追加情報のDataRowを作成
                        var recordRow = FileLoadClass.SetRecordRow(_fileLoadProperties, record, fileName, ProgressValue, excel);
                        // 項目チェック
                        if (!FileLoadClass.CheckColumn(_fileLoadProperties, _check, recordRow, fileName, ProgressValue)) continue;
                        // エラーが無ければDataTableに追加
                        _fileLoadProperties.LoadData.Rows.Add(recordRow);
                    }
                }
            }
        }
    }
}
