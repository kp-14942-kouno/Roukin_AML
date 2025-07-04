using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.ImportClass
{
    /// <summary>
    /// テーブル設定プロパティ
    /// </summary>
    public class TableSetting
    {
        public int table_id { get; set; }                           // テーブル設定ID
        public string table_name { get; set; } = string.Empty;      // テーブル名
        public string table_caption { get; set; } = string.Empty;   // テーブル名称
        public string schema { get; set; } = string.Empty;          // プロバイダ名
        public byte search_flg { get; set; }                        // 検索フラグ
    }

    /// <summary>
    /// テーブルフィールド設定プロパティ
    /// </summary>
    public class TableFields
    {
        public int table_id { get; set; }                           // テーブル設定ID   
        public int num { get; set; }                                // 項目番号
        public string field_name { get; set; } = string.Empty;      // フィールド名
        public string field_caption { get; set; } = string.Empty;   // フィールド名称
        public byte field_type { get; set; }                        // フィールド種別
        public byte data_type { get; set; }                         // データ種別
        public int data_length { get; set; }                        // データ長
        public byte fix_flg { get; set; }                           // 固定長フラグ
        public byte null_flg { get; set; }                          // NULL許可フラグ
        public string regpattern { get; set; } = string.Empty;      // 正規表現
        public byte primary_flg { get; set; }                       // 主キー
        public byte index_flg { get; set; }                         // インデックス
    }

    /// <summary>
    /// ファイル読込み設定プロパティ
    /// </summary>
    public class FileSetting
    {
        public int file_id { get; set; }                                // ファイル読込み設定ID
        public string process_name { get; set; } = string.Empty;        // 処理名
        public byte file_type { get; set; }                             // ファイル種別
        public int columns_id { get; set; }                             // カラム設定ID
        public int add_id { get; set; }                                 // 追加ID
        public string file_name_reg { get; set; } = string.Empty;       // ファイル名正規表現
        public byte moji_code { get; set; }                             // 文字コード
        public byte delimiter { get; set; }                             // 区切り文字
        public byte separator { get; set; }                             // 囲み文字
        public string jis_range_1byte { get; set; } = string.Empty;     // 1バイト文字の範囲
        public string jis_range_2byte { get; set; } = string.Empty;     // 2バイト文字の範囲
        public string unicode_range { get; set; } = string.Empty;       // Unicode文字の範囲
        public int start_record { get; set; }                           // 開始レコード
        public byte select_type { get; set; }                           // 選択種別
        public string file_path { get; set; } = string.Empty;           // ファイルパス
        public byte sheet_flg { get; set; }                             // シート指定フラグ
        public string sheet_name { get; set; } = string.Empty;          // シート名
        public byte file_move_flg { get; set; }                         // ファイル移動フラグ
        public string file_move_path { get; set; } = string.Empty;      // ファイル移動先パス
        public string js_node_key { get; set; } = string.Empty;         // JavaScriptのファイル名のKEY
    }

    /// <summary>
    /// カラム設定プロパティ
    /// </summary>
    public class FileColumns
    {
        public int columns_id { get; set; }                             // カラム設定ID
        public int num { get; set; }                                    // カラム番号
        public string column_name { get; set; } = string.Empty;         // カラム名
        public string column_caption { get; set; } = string.Empty;      // カラム名称
        public string column_range { get; set; } = string.Empty;        // カラム範囲
        public byte column_type { get; set; }                           // カラム種別
        public int column_length { get; set; }                          // カラム長
        public byte column_fix { get; set; }                            // カラム固定長
        public byte column_null { get; set; }                           // カラムNULL許可
        public byte column_trim { get; set; }                           // カラムトリム
        public string column_reg { get; set; } = string.Empty;          // カラム正規表現
        public int start_position { get; set; }                         // 開始位置
        public int length { get; set; }                                 // 長さ
        public string method_name { get; set; } = string.Empty;         // メソッド
        public byte check_flg { get; set; }                             // チェックフラグ

    }

    /// <summary>
    /// 追加情報プロパティ
    /// </summary>
    public class AddSetting
    {
        public int add_id { get; set; }                             // 追加ID
        public int num { get; set; }                                // 項目番号
        public string column_name { get; set; } = string.Empty;     // カラム名
        public string column_caption { get; set; } = string.Empty;  // カラム名称
        public byte add_type { get; set; }                          // 追加種別
        public string add_value { get; set; } = string.Empty;       // 追加値
        public byte column_type { get; set; }                       // カラム種別
        public int column_length { get; set; }                      // カラム長
        public byte column_null { get; set; }                       // カラムNULL許可
        public byte column_fix { get; set; }                        // カラム固定長
        public byte column_trim { get; set; }                       // カラムトリム
        public string column_reg { get; set; } = string.Empty;  // カラム正規表現
        public string method_name { get; set; } = string.Empty; // メソッド
        public string value { get; set; } = string.Empty;       // 値
    }

    /// <summary>
    /// インポート設定プロパティ
    /// </summary>
    internal class ImportSetting
    {
        public int import_id { get; set; }                          // インポートID
        public string process_name { get; set; } = string.Empty;    // 処理名
        public int table_id { get; set; }                           // テーブル設定ID
        public int field_id { get; set; }                           // フィールド設定ID
        public byte data_type { get; set; }                         // 取込みデータ種別
        public byte add_id { get; set; }                            // 追加ID
        public string data_name { get; set; } = string.Empty;       // データ名
        public int load_id { get; set; }                            // 読込設定ID
        public byte import_type { get; set; }                       // 取込種別
        public string where { get; set; } = string.Empty;           // 条件文
        public string js_node_key { get; set; } = string.Empty;     // JavaScriptのファイル名のKEY
        public string jis_range_1byte { get; set; } = string.Empty; // 1バイト文字の範囲
        public string jis_range_2byte { get; set; } = string.Empty; // 2バイト文字の範囲
        public string unicode_range { get; set; } = string.Empty;   // Unicode文字の範囲
    }

    /// <summary>
    /// インポートフィールド設定プロパティ
    /// </summary>
    internal class ImportFields : TableFields
    {
        public int field_id { get; set; }                           // フィールド設定ID
        public new int num { get; set; }                            // 項目番号
        public new string field_name { get; set; } = string.Empty;  // フィールド名
        public byte item_type { get; set; }                         // Item種別
        public string column_name { get; set; } = string.Empty;     // カラム名
        public string separator { get; set; } = string.Empty;       // 区切り文字
        public byte trim_flg { get; set; }                          // トリムフラグ
        public int start_position { get; set; }                     // 開始位置
        public int length { get; set; }                             // 長さ
        public string method_name { get; set; } = string.Empty;     // メソッド
        public byte check { get; set; }                             // チェックフラグ
        public byte null_update { get; set; }                       // NULL更新フラグ
    }
}
