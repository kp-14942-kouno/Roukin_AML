using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using MyLibrary;
using MyLibrary.MyModules;
using Org.BouncyCastle.Pqc.Crypto.Frodo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MyTemplate.ImportClass
{
    /// <summary>
    /// SQL文を作成する（SELECT分）パラメータクエリ
    /// </summary>
    internal static class ImportModules
    {
        /// <summary>
        /// SQL文を生成する(SELECT文) パラメータクエリ DataRowを使用
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="primaryFields"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static (string, Dictionary<string, object>) SelectParam(string tableName, List<ImportFields> primaryFields, DataRow row, bool isAccess =false)
        {
            var where = string.Empty;
            var param = new Dictionary<string, object>();

            foreach(var item in primaryFields)
            {
                where += isAccess == false ? $"{item.field_name}=@{item.column_name} and " : $"{item.field_name}=? and ";
                param.Add(item.column_name, row[item.column_name]);
            }
            where = where.TrimEnd(" and ".ToCharArray());
            var sql = $"select * from {tableName} where {where};";
            return (sql, param);
        }


        /// <summary>
        /// SQL文を生成する(SELECT文) パラメータクエリ Dictionary<string, string>を使用
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="primaryFields"></param>
        /// <param name="row"></param>
        /// <param name="isAccess"></param>
        /// <returns></returns>
        public static (string, Dictionary<string, object>) SelectParam(string tableName, List<ImportFields> primaryFields, Dictionary<string, string> row, bool isAccess = false)
        {
            var where = string.Empty;
            var param = new Dictionary<string, object>();

            foreach (var item in primaryFields)
            {
                where += isAccess == false ? $"{item.field_name}=@{item.column_name} and " : $"{item.field_name}=? and ";
                param.Add(item.column_name, row[item.column_name]);
            }
            where = where.TrimEnd(" and ".ToCharArray());
            var sql = $"select * from {tableName} where {where};";
            return (sql, param);
        }


        /// <summary>
        /// SQL文を生成する(INSERT文) パラメータクエリ
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="importFields"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static (string, Dictionary<string, object>) InsertParam(string tableName, List<ImportFields> importFields, DataRow row, bool isAccess =false)
        {
            var sql = string.Empty;
            var fields = string.Empty;
            var columns = string.Empty;
            var param = new Dictionary<string, object>();
            // ImportFieldsのフィールド名とDataRowのカラム名をマッピングして、SQL文を生成する
            foreach (var item in importFields)
            {
                fields += item.field_name + ",";
                columns += (isAccess == false ? "@" + item.column_name : "?") + ",";
                param.Add(item.field_name, row[item.column_name]);
            }
            fields = fields.TrimEnd(',');
            columns = columns.TrimEnd(',');
            sql = $"insert into {tableName} ({fields}) values ({columns});";
            return (sql, param);
        }

        /// <summary>
        /// SQL文を生成する(UPDATE文) パラメータクエリ
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="primaryFields"></param>
        /// <param name="importFields"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static (string, Dictionary<string, dynamic>) UpdatePram(string tableName, List<ImportFields> primaryFields, List<ImportFields> importFields, DataRow row, bool isAccess =false)
        {
            var sql = string.Empty;
            var fields = string.Empty;
            var where = string.Empty;
            var param = new Dictionary<string, dynamic>();
            // ImportFields（主キー以外）のフィールド名とDataRowのカラム名をマッピングして、SQL文を生成する
            foreach (var item in importFields)
            {
                // 主キーのフィールド名は除外する
                if (!primaryFields.Any(x => x.field_name == item.field_name))
                {
                    fields += item.field_name + " = " + (isAccess == false ? $"@{item.column_name}" : "?") + ",";
                    param.Add(item.field_name, row[item.column_name]);
                }
            }
            fields = fields.TrimEnd(',');
            // ImportFields(主キー）のフィールド名とDataRowのカラム名をマッピングして、Where文を生成する
            foreach (var item in primaryFields)
            {
                where += item.field_name + " = " + (isAccess == false ? $"@{item.column_name}" : "?") + ",";
                param.Add(item.field_name, row[item.column_name]);
            }
            where = where.TrimEnd(" and ".ToCharArray());
            sql = $"update {tableName} set {fields} where {where};";
            return (sql, param);
        }

        /// <summary>
        /// SQL文を生成する(MERGE文) 
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="tempTableName"></param>
        /// <param name="primaryFields"></param>
        /// <param name="importFields"></param>
        /// <returns></returns>
        public static string GenarateMergeQuery(string tableName, string tempTableName, List<ImportFields>primaryFields, List<ImportFields> importFields)
        {
            // 主キー条件を生成
            var onConditions = string.Join(" and ", primaryFields.Select(x => $"t.{x.field_name} = s.{x.field_name}"));
            // 更新項目を生成
            var updateFields = string.Join(", ", importFields.Select(x => $"t.{x.field_name} = s.{x.field_name}"));

            // 挿入項目を生成
            var insertColumn = string.Join(", ", importFields.Select(x => $"{x.field_name}"));
            var insertValues = string.Join(", ", importFields.Select(x => $"{x.field_name}"));

            var mergeQuery = $@"MERGE INTO {tableName} AS t
                USING {tempTableName} AS s 
                    ON {onConditions}  
                WHEN MATCHED THEN
                    UPDATE SET {updateFields}
                WHEN NOT MATCHED THEN 
                    INSERT ({insertColumn}) 
                    VALUES ({insertValues});";

            return mergeQuery;
        }

        /// <summary>
        /// Encoding
        /// </summary>
        /// <param name="encode"></param>
        /// <returns></returns>
        internal static Encoding GetEncoding(byte encode)
        {
            Encoding encoding = null;

            switch (encode)
            {
                // SJIS
                case 0:
                    encoding = Encoding.GetEncoding("Shift_JIS");
                    break;
                // UTF8
                case 1:
                    encoding = new UTF8Encoding(false);
                    break;
                // UTF8BOM
                case 2:
                    encoding = new UTF8Encoding(true);
                    break;
                // UTF16LE
                case 3:
                    encoding = new UnicodeEncoding(false, false);
                    break;
                // UTF16LEBOM
                case 4:
                    encoding = new UnicodeEncoding(true, false);
                    break;
                // UTF16BE
                case 5:
                    encoding = new UnicodeEncoding(false, true);
                    break;
                // UTF16BEBOM
                case 6:
                    encoding = new UnicodeEncoding(true, true);
                    break;
            }
            return encoding;
        }

        /// <summary>
        /// 区切り文字
        /// </summary>
        /// <param name="delimiter"></param>
        /// <returns></returns>
        internal static MyEnum.Delimiter GetDelimiter(byte delimiter)
        {
            switch (delimiter)
            {
                // TAB
                case 1:
                    return MyEnum.Delimiter.Tab;
                // 固定長（文字数）
                case 2:
                    return MyEnum.Delimiter.NoneLen;
                // 固定長（バイト数）
                case 3:
                    return MyEnum.Delimiter.NoneByte;
                // カンマ
                default:
                    return MyEnum.Delimiter.Comma;
            }
        }

        /// <summary>
        /// 囲い文字
        /// </summary>
        /// <param name="separator"></param>
        /// <returns></returns>
        internal static MyEnum.Separator GetSeparator(byte separator)
        {
            switch (separator)
            {
                // ダブルクォーテーション
                case 1:
                    return MyEnum.Separator.DoubleQuotation;
                // シングルクォーテーション
                case 2:
                    return MyEnum.Separator.SingleQuotation;
                // なし
                default:
                    return MyEnum.Separator.None;
            }
        }

        /// <summary>
        /// DbType変換
        /// </summary>
        /// <param name="fieldType"></param>
        /// <returns></returns>
        internal static DbType GetDbType(byte fieldType)
        {
            switch (fieldType)
            {
                case 0:
                    return DbType.String;   // 文字列
                case 1:
                    return DbType.Byte;     // バイト
                case 2:
                    return DbType.Int16;    // 整数 int
                case 3: 
                    return DbType.Int32;    // 整数 long
                case 4:
                    return DbType.String;   // Memo
                case 6:
                    return DbType.String;   // 日時
                case 7:
                    return DbType.Int32;    // AutoNumber
                default:
                    return DbType.String;
            }
        }

        /// <summary>
        /// JavaScript呼び出し
        /// </summary>
        /// <param name="script">JavaScriptファイル名の設定名</param>
        /// <param name="methodName"></param>
        /// <param name="index">methodNameが空欄時の戻り値のparametersの要素番号</param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        internal static string MethodRun(string methodName, int index, params object[] parameters)
        {
            if (string.IsNullOrEmpty(methodName)) return parameters[index - 1].ToString();

            // Script名とMethod名に分離
            var var = methodName.Split('.');

            // JavaScriptのファイル情報
            string dir = MyUtilityModules.AppSetting("script", "dir");
            string file = MyUtilityModules.AppSetting("script", var[0]);
            // JavaScript実行呼び出し
            var result = Modules.ExcuteJavaScript(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dir, file), var[1], parameters);
            return result.ToString();

        }

    }
}
